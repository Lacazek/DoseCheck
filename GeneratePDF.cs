using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Input;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace DoseCheck
{
    internal class GeneratePDF
    {

        public class StructureObjective
        {
            public string ID { get; set; }
            public string[] Aliases { get; set; }
            public string DVHObjective { get; set; }
            public string ExpectedValue { get; set; }
            public string RealValue { get; set; }
            public string Evaluator { get; set; }
            public bool Variation { get; set; }
            public bool Met { get; set; }
            public bool OverLimit { get; set; }
            public string FoundStructureID { get; set; }
            public string Referentiel { get; set; }
            public string Vol { get; set; }

        }
        #region Extraction
        public static double ExtractNumber(string input)
        {
            //Problème pour gérer le relatif (exemple PTV 95%)
            Regex regexD = new Regex(@"(?i)^(D)?(?<evalpt>\d+(\.\d+)?)(?<unit>(Gy|cc|%))$");
            Regex regexV = new Regex(@"(?i)^(V)?(?<evalpt>\d+(\.\d+)?)(?<unit>(Gy|cc|%))$");
            Regex regexSimple = new Regex(@"(?i)(?<evalpt>^\d+(\.\d+)?)(?<unit>(%|cc|Gy))?$");

            try
            {
                Match matchD = regexD.Match(input.Trim());
                if (matchD.Success)
                {
                    return Convert.ToDouble(matchD.Groups["evalpt"].Value, CultureInfo.InvariantCulture);
                }

                Match matchV = regexV.Match(input.Trim());
                if (matchV.Success)
                {
                    return Convert.ToDouble(matchV.Groups["evalpt"].Value, CultureInfo.InvariantCulture);
                }

                Match matchSimple = regexSimple.Match(input.Trim());
                if (matchSimple.Success)
                {
                    return Convert.ToDouble(matchSimple.Groups["evalpt"].Value, CultureInfo.InvariantCulture);
                }

                return -1.00;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Aucun regex ne correspond pour {input}\n" + ex.Message);
                return -1.00;
            }
        }

        static int CompareValues(double x, string obj, double tol, string name)
        {
            if (obj.Trim().ToLower().Contains("index"))
                obj = obj.Trim().Substring(0, 5);

            double y = ExtractNumber(obj.Trim());

            if (name.Contains("PTV") || name.Contains("CTV") || name.Contains("ITV") || name.Contains("GTV") && !name.ToUpper().Contains("z_"))
            {
                if (x >= y) return -1;
                if (x > y - tol && x < y) return 0;
            }
            else
            {
                if (x <= y) return -1;
                if (x <= y + tol && x > y + tol) return 0;

            }
            
            return 1; // x > y ou x tgt < y
        }
        #endregion

        private Model _model;
        private Dictionary<string, string> _results;
        private string WORKBOOK_TEMPLATE_DIR;
        private string WORKBOOK_RESULT_DIR;
        private List<StructureObjective> m_objectives; //all the dose values are internally in Gy.

        // Class pour assembler le fichier html avec les datas calculés dans GetMyData. Calcul également les informations nécessaire à ajouter au fichier final.
        internal GeneratePDF(Model model)
        {
            _results = new Dictionary<string, string>();

            try
            {
                _model = model;
                WORKBOOK_TEMPLATE_DIR = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), "Template_dosi");
                if (!Directory.Exists(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), "Resultats")))
                {
                    Directory.CreateDirectory(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), "Resultats"));
                    WORKBOOK_RESULT_DIR = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), "Resultats");
                }
                else
                    WORKBOOK_RESULT_DIR = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), "Resultats");

            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur dans la construction du PDF\n" + ex.Message);
            }

        }
        /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        #region Execute
        public void Execute(Dictionary<string, string> results)
        {
            {
                PlanningItem pitem = null;
                _results = results;
                List<string[]> CSVSheet = new List<string[]>();
                List<string[]> DataOut = new List<string[]>();
                int _decimal = 2;
                double tol = 1; // Gy

                //foreach (var item in _results.Keys) { MessageBox.Show(item); }

                // Permet de créer la list d'objet qui va alimenter le tableau HDV OAR du fichier html
                try
                {
                    m_objectives = new List<StructureObjective>();

                    foreach (var item in _results)
                    {
                        if (!item.Key.Split('/')[2].Contains("Volume"))
                        {
                            StructureObjective obj = new StructureObjective();
                            {
                                obj.Vol = _results.FirstOrDefault(x => x.Key.Split('/')[0] == item.Key.Split('/')[0] && x.Key.Contains("Volume")).Value;
                                obj.Referentiel = item.Key.Split('/')[4];
  
                            if (item.Key.Split('/')[3].Trim().ToLower() != "no tol" && !item.Key.Split('/')[3].Trim().ToLower().Contains("index"))
                                {
                                    obj.ID = item.Key.Split('/')[0];
                                    obj.DVHObjective = item.Key.Split('/')[3];
                                    obj.ExpectedValue = item.Key.Split('/')[2];
                                    obj.RealValue = item.Value;

                                switch (CompareValues(Convert.ToDouble(item.Value.Substring(0, item.Value.Length-2).Trim()),
                                        item.Key.Split('/')[3],
                                        tol,
                                        item.Key.Split('/')[0]))
                                    {
                                        case -1: // x < y
                                            obj.Met = true;
                                            obj.Variation = false;
                                            obj.OverLimit = false;
                                            obj.Evaluator = "OK";
                                            break;
                                        case 0: // x <= y + tol et x > y
                                            obj.Met = false;
                                            obj.Variation = true;
                                            obj.OverLimit = false;
                                            obj.Evaluator = "Warning";
                                            break;
                                        case 1: // x > y + tol
                                            obj.Met = false;
                                            obj.Variation = false;
                                            obj.OverLimit = true;
                                            obj.Evaluator = "Over Limit";
                                            break;
                                        default:
                                            throw new ArgumentException("Unexpected comparison result");
                                    }
                                }
                            else if (item.Key.Split('/')[3].Trim().ToLower().Contains( "dose max") || item.Key.Split('/')[3].Trim().ToLower().Contains("dose moyenne"))
                                {
                                    obj.ID = item.Key.Split('/')[0];
                                    obj.DVHObjective = item.Key.Split('/')[3];
                                    obj.ExpectedValue = item.Key.Split('/')[2];
                                    obj.RealValue = item.Value;
                                    obj.Variation = false;

                                    if (Convert.ToDouble(obj.RealValue) < Convert.ToDouble(obj.DVHObjective))
                                    {
                                        obj.Met = true;
                                        obj.OverLimit = false;
                                        obj.Evaluator = "OK";
                                    }
                                    else
                                    {
                                        obj.Met = false;
                                        obj.OverLimit = true;
                                        obj.Evaluator = "Over Limit";
                                    }
                                }
                                else
                                {
                                    obj.ID = item.Key.Split('/')[0];
                                    obj.DVHObjective = item.Key.Split('/')[3];
                                    obj.ExpectedValue = item.Key.Split('/')[2];
                                    obj.RealValue = item.Value;
                                    obj.Met = false;
                                    obj.Variation = false;
                                    obj.OverLimit = false;
                                    obj.Evaluator = "No Objectif";
                                }
                                m_objectives.Add(obj);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur dans la génération du PDF\n" + ex.Message);
                }

                try
                {
                    // Test les sommes de plan
                    if (_model.PlanSetup == null && _model.Course.PlanSums == null)
                    { throw new ApplicationException("Please load a plan or plansum."); }
                    if (_model.PlanSetup == null && _model.Course.PlanSums.Count() > 1)
                    { throw new ApplicationException("Please close all but one plan sum."); }
                    else if (_model.PlanSetup == null && _model.Course.PlanSums.Count() == 1)
                    {
                        PlanSum plansum = _model.Course.PlanSums.Single();
                        pitem = plansum;
                        StructureSet ss = plansum.StructureSet;
                    }
                    if (pitem == null) pitem = _model.PlanSetup;
                    if (_model.Patient == null || _model.StructureSet == null || pitem == null) { MessageBox.Show("Please load a patient and plan or plansum before running this script."); return; }

                    if (!_model.PlanSetup.IsDoseValid) { MessageBox.Show("Please calculate dose for the active plan before running this script."); return; } // a voir ici !!!!!!!!!!!!!!!

                    // make sure the workbook directory exists
                    if (!System.IO.Directory.Exists(WORKBOOK_TEMPLATE_DIR)) { MessageBox.Show(string.Format("The default template file directory '{0}' defined by the script does not exist.", WORKBOOK_TEMPLATE_DIR)); return; }

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Problème lors du chargement du plan\n" + ex.Message); // erreur ici aussi
                }


                // Calcul des paramètres de faisceau utilisés pour le fichier html uniquement (ne concerne pas la partie dose)
                double Um = 0, Xb = 0, Yb = 0, Zb = 0, XMax = 0;
                double[] X1 = { 0, 0, 0, 0, 0, 0 };
                double[] Y1 = { 0, 0, 0, 0, 0, 0 };
                double[] X2 = { 0, 0, 0, 0, 0, 0 };
                double[] Y2 = { 0, 0, 0, 0, 0, 0 };
                int i = 0, y = 0;

                try
                {
                    foreach (var b in _model.PlanSetup.Beams)
                    {
                        double XD = _model.Image.UserOrigin.x;
                        double YD = _model.Image.UserOrigin.y;
                        double ZD = _model.Image.UserOrigin.z;

                        // COORDONNEES DICOM DU FAISCEAU + TAILLE DE CHAMP + UM
                        if (b.Meterset.Value.ToString() == "Non Numérique") { }
                        else
                        {

                            //VERIFICATION ET COMPARAISON POSITION MAX MACHOIRES DE CHAQUE FAISCEAU
                            X1[i] = Math.Round(b.ControlPoints.First().JawPositions.X1 / 10, _decimal);
                            X2[i] = Math.Round(b.ControlPoints.First().JawPositions.X2 / 10, _decimal);
                            double xmaxi = (Math.Abs(X1[i] - X2[i]));
                            if (xmaxi >= XMax) { XMax = xmaxi; }
                            // POSITION DE CHAQUE FAISCEAU
                            Xb = Math.Round(b.IsocenterPosition.x, _decimal);
                            Yb = Math.Round(b.IsocenterPosition.y, _decimal);
                            Zb = Math.Round(b.IsocenterPosition.z, _decimal);
                            // ADDITION UM POUR CHAQUE FAISCEAU
                            Um += b.Meterset.Value;
                            i++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Problème lors du chargement des faisceaux\n" + ex.Message);
                }

                string Iso = string.Format("X : {0} cm  Y : {2} cm  Z : {1} cm", Math.Abs(Math.Round((Xb - _model.Image.UserOrigin.x) / 10, _decimal)), Math.Abs(Math.Round((Yb - _model.Image.UserOrigin.y) / 10, _decimal)), Math.Abs(Math.Round((_model.Image.UserOrigin.z - Zb) / 10, _decimal)));
                string[] InfoPlan = new string[] { _model.Patient.Name.ToString(), DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"), _model.PlanSetup.Id.ToString(), _model.PlanSetup.Beams.First().TreatmentUnit.ToString(), _model.PlanSetup.CreationUserName.ToString() };
                string[] EvalPlan = new string[] { _model.PlanSetup.TargetVolumeID.ToString(),_model.PlanSetup.Dose.DoseMax3D.ToString(), Math.Round(Um, _decimal).ToString(), Math.Round(Um / (_model.PlanSetup.TotalDose.Dose * 100), _decimal).ToString(), XMax.ToString(), Iso, _model.PlanSetup.PlanNormalizationMethod.ToString() };
                var PrescriptionList = new List<string[]>();

                try
                {
                    foreach (var target in _model.PlanSetup.RTPrescription.Targets)
                    {
                        // Ajouter un tableau de chaînes contenant les informations de chaque cible à la liste
                        PrescriptionList.Add(new string[]
                        {
        target.TargetId,
        target.DosePerFraction.ToString(),
        target.NumberOfFractions.ToString(),
        Math.Round(target.NumberOfFractions * target.DosePerFraction.Dose,_decimal).ToString(),
                        });
                    }
                }
                catch
                {
                    PrescriptionList.Add(new string[]
                {
        _model.PlanSetup.TargetVolumeID ?? "No target",
        _model.PlanSetup.DosePerFraction.ToString(),
        _model.PlanSetup.NumberOfFractions.ToString(),
        Math.Round((_model.PlanSetup.DosePerFraction.Dose * (int)_model.PlanSetup.NumberOfFractions),_decimal).ToString(),
                });

                }


                string[][] Prescription = PrescriptionList.ToArray();
                string[] Header = new string[] { "ID", "Volume [cc]", "D max [%]", "D99% [%]", "D95% [%]", "D90% [%]", "D moyenne [Gy]", "D m&#233diane [Gy]", "D min [Gy]", "Validation" };
                string[] Header_OARs = new string[] { "ID", "Volume [cc]", "Objectif", "Contrainte", "R&#233sultats", "R&#233f&#233rentiel", "Validation" };
                double PADDICK = -1, HI = -1, CI = -1, GI = -1, RCI = -1;
                string[,] PTVSTEREO = new string[150, 10];
                string[,] PTV = new string[150, 10];
                double V100, V95, V50;
                double D100, D50;
                int NbFraction; DoseValue DosePerFraction;

                // Mise en forme des données pour le fichier html
                try
                {
                    foreach (Structure scan in _model.StructureSet.Structures)
                    {
                        if (scan.Id.ToUpper().Contains("PTV") || scan.Id.ToUpper().Contains("CTV") || scan.Id.ToUpper().Contains("ITV") || scan.Id.ToUpper().Contains("GTV") && scan.Id == _model.PlanSetup.TargetVolumeID)
                        {
                            try
                            {
                                NbFraction = _model.RTPrescription.Targets.FirstOrDefault(x => x.TargetId == scan.Id).NumberOfFractions;
                                DosePerFraction = _model.RTPrescription.Targets.FirstOrDefault(x => x.TargetId == scan.Id).DosePerFraction;
                            }
                            catch
                            {
                                NbFraction = (int)_model.PlanSetup.NumberOfFractions;
                                DosePerFraction = _model.PlanSetup.DosePerFraction;
                            }

                            V100 = Math.Round(_model.PlanSetup.GetVolumeAtDose(scan, (DosePerFraction * NbFraction), VolumePresentation.AbsoluteCm3), _decimal);
                            V95 = Math.Round(_model.PlanSetup.GetVolumeAtDose(scan, 0.95 * (DosePerFraction * NbFraction), VolumePresentation.AbsoluteCm3), _decimal);
                            V50 = Math.Round(_model.PlanSetup.GetVolumeAtDose(scan, 0.5 * (DosePerFraction * NbFraction), VolumePresentation.AbsoluteCm3), _decimal);
                            D100 = Math.Round(_model.PlanSetup.GetDoseAtVolume(scan, 100, VolumePresentation.Relative, DoseValuePresentation.Relative).Dose, _decimal);
                            D50 = Math.Round(_model.PlanSetup.GetDoseAtVolume(scan, 50, VolumePresentation.Relative, DoseValuePresentation.Relative).Dose, _decimal);

                            try
                            {
                                HI = Convert.ToDouble(_results.FirstOrDefault(x => x.Key.Contains(scan.Id) && x.Key.Contains("HI")).Value);
                                CI = Convert.ToDouble(_results.FirstOrDefault(x => x.Key.Contains(scan.Id) && x.Key.Contains("CI")).Value);
                                RCI = Convert.ToDouble(_results.FirstOrDefault(x => x.Key.Contains(scan.Id) && x.Key.Contains("RCI")).Value);
                                PADDICK = Convert.ToDouble(_results.FirstOrDefault(x => x.Key.Contains(scan.Id) && x.Key.Contains("PADDICK")).Value);
                                GI = Convert.ToDouble(_results.FirstOrDefault(x => x.Key.Contains(scan.Id) && x.Key.Contains("GI")).Value);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Erreur lors de l'extraction des résultats : " + ex.Message);
                            }
                            PTVSTEREO[y, 0] = scan.Id;
                            PTVSTEREO[y, 1] = Math.Round(scan.Volume, 2).ToString();
                            PTVSTEREO[y, 2] = Math.Round(_model.PlanSetup.GetVolumeAtDose(_model.StructureSet.Structures.FirstOrDefault(x => x.DicomType.ToUpper() == "EXTERNAL"), (DosePerFraction * NbFraction), VolumePresentation.AbsoluteCm3), _decimal).ToString();
                            PTVSTEREO[y, 3] = Math.Round(_model.PlanSetup.GetVolumeAtDose(scan, (DosePerFraction * NbFraction), VolumePresentation.AbsoluteCm3), _decimal).ToString();
                            PTVSTEREO[y, 4] = Math.Round(_model.PlanSetup.GetDoseAtVolume(scan, 100, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose, _decimal).ToString();
                            PTVSTEREO[y, 5] = HI.ToString();
                            PTVSTEREO[y, 6] = CI.ToString();
                            PTVSTEREO[y, 7] = RCI.ToString();
                            PTVSTEREO[y, 8] = PADDICK.ToString();
                            PTVSTEREO[y, 9] = GI.ToString();

                            PTV[y, 0] = scan.Id;
                            PTV[y, 1] = Math.Round(scan.Volume, 2).ToString();
                            PTV[y, 2] = Math.Round(_model.PlanSetup.GetDoseAtVolume(scan, 0.01, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Relative).Dose, _decimal).ToString();
                            PTV[y, 3] = Math.Round(_model.PlanSetup.GetDoseAtVolume(scan, 99, VolumePresentation.Relative, DoseValuePresentation.Relative).Dose, _decimal).ToString();
                            PTV[y, 4] = Math.Round(_model.PlanSetup.GetDoseAtVolume(scan, 95, VolumePresentation.Relative, DoseValuePresentation.Relative).Dose, _decimal).ToString();
                            PTV[y, 5] = Math.Round(_model.PlanSetup.GetDoseAtVolume(scan, 90, VolumePresentation.Relative, DoseValuePresentation.Relative).Dose, _decimal).ToString();
                            PTV[y, 6] = Math.Round(_model.PlanSetup.GetDVHCumulativeData(scan, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1).MeanDose.Dose, _decimal).ToString();
                            PTV[y, 7] = Math.Round(_model.PlanSetup.GetDVHCumulativeData(scan, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1).MedianDose.Dose, _decimal).ToString();
                            PTV[y, 8] = Math.Round(_model.PlanSetup.GetDVHCumulativeData(scan, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1).MinDose.Dose, _decimal).ToString();
                        }
                        y++;
                    }
                }

                catch (Exception ex)
                {
                    MessageBox.Show("Problème lors du calcul des indices stéréotaxiques\n" + ex.Message);
                }
                string HtmlBody = ExportToHtml(Header, Header_OARs, InfoPlan, Prescription, EvalPlan, PTVSTEREO, PTV);
                string outputpath = System.IO.Path.Combine(WORKBOOK_RESULT_DIR + DateTime.Now.ToString("yyyy-MM-dd") + "-" + _model.Patient.Name.ToString() + pitem.Id.ToString() + ".html");

                System.IO.File.WriteAllText(outputpath, HtmlBody);
                System.Diagnostics.Process.Start(outputpath);
                System.Threading.Thread.Sleep(3000);

            }
        }

        ////////////////////////////////////////////////////////////// FIN EXECUTE ///////////////////////////////////////////////////
        #endregion

        #region ExportToHtml
        // Création et mise en forme du fichier html
        protected string ExportToHtml(string[] header, string[] header_OAR, string[] InfoPlan, string[][] Prescription, string[] EvaPlan, string[,] PTVSTEREO, string[,] PTV)
        {
            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html>");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body style='font-family:arial; font-size:medium'>");

            // Conteneur pour l'image et du tableau
            strHTMLBuilder.Append("<div style='display: flex; align-items: flex-start;'>");

            // Ajouter une image en haut à gauche
            strHTMLBuilder.Append("<div style='margin-right: 20px;'>");
            strHTMLBuilder.Append("<img src='B:\\RADIOTHERAPIE\\Killian\\Dosi\\Script\\DoseCheck\\Projects\\DoseCheck\\fisherMan4.png' alt='Description de l'image' style='width:300px; height:auto;'>");
            strHTMLBuilder.Append("</div>");

            // Conteneur pour les tableaux
            strHTMLBuilder.Append("<div>");

            //////////////////////////////////// TABLEAU INFO PLAN /////////////////////////////////////////////////////////////
            // INIT EN-TETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#ADD8E6' align='center' width='900' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerInfoPlan1 = new string[] { "INFORMATION DU DOSSIER" };
            foreach (string myColumn3 in headerInfoPlan1)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#ADD8E6' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            // FIN ENTETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' width='900' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerInfoPlan2 = new string[] { "NOM, Pr&#233nom (ID) ", "Date", "Plan", "Machine de traitement", "Op&#233rateur" };
            foreach (string myColumn3 in headerInfoPlan2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FAFAD2' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("<tr>");
            foreach (string myColumn3 in InfoPlan)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            //////////////////////////////////// FIN TABLEAU INFO PLAN /////////////////////////////////////////////////////////////

            //////////////////////////////////// TABLEAU PRESCRIPTION /////////////////////////////////////////////////////////////
            // INIT EN-TETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#ADD8E6' align='center' width='900' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerPrescription1 = new string[] { "PRESCRIPTION" };
            foreach (string myColumn3 in headerPrescription1)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#ADD8E6' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            // FIN ENTETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor=''#FAFAD2'' align='center' width='900' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerPrescription2 = new string[] { "Volume cible", "Dose/fr [Gy]", "Fractions", "Dose totale [Gy]" };
            foreach (string myColumn3 in headerPrescription2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("<tr>");

            foreach (string[] myColumn3 in Prescription)
            {
                strHTMLBuilder.Append("<tr>");
                foreach (string myColumn4 in myColumn3)
                {
                    strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                    strHTMLBuilder.Append(myColumn4);
                    strHTMLBuilder.Append("</td>");
                }
                strHTMLBuilder.Append("</tr>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            //////////////////////////////////// FIN TABLEAU PRESCRIPTION /////////////////////////////////////////////////////////////

            //////////////////////////////////// TABLEAU Evaluation de plan /////////////////////////////////////////////////////////////
            // INIT EN-TETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#ADD8E6' align='center' width='900' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerEval1 = new string[] { "EVALUATION DU PLAN" };
            foreach (string myColumn3 in headerEval1)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#ADD8E6' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            // FIN ENTETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' width='900' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerEval2 = new string[] { "Volume de normalisation","D max [%]", "UM [UM]", "Facteur de modulation [UM/cGy]", "Taille de champ max en X [cm]", "Iso Faisceaux (&#916)", "Normalisation" };
            foreach (string myColumn301 in headerEval2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                strHTMLBuilder.Append(myColumn301);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("<tr>");
            foreach (string myColumn301 in EvaPlan)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                strHTMLBuilder.Append(myColumn301);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            //////////////////////////////////// FIN TABLEAU Evaluation de plan /////////////////////////////////////////////////////////////

            //////////////////////////////////// TABLEAU STEREOTAXIE /////////////////////////////////////////////////////////////
            // INIT EN-TETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#ADD8E6' align='center' width='900' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerPTVEVAL2 = new string[] { "CALCUL DES INDICES STEREOTAXIQUES" };
            foreach (string myColumn3 in headerPTVEVAL2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#ADD8E6' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            // FIN ENTETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' width='900' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerPTVEVAL3 = new string[] { "ID", "Volume [cc]", "Vir [cc]", "VTir [cc]", "Isodose minimale [Gy]", "HI", "CI", "RCI", "Paddick", "GI" };
            foreach (string myColumn301 in headerPTVEVAL3)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                strHTMLBuilder.Append(myColumn301);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("<tr>");

            string[] headerPTVEVAL4 = new string[] { string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, "[<2.5]", "[0.7-1]", "[0.9-2.5]", "[0.7-1]", "[<3]" };
            foreach (string myColumn3 in headerPTVEVAL4)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FAFAD2' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("<tr");

            for (int i = 0; i < PTVSTEREO.Length / 10; i++)
            {
                strHTMLBuilder.Append("<tr>");
                for (int j = 0; j < 10; j++)
                {
                    if (PTVSTEREO[i, 0] != null && !PTVSTEREO[i, 0].Contains("z_") && !PTVSTEREO[i, 0].Contains("z-"))
                    {
                        strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                        strHTMLBuilder.Append(PTVSTEREO[i, j]);
                        strHTMLBuilder.Append("</td>");
                    }
                }
                strHTMLBuilder.Append("</tr>");
            }


            //////////////////////////////////// FIN TABLEAU STEREOTAXIE /////////////////////////////////////////////////////////////

            //////////////////////////////////////// TABLEAU CIBLES ///////////////////////////////////////////////////////////////////////

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#ADD8E6' align='center' WIDTH ='900' ;  style='border:dotted 1px Silver; font-family:arial; font-size:small'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerPTVEVAL_PTV = new string[] { "ANALYSE DES HDV : CIBLES" };

            foreach (string myColumn3 in headerPTVEVAL_PTV)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#ADD8E6' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' WIDTH ='900'; style='border:dotted 1px Silver; font-family:arial; font-size:small'>");
            strHTMLBuilder.Append("<tr>");
            try
            {
                foreach (string myColumn301 in header)
                {
                    strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                    strHTMLBuilder.Append(myColumn301);
                    strHTMLBuilder.Append("</td>");

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problème sur la construction du tableau : Header PTV Eval \n" + ex.Message);
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</tr>");

            //strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' WIDTH ='900'; style='border:dotted 1px Silver; font-family:arial; font-size:small'>");
            try
            {
                for (int PTVtabl = 0; PTVtabl < (PTV.Length / 10) - 1; PTVtabl++)
                {
                    if (PTVSTEREO[PTVtabl, 0] == null || PTVSTEREO[PTVtabl, 0].Contains("z_") || PTVSTEREO[PTVtabl, 0].Contains("z-")) { }
                    else
                    {
                        string[] headerPTVinfo2 = new string[] { PTV[PTVtabl, 0], PTV[PTVtabl, 1], PTV[PTVtabl, 2], PTV[PTVtabl, 3], PTV[PTVtabl, 4], PTV[PTVtabl, 5], PTV[PTVtabl, 6], PTV[PTVtabl, 7], PTV[PTVtabl, 8], string.Empty };
                        strHTMLBuilder.Append("</tr>");
                        foreach (string myColumn301 in headerPTVinfo2)
                        {
                            strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                            strHTMLBuilder.Append(myColumn301);
                            strHTMLBuilder.Append("</td>");

                        }
                        strHTMLBuilder.Append("</tr>");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problème sur les cibles \n" + ex.Message);
            }
            strHTMLBuilder.Append("</tr>");

            //////////////////////////////////////// FIN TABLEAU CIBLES ///////////////////////////////////////////////////////////////////////


            //////////////////////////////////////// TABLEAU 3 ///////////////////////////////////////////////////////////////////////
            //INIT EN-TETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#ADD8E6' align='center' WIDTH ='900' ;  style='border:dotted 1px Silver; font-family:arial; font-size:small'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerHDV = new string[] { "ANALYSE DES HDV : OARs" };

            foreach (string myColumn3 in headerHDV)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#ADD8E6' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");
            // FIN ENTETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#F8F8F8' align='center' WIDTH ='900'; style='border:dotted 1px Silver; font-family:arial; font-size:small'>");
            strHTMLBuilder.Append("<tr >");

            foreach (string myColumn in header_OAR)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                strHTMLBuilder.Append(myColumn);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");

            try
            {
                foreach (var obj in m_objectives)
                {
                    if (!(obj.ID.Contains("PTV") || obj.ID.Contains("CTV") || obj.ID.Contains("ITV") || obj.ID.Contains("GTV")))
                    {
                        if (!(obj.ExpectedValue.Contains("Min")) && !(obj.ExpectedValue.Contains("Median")))
                        {
                            strHTMLBuilder.Append("<tr >");
                            strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#F8F8F8' align='center'>");
                            strHTMLBuilder.Append(obj.ID);
                            strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#F8F8F8' align='center'>");
                            strHTMLBuilder.Append(obj.Vol);
                            strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#F8F8F8' align='center'>");
                            strHTMLBuilder.Append(obj.ExpectedValue);
                            strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#F8F8F8' align='center'>");
                            strHTMLBuilder.Append(obj.DVHObjective);

                            string starttag = "";
                            if (obj.Evaluator.ToString().Contains("Over Limit"))
                            {
                                starttag = "<td bgcolor='Red' align='center' style='font-family:arial; font-size:small;'>";
                            }
                            else if (obj.Evaluator.ToString().Contains("Warning"))
                            {
                                starttag = "<td bgcolor='Yellow' align='center' style='font-family:arial; font-size:small;'>";
                            }
                            else if (obj.Evaluator.ToString().Contains("OK"))
                            {
                                starttag = "<td bgcolor='LightGreen' align='center' style='font-family:arial; font-size:small;'>";
                            }
                            else
                            {
                                starttag = "<td style='font-family:arial align='center'; align='center' font-size:small;'>";
                            }

                            strHTMLBuilder.Append(starttag);
                            strHTMLBuilder.Append(obj.RealValue);
                            strHTMLBuilder.Append("</td>");

                            strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#F8F8F8' align='center'>");
                            strHTMLBuilder.Append(obj.Referentiel);
                            strHTMLBuilder.Append("</td>");

                            strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#F8F8F8' align='center'>");
                            strHTMLBuilder.Append(string.Empty);

                            strHTMLBuilder.Append("</td>");
                            strHTMLBuilder.Append("</tr>");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problème sur les objets : Objectifs \n" + ex.Message);
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            ////////////////////////////////////FIN TABLEAU 2 /////////////////////////////////////////////////////////////      
            strHTMLBuilder.Append("</div>"); // Fin conteneur tableaux
            strHTMLBuilder.Append("</div>"); // Fin conteneur flex

            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");

            string Htmltext = strHTMLBuilder.ToString();

            return Htmltext;

        }
        #endregion
    }
}







