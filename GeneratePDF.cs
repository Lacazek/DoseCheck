using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.Remoting.Contexts;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Shapes;
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
            public string Prescription { get; set; }
        }

        static int CompareValues(double x, double y, double tol, string name)
        {
            if (name.Contains(@"(?i)\bt\s*v\s\b"))
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


        private Model _model;
        private Dictionary<string, string> _results;
        private string WORKBOOK_TEMPLATE_DIR;
        private string WORKBOOK_RESULT_DIR;
        private List<StructureObjective> m_objectives; //all the dose values are internally in Gy.


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

                try
                {
                    foreach (var item in _results)
                    {
                        m_objectives = new List<StructureObjective>();
                        {
                            StructureObjective obj = new StructureObjective();
                            {
                                obj.ID = item.Key.Split('/')[0];
                                obj.DVHObjective = item.Key.Split('/')[3];
                                obj.ExpectedValue = item.Key.Split('/')[2];
                                obj.RealValue = item.Value;
                                switch (CompareValues(Convert.ToDouble(item.Value), Convert.ToDouble(item.Key.Split('/')[2]), tol, item.Key.Split('/')[0]))
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
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Erreur dans la génération du PDF\n" + ex.Message);
                }

                try
                {
                    // if a plansum is loaded take the first plansum to work on, otherwise take the active plansetup.

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

                double Um = 0, Xb = 0, Yb = 0, Zb = 0, XMax = 0;
                double[] X1 = { 0, 0, 0, 0, 0, 0 };
                double[] Y1 = { 0, 0, 0, 0, 0, 0 };
                double[] X2 = { 0, 0, 0, 0, 0, 0 };
                double[] Y2 = { 0, 0, 0, 0, 0, 0 };
                int i = 0, y = 0;

                foreach (var b in _model.PlanSetup.Beams)
                {
                    double XD = _model.Image.UserOrigin.x;
                    double YD = _model.Image.UserOrigin.y;
                    double ZD = _model.Image.UserOrigin.z;

                    // COORDONNEES DICOM DU FAISCEAU + TAILLE DE CHAMP + UM
                    if (b.Meterset.Value.ToString() == "Non Numérique") { }
                    else
                    {
                        MessageBox.Show(Um.ToString());
                        // ADDITION UM POUR CHAQUE FAISCEAU
                        Um = b.Meterset.Value + Um;
                        //VERIFICATION ET COMPARAISON POSITION MAX MACHOIRES DE CHAQUE FAISCEAU
                        X1[i] = Math.Round(b.ControlPoints.First().JawPositions.X1 / 10, _decimal);
                        X2[i] = Math.Round(b.ControlPoints.First().JawPositions.X2 / 10, _decimal);
                        double xmaxi = (Math.Abs(X1[i] - X2[i]));
                        if (xmaxi >= XMax) { XMax = xmaxi; }
                        // POSITION DE CHAQUE FAISCEAU
                        Xb = Math.Round(b.IsocenterPosition.x, _decimal);
                        Yb = Math.Round(b.IsocenterPosition.y, _decimal);
                        Zb = Math.Round(b.IsocenterPosition.z, _decimal);
                        i++;
                    }
                }
                string Iso = string.Format("X : {0} cm  Y : {2} cm  Z : {1} cm", Math.Abs(Math.Round((Xb - _model.Image.UserOrigin.x) / 10, _decimal)), Math.Abs(Math.Round((Yb - _model.Image.UserOrigin.y) / 10, _decimal)), Math.Abs(Math.Round((_model.Image.UserOrigin.z - Zb) / 10, _decimal)));
                string[] InfoPlan = new string[] { _model.Patient.Name.ToString(), DateTime.Now.ToString("yyyy-MM-dd hh:mm:ss"), _model.PlanSetup.Id.ToString(), _model.PlanSetup.Beams.First().TreatmentUnit.ToString(), _model.PlanSetup.CreationUserName.ToString() };
                string[] EvalPlan = new string[] { _model.PlanSetup.Dose.DoseMax3D.ToString(), Math.Round(Um, _decimal).ToString(), Math.Round((Um / _model.PlanSetup.TotalDose.Dose * 100), _decimal).ToString(), XMax.ToString(), Iso };
                string[] Prescription = new string[] { _model.PlanSetup.TotalDose.ToString(), _model.PlanSetup.DosePerFraction.ToString(), _model.PlanSetup.NumberOfFractions.ToString(), _model.PlanSetup.TargetVolumeID.ToString() };
                string[] Header = new string[] { "ID Cible", "Volume cc" , "D95%", "D50%", "Dmax", "Dmean", "Dmoy", "Dmin" };
                
                string[] PTV_Name = new string[100], PTV_95s = new string[100], PTV_MAXs = new string[100];
                double[] PTV_95 = new double[100], PTV_MAX = new double[100], TVpi = new double[100], Vptv = new double[100];
                double[] R100 = new double[100], R50 = new double[100];
                double[] PADDICK = new double[100], HI = new double[100], CI = new double[100] , GI = new double[100], RCI = new double[100];
                string[,] PTVSTEREO = new string[100, 6], PTVSTEREOCalc = new string[100, 9];

                //Structure[] PTV_Structure=new structure [100];
                Structure structu = null;

                foreach (Structure scan in _model.StructureSet.Structures)
                {
                    //if (scan.Id.Contains(@"(?i)\bt\s*v\s\b") && scan.Id == _model.PlanSetup.TargetVolumeID)
                    if (scan.Id.Contains("PTV") || scan.Id.Contains("CTV") || scan.Id.Contains("ITV") || scan.Id.Contains("GTV") && scan.Id == _model.PlanSetup.TargetVolumeID)
                    {
                        PTV_Name[y] = _model.PlanSetup.TargetVolumeID.ToString();
                        structu = scan;
                        

                        // Problème dans la calcul de la dose prescrite par cible
                        double V100 = Math.Round(_model.PlanSetup.GetVolumeAtDose(structu, 100 * (DoseValue)(_model.RTPrescription.Targets.Where(x => x.Name == structu.Id).FirstOrDefault().DosePerFraction * _model.RTPrescription.NumberOfFractions), VolumePresentation.Relative),_decimal) ;
                        double V95 = Math.Round(_model.PlanSetup.GetVolumeAtDose(structu, 95 * (DoseValue)(_model.RTPrescription.Targets.Where(x => x.Name == structu.Id).FirstOrDefault().DosePerFraction * _model.RTPrescription.NumberOfFractions), VolumePresentation.Relative),_decimal);                
                        double V50 = Math.Round(_model.PlanSetup.GetVolumeAtDose(structu, 50 * (DoseValue)(_model.RTPrescription.Targets.Where(x => x.Name == structu.Id).FirstOrDefault().DosePerFraction * _model.RTPrescription.NumberOfFractions), VolumePresentation.Relative),_decimal);                      
                        /*
                        double V100 = Math.Round(_model.PlanSetup.GetVolumeAtDose(structu, new DoseValue(100 *_model.PlanSetup.TotalDose.Dose, DoseValue.DoseUnit.Gy), VolumePresentation.Relative), _decimal);
                        double V95 = Math.Round(_model.PlanSetup.GetVolumeAtDose(structu, new DoseValue(95 * _model.PlanSetup.TotalDose.Dose, DoseValue.DoseUnit.Gy), VolumePresentation.Relative), _decimal);
                        double V50 = Math.Round(_model.PlanSetup.GetVolumeAtDose(structu, new DoseValue(2 * _model.PlanSetup.TotalDose.Dose, DoseValue.DoseUnit.Gy), VolumePresentation.Relative), _decimal);  
                        */
                        double V0 = Convert.ToDouble(_results.Where(x => x.Key.Split(';')[0] == scan.Id).First().Value);
                        //PTV_95[y] = _model.PlanSetup.GetVolumeAtDose(structu, new DoseValue( * _model.PlanSetup.TotalDose.Dose, DoseValue.DoseUnit.Gy), VolumePresentation.Relative);
                        PTV_95[y] = V95;
                        PTV_95s[y] = V95.ToString();
                        PTV_MAX[y] = V0;
                        PTV_MAXs[y] = V0.ToString();

                        TVpi[y] = V95;
                        Vptv[y] = Math.Round(structu.Volume, _decimal);

                        R100[y] = V100 / Vptv[y];
                        R50[y] = V50 / Vptv[y];

                        HI[y] = Convert.ToDouble(_results.Where(x => x.Key.Split(';')[0] == scan.Id || x.Key.Split(';')[2] == "HI").First().Value);
                        CI[y] = Convert.ToDouble(_results.Where(x => x.Key.Split(';')[0] == scan.Id || x.Key.Split(';')[2] == "CI").First().Value);
                        RCI[y] = Convert.ToDouble(_results.Where(x => x.Key.Split(';')[0] == scan.Id || x.Key.Split(';')[2] == "RCI").First().Value);
                        PADDICK[y] = Convert.ToDouble(_results.Where(x => x.Key.Split(';')[0] == scan.Id || x.Key.Split(';')[2] == "PADDICK").First().Value);
                        GI[y] = Convert.ToDouble(_results.Where(x => x.Key.Split(';')[0] == scan.Id || x.Key.Split(';')[2] == "GI").First().Value);

                        PTVSTEREO[y, 0] = PTV_Name[y];
                        PTVSTEREO[y, 1] = scan.Volume.ToString();
                        PTVSTEREO[y, 2] = PTV_95s[y];
                        PTVSTEREO[y, 3] = PTV_MAXs[y];
                        PTVSTEREO[y, 4] = Math.Round(HI[y], _decimal).ToString();
                        PTVSTEREO[y, 5] = Math.Round(CI[y], _decimal).ToString();
                        PTVSTEREO[y, 6] = Math.Round(CI[y], _decimal).ToString();
                        PTVSTEREO[y, 7] = Math.Round(PADDICK[y], _decimal).ToString();
                        PTVSTEREO[y, 8] = Math.Round(GI[y], _decimal).ToString();
                    }
                    y++;
                }
                string HtmlBody = ExportToHtml(Header, InfoPlan, Prescription, EvalPlan, PTVSTEREOCalc, PTVSTEREO);
                string outputpath = System.IO.Path.Combine(WORKBOOK_RESULT_DIR + DateTime.Now.ToString("yyyy-MM-dd") + "-" + _model.Patient.Name.ToString() + pitem.Id.ToString() + ".html");

                System.IO.File.WriteAllText(outputpath, HtmlBody);
                System.Diagnostics.Process.Start(outputpath);
                System.Threading.Thread.Sleep(3000);

            }
        }

        ////////////////////////////////////////////////////////////// FIN EXECUTE ///////////////////////////////////////////////////
        #endregion

        //convert the results data to something that can be HTML-ified

        /*
        #region ExportHTML
        protected string ExportToHtml(string[] header, string[] InfoPlan, string[] Prescription, string[] EvaPlan, string[,] PTVSTEREOCalc, string[,] PTVSTEREO)
        {

            StringBuilder strHTMLBuilder = new StringBuilder();
            strHTMLBuilder.Append("<html >");
            strHTMLBuilder.Append("<head>");
            strHTMLBuilder.Append("</head>");
            strHTMLBuilder.Append("<body style='font-family:arial; font-size:medium'>");
            strHTMLBuilder.Append("<br>");

            strHTMLBuilder.Append("<div style='position:relative;'>");
            strHTMLBuilder.Append("<img src='B:\\RADIOTHERAPIE\\Killian\\Dosi\\Script\\DoseCheck\\Projects\\DoseCheck\\fisherMan4.png' alt='Description de l'image' style='position:absolute; top:10; left:10; width:100px; height:auto;'>");
            strHTMLBuilder.Append("</div>");

            strHTMLBuilder.Append("<br>");

            ////////////////////////////////////TABLEAU INFO PLAN  /////////////////////////////////////////////////////////////
            //INIT EN-TETE

            //color #FFD700
            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FF95D1FF' align='center' WIDTH ='600' ;  style='border:dotted 1px Silver; font-family:arial; font-size:small'>");

            string[] headerInfoPlan1 = new string[] { "INFORMATION DU DOSSIER" };

            foreach (string myColumn3 in headerInfoPlan1)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FF95D1FF' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");
            // FIN ENTETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' WIDTH ='600' ;  style='border:dotted 1px Silver; font-family:arial; font-size:small'>");

            string[] headerInfoPlan2 = new string[] { "NOM, Pr&#233nom (ID) ", "Date", "Plan", "Machine de TTT", "Op&#233rateur" };

            foreach (string myColumn3 in headerInfoPlan2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FAFAD2' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");

            foreach (string myColumn3 in InfoPlan)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }



            strHTMLBuilder.Append("</table>");

            ////////////////////////////////////FIN TABLEAU INFO PLAN /////////////////////////////////////////////////////////////
            //	strHTMLBuilder.Append("<br>");



            ////////////////////////////////////TABLEAU PRESRIPTION  /////////////////////////////////////////////////////////////
            //INIT EN-TETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FFD700' align='center' WIDTH ='600' ;  style='border:dotted 1px Silver; font-family:arial; font-size:small'>");

            string[] headerPrescription1 = new string[] { "PRESCRIPTION" };

            foreach (string myColumn3 in headerPrescription1)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FFD700' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");
            // FIN ENTETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' WIDTH ='600'; style='border:dotted 1px Silver; font-family:arial; font-size:small'>");

            string[] headerPrescription2 = new string[] { "Dose totale (Gy)", "Dose/fr (Gy)", "Nfr", "Volume cible" };

            foreach (string myColumn3 in headerPrescription2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");


            foreach (string myColumn3 in Prescription)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }

            strHTMLBuilder.Append("</table>");

            ////////////////////////////////////FIN TABLEAU PRESCRIPTION /////////////////////////////////////////////////////////////


            ////////////////////////////////////TABLEAU Evaluation de plan  /////////////////////////////////////////////////////////////
            //INIT EN-TETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FFD700' align='center' WIDTH ='600' ;  style='border:dotted 1px Silver; font-family:arial; font-size:small'>");

            string[] headerEval1 = new string[] { "EVALUATION DU PLAN" };

            foreach (string myColumn3 in headerEval1)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FFD700' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");
            // FIN ENTETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' WIDTH ='600'; style='border:dotted 1px Silver; font-family:arial; font-size:small'>");

            string[] headerEval2 = new string[] { "Dmax", "UM", "Facteur de modulation", "Taille de champ max en X (cm)", "Iso Faisceaux (Beta)" };

            foreach (string myColumn301 in headerEval2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                strHTMLBuilder.Append(myColumn301);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");


            foreach (string myColumn301 in EvaPlan)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                strHTMLBuilder.Append(myColumn301);
                strHTMLBuilder.Append("</td>");
            }

            strHTMLBuilder.Append("</table>");

            ////////////////////////////////////FIN TABLEAU Evaluation de plan /////////////////////////////////////////////////////////////

            ////////////////////////////////////TABLEAU STEREOTAXIE  /////////////////////////////////////////////////////////////
            //INIT EN-TETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FFD700' align='center' WIDTH ='600' ;  style='border:dotted 1px Silver; font-family:arial; font-size:small'>");

            string[] headerPTVEVAL2 = new string[] { "CALCUL DES INDICES STEREOTAXIQUES" };

            foreach (string myColumn3 in headerPTVEVAL2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FFD700' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");
            // FIN ENTETE    

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' WIDTH ='600'; style='border:dotted 1px Silver; font-family:arial; font-size:small'>");
            string[] headerPTVEVAL3 = new string[] { "ID", "Volume [cc]", "V95%", "Dmax", "HI", "CI", "RCI", "Paddick", "GI" };
            //string[] refCI = new string[] { "PTV", "[%]", "[%]", "[< 2.5]", "[0.7-1]", "[0.9-2.5]", "[0.7-1]", "[< 3]" };

            try
            {
                foreach (string myColumn301 in headerPTVEVAL3)
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
            
            try
            {
                for (int PTVtabl = 0; PTVtabl < 100; PTVtabl++)
                {
                    if (PTVSTEREO[PTVtabl, 0] == null) { }
                    else
                    {
                        string[] headerPTVinfo2 = new string[] { PTVSTEREO[PTVtabl, 0], PTVSTEREO[PTVtabl, 1], PTVSTEREO[PTVtabl, 2], PTVSTEREO[PTVtabl, 3], PTVSTEREO[PTVtabl, 4], PTVSTEREO[PTVtabl, 5], PTVSTEREO[PTVtabl, 6], PTVSTEREO[PTVtabl, 7], PTVSTEREO[PTVtabl, 8] };

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
                MessageBox.Show("Problème sur les PTV \n" + ex.Message);
            }
            strHTMLBuilder.Append("</table>");

            ////////////////////////////////////FIN TABLEAU TABLEAU STEREOTAXIE  /////////////////////////////////////////////////////////////
            ///

            //////////////////////////////////////// TABLEAU CIBLES ///////////////////////////////////////////////////////////////////////

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FFD700' align='center' WIDTH ='600' ;  style='border:dotted 1px Silver; font-family:arial; font-size:small'>");

            string[] headerPTVEVAL_PTV = new string[] { "Recap' : CIBLES" };

            foreach (string myColumn3 in headerPTVEVAL_PTV)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FFD700' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' WIDTH ='600'; style='border:dotted 1px Silver; font-family:arial; font-size:small'>");
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
           
            try
            {
                for (int PTVtabl = 0; PTVtabl < 100; PTVtabl++)
                {
                    if (PTVSTEREO[PTVtabl, 0] == null) { }
                    else
                    {
                        string[] headerPTVinfo2 = new string[] { PTVSTEREO[PTVtabl, 0], PTVSTEREO[PTVtabl, 1], PTVSTEREO[PTVtabl, 2], PTVSTEREO[PTVtabl, 3], PTVSTEREO[PTVtabl, 4], PTVSTEREO[PTVtabl, 5], PTVSTEREO[PTVtabl, 6], PTVSTEREO[PTVtabl, 7] };

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
                MessageBox.Show("Problème sur les PTV \n" + ex.Message);
            }
            strHTMLBuilder.Append("</table>");


            //////////////////////////////////////// FIN TABLEAU CIBLES ///////////////////////////////////////////////////////////////////////


            //////////////////////////////////////// TABLEAU 3 ///////////////////////////////////////////////////////////////////////
            //INIT EN-TETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#FFD700' align='center' WIDTH ='600' ;  style='border:dotted 1px Silver; font-family:arial; font-size:small'>");

            string[] headerHDV = new string[] { "ANALYSE DES HDV : Cibles et OARs" };

            foreach (string myColumn3 in headerHDV)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FFD700' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");
            // FIN ENTETE

            strHTMLBuilder.Append("<table border='1px' cellpadding='1' cellspacing='0' bgcolor='#F8F8F8' align='center' WIDTH ='600'; style='border:dotted 1px Silver; font-family:arial; font-size:small'>");
            strHTMLBuilder.Append("<tr >");

            foreach (string myColumn in header)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                strHTMLBuilder.Append(myColumn);
                strHTMLBuilder.Append("</td>");

            }
            strHTMLBuilder.Append("</tr>");

            foreach (var obj in m_objectives)
            {
                MessageBox.Show("OK   KK");
            }
           
            try
            {
                foreach (var obj in m_objectives)
                {
                    strHTMLBuilder.Append("<tr >");
                    strHTMLBuilder.Append(obj.ID);
                    strHTMLBuilder.Append(obj.ExpectedValue);
                    strHTMLBuilder.Append(obj.DVHObjective);
                    strHTMLBuilder.Append(obj.RealValue);

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
                    //strHTMLBuilder.Append(myRow[myColumn.ColumnName].ToString());
                    strHTMLBuilder.Append("</td>");

                    strHTMLBuilder.Append("</tr>");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Problème sur les objets : Objectifs \n" + ex.Message);
            }

            ////////////////////////////////////FIN TABLEAU 2 /////////////////////////////////////////////////////////////      


            //Close tags.  
            strHTMLBuilder.Append("</table>");
            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");

            string Htmltext = strHTMLBuilder.ToString();

            return Htmltext;

        }*/

        #region ExportToHtml
        protected string ExportToHtml(string[] header, string[] InfoPlan, string[] Prescription, string[] EvaPlan, string[,] PTVSTEREOCalc, string[,] PTVSTEREO)
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
            strHTMLBuilder.Append("<img src='B:\\RADIOTHERAPIE\\Killian\\Dosi\\Script\\DoseCheck\\Projects\\DoseCheck\\fisherMan4.png' alt='Description de l'image' style='width:100px; height:auto;'>");
            strHTMLBuilder.Append("</div>");

            // Conteneur pour les tableaux
            strHTMLBuilder.Append("<div>");

            //////////////////////////////////// TABLEAU INFO PLAN /////////////////////////////////////////////////////////////
            // INIT EN-TETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FF95D1FF' align='center' width='600' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerInfoPlan1 = new string[] { "INFORMATION DU DOSSIER" };
            foreach (string myColumn3 in headerInfoPlan1)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FF95D1FF' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            // FIN ENTETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' width='600' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerInfoPlan2 = new string[] { "NOM, Prénom (ID) ", "Date", "Plan", "Machine de TTT", "Opérateur" };
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
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FFD700' align='center' width='600' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerPrescription1 = new string[] { "PRESCRIPTION" };
            foreach (string myColumn3 in headerPrescription1)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FFD700' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            // FIN ENTETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' width='600' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerPrescription2 = new string[] { "Dose totale (Gy)", "Dose/fr (Gy)", "Nfr", "Volume cible" };
            foreach (string myColumn3 in headerPrescription2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("<tr>");
            foreach (string myColumn3 in Prescription)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            //////////////////////////////////// FIN TABLEAU PRESCRIPTION /////////////////////////////////////////////////////////////

            //////////////////////////////////// TABLEAU Evaluation de plan /////////////////////////////////////////////////////////////
            // INIT EN-TETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FFD700' align='center' width='600' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerEval1 = new string[] { "EVALUATION DU PLAN" };
            foreach (string myColumn3 in headerEval1)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FFD700' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            // FIN ENTETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' width='600' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerEval2 = new string[] { "Dmax", "UM", "Facteur de modulation", "Taille de champ max en X (cm)", "Iso Faisceaux (Beta)" };
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
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FFD700' align='center' width='600' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerPTVEVAL2 = new string[] { "CALCUL DES INDICES STEREOTAXIQUES" };
            foreach (string myColumn3 in headerPTVEVAL2)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' bgcolor='#FFD700' align='center'>");
                strHTMLBuilder.Append(myColumn3);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");
            strHTMLBuilder.Append("</table>");

            // FIN ENTETE
            strHTMLBuilder.Append("<table border='1' cellpadding='1' cellspacing='0' bgcolor='#FAFAD2' align='center' width='600' style='border:dotted 1px Silver; font-family:arial; font-size:small;'>");
            strHTMLBuilder.Append("<tr>");
            string[] headerPTVEVAL3 = new string[] { "ID", "Volume [cc]", "V95%", "Dmax", "HI", "CI", "RCI", "Paddick", "GI" };
            foreach (string myColumn301 in headerPTVEVAL3)
            {
                strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#FAFAD2'>");
                strHTMLBuilder.Append(myColumn301);
                strHTMLBuilder.Append("</td>");
            }
            strHTMLBuilder.Append("</tr>");

            for (int i = 0; i <= PTVSTEREOCalc.GetUpperBound(0); i++)
            {
                strHTMLBuilder.Append("<tr>");
                for (int j = 0; j <= PTVSTEREOCalc.GetUpperBound(1); j++)
                {
                    strHTMLBuilder.Append("<td style='font-family:arial' align='center' bgcolor='#F8F8F8'>");
                    strHTMLBuilder.Append(PTVSTEREOCalc[i, j]);
                    strHTMLBuilder.Append("</td>");
                }
                strHTMLBuilder.Append("</tr>");
            }
            strHTMLBuilder.Append("</table>");

            //////////////////////////////////// FIN TABLEAU STEREOTAXIE /////////////////////////////////////////////////////////////

            strHTMLBuilder.Append("</div>"); // Fin conteneur tableaux
            strHTMLBuilder.Append("</div>"); // Fin conteneur flex

            strHTMLBuilder.Append("</body>");
            strHTMLBuilder.Append("</html>");

            return strHTMLBuilder.ToString();
        }
#endregion
    }
}







