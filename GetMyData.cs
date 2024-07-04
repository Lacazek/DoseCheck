using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Documents;
using System.Windows.Shapes;
using System.Xml.Linq;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;

namespace DoseCheck
{
    internal class GetMyData : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private string _userFile;
        private StreamReader _streamReader;
        private StreamWriter _streamWriter;
        private Model _model;
        private GeneratePDF _generatePDF;
        private CreateExcelForStats _createExcelForStats;
        private Dictionary<string, string> _results;


        internal GetMyData(Model _m)
        {
            _userFile = string.Empty;
            _model = _m;
            _generatePDF = new GeneratePDF(_model);
            _results = new Dictionary<string, string>();
            _createExcelForStats = new CreateExcelForStats(_model);
        }

        public void MyData()
        {
            string line, _s = "";
            bool isDosePrescribed = true;
            int i = 1;
            int addNLines = 0;
            PlanSetup myPlan = _model.PlanSetup;
            Structure Body = myPlan.StructureSet.Structures.FirstOrDefault(x => x.DicomType.ToUpper() == "EXTERNAL");
            Structure st = Body;
            if (_model.RTPrescription == null)
            {
                MessageBox.Show("Il n'y a aucune prescription numérique rattachée à ce plan");
                isDosePrescribed = false;
            }

            string d_at_v_pattern = @"^D(?<evalpt>\d+\p{P}\d+|\d+)(?<unit>(%|cc))$"; // matches D95%, D2cc
            string v_at_d_pattern = @"^V(?<evalpt>\d+\p{P}\d+|\d+)(?<unit>(%|cc))$"; // matches V50.4cc or V50.4% 

            #region Modification du fichier template
            // Permet d'ajouter dans le fichier texte les contraintes spécifiques de la prescription 
            // Ces lignes seront ajoutés au template puis supprimées à la fin pour ne laisser que le template

            try
            {
                _streamWriter = new StreamWriter(_userFile, true);

                if (isDosePrescribed)
                {
                    foreach (var index in _model.RTPrescription.Targets)
                    {
                        foreach (var constraint in index.Constraints)
                        {
                            _streamWriter.WriteLine(index.TargetId + ";" + constraint.Value1 + "  " + constraint.Unit1 + "," + constraint.Value2 + "  " + constraint.Unit2);
                            addNLines++;
                        }
                    }
                    foreach (var OAR in _model.RTPrescription.OrgansAtRisk)
                    {
                        foreach (var constraint in OAR.Constraints)
                        {
                            _streamWriter.WriteLine(OAR.OrganAtRiskId + ";" + constraint.Value1 + "  " + constraint.Unit1 + "," + constraint.Value2 + "  " + constraint.Unit2);
                            addNLines++;

                            //Tuple<string, string> values = Tuple.Create(cons.Value1, cons.Value2);
                            //_oar_results.Add(_s + "/ Contrainte /" + cons.Value1, values);
                        }
                    }
                }
                _streamWriter.Close();
                _streamReader = new StreamReader(_userFile);
                _streamReader.ReadLine(); // ignore la 1ere ligne
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            #endregion

            #region lecture du fichier
            while ((line = _streamReader.ReadLine()) != null)
            {
                var testMatchD_at_V = Regex.Matches((line.Split(';')[1]).Split(',')[0], d_at_v_pattern);
                var testMatchV_at_D = Regex.Matches((line.Split(';')[1]).Split(',')[0], v_at_d_pattern);
                double similarity, maxSimilarity = 0;


                // Matching ID structure dans le fichier et dans le SS.
                // L'ID récupéré correspond au nom de la structure présente dans le SS (modification de l'ID de la structure du template)

                foreach (Structure structure in myPlan.StructureSet.Structures)
                {
                    similarity = CalculateSimilarity(structure.Id, line.Split(';')[0]);
                    if (similarity > maxSimilarity)
                    {
                        maxSimilarity = similarity;
                        _s = structure.Id;
                        st = structure;
                    }
                }

                // Change de nom local pour pouvoir différencier les clés (!!! Attention pas de la structure !!!) 
                if (!_results.ContainsKey(_s))
                {
                    _s = _s + " / " + i;
                    i++;
                }
                DVHData myDVH = myPlan.GetDVHCumulativeData(st, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
                //Permet d'ajouter une fois à chaque structure la dose max,moyenne médiane et min
                if (!_results.Keys.Any(x => x.Contains(_s.Split('/')[0])))
                {
                    _results.Add(_s + " / Max dose", Math.Round(myDVH.MaxDose.Dose, 2).ToString() + " Gy");
                    _results.Add(_s + " / Mean Dose", Math.Round(myDVH.MeanDose.Dose, 2).ToString() + " Gy");
                    _results.Add(_s + " / Median Dose", Math.Round(myDVH.MedianDose.Dose, 2).ToString() + " Gy");
                    _results.Add(_s + " / Min Dose", Math.Round(myDVH.MinDose.Dose, 2).ToString() + " Gy");
                }


                if (testMatchD_at_V.Count != 0) // count is 1 if D95% or D2cc
                {
                    Group eval = testMatchD_at_V[0].Groups["evalpt"];
                    Group unit = testMatchD_at_V[0].Groups["unit"];
                    DoseValue.DoseUnit du = DoseValue.DoseUnit.Gy;
                    DoseValue myD_something = new DoseValue(1000.1000, du);
                    double myD = Convert.ToDouble(eval.Value);
                    if (unit.Value == "%")
                    {
                        _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1], Math.Round(myPlan.GetDoseAtVolume(st, myD, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose, 3).ToString() + " Gy");
                        // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif] et le résultats au format [ résultat unité]

                    }
                    else if (unit.Value == "cc")
                    {
                        _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1], Math.Round(myPlan.GetDoseAtVolume(st, myD, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose, 3).ToString() + " Gy");
                        // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif] et le résultats au format [ résultat unité]
                    }
                    else
                        _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1], "-1.00");
                    // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif] et le résultats au format [ résultat unité]
                }

                if (testMatchV_at_D.Count != 0) // count is 1
                {
                    Group eval = testMatchV_at_D[0].Groups["evalpt"];
                    Group unit = testMatchV_at_D[0].Groups["unit"];
                    DoseValue.DoseUnit du = DoseValue.DoseUnit.Gy;
                    DoseValue myRequestedDose = new DoseValue(Convert.ToDouble(eval.Value), du);

                    if (unit.Value == "cc")
                        _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1], Math.Round(myPlan.GetVolumeAtDose(st, myRequestedDose, VolumePresentation.AbsoluteCm3), 3).ToString() + " cc");
                    // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif] et le résultats au format [ résultat unité]
                    else if (unit.Value == "%")
                        _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1], Math.Round(myPlan.GetVolumeAtDose(st, myRequestedDose, VolumePresentation.Relative), 2).ToString() + " %");
                    // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif] et le résultats au format [ résultat unité]
                    else
                        _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1], "-1.00");
                    // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif] et le résultats au format [ résultat unité]
                }
                #endregion
                MessageBox.Show("pre - test ");
                MessageBox.Show( Math.Round(Convert.ToDouble(myPlan.GetVolumeAtDose(st, 95 * (DoseValue)(_model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction * _model.RTPrescription.NumberOfFractions), VolumePresentation.Relative)), 2).ToString() + " %");
                MessageBox.Show("post - test ");
                //
                // ici l'objet ne s'initialise pas mais la prescription est bien instanciée
                //

                #region Indice
                try
                {
                    if (isDosePrescribed)
                    {
                        if (st.Id.Contains("PTV") || st.Id.Contains("CTV") || st.Id.Contains("ITV") || st.Id.Contains("GTV") && st.Id == _model.PlanSetup.TargetVolumeID && !st.Id.ToUpper().Contains("z_"))
                        {
                            if(!_results.Keys.Contains(_s))
                            {
                                _results.Add(_s + "/ D95%", Math.Round(Convert.ToDouble(myPlan.GetDoseAtVolume(st, 95 * (double)(_model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction.Dose * _model.RTPrescription.NumberOfFractions), VolumePresentation.Relative, DoseValuePresentation.Absolute)), 2).ToString() + " Gy");
                                _results.Add(_s + "/ V95%", Math.Round(Convert.ToDouble(myPlan.GetVolumeAtDose(st, 95 * (DoseValue)(_model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction * _model.RTPrescription.NumberOfFractions), VolumePresentation.Relative)), 2).ToString() + " %");

                                #region HomogenityIndex
                                double d02 = Convert.ToDouble(myPlan.GetDoseAtVolume(st, 2 * (double)(_model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction.Dose * _model.RTPrescription.NumberOfFractions), VolumePresentation.AbsoluteCm3, DoseValuePresentation.Relative));
                                double d98 = Convert.ToDouble(myPlan.GetDoseAtVolume(st, 98 * (double)(_model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction.Dose * _model.RTPrescription.NumberOfFractions), VolumePresentation.AbsoluteCm3, DoseValuePresentation.Relative));
                                double d50 = Convert.ToDouble(myPlan.GetDoseAtVolume(st, 50 * (double)(_model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction.Dose * _model.RTPrescription.NumberOfFractions), VolumePresentation.AbsoluteCm3, DoseValuePresentation.Relative));
                                _results.Add(_s + "/ HI", (Math.Round((d02 - d98) / d50, 3)).ToString());
                                #endregion
                                #region ConformityIndex
                                //Conformity Index requres Body as input structure for dose calc and volume of target 
                                double volIsodoseLvl = myPlan.GetVolumeAtDose(Body, _model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction * (int)_model.RTPrescription.NumberOfFractions, VolumePresentation.AbsoluteCm3);
                                _results.Add(_s + "/ CI", Math.Round(volIsodoseLvl / st.Volume, 3).ToString());
                                #endregion
                                #region PaddickConformityIndex
                                double PIV = myPlan.GetVolumeAtDose(Body, _model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction * (int)_model.RTPrescription.NumberOfFractions, VolumePresentation.AbsoluteCm3);
                                double TV_PIV = myPlan.GetVolumeAtDose(st, _model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction * (int)_model.RTPrescription.NumberOfFractions, VolumePresentation.AbsoluteCm3);
                                _results.Add(_s + "/ PADDICK", Math.Round((TV_PIV * TV_PIV) / (st.Volume * PIV), 3).ToString());
                                #endregion
                                #region GradientIndex
                                double v50 = myPlan.GetVolumeAtDose(Body, _model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction * (int)_model.RTPrescription.NumberOfFractions * 0.5, VolumePresentation.AbsoluteCm3);
                                double v100 = myPlan.GetVolumeAtDose(Body, _model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction * (int)_model.RTPrescription.NumberOfFractions, VolumePresentation.AbsoluteCm3);
                                _results.Add(_s + "/ GI", Math.Round((v50 / v100), 2).ToString());
                                #endregion
                                #region RCI
                                double volTIP = myPlan.GetVolumeAtDose(st, _model.RTPrescription.Targets.Where(x => x.Name == st.Id).FirstOrDefault().DosePerFraction * (int)_model.RTPrescription.NumberOfFractions, VolumePresentation.AbsoluteCm3);
                                _results.Add(_s + "/ RCI", Math.Round(volTIP / st.Volume, 3).ToString());
                                #endregion
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            #endregion
            _results.OrderBy(kvp => kvp.Key);
            _createExcelForStats.Fill(_results);
            _createExcelForStats.Close();
            _generatePDF.Execute(_results);
            _streamReader.Close();
            RemoveLastNLines(_userFile, addNLines); // ligne à évaluer
        }
        #region Calcul of similarity (Distance de Levenshtein)
        internal double CalculateSimilarity(string name, string key)
        {

            int n = name.Length;
            int m = key.Length;
            int[,] d = new int[n + 1, m + 1];

            int maxLength = Math.Max(n, m);

            if (n == 0) return m;
            if (m == 0) return n;

            for (int i = 0; i <= n; d[i, 0] = i++) ;
            for (int j = 0; j <= m; d[0, j] = j++) ;

            for (int i = 1; i <= n; i++)
            {
                for (int j = 1; j <= m; j++)
                {
                    int cost = (key[j - 1] == name[i - 1]) ? 0 : 1;
                    d[i, j] = Math.Min(
                        Math.Min(d[i - 1, j] + 1, d[i, j - 1] + 1),
                            d[i - 1, j - 1] + cost);
                }
            }
            return 1.0 - (double)d[n, m] / maxLength;
        }
        #endregion

        #region Remove lines
        static void RemoveLastNLines(string filePath, int n)
        {
            try
            {
                var lines = File.ReadAllLines(filePath);
                if (lines.Length >= n)
                {
                    File.WriteAllLines(filePath, lines.Take(lines.Length - n));
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        #endregion

        #region Get and Set
        internal string UserFile
        {
            get { return _userFile; }
            set
            {
                _userFile = System.IO.Path.Combine(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), "Template_dosi"), value + ".txt");
                //OnPropertyChanged(nameof(_userFile));
                MyData();
                //_generatePDF.PlanReport(_results, _oar_results, _model.Patient, _model.PlanSetup);
            }
        }
        internal Dictionary<string, string> Results
        {
            get { return _results; }
        }

        protected virtual void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }
        #endregion

    }
}