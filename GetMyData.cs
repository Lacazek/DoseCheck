using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using VMS.TPS.Common.Model.API;
using VMS.TPS.Common.Model.Types;


// Class pour récupérer les résultats du fichier txt et de la prescription 
// Permet de générer tous les calculs 
// Appelle GeneratePDF pour réaliser le fichier de sortie (html)
// Appelle CreateExcelForStatss afin de réaliser des études par la suite (si besoin)

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
        private Dictionary<string, string> _index;
        private PairOrgan _pairOrgan;


        internal GetMyData(Model _m)
        {
            _userFile = string.Empty;
            _model = _m;
            _generatePDF = new GeneratePDF(_model);
            _results = new Dictionary<string, string>();
            _index = new Dictionary<string, string>();
            _createExcelForStats = new CreateExcelForStats(_model);
            _pairOrgan = new PairOrgan();
        }

        public void MyData()
        {
            string line, _s = "";
            int i = 1;
            int addNLines = 0;
            PlanSetup myPlan = _model.PlanSetup;
            Structure Body = myPlan.StructureSet.Structures.FirstOrDefault(x => x.DicomType.ToUpper() == "EXTERNAL");
            Structure st = Body;

            string d_at_v_pattern = @"^D(?<evalpt>\d+\p{P}\d+|\d+)(?<unit>(%|cc))$"; // matches D95%, D2cc
            string v_at_d_pattern = @"^V(?<evalpt>\d+\p{P}\d+|\d+)(?<unit>(%|cc|Gy))$"; // matches V50.4% or V50.4Gy 
            int lastLineNumber = File.ReadLines(_userFile).Count();
            bool _token = true;


            #region Modification du fichier template
            // Permet d'ajouter dans le fichier texte les contraintes spécifiques de la prescription 
            // Ces lignes seront ajoutés au template puis supprimées à la fin pour ne laisser que le template

            _streamWriter = new StreamWriter(_userFile, true);
            try
            {
                foreach (var index in _model.RTPrescription.Targets)
                {
                    foreach (var constraint in index.Constraints)
                    {
                        if (!constraint.Value1.Length.Equals(0))
                        {
                            _streamWriter.WriteLine(index.TargetId + ";" + constraint.Value1 + constraint.Unit1 + "," + constraint.Value2 + constraint.Unit2 + "," + "Prescription_num&#233rique");
                            addNLines++;
                        }
                    }
                }
                string pairstruct = string.Empty;
                foreach (var OAR in _model.RTPrescription.OrgansAtRisk)
                {
                    foreach (var constraint in OAR.Constraints)
                    {
                        if (!constraint.Value1.Length.Equals(0))
                        {
                            if (string.IsNullOrEmpty(constraint.Value2))
                            {
                                if (!constraint.ConstraintType.ToString().ToLower().Contains("mean"))
                                {
                                    //if (_pairOrgan.GetOrgan.ContainsKey(OAR.OrganAtRiskId.ToLower()))
                                    if (_pairOrgan.GetOrgan.Keys.Any(key => Regex.IsMatch(key, Regex.Escape(OAR.OrganAtRiskId), RegexOptions.IgnoreCase)))
                                    {
                                        pairstruct =
                                            (OAR.OrganAtRiskId.ToLower().Contains("@\"(?i)\\bpoumon\\s(homolateral|homolatéral)\\b\"")) ? "homolateral" :
                                            (OAR.OrganAtRiskId.ToLower().Contains("@\"(?i)\\bpoumon\\s(controlateral|controlatéral)\\b\"")) ? "controlateral" :
                                            OAR.OrganAtRiskId.ToLower();

                                        _streamWriter.WriteLine(_pairOrgan.GetOrgan[pairstruct].Item1 + "; Dose max ," + constraint.Value1 + constraint.Unit1 + "," + "Prescription_num&#233rique");
                                        addNLines++;
                                        _streamWriter.WriteLine(_pairOrgan.GetOrgan[pairstruct].Item2 + "; Dose max ," + constraint.Value1 + constraint.Unit1 + "," + "Prescription_num&#233rique");
                                        addNLines++;
                                    }
                                    else
                                    {
                                        _streamWriter.WriteLine(OAR.OrganAtRiskId + "; Dose max ," + constraint.Value1 + constraint.Unit1 + "," + "Prescription_num&#233rique");
                                        addNLines++;
                                    }
                                }
                                else
                                {
                                    pairstruct =
                                        (OAR.OrganAtRiskId.ToLower().Contains("@\"(?i)\\bpoumon\\s(homolateral|homolatéral)\\b\"")) ? "homolateral" :
                                        (OAR.OrganAtRiskId.ToLower().Contains("@\"(?i)\\bpoumon\\s(controlateral|controlatéral)\\b\"")) ? "controlateral" :
                                        OAR.OrganAtRiskId.ToLower();

                                    //if (_pairOrgan.GetOrgan.ContainsKey(OAR.OrganAtRiskId.ToLower()))
                                    if (_pairOrgan.GetOrgan.Keys.Any(key => Regex.IsMatch(key, Regex.Escape(OAR.OrganAtRiskId), RegexOptions.IgnoreCase)))
                                    {
                                        _streamWriter.WriteLine(_pairOrgan.GetOrgan[pairstruct].Item1 + "; Dose moyenne ," + constraint.Value1 + constraint.Unit1 + "," + "Prescription_num&#233rique");
                                        addNLines++;
                                        _streamWriter.WriteLine(_pairOrgan.GetOrgan[pairstruct].Item2 + "; Dose moyenne ," + constraint.Value1 + constraint.Unit1 + "," + "Prescription_num&#233rique");
                                        addNLines++;
                                    }
                                    else
                                    {
                                        _streamWriter.WriteLine(OAR.OrganAtRiskId + "; Dose moyenne ," + constraint.Value1 + constraint.Unit1 + "," + "Prescription_num&#233rique");
                                        addNLines++;
                                    }


                                }
                            }
                            else
                            {
                                _streamWriter.WriteLine(OAR.OrganAtRiskId + ";" + constraint.Value1 + constraint.Unit1 + "," + constraint.Value2 + constraint.Unit2 + "," + "Prescription_num&#233rique");
                                addNLines++;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur sur la prescription : aucune prescription n'est rattachée \n" + ex.Message);
            }
            #endregion

            #region lecture du fichier
            int lineNumber = 1;

            _streamWriter.Close();
            _streamReader = new StreamReader(_userFile);
            _streamReader.ReadLine(); // ignore la 1ere ligne

            while ((line = _streamReader.ReadLine()) != null)
            {
                line = line.Replace(" ", "");
                var testMatchD_at_V = Regex.Matches((line.Split(';')[1]).Split(',')[0], d_at_v_pattern);
                var testMatchV_at_D = Regex.Matches((line.Split(';')[1]).Split(',')[0], v_at_d_pattern);
                double similarity, maxSimilarity = 0;
                string referentiel = "x";

                try
                {
                    if (line.Split(';')[1].Split(',')[2].Equals(string.Empty))
                        referentiel = "x";
                    else
                        referentiel = line.Split(';')[1].Split(',')[2];
                }
                catch
                {
                    referentiel = "Prescription";
                }

                // Matching ID structure dans le fichier et dans le SS.
                // L'ID récupéré correspond au nom de la structure présente dans le SS (modification de l'ID de la structure du template)

                foreach (Structure structure in myPlan.StructureSet.Structures)
                {
                    //similarity = CalculateSimilarity(structure.Id, line.Split(';')[0]);
                    similarity = CalculateSimilarity(Regex.Replace(structure.Id.ToLower(), @"[\s\r\n]+", "").ToLower().Trim(), Regex.Replace(line.Split(';')[0].ToLower(), @"[\s\r\n]+", "").ToLower().Trim());
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

                int NbFraction; DoseValue DosePerFraction;

                // Dose de prescription qui sera dépendante de la prescription numérique
                try
                {
                    NbFraction = _model.RTPrescription.Targets.FirstOrDefault(x => x.TargetId == st.Id).NumberOfFractions;
                    DosePerFraction = _model.RTPrescription.Targets.FirstOrDefault(x => x.TargetId == st.Id).DosePerFraction;
                }
                catch
                {
                    NbFraction = (int)_model.PlanSetup.NumberOfFractions;
                    DosePerFraction = _model.PlanSetup.DosePerFraction;
                }

                DVHData myDVH = myPlan.GetDVHCumulativeData(st, DoseValuePresentation.Absolute, VolumePresentation.Relative, 0.1);
                // Permet d'ajouter une fois à chaque structure la dose max,moyenne médiane et min
                // Permet de vérifier les doses moyennes et max en fonction de la prescription
                try
                {
                    if (!_results.Keys.Any(x => x.ToLower().Contains(_s.Split('/')[0].ToLower())))
                    {
                        _results.Add(_s + " / Max Dose / no tol / CHU", Math.Round(myDVH.MaxDose.Dose, 2).ToString() + " Gy");
                        _results.Add(_s + " / Mean Dose / no tol / CHU", Math.Round(myDVH.MeanDose.Dose, 2).ToString() + " Gy");
                        _results.Add(_s + " / Median Dose / no tol / CHU", Math.Round(myDVH.MedianDose.Dose, 2).ToString() + " Gy");
                        _results.Add(_s + " / Min Dose / no tol / CHU", Math.Round(myDVH.MinDose.Dose, 2).ToString() + " Gy");
                        _results.Add(_s + " / Volume / no tol / CHU", Math.Round(st.Volume, 2).ToString() + " cc");
                        // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif / référentiel] et le résultats au format [ résultat unité]
                    }
                    if (line.Split(';')[1].Split(',')[0].Trim().ToLower().Contains("dosemax"))
                    {
                        _results.Add(_s + " / Max Dose / " + (line.Split(';')[1]).Split(',')[1] + " /" + referentiel, Math.Round(myDVH.MaxDose.Dose, 2).ToString() + " Gy");
                    }
                    if (line.Split(';')[1].Split(',')[0].Trim().ToLower().Contains("dosemoyenne"))
                    {
                        _results.Add(_s + " / Mean Dose / " + (line.Split(';')[1]).Split(',')[1] + " /" + referentiel, Math.Round(myDVH.MeanDose.Dose, 2).ToString() + " Gy");
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show($"Erreur sur le calcul des doses : La structure {st.Id} a été trouvé à la place de {line.Split(';')[0]} \n" + ex.Message);
                }

                // Calcul de la partie VxxGy, Vxx%, Dxxcc, Dxx%
                try
                {
                    if (testMatchD_at_V.Count != 0) // count is 1 if D95% or D2cc
                    {
                        Group eval = testMatchD_at_V[0].Groups["evalpt"];
                        Group unit = testMatchD_at_V[0].Groups["unit"];
                        DoseValue.DoseUnit du = DoseValue.DoseUnit.Gy;
                        DoseValue myD_something = new DoseValue(1000.1000, du);
                        double myD = Convert.ToDouble(eval.Value);

                        if (unit.Value == "%")
                        {
                            _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1] + " / " + referentiel, Math.Round(myPlan.GetDoseAtVolume(st, myD, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose, 3).ToString() + " Gy");
                            // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif / référentiel] et le résultats au format [ résultat unité]
                        }
                        else if (unit.Value == "cc")
                        {
                            _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1] + " / " + referentiel, Math.Round(myPlan.GetDoseAtVolume(st, myD, VolumePresentation.AbsoluteCm3, DoseValuePresentation.Absolute).Dose, 3).ToString() + " Gy");
                            // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif / référentiel] et le résultats au format [ résultat unité]
                        }
                        else
                            _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1] + " / " + referentiel, "-1.00");
                        // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif / référentiel] et le résultats au format [ résultat unité]
                    }

                    if (testMatchV_at_D.Count != 0) // count is 1
                    {
                        Group eval = testMatchV_at_D[0].Groups["evalpt"];
                        Group unit = testMatchV_at_D[0].Groups["unit"];
                        DoseValue.DoseUnit du = DoseValue.DoseUnit.Gy;
                        DoseValue myRequestedDose = new DoseValue(Convert.ToDouble(eval.Value), du);

                        //if (unit.Value == "cc")
                        if (line.Split(';')[1].Split(',')[1].Substring(line.Split(';')[1].Split(',')[1].Length - 2) == "cc")
                        {
                            _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1] + "/ " + referentiel, Math.Round(myPlan.GetVolumeAtDose(st, myRequestedDose, VolumePresentation.AbsoluteCm3), 2).ToString() + " cc");
                            // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif / référentiel] et le résultats au format [ résultat unité]
                        }
                        else if (unit.Value == "%")
                            _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1] + "/" + referentiel, Math.Round(myPlan.GetVolumeAtDose(st, (DoseValue)(myRequestedDose.Dose * NbFraction * DosePerFraction / 100), VolumePresentation.Relative), 2).ToString() + " %");
                        // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif / référentiel] et le résultats au format [ résultat unité]
                        else if (unit.Value == "Gy")
                            _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1] + "/" + referentiel, Math.Round(myPlan.GetVolumeAtDose(st, myRequestedDose, VolumePresentation.Relative), 2).ToString() + " %");
                        // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif / référentiel] et le résultats au format [ résultat unité]
                        else
                            _results.Add(_s + " / " + (line.Split(';')[1]).Split(',')[0] + "/" + (line.Split(';')[1]).Split(',')[1] + "/" + referentiel, "-1.00");
                        // Deux parties : la clé au format [ organe / indice / valeur recherchée / objectif / référentiel] et le résultats au format [ résultat unité]
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Problème dans la récupération des doses (exemple : Vxx% ou Dxx%\n {ex.Message}");
                }
                #endregion

                #region Indice  
                if (_token)
                {
                    foreach (var _target in myPlan.StructureSet.Structures)
                    {
                        // Calcul des indices de stéréotaxie
                        try
                        {
                            if (_target.Id.ToUpper().Contains("PTV") || _target.Id.ToUpper().Contains("CTV") || _target.Id.ToUpper().Contains("ITV") || _target.Id.ToUpper().Contains("GTV") && !st.Id.ToUpper().Contains("z_"))// && !_results.Keys.Any(x => x.Contains(st.Id)))
                            {
                                _index.Add(_s + "/ V107% " + _target.Id + "  / 107% / " + referentiel, Math.Round(myPlan.GetVolumeAtDose(_target, 1.07 * (DosePerFraction * NbFraction), VolumePresentation.AbsoluteCm3), 2).ToString() + " cc");
                                _index.Add(_s + "/ D95%  " + _target.Id + " / 95% / " + referentiel, Math.Round(myPlan.GetDoseAtVolume(_target, 0.95 * (DosePerFraction.Dose * NbFraction), VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose, 2).ToString() + " Gy");
                                _index.Add(_s + "/ V95%  " + _target.Id + " / 95% / " + referentiel, Math.Round(myPlan.GetVolumeAtDose(_target, 0.95 * (DosePerFraction * NbFraction), VolumePresentation.Relative), 2).ToString() + " %");
                                _index.Add(_s + "/ D50%  " + _target.Id + " / 50% / " + referentiel, Math.Round(myPlan.GetDoseAtVolume(_target, 0.50 * (DosePerFraction.Dose * NbFraction), VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose, 2).ToString() + " Gy");

                                double stx = _model.PlanSetup.Id.ToUpper().Contains("STX") ? 1 : 0.95;
                                #region HomogenityIndex
                                double d02 = myPlan.GetDoseAtVolume(_target, 2, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
                                double d98 = myPlan.GetDoseAtVolume(_target, 98, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
                                double d50 = myPlan.GetDoseAtVolume(_target, 50, VolumePresentation.Relative, DoseValuePresentation.Absolute).Dose;
                                _index.Add(_s + "/ HI " + _target.Id + " / index <2.5 / " + referentiel, (Math.Round((d02 - d98) / d50, 3)).ToString());
                                #endregion

                                #region ConformityIndex
                                double volIsodoseLvl = myPlan.GetVolumeAtDose(Body, (DosePerFraction * NbFraction) * stx, VolumePresentation.AbsoluteCm3);
                                _index.Add(_s + "/ CI " + _target.Id + " / index 0.7-1 / " + referentiel, Math.Round(volIsodoseLvl / _target.Volume, 3).ToString());
                                #endregion

                                #region PaddickConformityIndex
                                double PIV = myPlan.GetVolumeAtDose(Body, (DosePerFraction * NbFraction) * stx, VolumePresentation.AbsoluteCm3);
                                double TV_PIV = myPlan.GetVolumeAtDose(_target, (DosePerFraction * NbFraction) * stx, VolumePresentation.AbsoluteCm3);
                                _index.Add(_s + "/ PADDICK " + _target.Id + " / index 0.7-1 / " + referentiel, Math.Round((TV_PIV * TV_PIV) / (_target.Volume * PIV), 3).ToString());
                                #endregion

                                #region GradientIndex
                                double v50 = myPlan.GetVolumeAtDose(Body, 0.5 * (DosePerFraction * NbFraction), VolumePresentation.AbsoluteCm3);
                                double v100 = myPlan.GetVolumeAtDose(Body, (DosePerFraction * NbFraction), VolumePresentation.AbsoluteCm3);
                                _index.Add(_s + "/ GI " + _target.Id + " / index <3 / " + referentiel, Math.Round((v50 / v100), 2).ToString());
                                #endregion

                                #region RCI
                                double volTIP = myPlan.GetVolumeAtDose(_target, (DosePerFraction * NbFraction)* stx, VolumePresentation.AbsoluteCm3);
                                _index.Add(_s + "/ RCI " + _target.Id + " / index 0.9-2.5 / " + referentiel, Math.Round(volTIP / _target.Volume, 3).ToString());
                                #endregion
                                _token = false;
                            }
                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show("Erreur dans le calcul des indices stéréotaxiques\n" + ex.Message);
                        }

                    }

                }

                lineNumber++;
            }

            #endregion

            _results = _results.OrderBy(kvp => kvp.Key).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            _createExcelForStats.Fill(_results);
            _createExcelForStats.Close();
            _generatePDF.Execute(_results, _index);
            _streamReader.Close();
            RemoveLastNLines(_userFile, addNLines);
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
                MessageBox.Show("Problème dans la suppression des lignes du fichier txt\n" + ex.Message);

            }
        }
        #endregion

        #region Get and Set
        internal string UserFile
        {
            get { return _userFile; }
            set
            {
                //_userFile = System.IO.Path.Combine(System.IO.Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), "Template_dosi"), value + ".txt");
                _userFile = System.IO.Path.Combine(System.IO.Path.Combine(_model.Path, "Template_dosi"), value + ".txt");
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