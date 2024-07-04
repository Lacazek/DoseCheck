using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using VMS.TPS.Common.Model.API;

namespace DoseCheck
{

    // A améliorer : 
    // assurer que les valeurs recherchées apparaissent dans les colonnes en haut
    // et que les résultats sont bien associées aux résultats.
    // diminuer le nombre de chiffres significatifs

    internal class CreateExcelForStats
    {
        private StreamWriter _excelForStats;
        private Patient _patient;
        private Course _course;
        private PlanSetup _plan;

        internal CreateExcelForStats(Model _model)
        {
            try
            {
                _patient = _model.Patient;
                _course = _model.Course;
                _plan = _model.PlanSetup;
                if (!Directory.Exists("./out"))
                    Directory.CreateDirectory("./out");
                _excelForStats = new StreamWriter("out/Data_" + _patient.LastName + "_" + _patient.FirstName + ".csv");
                _excelForStats.WriteLine("patientID;courseID;planID;TotalDose;Dose/#;Fractions;Structure;Dose max; Dose moyenne; Dose mediane; Dose min");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Erreur dans la construction du fichier excel\n" + ex.Message);
            }
        }

        internal void Fill(Dictionary<string, string> _results)
        {

            string currentOrgan = "";
            try
            {
                foreach (var value in _results)
                {
                    string organ = value.Key.Split('/')[0];

                    if (!organ.Equals(currentOrgan))
                    {
                        if (!string.IsNullOrEmpty(currentOrgan))
                        {
                            _excelForStats.WriteLine();
                        }
                        _excelForStats.Write("{0};{1};{2};{3};{4};{5};{6}",
                            _patient.Id, _course.Id, _plan.Id, _plan.TotalDose, _plan.DosePerFraction, _plan.NumberOfFractions, organ);
                        _excelForStats.Write(";");
                        currentOrgan = organ;
                    }
                    else
                    {
                        _excelForStats.Write(";");
                    }
                    _excelForStats.Write("{0:0.00}", value.Value);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }


            _excelForStats.WriteLine();
        }

        internal void Close()
        {
            _excelForStats.Close();
        }
    }
}

