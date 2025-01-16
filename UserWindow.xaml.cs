using System;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using VMS.TPS.Common.Model.API;


namespace DoseCheck
{
    /// <summary>
    /// Logique d'interaction pour DoseCheck.xaml
    /// </summary>
    public partial class UserWindow : Window
    {
        private Model _model;

        public UserWindow(Patient patient, Course course, PlanSetup plansetup, RTPrescription RTPrescription, VMS.TPS.Common.Model.API.Image image)
        {
            InitializeComponent();
            _model = new Model(patient, course,plansetup,RTPrescription,image);
            DataContext = _model;

            try
            {
                Patient_Info.Text = $" Patient : {_model.Patient.Name} {_model.Patient.DateOfBirth}\n" +
                            $"Oncologue principal : {_model.Patient.PrimaryOncologistName} {_model.Patient.PrimaryOncologistId}\n" +
                            $"Intention du course : {_model.Course.Intent}\n" +
                            $"Id du course : {_model.Course.Id}\n" +
                            $"Statut du course : {_model.Course.ClinicalStatus}\n" +
                            $"Nom du plan : {_model.PlanSetup.Id}\n" +
                            $"Commentaire : {_model.Patient.Comment}";

                OK_Button.Visibility = Visibility.Collapsed;

                //foreach (var item in Directory.GetFiles("B:\\RADIOTHERAPIE\\Physique\\43 - Routine\\scripting\\Template_dosi"))
                //foreach (var item in Directory.GetFiles(Path.Combine(System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString(), _model.Path + "Template_dosi")))
                    foreach (var item in Directory.GetFiles(Path.Combine( _model.Path , "Template_dosi")))
                    {
                    _model.AddFile = System.IO.Path.GetFileNameWithoutExtension(item);
                }

                _model.File.Sort();

                foreach (var file in _model.File)
                {
                    Box_File.Items.Add(file);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                CloseLog();
            }
        }

        internal void CloseLog()
        {
            _model.CloseLog();
        }

        private void ComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            if (Box_File.SelectedItem != null)
            {
                OK_Button.Visibility = Visibility.Visible;
            }
            else
                OK_Button.Visibility = Visibility.Collapsed;

        }
        private void Button_Close(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
        private void Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                _model.UserFile = (string)Box_File.SelectedItem;
                this.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                this.Close();
            }
        }
    }
}
