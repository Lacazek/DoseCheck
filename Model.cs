using System;
using System.Collections.Generic;
using VMS.TPS.Common.Model.API;
using System.IO;

namespace DoseCheck
{

    internal class Model
    {
        private Patient _patient;
        private Course _course;
        private PlanSetup _planSetup;
        private RTPrescription _rTPrescription;
        private Image _image;
        private StreamWriter _logFile;
        private List<String> _list;
        private GetMyData _getMyData;
        private readonly string _path;


        internal Model(Patient patient, Course course, PlanSetup plansetup, RTPrescription RTPrescription, VMS.TPS.Common.Model.API.Image image)
        {
            // A ajuster: première ligne dans le dossier du script, deuxième ligne dans le dossier de travail du CHU
            //_path = System.IO.Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location).ToString();
            _path = @"B:\RADIOTHERAPIE\Killian\Dosi\Script\1.fini\DoseCheck";
            _logFile = new StreamWriter(System.IO.Path.Combine(_path,"Log.txt"), true);
            _logFile.WriteLine($"\n**********************************");
            _logFile.WriteLine($"Debut de programme : {DateTime.Now}");
            _logFile.WriteLine($"Ordinateur utilisé : {Environment.MachineName}");
            _logFile.WriteLine($"User : {Environment.UserName}\n");
            _logFile.WriteLine($"Fichier ouvert\n");
            _patient = patient;
            _course = course;
            _planSetup = plansetup;
            _rTPrescription = RTPrescription;
            _image = image;
            _list = new List<String>();
            _getMyData = new GetMyData(this);
        }

        internal void CloseLog()
        {
            _logFile.WriteLine($"Fin de programme : {DateTime.Now}");
            _logFile.WriteLine($"Fichier Log ferme");
            _logFile.WriteLine($"**********************************");
            _logFile.Close();
        }

        #region Get and Set
        internal Patient Patient
        {
            get { return _patient; }
        }
        internal Course Course
        {
            get { return _course; }
        }
        internal PlanSetup PlanSetup
        {
            get { return _planSetup; }
        }
        internal RTPrescription RTPrescription
        {
            get { return _rTPrescription; }
        }
        internal StructureSet StructureSet
        {
            get { return _planSetup.StructureSet; }
        }
        internal Image Image
        {
            get { return _image; }
        }
        internal List<string> File
        {
            get { return _list; }
        }
        internal string AddFile
        {
            set { _list.Add(value); }
        }
        internal string UserFile
        {
            get { return _getMyData.UserFile; }
            set { _getMyData.UserFile = value; }
        }
        internal string Path
        {
            get { return _path; }
        }
    }

}
#endregion
