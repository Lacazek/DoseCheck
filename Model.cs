using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using VMS.TPS.Common.Model.API;
using DoseCheck;
using System.IO;
using System.Runtime.InteropServices;
using System.Runtime.CompilerServices;
using static System.Net.WebRequestMethods;
using System.Windows;
using System.ComponentModel;
using System.Reflection;

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


        internal Model(Patient patient, Course course, PlanSetup plansetup, RTPrescription RTPrescription, VMS.TPS.Common.Model.API.Image image)
        {
            _logFile = new StreamWriter("Log.txt", true);
            _logFile.WriteLine($"\n**********************************");
            _logFile.WriteLine($"Debut de programme : {DateTime.Now}");
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
    }

}
#endregion
