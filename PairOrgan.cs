using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DoseCheck
{
    internal class PairOrgan
    {
        private readonly Dictionary<string, Tuple<string, string>> _pairOrgan;
        internal PairOrgan()
        {
            _pairOrgan = new Dictionary<string, Tuple<string, string>>
            {
                { "plexus", new Tuple<string, string>("plexus G", "plexus D") },
                { "plexus brachial", new Tuple<string, string>("plexus G", "plexus D") },
                { "poumon", new Tuple<string, string>("poumon G", "poumon D")},
                { "poumon homolateral", new Tuple<string, string>("poumon G", "poumon D") },
                { "poumon controlateral", new Tuple<string, string>("poumon G", "poumon D") },
                { "tete femoral", new Tuple<string, string>("tete femoral G", "tete femoral G") },
                { "femur", new Tuple<string, string>("tete femoral G", "tete femoral G") },
                { "iliaque", new Tuple<string, string>("iliaque G", "iliaque D") },
                { "tete humerale", new Tuple<string, string>("tete humerale G", "tete humerale D") },
                { "humerus", new Tuple<string, string>("humerus G", "humerus D") },
                { "oeil", new Tuple<string, string>("oeil G", "oeil D") },
                { "retine", new Tuple<string, string>("retine G", "retine D") },
                { "cristallin", new Tuple<string, string>("cristallin G", "cristallin D") },                           
                { "cochlée", new Tuple<string, string>("cochlee G", "cochlee D") },
                { "hippocampe", new Tuple<string, string>("hippocampe G", "hippocampe D") },
                { "nerf optique", new Tuple<string, string>("Nerf optique D", "Nerf optique G") },
                { "sous max", new Tuple<string, string>("Glande SousMax D", "Glande SousMax G") },
                { "parotide", new Tuple<string, string>("parotide D", "parotide G") }
            };
        }

        internal Dictionary<string, Tuple<string, string>> GetOrgan
        {
            get { return _pairOrgan; }

        }
    }
}
