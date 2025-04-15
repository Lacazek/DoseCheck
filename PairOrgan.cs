/******************************************************************************
 * Nom du fichier : PairOrgan.cs
 * Auteur         : LACAZE Killian
 * Date de création : [02/10/2024]
 * Description    : [Brève description du contenu ou de l'objectif du code]
 *
 * Droits d'auteur © [2024], [LACAZE.K].
 * Tous droits réservés.
 * 
 * Ce code a été développé exclusivement par LACAZE Killian. Toute utilisation de ce code 
 * est soumise aux conditions suivantes :
 * 
 * 1. L'utilisation de ce code est autorisée uniquement à titre personnel ou professionnel, 
 *    mais sans modification de son contenu.
 * 2. Toute redistribution, copie, ou publication de ce code sans l'accord explicite 
 *    de l'auteur est strictement interdite.
 * 3. L'auteur assume la responsabilité de l'utilisation de ce code dans ses propres projets.
 * 
 * CE CODE EST FOURNI "EN L'ÉTAT", SANS AUCUNE GARANTIE, EXPRESSE OU IMPLICITE. 
 * L'AUTEUR DÉCLINE TOUTE RESPONSABILITÉ POUR TOUT DOMMAGE OU PERTE RÉSULTANT 
 * DE L'UTILISATION DE CE CODE.
 *
 * Toute utilisation non autorisée ou attribution incorrecte de ce code est interdite.
 ******************************************************************************/


using System;
using System.Collections.Generic;

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
                { "plexus brachial", new Tuple<string, string>("plexus_G", "plexus_D") },
                { "poumon", new Tuple<string, string>("poumon_G", "poumon_D")},
                { "poumon homolateral", new Tuple<string, string>("poumon_G", "poumon_D") },
                { "poumon controlateral", new Tuple<string, string>("poumon_G", "poumon_D") },
                { "tete femoral", new Tuple<string, string>("tete femoral G", "tete femoral G") },
                { "femur", new Tuple<string, string>("tete femoral G", "tete femoral G") },
                { "iliaque", new Tuple<string, string>("iliaque G", "iliaque D") },
                { "tete humerale", new Tuple<string, string>("tete humerale G", "tete humerale D") },
                { "humerus", new Tuple<string, string>("humerus G", "humerus D") },
                { "oeil", new Tuple<string, string>("oeil_G", "oeil_D") },
                { "retine", new Tuple<string, string>("retine G", "retine D") },
                { "cristallin", new Tuple<string, string>("cristallin_G", "cristallin_D") },                           
                { "cochlée", new Tuple<string, string>("cochlee G", "cochlee D") },
                { "hippocampe", new Tuple<string, string>("hippocampe G", "hippocampe D") },
                { "nerf optique", new Tuple<string, string>("Nerf_optique_D", "Nerf_optique_G") },
                { "sous max", new Tuple<string, string>("Glande SousMax D", "Glande SousMax G") },
                { "parotide", new Tuple<string, string>("parotide_D", "parotide_G") }
            };
        }

        internal Dictionary<string, Tuple<string, string>> GetOrgan
        {
            get { return _pairOrgan; }

        }
    }
}
