using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace BOQExporter.Revit.Utils
{
    internal static class RvtUtils
    {
        public static class CategoryUtils
        {


            internal static class SelectionMemory
            {

                //category selection fields
                public static HashSet<BuiltInCategory> SavedCategories { get; set; } = new HashSet<BuiltInCategory>();
                public static bool ArchChecked { get; set; }
                public static bool StruChecked { get; set; }
                public static bool MechChecked { get; set; }
                public static bool ElecChecked { get; set; }


                //meta date fields
                public static string ProjectName { get; set; } = "";
                public static string ClientName { get; set; } = "";
                public static string Location { get; set; } = ""; 
                public static string IssueDate { get; set; } = "";

                //logo path field
                public static string LogoPath { get; set; } = "";

            }





        }
    }
}
