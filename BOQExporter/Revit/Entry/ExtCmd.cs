using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using BOQExporter.Excel.Utits;
using BOQExporter.Revit.Modeles;
using BOQExporter.UI;




namespace BOQExporter.Revit.Entry
{
    [Transaction(TransactionMode.Manual)]
    public class ExtCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                //TaskDialog.Show("BOQ Exporter", "ExtCmd.Execute triggered.");

                // Get Revit context
                UIApplication uiApp = commandData.Application;
                UIDocument uiDoc = uiApp.ActiveUIDocument;
                Document doc = uiDoc.Document;

                
                

                // ✅ Set BOQData.App and BOQData.Doc
                BOQData.App = uiApp.Application;
                BOQData.Doc = uiDoc?.Document;
                

                // Safety check
                if (BOQData.Doc == null)
                { 
                    message = "Active document is null. Please open a project before running the add-in.";
                    return Result.Failed;
                }


                // Show UI
                using (MainFrom form = new MainFrom()) 
                {
                    form.ShowDialog(); 
                }
                return Result.Succeeded;
            } 
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }

}

