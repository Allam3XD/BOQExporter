using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace BOQExporter.Revit.Utils
{

    /// <summary>
    /// a wrapper class for BuiltInCategory to show in UI
    /// </summary>
    public class CategoryItem
    {
        public BuiltInCategory Bic { get; }
        public string Label { get; }
        // constructor for CategoryItem class
        public CategoryItem(BuiltInCategory bic)
        {
            Bic = bic;
            Label = LabelUtils.GetLabelFor(bic);
        }

        public override string ToString() => Label; // ensures label shows in UI
    }

}
