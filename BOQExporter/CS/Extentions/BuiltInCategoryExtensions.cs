using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace BOQExporter.CS.Extentions
{
    public static class BuiltInCategoryExtensions
    {
        /// <summary>
        /// returns the name of the BuiltInCategory
        /// </summary>
        /// <param name="bic"></param>
        /// <returns></returns>
        public static string Name(this BuiltInCategory bic)
        {
            return LabelUtils.GetLabelFor(bic);
        }
    }

}
