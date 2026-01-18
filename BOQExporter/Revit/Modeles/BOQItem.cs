using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace BOQExporter.Revit.Modeles
{
    public class BOQItem
    {
        /// <summary>
        /// Number of the BOQ item in the list.
        /// </summary>
        public int ItemNo { get; set; }

        /// <summary>
        /// Allocated category of the BOQ item.
        /// </summary>
        public string Category { get; set; }

        /// <summary>
        /// MasterFormat item code (e.g., "03 30 00").
        /// </summary>
        public string ItemCode { get; set; }

        /// <summary>
        /// Description of the BOQ item.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Quantity of the BOQ item.
        /// </summary>
        public double Quantity { get; set; }

        /// <summary>
        /// Unit of measurement for the BOQ item.
        /// </summary>
        public string Unit { get; set; }

        /// <summary>
        /// The unit rate of the BOQ item.
        /// </summary>
        public double UnitRate { get; set; }

        /// <summary>
        /// The total cost of the BOQ item (Quantity × UnitRate).
        /// </summary>
        public double TotalCost => Quantity * UnitRate;

        /// <summary>
        /// Remarks or notes for the BOQ item.
        /// </summary>
        public string Remarks { get; set; }

        /// <summary>
        /// The number of times this item is used in the project.
        /// </summary>
        public int UsageCount { get; set; } = 1;
        /// <summary>
        /// the Revit element associated with this BOQ item
        /// </summary>
        public Element Element { get; set; }

        /// <summary>
        /// constructor of BOQItem for the traceability table
        /// </summary>
        /// <param name="itemCode"></param>
        /// <param name="description"></param>
        /// <param name="quantity"></param>
        /// <param name="unit"></param>
        public BOQItem(string category, string itemCode, string description,
                        double quantity, string unit)
        {
            Category = category;
            ItemCode = itemCode;
            Description = description;
            Quantity = quantity;
            Unit = unit;
        }

        /// <summary>
        /// an overloaded constructor of BOQItem for the main BOQ table
        /// </summary>
        /// <param name="itemNo"></param>
        /// <param name="category"></param>
        /// <param name="itemCode"></param>
        /// <param name="description"></param>
        /// <param name="quantity"></param>
        /// <param name="unit"></param>
        /// <param name="unitRate"></param>
        /// <param name="remarks"></param>
        public BOQItem(int itemNo, string category, string itemCode,
                        string description, double quantity,
                        string unit, double unitRate,
                        string remarks,
                        Element element = null)
        {
            ItemNo = itemNo;
            Category = category;
            ItemCode = itemCode;
            Description = description;
            Quantity = quantity;
            Unit = unit;
            UnitRate = unitRate;
            Remarks = remarks;
            Element = element;
        }
        /// <summary>
        /// converts the BOQItem to an object array for excel export for the main tabel
        /// </summary>
        /// <returns></returns>
        public object[] ToArray()
        {
            return new object[]
            {
                ItemNo,
                ItemCode,
                Description,
                Quantity,
                Unit,
                UnitRate,
                TotalCost,
                Remarks
            };
        }
        /// <summary>
        /// Converts the BOQItem to an object array for trace tabel
        /// </summary>
        /// <returns></returns>
        public object[] ToTraceArray()
        {
            return new object[]
            {
                ItemCode,
                Category,   
                Description,
                Quantity,
                Unit,
                UsageCount
            };
        }


    }
}
