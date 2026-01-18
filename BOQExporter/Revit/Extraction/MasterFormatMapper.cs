using Autodesk.Revit.DB;
using System.Collections.Generic;

namespace BOQExporter.Revit.Extraction
{
    /// <summary>
    /// this static class provides functionality to map Revit element categories and type names
    /// </summary>
    internal static class MasterFormatMapper
    {
        #region MasterFormat Supported Categories
        // Dictionary: BuiltInCategory → (keyword → MasterFormat code)
        private static readonly Dictionary<BuiltInCategory, Dictionary<string, string>> MappingRules =
            new Dictionary<BuiltInCategory, Dictionary<string, string>>
        {
            // ---------------- ARCHITECTURE ----------------
            { BuiltInCategory.OST_Walls, new Dictionary<string, string>
                {
                    { "structural", "03 30 00" }, // Cast-in-Place Concrete (structural walls)
                    { "concrete", "03 30 00" },
                    { "masonry", "04 20 00" },
                    { "brick", "04 21 00" },
                    { "curtain", "08 44 00" },
                    { "gypsum", "09 29 00" },
                    { "drywall", "09 29 00" },
                    { "metal stud", "09 22 00" },
                    { "cmu", "04 22 00" },
                    { "default", "09 00 00" }
                }
            },

            { BuiltInCategory.OST_Floors, new Dictionary<string, string>
                {
                    { "concrete", "03 30 00" },
                    { "precast", "03 40 00" },
                    { "wood", "06 10 00" },
                    { "tile", "09 30 00" },
                    { "carpet", "09 68 00" },
                    { "raised", "09 69 00" },
                    { "default", "09 60 00" }
                }
            },

            { BuiltInCategory.OST_Doors, new Dictionary<string, string>
                {
                    { "wood", "08 14 00" },
                    { "metal", "08 11 00" },
                    { "glass", "08 12 00" },
                    { "fire", "08 11 13" },
                    { "default", "08 10 00" }
                }
            },

            { BuiltInCategory.OST_Windows, new Dictionary<string, string>
                {
                    { "aluminum", "08 51 00" },
                    { "wood", "08 52 00" },
                    { "steel", "08 53 00" },
                    { "default", "08 50 00" }
                }
            },

            { BuiltInCategory.OST_CurtainWallPanels, new Dictionary<string, string>
                {
                    { "glass", "08 44 13" },       // Glazed curtain panels
                    { "glazed", "08 44 13" },      // Explicit match for "Glazed"
                    { "metal", "08 44 23" },       // Metal curtain panels
                    { "spandrel", "08 44 33" },    // Spandrel panels
                    { "solid", "08 44 33" },       // Solid panels treated like spandrel
                    { "default", "08 44 00" }      // Generic curtain wall panel
                }
            },


            { BuiltInCategory.OST_CurtainWallMullions, new Dictionary<string, string>
                {
                    { "aluminum", "08 44 23" },
                    { "steel", "08 44 23" },
                    { "default", "08 44 00" }
                }
            },

            { BuiltInCategory.OST_Ceilings, new Dictionary<string, string>
                {
                    { "acoustic", "09 51 00" },
                    { "gypsum", "09 29 00" },
                    { "suspended", "09 51 13" },
                    { "default", "09 50 00" }
                }
            },

            { BuiltInCategory.OST_Roofs, new Dictionary<string, string>
                {
                    { "metal", "07 41 00" },
                    { "tile", "07 32 00" },
                    { "membrane", "07 54 00" },
                    { "shingle", "07 31 13" },
                    { "default", "07 30 00" }
                }
            },

            { BuiltInCategory.OST_Stairs, new Dictionary<string, string>
                {
                    { "steel", "05 51 00" },
                    { "concrete", "03 30 00" },
                    { "wood", "06 11 00" },
                    { "default", "05 50 00" }
                }
            },

           

            { BuiltInCategory.OST_Furniture, new Dictionary<string, string>
                {
                    { "chair", "12 52 00" },
                    { "table", "12 53 00" },
                    { "desk", "12 55 00" },
                    { "casework", "12 30 00" },
                    { "default", "12 50 00" }
                }
            },

            { BuiltInCategory.OST_Casework, new Dictionary<string, string>
                {
                    { "wood", "12 32 00" },
                    { "metal", "12 33 00" },
                    { "plastic", "12 35 00" },
                    { "default", "12 30 00" }
                }
            },

            

            { BuiltInCategory.OST_DetailComponents, new Dictionary<string, string>
                {
                    { "rebar", "03 20 00" },
                    { "flashing", "07 60 00" },
                    { "insulation", "07 21 00" },
                    { "default", "01 00 00" }
                }
            },

            { BuiltInCategory.OST_Site, new Dictionary<string, string>
                {
                    { "planting", "32 90 00" },
                    { "furniture", "32 30 00" },
                    { "paving", "32 13 00" },
                    { "curb", "32 16 00" },
                    { "default", "32 00 00" }
                }
            },

            { BuiltInCategory.OST_Topography, new Dictionary<string, string>
                {
                    { "cut", "31 23 00" },
                    { "fill", "31 23 00" },
                    { "grading", "31 22 00" },
                    { "default", "31 20 00" }
                }
            },

            { BuiltInCategory.OST_Parking, new Dictionary<string, string>
                {
                    { "striping", "32 17 23" },
                    { "curb", "32 16 00" },
                    { "default", "32 90 00" }
                }
            },

            // ---------------- STRUCTURE ----------------
            { BuiltInCategory.OST_Columns, new Dictionary<string, string>
                {
                    { "steel", "05 12 00" },
                    { "concrete", "03 40 00" },
                    { "wood", "06 13 00" },
                    { "default", "05 10 00" }
                }
            },

            { BuiltInCategory.OST_StructuralFraming, new Dictionary<string, string>
                {
                    { "steel", "05 12 00" },
                    { "concrete", "03 40 00" },
                    { "wood", "06 13 00" },
                    { "default", "05 10 00" }
                }
            },

            { BuiltInCategory.OST_StructuralFoundation, new Dictionary<string, string>
                {
                    { "pile", "31 62 00" },
                    { "mat", "03 30 00" },
                    { "footing", "03 30 00" },
                    { "default", "03 10 00" }
                }
            },

           

            { BuiltInCategory.OST_Rebar, new Dictionary<string, string>
                {
                    { "mesh", "03 21 00" },
                    { "bar", "03 21 00" },
                    { "default", "03 20 00" }
                }
            },

            { BuiltInCategory.OST_FabricAreas, new Dictionary<string, string>
                {
                    { "mesh", "03 21 00" },
                    { "default", "03 21 00" }
                }
            },

            { BuiltInCategory.OST_FabricReinforcement, new Dictionary<string, string>
                {
                    { "mesh", "03 21 00" },
                    { "sheet", "03 21 00" },
                    { "default", "03 21 00" }
                }
            },

           

            // ---------------- ELECTRICAL ----------------
            { BuiltInCategory.OST_ElectricalFixtures, new Dictionary<string, string>
                {
                    { "light", "26 50 00" },
                    { "switch", "26 27 00" },
                    { "receptacle", "26 27 26" },
                    { "panel", "26 24 00" },
                    { "default", "26 00 00" }
                }
            },

            { BuiltInCategory.OST_ElectricalEquipment, new Dictionary<string, string>
                {
                    { "transformer", "26 12 00" },
                    { "panel", "26 24 00" },
                    { "ups", "26 33 53" },
                    { "default", "26 20 00" }
                }
            },

            { BuiltInCategory.OST_LightingFixtures, new Dictionary<string, string>
                {
                    { "recessed", "26 51 13" },
                    { "surface", "26 51 16" },
                    { "emergency", "26 52 00" },
                    { "default", "26 51 00" }
                }
            },

            { BuiltInCategory.OST_LightingDevices, new Dictionary<string, string>
                {
                    { "switch", "26 27 00" },
                    { "sensor", "26 27 26" },
                    { "dimmer", "26 27 26" },
                    { "default", "26 27 00" }
                }
            },

            

            { BuiltInCategory.OST_FireAlarmDevices, new Dictionary<string, string>
                {
                    { "smoke", "28 31 00" },
                    { "heat", "28 31 00" },
                    { "pull", "28 31 00" },
                    { "default", "28 31 00" }
                }
            },

            { BuiltInCategory.OST_SecurityDevices, new Dictionary<string, string>
                {
                    { "camera", "28 20 00" },
                    { "card", "28 13 00" },
                    { "motion", "28 20 00" },
                    { "default", "28 10 00" }
                }
            },

            { BuiltInCategory.OST_DataDevices, new Dictionary<string, string>
                {
                    { "outlet", "27 15 00" },
                    { "jack", "27 15 00" },
                    { "default", "27 15 00" }
                }
            },

           

            { BuiltInCategory.OST_ElectricalCircuit, new Dictionary<string, string>
                {
                    { "branch", "26 05 00" },
                    { "feeder", "26 05 00" },
                    { "default", "26 05 00" }
                }
            },

           

            // ---------------- MECHANICAL / MEP ----------------
            { BuiltInCategory.OST_DuctCurves, new Dictionary<string, string>
                {
                    { "rectangular", "23 31 13" },
                    { "round", "23 31 14" },
                    { "spiral", "23 31 14" },
                    { "default", "23 31 00" }
                }
            },

            { BuiltInCategory.OST_DuctFitting, new Dictionary<string, string>
                {
                    { "elbow", "23 33 00" },
                    { "tee", "23 33 00" },
                    { "transition", "23 33 00" },
                    { "default", "23 33 00" }
                }
            },

            { BuiltInCategory.OST_FlexDuctCurves, new Dictionary<string, string>
                {
                    { "insulated", "23 31 13" },
                    { "uninsulated", "23 31 13" },
                    { "default", "23 31 13" }
                }
            },

            

            { BuiltInCategory.OST_MechanicalEquipment, new Dictionary<string, string>
                {
                    { "ahu", "23 73 00" },
                    { "fan", "23 34 00" },
                    { "chiller", "23 64 00" },
                    { "boiler", "23 52 00" },
                    { "pump", "23 21 23" },
                    { "default", "23 05 00" }
                }
            },

            { BuiltInCategory.OST_PipeCurves, new Dictionary<string, string>
                {
                    { "copper", "22 11 16" },
                    { "steel", "22 11 19" },
                    { "plastic", "22 11 23" },
                    { "default", "22 10 00" }
                }
            },

            { BuiltInCategory.OST_PipeFitting, new Dictionary<string, string>
                {
                    { "elbow", "22 11 00" },
                    { "tee", "22 11 00" },
                    { "coupling", "22 11 00" },
                    { "default", "22 11 00" }
                }
            },

            { BuiltInCategory.OST_PipeAccessory, new Dictionary<string, string>
                {
                    { "valve", "22 12 00" },
                    { "meter", "22 12 00" },
                    { "strainer", "22 12 00" },
                    { "default", "22 12 00" }
                }
            },

            { BuiltInCategory.OST_Sprinklers, new Dictionary<string, string>
                {
                    { "upright", "21 13 13" },
                    { "pendant", "21 13 13" },
                    { "sidewall", "21 13 13" },
                    { "default", "21 13 00" }
                }
            },

            { BuiltInCategory.OST_Conduit, new Dictionary<string, string>
                {
                    { "emt", "26 05 33" },
                    { "rigid", "26 05 33" },
                    { "flex", "26 05 33" },
                    { "default", "26 05 33" }
                }
            },

            { BuiltInCategory.OST_ConduitFitting, new Dictionary<string, string>
                {
                    { "elbow", "26 05 33" },
                    { "coupling", "26 05 33" },
                    { "default", "26 05 33" }
                }
            },

            { BuiltInCategory.OST_CableTray, new Dictionary<string, string>
                {
                    { "ladder", "26 36 00" },
                    { "solid", "26 36 00" },
                    { "default", "26 36 00" }
                }
            },

            { BuiltInCategory.OST_CableTrayFitting, new Dictionary<string, string>
                {
                    { "elbow", "26 36 00" },
                    { "tee", "26 36 00" },
                    { "default", "26 36 00" }
                }
            },

           
        };
        #endregion


        /// <summary>
        /// this method gets the MasterFormat code for a given Revit element based on its category and type name.
        /// </summary>
        /// <param name="element"></param>
        /// <returns></returns>
        public static string GetCode(Element element)
        {
            var cat = (BuiltInCategory)element.Category.Id.IntegerValue; // for Revit 2022-2025 API
            //var cat = (BuiltInCategory)element.Category.Id.Value; // for Revit 2026 API


            string typeName = element.Name?.ToLower() ?? "";

            if (MappingRules.ContainsKey(cat))
            {
                var rules = MappingRules[cat];

                foreach (var kvp in rules)
                {
                    if (kvp.Key != "default" && typeName.Contains(kvp.Key))
                        return kvp.Value;
                }

                if (rules.ContainsKey("default"))
                    return rules["default"];
            }

            return "00 00 00"; // uncategorized fallback
        }
        /// <summary>
        /// Gets all supported BuiltInCategories in the MasterFormat mapping.
        /// </summary>
        /// <returns></returns>
        public static IEnumerable<BuiltInCategory> SupportedCategories()
        {
            return MappingRules.Keys; // all categories you mapped
        }

    }
}
