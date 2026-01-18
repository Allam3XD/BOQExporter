using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using BOQExporter.Revit.Modeles;
using BOQExporter.CS.Extentions;

namespace BOQExporter.Revit.Extraction
{
    internal static class BOQExtractor
    {

              
        
        /// <summary>
        /// extract BOQ items from the Revit document over 58 cat.
        /// </summary>
        /// <param name="doc"></param>
        public static void Extract(Document doc , HashSet<BuiltInCategory> allowedCategories)
        {
            BOQData.Initialize();

            // to check if the category is allowed
            bool IsAllowed(BuiltInCategory bic) => allowedCategories.Contains(bic);
            

            //------Architectural Elements Extraction------20

            #region Walls Extraction
            if (IsAllowed(BuiltInCategory.OST_Walls))
            {
                var walls = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .Cast<Wall>();

                foreach (var wall in walls)
                {
                    double area = wall.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)?.AsDouble() ?? 0;
                    double areaM2 = QuantityUtils.ToSquareMeters(area);

                    string code = MasterFormatMapper.GetCode(wall);
                    string desc = wall.Name;
                    string cate = wall.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(areaM2, 2), "m²"));
                }
            }
            #endregion

            #region Floors Extraction
            if (IsAllowed(BuiltInCategory.OST_Floors))
            {
                var floors = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Floors)
                    .WhereElementIsNotElementType()
                    .Cast<Floor>();

                foreach (var floor in floors)
                {
                    double area = floor.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)?.AsDouble() ?? 0;
                    double areaM2 = QuantityUtils.ToSquareMeters(area);

                    string code = MasterFormatMapper.GetCode(floor);
                    string desc = floor.Name;
                    string cate = floor.Category?.Name ?? "_";

                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(areaM2, 2), "m²"));
                }
            }
            #endregion

            #region Doors Extraction
            if (IsAllowed(BuiltInCategory.OST_Doors))
            {
                var doors = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Doors)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>();

                foreach (var door in doors)
                {
                    string code = MasterFormatMapper.GetCode(door);
                    string desc = door.Name;
                    string cate = door.Category?.Name ?? "_";

                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Windows Extraction
            if (IsAllowed(BuiltInCategory.OST_Windows))
            {
                var windows = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Windows)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>();

                foreach (var window in windows)
                {
                    string code = MasterFormatMapper.GetCode(window);
                    string desc = window.Name;
                    string cate = window.Category?.Name ?? "_";

                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Curtain Wall Panels Extraction
            if (IsAllowed(BuiltInCategory.OST_CurtainWallPanels))
            {
                var panels = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_CurtainWallPanels)
                    .WhereElementIsNotElementType();

                foreach (var panel in panels)
                {
                    string code = MasterFormatMapper.GetCode(panel);
                    string desc = panel.Name;
                    string cate = panel.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Ceilings Extraction
            if (IsAllowed(BuiltInCategory.OST_Ceilings))
            {
                var ceilings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Ceilings)
                    .WhereElementIsNotElementType();

                foreach (var ceiling in ceilings)
                {
                    double area = ceiling.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)?.AsDouble() ?? 0;
                    double areaM2 = QuantityUtils.ToSquareMeters(area);
                    string code = MasterFormatMapper.GetCode(ceiling);
                    string desc = ceiling.Name;
                    string cate = ceiling.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(areaM2, 2), "m²"));
                }
            }
            #endregion

            #region Roofs Extraction 
            if (IsAllowed(BuiltInCategory.OST_Roofs ))
            {
                var roofs = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Roofs)
                    .WhereElementIsNotElementType();

                foreach (var roof in roofs)
                {
                    double area = roof.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)?.AsDouble() ?? 0;
                    double areaM2 = QuantityUtils.ToSquareMeters(area);
                    string code = MasterFormatMapper.GetCode(roof);
                    string desc = roof.Name;
                    string cate = roof.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(areaM2, 2), "m²"));
                }
            }
            #endregion

            #region Stairs Extraction
            if (IsAllowed(BuiltInCategory.OST_Stairs ))
            {
                var stairs = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Stairs)
                    .WhereElementIsNotElementType();

                foreach (var stair in stairs)
                {
                    string code = MasterFormatMapper.GetCode(stair);
                    string desc = stair.Name;
                    string cate = stair.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Furniture Extraction
            if (IsAllowed(BuiltInCategory.OST_Furniture))
            {
                var furniture = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Furniture)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>();

                foreach (var item in furniture)
                {
                    string code = MasterFormatMapper.GetCode(item);
                    string desc = item.Name;
                    string cate = item.Category?.Name ?? "_";

                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Plumbing Fixtures Extraction
            if (IsAllowed(BuiltInCategory.OST_PlumbingFixtures))
            {
                var plumbing = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PlumbingFixtures)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>();

                foreach (var fixture in plumbing)
                {
                    string code = MasterFormatMapper.GetCode(fixture);
                    string desc = fixture.Name;
                    string cate = fixture.Category?.Name ?? "_";

                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Casework Extraction
            if (IsAllowed(BuiltInCategory.OST_Casework))
            {
                var casework = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Casework)
                    .WhereElementIsNotElementType();

                foreach (var cw in casework)
                {
                    string code = MasterFormatMapper.GetCode(cw);
                    string desc = cw.Name;
                    string cate = cw.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Rooms Extraction
            if (IsAllowed(BuiltInCategory.OST_Rooms))
            {
                var rooms = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Rooms)
                    .WhereElementIsNotElementType();

                foreach (var room in rooms)
                {
                    double area = room.get_Parameter(BuiltInParameter.ROOM_AREA)?.AsDouble() ?? 0;
                    double areaM2 = QuantityUtils.ToSquareMeters(area);
                    string code = MasterFormatMapper.GetCode(room);
                    string desc = room.Name;
                    string cate = room.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(areaM2, 2), "m²"));
                }
            }
            #endregion

            #region Detail Items Extraction
            if (IsAllowed(BuiltInCategory.OST_DetailComponents))
            {
                var details = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DetailComponents)
                    .WhereElementIsNotElementType();

                foreach (var di in details)
                {
                    string code = MasterFormatMapper.GetCode(di);
                    string desc = di.Name;
                    string cate = di.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Curtain Wall Mullions Extraction
            if (IsAllowed(BuiltInCategory.OST_CurtainWallMullions))
            {
                var mullions = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_CurtainWallMullions)
                    .WhereElementIsNotElementType();

                foreach (var mullion in mullions)
                {
                    double len = mullion.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;
                    double lenM = QuantityUtils.ToMeters(len);
                    string code = MasterFormatMapper.GetCode(mullion);
                    string desc = mullion.Name;
                    string cate = mullion.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(lenM, 2), "m"));
                }
            }
            #endregion

            #region Structural Walls Extraction 
            if (IsAllowed(BuiltInCategory.OST_Walls))
            {
                var structWalls = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Walls)
                    .WhereElementIsNotElementType()
                    .Cast<Wall>()
                    .Where(w => w.get_Parameter(BuiltInParameter.WALL_STRUCTURAL_SIGNIFICANT)?.AsInteger() == 1);

                foreach (var sw in structWalls)
                {
                    double area = sw.get_Parameter(BuiltInParameter.HOST_AREA_COMPUTED)?.AsDouble() ?? 0;
                    double areaM2 = QuantityUtils.ToSquareMeters(area);
                    string code = MasterFormatMapper.GetCode(sw);
                    string desc = sw.Name;
                    string cate = sw.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(areaM2, 2), "m²"));
                }
            }
            #endregion

            #region Topography Extraction
            if (IsAllowed(BuiltInCategory.OST_Topography))
            {
                var topo = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Topography)
                    .WhereElementIsNotElementType();

                foreach (var t in topo)
                {
                    string code = MasterFormatMapper.GetCode(t);
                    string desc = t.Name;
                    string cate = t.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Site Components Extraction 
            if (IsAllowed(BuiltInCategory.OST_Site))
            {
                var siteComps = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Site)
                    .WhereElementIsNotElementType();

                foreach (var sc in siteComps)
                {
                    string code = MasterFormatMapper.GetCode(sc);
                    string desc = sc.Name;
                    string cate = sc.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Parking Extraction
            if (IsAllowed(BuiltInCategory.OST_Parking))
            {
                var parking = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Parking)
                    .WhereElementIsNotElementType();

                foreach (var pk in parking)
                {
                    string code = MasterFormatMapper.GetCode(pk);
                    string desc = pk.Name;
                    string cate = pk.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion



            


            //------Structural Elements Extraction------9

            #region Columns Extraction
            if (IsAllowed(BuiltInCategory.OST_Columns))
            {
                var columns = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Columns)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>();

                foreach (var column in columns)
                {
                    string code = MasterFormatMapper.GetCode(column);
                    string desc = column.Name;
                    string cate = column.Category?.Name ?? "_";

                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Structural Framing Extraction
            if (IsAllowed(BuiltInCategory.OST_StructuralFraming))
            {
                var framing = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFraming)
                    .WhereElementIsNotElementType();

                foreach (var beam in framing)
                {
                    double len = beam.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;
                    double lenM = QuantityUtils.ToMeters(len);
                    string code = MasterFormatMapper.GetCode(beam);
                    string desc = beam.Name;
                    string cate = beam.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(lenM, 2), "m"));
                }
            }
            #endregion

            #region Structural Foundations Extraction
            if (IsAllowed(BuiltInCategory.OST_StructuralFoundation))
            {
                var foundations = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructuralFoundation)
                    .WhereElementIsNotElementType();

                foreach (var fnd in foundations)
                {
                    string code = MasterFormatMapper.GetCode(fnd);
                    string desc = fnd.Name;
                    string cate = fnd.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Structural Connections Extraction 
            if (IsAllowed(BuiltInCategory.OST_StructConnections))
            {
                var connections = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_StructConnections)
                    .WhereElementIsNotElementType();

                foreach (var conn in connections)
                {
                    string code = MasterFormatMapper.GetCode(conn);
                    string desc = conn.Name;
                    string cate = conn.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Rebar Extraction 
            if (IsAllowed(BuiltInCategory.OST_Rebar))
            {
                var rebar = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Rebar)
                    .WhereElementIsNotElementType();

                foreach (var rb in rebar)
                {
                    double len = rb.get_Parameter(BuiltInParameter.REBAR_ELEM_TOTAL_LENGTH)?.AsDouble() ?? 0;
                    double lenM = QuantityUtils.ToMeters(len);
                    string code = MasterFormatMapper.GetCode(rb);
                    string desc = rb.Name;
                    string cate = rb.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(lenM, 2), "m"));
                }
            }
            #endregion

            #region Fabric Areas Extraction 
            if (IsAllowed(BuiltInCategory.OST_FabricAreas))
            {
                var fabricAreas = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_FabricAreas)
                    .WhereElementIsNotElementType();

                foreach (var fa in fabricAreas)
                {
                    string code = MasterFormatMapper.GetCode(fa);
                    string desc = fa.Name;
                    string cate = fa.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Fabric Sheets Extraction 
            if (IsAllowed(BuiltInCategory.OST_FabricReinforcement))
            {
                var fabricSheets = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_FabricReinforcement)
                    .WhereElementIsNotElementType();

                foreach (var fs in fabricSheets)
                {
                    string code = MasterFormatMapper.GetCode(fs);
                    string desc = fs.Name;
                    string cate = fs.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion



           


            //------Electrical Elements Extraction------10

            #region Electrical Fixtures Extraction
            if (IsAllowed(BuiltInCategory.OST_ElectricalFixtures))
            {
                var electrical = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_ElectricalFixtures)
                    .WhereElementIsNotElementType()
                    .Cast<FamilyInstance>();

                foreach (var fixture in electrical)
                {
                    string code = MasterFormatMapper.GetCode(fixture);
                    string desc = fixture.Name;
                    string cate = fixture.Category?.Name ?? "_";

                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Electrical Equipment Extraction
            if (IsAllowed(BuiltInCategory.OST_ElectricalEquipment))
            {
                var elecEquip = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_ElectricalEquipment)
                    .WhereElementIsNotElementType();

                foreach (var eq in elecEquip)
                {
                    string code = MasterFormatMapper.GetCode(eq);
                    string desc = eq.Name;
                    string cate = eq.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Lighting Fixtures Extraction
            if (IsAllowed(BuiltInCategory.OST_LightingFixtures))
            {
                var lightFixtures = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_LightingFixtures)
                    .WhereElementIsNotElementType();

                foreach (var lf in lightFixtures)
                {
                    string code = MasterFormatMapper.GetCode(lf);
                    string desc = lf.Name;
                    string cate = lf.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Lighting Devices Extraction
            if (IsAllowed(BuiltInCategory.OST_LightingDevices))
            {
                var lightDevices = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_LightingDevices)
                    .WhereElementIsNotElementType();

                foreach (var ld in lightDevices)
                {
                    string code = MasterFormatMapper.GetCode(ld);
                    string desc = ld.Name;
                    string cate = ld.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Communication Devices Extraction
            if (IsAllowed(BuiltInCategory.OST_CommunicationDevices))
            {
                var commDevices = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_CommunicationDevices)
                    .WhereElementIsNotElementType();

                foreach (var cd in commDevices)
                {
                    string code = MasterFormatMapper.GetCode(cd);
                    string desc = cd.Name;
                    string cate = cd.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Fire Alarm Devices Extraction 
            if (IsAllowed(BuiltInCategory.OST_FireAlarmDevices))
            {
                var fireDevices = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_FireAlarmDevices)
                    .WhereElementIsNotElementType();

                foreach (var fd in fireDevices)
                {
                    string code = MasterFormatMapper.GetCode(fd);
                    string desc = fd.Name;
                    string cate = fd.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Security Devices Extraction 
            if (IsAllowed(BuiltInCategory.OST_SecurityDevices))
            {
                var secDevices = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_SecurityDevices)
                    .WhereElementIsNotElementType();

                foreach (var sd in secDevices)
                {
                    string code = MasterFormatMapper.GetCode(sd);
                    string desc = sd.Name;
                    string cate = sd.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Data Devices Extraction 
            if (IsAllowed(BuiltInCategory.OST_DataDevices))
            {
                var dataDevices = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DataDevices)
                    .WhereElementIsNotElementType();

                foreach (var dd in dataDevices)
                {
                    string code = MasterFormatMapper.GetCode(dd);
                    string desc = dd.Name;
                    string cate = dd.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Nurse Call Devices Extraction 
            if (IsAllowed(BuiltInCategory.OST_NurseCallDevices))
            {
                var nurseDevices = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_NurseCallDevices)
                    .WhereElementIsNotElementType();

                foreach (var nd in nurseDevices)
                {
                    string code = MasterFormatMapper.GetCode(nd);
                    string desc = nd.Name;
                    string cate = nd.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion



           



            //------Mechanical Elements Extraction------19

            #region Ducts Extraction
            if (IsAllowed(BuiltInCategory.OST_DuctCurves))
            {
                var ducts = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DuctCurves)
                    .WhereElementIsNotElementType();

                foreach (var d in ducts)
                {
                    double len = d.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;
                    double lenM = QuantityUtils.ToMeters(len);
                    string code = MasterFormatMapper.GetCode(d);
                    string desc = d.Name;
                    string cate = d.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(lenM, 2), "m"));
                }
            }
            #endregion

            #region Duct Fittings Extraction 
            if (IsAllowed(BuiltInCategory.OST_DuctFitting))
            {
                var ductFittings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_DuctFitting)
                    .WhereElementIsNotElementType();

                foreach (var df in ductFittings)
                {
                    string code = MasterFormatMapper.GetCode(df);
                    string desc = df.Name;
                    string cate = df.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Flex Ducts Extraction 
            if (IsAllowed(BuiltInCategory.OST_FlexDuctCurves))
            {
                var flexDucts = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_FlexDuctCurves)
                    .WhereElementIsNotElementType();

                foreach (var fd in flexDucts)
                {
                    double len = fd.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;
                    double lenM = QuantityUtils.ToMeters(len);
                    string code = MasterFormatMapper.GetCode(fd);
                    string desc = fd.Name;
                    string cate = fd.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(lenM, 2), "m"));
                }
            }
            #endregion

            #region Mechanical Equipment Extraction 
            if (IsAllowed(BuiltInCategory.OST_MechanicalEquipment))
            {
                var mechEquip = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_MechanicalEquipment)
                    .WhereElementIsNotElementType();

                foreach (var me in mechEquip)
                {
                    string code = MasterFormatMapper.GetCode(me);
                    string desc = me.Name;
                    string cate = me.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Pipes Extraction 
            if (IsAllowed(BuiltInCategory.OST_PipeCurves))
            {
                var pipes = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeCurves)
                    .WhereElementIsNotElementType();

                foreach (var p in pipes)
                {
                    double len = p.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;
                    double lenM = QuantityUtils.ToMeters(len);
                    string code = MasterFormatMapper.GetCode(p);
                    string desc = p.Name;
                    string cate = p.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(lenM, 2), "m"));
                }
            }
            #endregion

            #region Pipe Fittings Extraction 
            if (IsAllowed(BuiltInCategory.OST_PipeFitting))
            {
                var pipeFittings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeFitting)
                    .WhereElementIsNotElementType();

                foreach (var pf in pipeFittings)
                {
                    string code = MasterFormatMapper.GetCode(pf);
                    string desc = pf.Name;
                    string cate = pf.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Conduits Extraction 
            if (IsAllowed(BuiltInCategory.OST_Conduit))
            {
                var conduits = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Conduit)
                    .WhereElementIsNotElementType();

                foreach (var c in conduits)
                {
                    double len = c.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;
                    double lenM = QuantityUtils.ToMeters(len);
                    string code = MasterFormatMapper.GetCode(c);
                    string desc = c.Name;
                    string cate = c.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(lenM, 2), "m"));
                }
            }
            #endregion

            #region Conduit Fittings Extraction 
            if (IsAllowed(BuiltInCategory.OST_ConduitFitting))
            {
                var conduitFittings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_ConduitFitting)
                    .WhereElementIsNotElementType();

                foreach (var cf in conduitFittings)
                {
                    string code = MasterFormatMapper.GetCode(cf);
                    string desc = cf.Name;
                    string cate = cf.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Cable Trays Extraction
            if (IsAllowed(BuiltInCategory.OST_CableTray))
            {
                var cableTrays = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_CableTray)
                    .WhereElementIsNotElementType();

                foreach (var ct in cableTrays)
                {
                    double len = ct.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH)?.AsDouble() ?? 0;
                    double lenM = QuantityUtils.ToMeters(len);
                    string code = MasterFormatMapper.GetCode(ct);
                    string desc = ct.Name;
                    string cate = ct.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(lenM, 2), "m"));
                }
            }
            #endregion

            #region Cable Tray Fittings Extraction 
            if (IsAllowed(BuiltInCategory.OST_CableTrayFitting))
            {
                var cableTrayFittings = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_CableTrayFitting)
                    .WhereElementIsNotElementType();

                foreach (var ctf in cableTrayFittings)
                {
                    string code = MasterFormatMapper.GetCode(ctf);
                    string desc = ctf.Name;
                    string cate = ctf.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region HVAC Zones Extraction 
            if (IsAllowed(BuiltInCategory.OST_HVAC_Zones))
            {
                var hvacZones = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_HVAC_Zones)
                    .WhereElementIsNotElementType();

                foreach (var hz in hvacZones)
                {
                    string code = MasterFormatMapper.GetCode(hz);
                    string desc = hz.Name;
                    string cate = hz.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Spaces Extraction
            if (IsAllowed(BuiltInCategory.OST_MEPSpaces))
            {
                var spaces = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_MEPSpaces)
                    .WhereElementIsNotElementType();

                foreach (var sp in spaces)
                {
                    double area = sp.get_Parameter(BuiltInParameter.ROOM_AREA)?.AsDouble() ?? 0;
                    double areaM2 = QuantityUtils.ToSquareMeters(area);
                    string code = MasterFormatMapper.GetCode(sp);
                    string desc = sp.Name;
                    string cate = sp.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, Math.Round(areaM2, 2), "m²"));
                }
            }
            #endregion

            #region Sprinklers Extraction 
            if (IsAllowed(BuiltInCategory.OST_Sprinklers))
            {
                var sprinklers = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_Sprinklers)
                    .WhereElementIsNotElementType();

                foreach (var spr in sprinklers)
                {
                    string code = MasterFormatMapper.GetCode(spr);
                    string desc = spr.Name;
                    string cate = spr.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Pipe Accessories Extraction 
            if (IsAllowed(BuiltInCategory.OST_PipeAccessory))
            {
                var pipeAcc = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_PipeAccessory)
                    .WhereElementIsNotElementType();

                foreach (var pa in pipeAcc)
                {
                    string code = MasterFormatMapper.GetCode(pa);
                    string desc = pa.Name;
                    string cate = pa.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion

            #region Fabrication Parts Extraction
            if (IsAllowed(BuiltInCategory.OST_FabricationPipework))
            {
                var fabrication_Parts = new FilteredElementCollector(doc)
                    .OfCategory(BuiltInCategory.OST_FabricationPipework) // or Ductwork, Containment
                    .WhereElementIsNotElementType();

                foreach (var fp in fabrication_Parts)
                {
                    string code = MasterFormatMapper.GetCode(fp);
                    string desc = fp.Name;
                    string cate = fp.Category?.Name ?? "_";
                    BOQData.Items.Add(new BOQItem(cate, code, desc, 1, "ea"));
                }
            }
            #endregion


           
        }

    }
}
