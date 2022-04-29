using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CreatingModel_Task4
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreatingModel : IExternalCommand
    {
        List<Wall> walls = new List<Wall>();
        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            List<Level> listLevel = new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .ToList();

            Level level1 = listLevel
                .Where(x => x.Name.Equals("Уровень 1"))
                .FirstOrDefault();

            Level level2 = listLevel
                .Where(x => x.Name.Equals("Уровень 2"))
                .FirstOrDefault();
            List<XYZ> points = new List<XYZ>();
            Transaction transaction = new Transaction(doc, "Построение дома");
            transaction.Start();           

            CreateWalls(doc, points, level1, level2);

            AddDoor(doc, level1, walls[0]);

            for (int i = 1; i < 4; i++)
            {
                AddWindows(doc, level1, walls[i]);
            }
            
            //AddRoof(doc, level2, walls);
            AddRoof2(doc, level2);
            transaction.Commit();

            return Result.Succeeded;
        }

        private void AddRoof2(Document doc, Level level2)
        {
            RoofType roofType = new FilteredElementCollector(doc)
               .OfClass(typeof(RoofType))
               .OfType<RoofType>()
               .Where(x => x.Name.Equals("Типовой - 400мм"))
               .Where(x => x.FamilyName.Equals("Базовая крыша"))
               .FirstOrDefault();

            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);
            double dz = UnitUtils.ConvertToInternalUnits(4000, UnitTypeId.Millimeters);
            double dx = width / 2;
            double dy = depth / 2;
            double wallwidth = walls[0].Width;
            double dt = wallwidth / 2;           
           
            CurveArray curveArray = new CurveArray();
            curveArray.Append(Line.CreateBound(new XYZ(-dx, -dy - dt, dz), new XYZ(-dx, 0, 2 * dz)));
            curveArray.Append(Line.CreateBound(new XYZ(-dx, 0, 2 * dz), new XYZ(-dx, dy + dt, dz)));
           
            ReferencePlane plane = doc.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(0, 0, 20), new XYZ(0, 20, 0), doc.ActiveView);
            ExtrusionRoof extrusionRoof = doc.Create.NewExtrusionRoof(curveArray, plane, level2, roofType, -dx - dt, dx + dt);
            extrusionRoof.EaveCuts = EaveCutterType.TwoCutSquare;
        }

        //private void AddRoof(Document doc, Level level2, List<Wall> walls)
        //{
        //    RoofType roofType = new FilteredElementCollector(doc)
        //       .OfClass(typeof(RoofType))               
        //       .OfType<RoofType>()
        //       .Where(x => x.Name.Equals("Типовой - 400мм"))
        //       .Where(x => x.FamilyName.Equals("Базовая крыша"))
        //       .FirstOrDefault();

        //    double wallwidth = walls[0].Width;
        //    double dt = wallwidth / 2;
        //    List<XYZ> points = new List<XYZ>();
        //    points.Add(new XYZ(-dt, -dt, 0));
        //    points.Add(new XYZ(dt, -dt, 0));
        //    points.Add(new XYZ(dt, dt, 0));
        //    points.Add(new XYZ(-dt, dt, 0));
        //    points.Add(new XYZ(-dt, -dt, 0));

        //    Application application = doc.Application;
        //    CurveArray footprint = application.Create.NewCurveArray();
        //    for (int i = 0; i < 4; i++)
        //    {
        //        LocationCurve curve = walls[i].Location as LocationCurve;
        //        //footprint.Append(curve.Curve);
        //        XYZ p1 = curve.Curve.GetEndPoint(0);
        //        XYZ p2 = curve.Curve.GetEndPoint(1);
        //        Line line = Line.CreateBound(p1 + points[i], p2 + points[i + 1]);
        //        footprint.Append(line);
        //    }
        //    ModelCurveArray footPrintToModelCurveMapping = new ModelCurveArray();
        //    FootPrintRoof footprintRoof = doc.Create.NewFootPrintRoof(footprint, level2, roofType, out footPrintToModelCurveMapping);
        //    ModelCurveArrayIterator iterator = footPrintToModelCurveMapping.ForwardIterator();
        //    iterator.Reset();
        //    while (iterator.MoveNext())
        //    {
        //        ModelCurve modelCurve = iterator.Current as ModelCurve;
        //        footprintRoof.set_DefinesSlope(modelCurve, true);
        //        footprintRoof.set_SlopeAngle(modelCurve, 0.5);
        //    }
        //    foreach (ModelCurve m in footPrintToModelCurveMapping)
        //    {
        //        footprintRoof.set_DefinesSlope(m, true);
        //        footprintRoof.set_SlopeAngle(m, 0.5);
        //    }

        //}
                
        private void AddWindows(Document doc, Level level1, Wall wall)
        {
            FamilySymbol windowType1 = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 1830 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);            
            XYZ point = (point1 + point2) / 2;
                        
            if (!windowType1.IsActive)
                windowType1.Activate();

            FamilyInstance windowHeight = doc.Create.NewFamilyInstance(point, windowType1, wall, level1, StructuralType.NonStructural);
            Parameter offset = windowHeight.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);
            offset.Set(UnitUtils.ConvertToInternalUnits(1500, UnitTypeId.Millimeters));           
        }

        private void AddDoor(Document doc, Level level1, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 2134 мм"))
                .Where(x => x.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;
            
            if (!doorType.IsActive)
                doorType.Activate();

            doc.Create.NewFamilyInstance(point, doorType, wall, level1, StructuralType.NonStructural);            
        }

        private void CreateWalls(Document doc, List<XYZ> points, Level level1, Level level2)
        {
            double width = UnitUtils.ConvertToInternalUnits(10000, UnitTypeId.Millimeters);
            double depth = UnitUtils.ConvertToInternalUnits(5000, UnitTypeId.Millimeters);
            double dx = width / 2;
            double dy = depth / 2;
                        
            points.Add(new XYZ(-dx, -dy, 0));
            points.Add(new XYZ(dx, -dy, 0));
            points.Add(new XYZ(dx, dy, 0));
            points.Add(new XYZ(-dx, dy, 0));
            points.Add(new XYZ(-dx, -dy, 0));
                        
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(points[i], points[i + 1]);
                Wall wall = Wall.Create(doc, line, level1.Id, false);
                walls.Add(wall);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(level2.Id);
            }
        }
    }
}
