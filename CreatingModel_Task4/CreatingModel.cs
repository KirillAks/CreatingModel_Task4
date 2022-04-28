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

            AddWindow1(doc, level1, walls[1]);
            AddWindow2(doc, level1, walls[2]);
            AddWindow3(doc, level1, walls[3]);

            transaction.Commit();

            return Result.Succeeded;
        }

        private void AddWindow3(Document doc, Level level1, Wall wall)
        {
            FamilySymbol windowType3 = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0915 x 0610 мм"))
                .Where(x => x.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();

            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ point1 = hostCurve.Curve.GetEndPoint(0);
            XYZ point2 = hostCurve.Curve.GetEndPoint(1);
            XYZ point = (point1 + point2) / 2;

            if (!windowType3.IsActive)
                windowType3.Activate();

            FamilyInstance windowHeight = doc.Create.NewFamilyInstance(point, windowType3, wall, level1, StructuralType.NonStructural);
            Parameter offset = windowHeight.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);
            offset.Set(UnitUtils.ConvertToInternalUnits(2000, UnitTypeId.Millimeters));
        }

        private void AddWindow2(Document doc, Level level1, Wall wall)
        {
            FamilySymbol windowType2 = new FilteredElementCollector(doc)
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

            if (!windowType2.IsActive)
                windowType2.Activate();

            FamilyInstance windowHeight = doc.Create.NewFamilyInstance(point, windowType2, wall, level1, StructuralType.NonStructural);
            Parameter offset = windowHeight.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM);
            offset.Set(UnitUtils.ConvertToInternalUnits(1000, UnitTypeId.Millimeters));
        }

        private void AddWindow1(Document doc, Level level1, Wall wall)
        {
            FamilySymbol windowType1 = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(x => x.Name.Equals("0406 x 1220 мм"))
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
