using System;
using System.IO;
using Aspose.Cells_FOSS;

namespace Aspose.Cells_FOSS.Samples.Shapes
{
    internal static class Program
    {
        private static void Main()
        {
            var outputPath = Path.Combine(AppContext.BaseDirectory, "shapes-sample.xlsx");

            var workbook = new Workbook();
            var sheet = workbook.Worksheets[0];
            sheet.Name = "Shapes";

            sheet.Cells["A1"].PutValue("Shape");
            sheet.Cells["B1"].PutValue("Description");

            var rectIndex = sheet.Shapes.Add(0, 2, 2, 3, AutoShapeType.Rectangle);
            var rectShape = sheet.Shapes[rectIndex];
            rectShape.Name = "Rectangle Box";

            var roundedRectIndex = sheet.Shapes.Add(0, 4, 2, 5, AutoShapeType.RoundedRectangle);
            var roundedRectShape = sheet.Shapes[roundedRectIndex];
            roundedRectShape.Name = "Rounded Rectangle";

            var ellipseIndex = sheet.Shapes.Add(0, 6, 2, 7, AutoShapeType.Ellipse);
            var ellipseShape = sheet.Shapes[ellipseIndex];
            ellipseShape.Name = "Oval Shape";

            var triangleIndex = sheet.Shapes.Add(3, 2, 5, 3, AutoShapeType.Triangle);
            var triangleShape = sheet.Shapes[triangleIndex];
            triangleShape.Name = "Triangle";

            var rightTriangleIndex = sheet.Shapes.Add(3, 4, 5, 5, AutoShapeType.RightTriangle);
            var rightTriangleShape = sheet.Shapes[rightTriangleIndex];
            rightTriangleShape.Name = "Right Triangle";

            var diamondIndex = sheet.Shapes.Add(3, 6, 5, 7, AutoShapeType.Diamond);
            var diamondShape = sheet.Shapes[diamondIndex];
            diamondShape.Name = "Diamond";

            var pentagonIndex = sheet.Shapes.Add(6, 2, 8, 3, AutoShapeType.Pentagon);
            var pentagonShape = sheet.Shapes[pentagonIndex];
            pentagonShape.Name = "Pentagon";

            var hexagonIndex = sheet.Shapes.Add(6, 4, 8, 5, AutoShapeType.Hexagon);
            var hexagonShape = sheet.Shapes[hexagonIndex];
            hexagonShape.Name = "Hexagon";

            var octagonIndex = sheet.Shapes.Add(6, 6, 8, 7, AutoShapeType.Octagon);
            var octagonShape = sheet.Shapes[octagonIndex];
            octagonShape.Name = "Octagon";

            var star5Index = sheet.Shapes.Add(9, 2, 11, 3, AutoShapeType.Star5Point);
            var star5Shape = sheet.Shapes[star5Index];
            star5Shape.Name = "5-Point Star";

            var star8Index = sheet.Shapes.Add(9, 4, 11, 5, AutoShapeType.Star8Point);
            var star8Shape = sheet.Shapes[star8Index];
            star8Shape.Name = "8-Point Star";

            var star12Index = sheet.Shapes.Add(9, 6, 11, 7, AutoShapeType.Star12Point);
            var star12Shape = sheet.Shapes[star12Index];
            star12Shape.Name = "12-Point Star";

            var rightArrowIndex = sheet.Shapes.Add(12, 2, 14, 3, AutoShapeType.RightArrow);
            var rightArrowShape = sheet.Shapes[rightArrowIndex];
            rightArrowShape.Name = "Right Arrow";

            var leftArrowIndex = sheet.Shapes.Add(12, 4, 14, 5, AutoShapeType.LeftArrow);
            var leftArrowShape = sheet.Shapes[leftArrowIndex];
            leftArrowShape.Name = "Left Arrow";

            var upDownArrowIndex = sheet.Shapes.Add(12, 6, 14, 7, AutoShapeType.UpDownArrow);
            var upDownArrowShape = sheet.Shapes[upDownArrowIndex];
            upDownArrowShape.Name = "Up-Down Arrow";

            var heartIndex = sheet.Shapes.Add(15, 2, 17, 3, AutoShapeType.Heart);
            var heartShape = sheet.Shapes[heartIndex];
            heartShape.Name = "Heart";

            var lightningIndex = sheet.Shapes.Add(15, 4, 17, 5, AutoShapeType.Lightning);
            var lightningShape = sheet.Shapes[lightningIndex];
            lightningShape.Name = "Lightning Bolt";

            var cloudIndex = sheet.Shapes.Add(15, 6, 17, 7, AutoShapeType.Cloud);
            var cloudShape = sheet.Shapes[cloudIndex];
            cloudShape.Name = "Cloud";

            var sunIndex = sheet.Shapes.Add(18, 2, 20, 3, AutoShapeType.Sun);
            var sunShape = sheet.Shapes[sunIndex];
            sunShape.Name = "Sun";

            var moonIndex = sheet.Shapes.Add(18, 4, 20, 5, AutoShapeType.Moon);
            var moonShape = sheet.Shapes[moonIndex];
            moonShape.Name = "Moon";

            var plusIndex = sheet.Shapes.Add(18, 6, 20, 7, AutoShapeType.Plus);
            var plusShape = sheet.Shapes[plusIndex];
            plusShape.Name = "Plus Sign";

            var cubeIndex = sheet.Shapes.Add(21, 2, 23, 3, AutoShapeType.Cube);
            var cubeShape = sheet.Shapes[cubeIndex];
            cubeShape.Name = "Cube";

            var cylinderIndex = sheet.Shapes.Add(21, 4, 23, 5, AutoShapeType.Cylinder);
            var cylinderShape = sheet.Shapes[cylinderIndex];
            cylinderShape.Name = "Cylinder";

            var mathPlusIndex = sheet.Shapes.Add(21, 6, 23, 7, AutoShapeType.MathPlus);
            var mathPlusShape = sheet.Shapes[mathPlusIndex];
            mathPlusShape.Name = "Math Plus";

            workbook.Save(outputPath);

            var loaded = new Workbook(outputPath);
            var loadedSheet = loaded.Worksheets["Shapes"];

            Console.WriteLine("Saved: " + outputPath);
            Console.WriteLine("Shape count: " + loadedSheet.Shapes.Count);
            Console.WriteLine();
            Console.WriteLine("Basic Shapes:");
            Console.WriteLine("  " + loadedSheet.Shapes[0].Name + " (" + loadedSheet.Shapes[0].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[1].Name + " (" + loadedSheet.Shapes[1].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[2].Name + " (" + loadedSheet.Shapes[2].AutoShapeType + ")");
            Console.WriteLine();
            Console.WriteLine("Triangles & Polygons:");
            Console.WriteLine("  " + loadedSheet.Shapes[3].Name + " (" + loadedSheet.Shapes[3].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[4].Name + " (" + loadedSheet.Shapes[4].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[5].Name + " (" + loadedSheet.Shapes[5].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[6].Name + " (" + loadedSheet.Shapes[6].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[7].Name + " (" + loadedSheet.Shapes[7].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[8].Name + " (" + loadedSheet.Shapes[8].AutoShapeType + ")");
            Console.WriteLine();
            Console.WriteLine("Stars:");
            Console.WriteLine("  " + loadedSheet.Shapes[9].Name + " (" + loadedSheet.Shapes[9].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[10].Name + " (" + loadedSheet.Shapes[10].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[11].Name + " (" + loadedSheet.Shapes[11].AutoShapeType + ")");
            Console.WriteLine();
            Console.WriteLine("Arrows:");
            Console.WriteLine("  " + loadedSheet.Shapes[12].Name + " (" + loadedSheet.Shapes[12].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[13].Name + " (" + loadedSheet.Shapes[13].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[14].Name + " (" + loadedSheet.Shapes[14].AutoShapeType + ")");
            Console.WriteLine();
            Console.WriteLine("Symbols & Icons:");
            Console.WriteLine("  " + loadedSheet.Shapes[15].Name + " (" + loadedSheet.Shapes[15].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[16].Name + " (" + loadedSheet.Shapes[16].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[17].Name + " (" + loadedSheet.Shapes[17].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[18].Name + " (" + loadedSheet.Shapes[18].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[19].Name + " (" + loadedSheet.Shapes[19].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[20].Name + " (" + loadedSheet.Shapes[20].AutoShapeType + ")");
            Console.WriteLine();
            Console.WriteLine("3D & Math:");
            Console.WriteLine("  " + loadedSheet.Shapes[21].Name + " (" + loadedSheet.Shapes[21].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[22].Name + " (" + loadedSheet.Shapes[22].AutoShapeType + ")");
            Console.WriteLine("  " + loadedSheet.Shapes[23].Name + " (" + loadedSheet.Shapes[23].AutoShapeType + ")");
        }
    }
}
