using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;

using Newtonsoft.Json;

[assembly: CommandClass(typeof(PolygonData.Commands))]
namespace PolygonData
{
    public sealed class Commands
    {
        /// <summary>
        /// Private constructor to keep this class from being instantiated.
        /// This class is sealed as well, since we don't intend other
        /// classes to inherit from it
        /// </summary>
        private Commands() {}

        /// <summary>
        /// Call from inside AutoCAD with PolyData.
        /// Generates a text file with JSON objects representing the polygons.
        /// </summary>
        [CommandMethod("polydata", "polydata", "polydata",
         CommandFlags.Modal & CommandFlags.Session)]
        public static void PolyDataCommand()
        {
            // Get Autocad document
            var acDoc = Application.DocumentManager.CurrentDocument;
            // Any direct access to the objects must be wrapped in a transaction
            // A DWG file is a specialised database.
            using (var acTrans = acDoc.Database.TransactionManager.StartTransaction())
            {
                // Get model space object for read. For this we have to use the
                // transaction object (acTrans) and pass it an ObjectId.
                var model = (BlockTableRecord)acTrans.GetObject(
                    acDoc.Database.CurrentSpaceId, OpenMode.ForRead);

                // We will put our polygons here
                var polygons = new List<AcadPolygon>();
                // Now we iterate through all objects in the model space.
                // Notice that we get ObjectIds and not the actual objects.
                foreach(ObjectId objId in model)
                {
                    // Get the actual object
                    var obj = acTrans.GetObject(objId, OpenMode.ForRead);
                    // Ignore this object if it is not a Polyline
                    var polyline = obj as Polyline;
                    if (polyline == null)
                        continue;

                    // We will save the points in this AcadPolygon object
                    var poly = new AcadPolygon();
                    // Here we iterate through the number of vertices
                    for (var i = 0; i < polyline.NumberOfVertices; ++i)
                    {
                        // Get the i-th vertex of this polyline.
                        var pt = polyline.GetPoint2dAt(i);
                        // Create a point suitable to our data structure
                        var point = new AcadPolygon.Point() {
                            X = pt.X,
                            Y = pt.Y
                        };
                        // And add it to our AcadPolygon container.
                        poly.Points.Add(point);
                    }
                    // Add this polygon to the collection
                    polygons.Add(poly);
                }
                // Convert the list of polygons to json
                var json = JsonConvert.SerializeObject(
                    polygons,
                    Formatting.Indented);
                // Create a temporary text file
                var filename = System.IO.Path.GetTempFileName() + ".txt";
                // Write this json representation to that file
                System.IO.File.WriteAllText(filename, json);
                // Open it with the default text editor
                System.Diagnostics.Process.Start(filename);
            }
        }
    }
}
