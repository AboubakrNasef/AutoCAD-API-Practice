using AAAS = Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.PlottingServices;
[assembly: CommandClass(typeof(AutoCAD_API_Practice.Drawing_Commands))]
namespace AutoCAD_API_Practice
{
    public class Drawing_Commands : IExtensionApplication
    {
        #region Properties

       
        private Document _doc
        {
            get
            {
                return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            }
        }
        private Database _DB
        {
            get
            {
                return _doc.Database;
            }
        }
        private Editor _editor
        {
            get
            {
                return _doc.Editor;
            }
        }
        #endregion

        [CommandMethod("AB_DrawRect")]
        public void AB_DrawRect()
        {
       
           

            PromptPointResult promptPoint1 = _editor.GetPoint("Pick First Point");
            if (promptPoint1.Status != PromptStatus.OK) return;
          
            Point3d point1 = promptPoint1.Value;
            PromptPointResult promptPoint2 = _editor.GetPoint("Pick Second Point");
            if (promptPoint2.Status != PromptStatus.OK) return;
            Point3d point2 = promptPoint2.Value;

            using (Transaction tr = _DB.TransactionManager.StartTransaction())
            {
                Polyline rectangle = new Polyline(4);
                Point2d p1 = new Point2d(point1.X,point1.Y);
                Point2d p2 = new Point2d(point2.X,point1.Y);
                Point2d p3 = new Point2d(point2.X,point2.Y);
                Point2d p4 = new Point2d(point1.X,point2.Y);


                rectangle.AddVertexAt(0, p1, 0, 0, 0);
                rectangle.AddVertexAt(0, p2, 0, 0, 0);
                rectangle.AddVertexAt(0, p3, 0, 0, 0);
                rectangle.AddVertexAt(0, p4, 0, 0, 0);

                rectangle.Closed= true;

                // Add the rectangle to the current space
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(_DB.CurrentSpaceId, OpenMode.ForWrite);
                btr.AppendEntity(rectangle);
                tr.AddNewlyCreatedDBObject(rectangle, true);

                tr.Commit();

            }
        }

        #region IExtensionApplication

        public void Initialize()
        {
            throw new NotImplementedException();
        }

        public void Terminate()
        {
            throw new NotImplementedException();
        }


        #endregion
    }
}
