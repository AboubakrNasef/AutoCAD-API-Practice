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

[assembly: CommandClass(typeof(AutoCAD_API_Practice.AutoCAD_Commands))]
namespace AutoCAD_API_Practice
{
    public class AutoCAD_Commands : IExtensionApplication
    {


        // Get the current document and database
        private Document _doc
        {
            get
            {
                return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            }
        }
        private Database _curDb
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
        [CommandMethod("AB_Hello")]
        public void AB_Hello()
        {
            AAAS.Document doc = AAAS.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.WriteMessage("Hello First Command");
        }
        [CommandMethod("AB_Zoom")]

        public void AB_Zoom()
        {
            Point3d p1 = default(Point3d);
            Point3d p2 = default(Point3d);

            PromptPointResult pr = _editor.GetPoint(new PromptPointOptions("Pick First Point"));

            if (pr.Status != PromptStatus.OK)

            {
                p1 = pr.Value;



                return;

            }
            pr = _editor.GetPoint(new PromptPointOptions("Pick Second Point"));

            if (pr.Status != PromptStatus.OK)

            {
                p2 = pr.Value;



                return;

            }
            _zoom(p1, p2, new Point3d(), 1);
        }


        public List<string> ListLayouts()
        {
            List<string> layoutNames = new List<string>();

            // Get the layout dictionary of the current database
            using (Transaction acTrans = _curDb.TransactionManager.StartTransaction())
            {
                DBDictionary lays =
                    acTrans.GetObject(_curDb.LayoutDictionaryId,
                        OpenMode.ForRead) as DBDictionary;



                // Step through and list each named layout and Model
                foreach (DBDictionaryEntry item in lays)
                {
                    layoutNames.Add(item.Key);

                }

                // Abort the changes to the database
                acTrans.Abort();
            }

            return layoutNames;
        }


        [CommandMethod("AB_CreateLayout")]
        public void CreateLayout()
        {
            var names = ListLayouts();

            PromptResult pr = _editor.GetString("Enter layout Name");
            if (pr.Status == PromptStatus.OK)
            {
                if (names.Contains(pr.StringResult))
                {
                    _editor.WriteMessage("Layout with the same name already exists");
                    return;
                }
            }

            // Get the layout and plot settings of the named pagesetup
            using (Transaction acTrans = _curDb.TransactionManager.StartTransaction())
            {
                // Reference the Layout Manager
                LayoutManager acLayoutMgr = LayoutManager.Current;

                // Create the new layout with default settings
                ObjectId objID = acLayoutMgr.CreateLayout(pr.StringResult);

                // Open the layout
                Layout acLayout = acTrans.GetObject(objID,
                                                    OpenMode.ForRead) as Layout;


                // Get the PlotInfo from the layout
                PlotInfo acPlInfo = new PlotInfo();
                acPlInfo.Layout = acLayout.ObjectId;

                // Get a copy of the PlotSettings from the layout
                PlotSettings acPlSet = new PlotSettings(acLayout.ModelType);
                acPlSet.CopyFrom(acLayout);

                // Update the PlotConfigurationName property of the PlotSettings object
                PlotSettingsValidator acPlSetVdr = PlotSettingsValidator.Current;
                acPlSetVdr.SetPlotConfigurationName(acPlSet, "AutoCAD PDF (General Documentation).pc3",
                                                    "ISO_A1_(594.00_x_841.00_MM)");

                // Update the layout
                acLayout.UpgradeOpen();
                acLayout.CopyFrom(acPlSet);

                // Output the name of the new device assigned to the layout
                _editor.WriteMessage("\nNew device name: " +
                                           acLayout.PlotConfigurationName);



                // Set the layout current if it is not already
                if (acLayout.TabSelected == false)
                {
                    acLayoutMgr.CurrentLayout = acLayout.LayoutName;
                }




                // Save the changes made
                acTrans.Commit();
            }

        }

        [CommandMethod("AB_FourFloatingViewports")]
        public static void FourFloatingViewports()
        {
            // Get the current document and database, and start a transaction
            Document acDoc = AAAS.Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                             OpenMode.ForRead) as BlockTable;

                // Open the Block table record Paper space for write
                BlockTableRecord acBlkTblRec;
                acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.PaperSpace],
                                                OpenMode.ForWrite) as BlockTableRecord;

                // Switch to the previous Paper space layout
                AAAS.Application.SetSystemVariable("TILEMODE", 0);
                acDoc.Editor.SwitchToPaperSpace();

                Point3dCollection acPt3dCol = new Point3dCollection();
                acPt3dCol.Add(new Point3d(2.5, 5.5, 0));
                acPt3dCol.Add(new Point3d(2.5, 2.5, 0));
                acPt3dCol.Add(new Point3d(5.5, 5.5, 0));
                acPt3dCol.Add(new Point3d(5.5, 2.5, 0));

                Vector3dCollection acVec3dCol = new Vector3dCollection();
                acVec3dCol.Add(new Vector3d(0, 0, 1));
                acVec3dCol.Add(new Vector3d(0, 1, 0));
                acVec3dCol.Add(new Vector3d(1, 0, 0));
                acVec3dCol.Add(new Vector3d(1, 1, 1));

                double dWidth = 2.5;
                double dHeight = 2.5;

                Viewport acVportLast = null;
                int nCnt = 0;

                foreach (Point3d acPt3d in acPt3dCol)
                {
                    using (Viewport acVport = new Viewport())
                    {
                        acVport.CenterPoint = acPt3d;
                        acVport.Width = dWidth;
                        acVport.Height = dHeight;

                        // Add the new object to the block table record and the transaction
                        acBlkTblRec.AppendEntity(acVport);
                        acTrans.AddNewlyCreatedDBObject(acVport, true);

                        // Change the view direction
                        acVport.ViewDirection = acVec3dCol[nCnt];

                        // Enable the viewport
                        acVport.On = true;

                        // Record the last viewport created
                        acVportLast = acVport;

                        // Increment the counter by 1
                        nCnt = nCnt + 1;
                    }
                }

                if (acVportLast != null)
                {
                    // Activate model space in the viewport
                    acDoc.Editor.SwitchToModelSpace();

                
                }

                // Save the new objects to the database
                acTrans.Commit();
            }
        }

        private void _zoom(Point3d pMin, Point3d pMax, Point3d pCenter, double dFactor)
        {



            int nCurVport = System.Convert.ToInt32(Autodesk.AutoCAD.ApplicationServices.Application.GetSystemVariable("CVPORT"));

            // Get the extents of the current space when no points 
            // or only a center point is provided
            // Check to see if Model space is current

            //in autocad TileMode == 0 means paperspace   ==1 means model Space
            //in code tile mode equal true when in model space
            if (_curDb.TileMode == true)
            {
                if (pMin.Equals(new Point3d()) == true &&
                    pMax.Equals(new Point3d()) == true)
                {
                    pMin = _curDb.Extmin;
                    pMax = _curDb.Extmax;
                }
            }
            else
            {
                // Check to see if Paper space is current
                if (nCurVport == 1)
                {
                    // Get the extents of Paper space
                    if (pMin.Equals(new Point3d()) == true &&
                        pMax.Equals(new Point3d()) == true)
                    {
                        pMin = _curDb.Pextmin;
                        pMax = _curDb.Pextmax;
                    }
                }
                else
                {
                    // Get the extents of Model space
                    if (pMin.Equals(new Point3d()) == true &&
                        pMax.Equals(new Point3d()) == true)
                    {
                        pMin = _curDb.Extmin;
                        pMax = _curDb.Extmax;
                    }
                }
            }

            // Start a transaction
            using (Transaction acTrans = _curDb.TransactionManager.StartTransaction())
            {
                // Get the current view
                using (ViewTableRecord acView = _doc.Editor.GetCurrentView())
                {
                    Extents3d eExtents;

                    // Translate WCS coordinates to DCS
                    Matrix3d matWCS2DCS;
                    matWCS2DCS = Matrix3d.PlaneToWorld(acView.ViewDirection);
                    matWCS2DCS = Matrix3d.Displacement(acView.Target - Point3d.Origin) * matWCS2DCS;
                    matWCS2DCS = Matrix3d.Rotation(-acView.ViewTwist,
                                                    acView.ViewDirection,
                                                    acView.Target) * matWCS2DCS;

                    // If a center point is specified, define the min and max 
                    // point of the extents
                    // for Center and Scale modes
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        pMin = new Point3d(pCenter.X - (acView.Width / 2),
                                            pCenter.Y - (acView.Height / 2), 0);

                        pMax = new Point3d((acView.Width / 2) + pCenter.X,
                                            (acView.Height / 2) + pCenter.Y, 0);
                    }

                    // Create an extents object using a line
                    using (Line acLine = new Line(pMin, pMax))
                    {
                        eExtents = new Extents3d(acLine.Bounds.Value.MinPoint,
                                                    acLine.Bounds.Value.MaxPoint);
                    }

                    // Calculate the ratio between the width and height of the current view
                    double dViewRatio;
                    dViewRatio = (acView.Width / acView.Height);

                    // Tranform the extents of the view
                    matWCS2DCS = matWCS2DCS.Inverse();
                    eExtents.TransformBy(matWCS2DCS);

                    double dWidth;
                    double dHeight;
                    Point2d pNewCentPt;

                    // Check to see if a center point was provided (Center and Scale modes)
                    if (pCenter.DistanceTo(Point3d.Origin) != 0)
                    {
                        dWidth = acView.Width;
                        dHeight = acView.Height;

                        if (dFactor == 0)
                        {
                            pCenter = pCenter.TransformBy(matWCS2DCS);
                        }

                        pNewCentPt = new Point2d(pCenter.X, pCenter.Y);
                    }
                    else // Working in Window, Extents and Limits mode
                    {
                        // Calculate the new width and height of the current view
                        dWidth = eExtents.MaxPoint.X - eExtents.MinPoint.X;
                        dHeight = eExtents.MaxPoint.Y - eExtents.MinPoint.Y;

                        // Get the center of the view
                        pNewCentPt = new Point2d(((eExtents.MaxPoint.X + eExtents.MinPoint.X) * 0.5),
                                                    ((eExtents.MaxPoint.Y + eExtents.MinPoint.Y) * 0.5));
                    }

                    // Check to see if the new width fits in current window
                    if (dWidth > (dHeight * dViewRatio)) dHeight = dWidth / dViewRatio;

                    // Resize and scale the view
                    if (dFactor != 0)
                    {
                        acView.Height = dHeight * dFactor;
                        acView.Width = dWidth * dFactor;
                    }

                    // Set the center of the view
                    acView.CenterPoint = pNewCentPt;

                    // Set the current view
                    _doc.Editor.SetCurrentView(acView);
                }

                // Commit the changes
                acTrans.Commit();
            }


        }



        #region IExtentionApplication

        public void Initialize()
        {
            AAAS.Document doc = AAAS.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            ed.WriteMessage("Hello First Command");
        }

        public void Terminate()
        {
            throw new NotImplementedException();
        }
        #endregion
    }
}
