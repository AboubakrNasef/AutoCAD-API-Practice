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
using System.Windows.Forms;
using msVB = Microsoft.VisualBasic;
using Autodesk.AutoCAD.GraphicsSystem;

[assembly: CommandClass(typeof(AutoCAD_API_Practice.Layout_Commands))]
namespace AutoCAD_API_Practice
{
    public class Layout_Commands : IExtensionApplication
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

        [CommandMethod("AB_RenameLayouts")]
        public void AB_RenameLayouts()
        {

            string oldName;
            string newName;


            List<DBDictionaryEntry> lstLayouts = ListLayouts();

            using (Transaction tr = _DB.TransactionManager.StartTransaction())
            {

                var inputResult = msVB.Interaction.InputBox("Enter Sheet Name");
                if (!string.IsNullOrEmpty(inputResult))
                {
                    for (int i = 0; i < lstLayouts.Count; i++)
                    {

                        Layout layout = (Layout)tr.GetObject(lstLayouts[i].Value, OpenMode.ForWrite);


                        layout.UpgradeOpen();

                        oldName = layout.LayoutName;
                        newName = inputResult + $" {i + 1}";


                        if (oldName != newName)
                        {
                            if (!layout.ModelType)
                            {
                                LayoutManager.Current.RenameLayout(oldName, newName);

                            }
                        }



                    }
                    tr.Commit();

                }
            }
        }

        public List<string> ListLayoutsNames()
        {
            List<string> layoutNames = new List<string>();

            // Get the layout dictionary of the current database
            using (Transaction acTrans = _DB.TransactionManager.StartTransaction())
            {
                DBDictionary lays =
                    acTrans.GetObject(_DB.LayoutDictionaryId,
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
        public List<DBDictionaryEntry> ListLayouts()
        {
            List<DBDictionaryEntry> layouts = new List<DBDictionaryEntry>();

            // Get the layout dictionary of the current database
            using (Transaction acTrans = _DB.TransactionManager.StartTransaction())
            {
                DBDictionary lays =
                    acTrans.GetObject(_DB.LayoutDictionaryId,
                        OpenMode.ForRead) as DBDictionary;



                // Step through and list each named layout and Model
                foreach (DBDictionaryEntry item in lays)
                {
                    layouts.Add(item);

                }

                // Abort the changes to the database
                acTrans.Abort();
            }

            return layouts;
        }

        [CommandMethod("AB_CreateLayout")]
        public void CreateLayout()
        {
            var names = ListLayoutsNames();

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
            using (Transaction acTrans = _DB.TransactionManager.StartTransaction())
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
