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

    [assembly:CommandClass(typeof(AutoCAD_API_Practice.AutoCAD_Commands))]
namespace AutoCAD_API_Practice
{
    public class AutoCAD_Commands:IExtensionApplication
    {
        [CommandMethod("Hello")]
        public void Hello()
        {
            AAAS.Document doc = AAAS.Application.DocumentManager.MdiActiveDocument;
            Editor ed=doc.Editor;
            ed.WriteMessage("Hello First Command");
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
