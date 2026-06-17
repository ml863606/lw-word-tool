using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace WordTool
{
    [ComVisible(true)]
    public class RibbonAssistant : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RibbonAssistant()
        {
        }

        #region IRibbonExtensibility 成员

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("WordTool.RibbonAssistant.xml");
        }

        #endregion

        #region 功能区回调

        private System.Collections.Generic.List<Models.AnalyzedParagraph> _currentAnalysis;

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void OnAutoFormat(Office.IRibbonControl control)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            Globals.ThisAddIn.PaneManager.ShowMainPanel(doc);
        }

        public void OnShowPanel(Office.IRibbonControl control)
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            Globals.ThisAddIn.PaneManager.ShowMainPanel(doc);
        }

        #endregion

        #region 帮助器

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
