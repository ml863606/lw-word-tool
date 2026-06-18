using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using WordTool.Workflows;

namespace WordTool.UI
{
    public class TaskPaneManager
    {
        private readonly CustomTaskPaneCollection _customTaskPanes;
        private readonly Dictionary<Word.Window, CustomTaskPane> _panes = new Dictionary<Word.Window, CustomTaskPane>();

        public TaskPaneManager(CustomTaskPaneCollection customTaskPanes)
        {
            _customTaskPanes = customTaskPanes;
        }

        public void ShowMainPanel(Word.Document doc, bool showAutoTab = true)
        {
            Word.Window activeWindow = doc.ActiveWindow;
            if (activeWindow == null) return;

            CustomTaskPane pane = null;

            // Remove closed/invalid windows to avoid memory leak
            CleanUpClosedPanes();

            // Find or create pane for the active window
            foreach (var kvp in _panes)
            {
                try
                {
                    if (kvp.Key == activeWindow)
                    {
                        pane = kvp.Value;
                        break;
                    }
                }
                catch { }
            }

            if (pane == null)
            {
                var mainControl = new MainPanelControl();
                pane = _customTaskPanes.Add(mainControl, "排版小助手控制台", activeWindow);
                pane.Width = 350;
                
                // Track it
                _panes[activeWindow] = pane;
            }

            var control = pane.Control as MainPanelControl;
            if (control != null)
            {
                var workflow = new FormatWorkflow(doc, control.LogMessage, control.UpdateProgress);
                workflow.YieldUICallback = () =>
                {
                    try
                    {
                        if (control != null && !control.IsDisposed && control.IsHandleCreated)
                        {
                            control.Invoke(new Action(() => System.Windows.Forms.Application.DoEvents()));
                        }
                    }
                    catch { }
                };
                control.BindWorkflow(workflow);
                control.SelectTab(showAutoTab);
            }

            pane.Visible = true;
        }

        public void Hide()
        {
            try
            {
                var activeWindow = Globals.ThisAddIn.Application.ActiveWindow;
                if (activeWindow != null && _panes.ContainsKey(activeWindow))
                {
                    _panes[activeWindow].Visible = false;
                }
            }
            catch { }
        }

        private void CleanUpClosedPanes()
        {
            var deadWindows = new List<Word.Window>();
            foreach (var kvp in _panes)
            {
                try
                {
                    // Accessing properties of closed windows throws COMException
                    var title = kvp.Key.Caption; 
                }
                catch
                {
                    deadWindows.Add(kvp.Key);
                }
            }

            foreach (var win in deadWindows)
            {
                try
                {
                    var pane = _panes[win];
                    _customTaskPanes.Remove(pane);
                }
                catch { }
                _panes.Remove(win);
            }
        }
    }
}
