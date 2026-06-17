using Microsoft.Office.Tools;
using System;
using WordTool.Workflows;

namespace WordTool.UI
{
    public class TaskPaneManager
    {
        private CustomTaskPane _mainTaskPane;
        private MainPanelControl _mainControl;

        public TaskPaneManager(CustomTaskPaneCollection customTaskPanes)
        {
            _mainControl = new MainPanelControl();
            _mainTaskPane = customTaskPanes.Add(_mainControl, "排版小助手控制台");
            _mainTaskPane.Width = 350;
        }

        public void ShowMainPanel(Microsoft.Office.Interop.Word.Document doc)
        {
            var workflow = new FormatWorkflow(doc, _mainControl.LogMessage, _mainControl.UpdateProgress);
            _mainControl.BindWorkflow(workflow);
            _mainTaskPane.Visible = true;
        }

        public void Hide()
        {
            _mainTaskPane.Visible = false;
        }
    }
}
