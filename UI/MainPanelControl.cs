using System;
using System.Drawing;
using System.Windows.Forms;

namespace WordTool.UI
{
    public partial class MainPanelControl : UserControl
    {
        private Workflows.FormatWorkflow _workflow;
        
        public MainPanelControl()
        {
            InitializeComponent();
        }

        public void BindWorkflow(Workflows.FormatWorkflow workflow)
        {
            _workflow = workflow;
        }

        public void LogMessage(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(LogMessage), message);
                return;
            }
            txtLog.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
            txtLog.ScrollToCaret();
        }

        public void UpdateProgress(int percent)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<int>(UpdateProgress), percent);
                return;
            }
            progressBar1.Value = Math.Max(0, Math.Min(100, percent));
        }

        private void btnRunAll_Click(object sender, EventArgs e)
        {
            txtLog.Clear();
            btnRunAll.Enabled = false;
            // Background thread to not freeze UI
            System.Threading.Tasks.Task.Run(() =>
            {
                _workflow?.RunAllSteps();
                this.Invoke(new Action(() => btnRunAll.Enabled = true));
            });
        }

        private void btnClean_Click(object sender, EventArgs e)
        {
            _workflow?.RunCleanData();
        }

        private void btnStyle_Click(object sender, EventArgs e)
        {
            _workflow?.RunStyleRebuild();
        }

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            var paragraphs = _workflow?.RunAnalysis();
            // 在实际产品中，这里可能需要弹出 OutlineControl 进行确认
            // 为了简单演示，我们直接接着排版正文
            _workflow?.RunFormatting(paragraphs);
        }

        private void btnMedia_Click(object sender, EventArgs e)
        {
            _workflow?.RunMediaFormatting();
        }

        private void btnLayout_Click(object sender, EventArgs e)
        {
            _workflow?.RunLayout();
        }
    }
}
