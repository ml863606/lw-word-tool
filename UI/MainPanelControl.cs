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

        public void SelectTab(bool autoTab)
        {
            if (autoTab)
            {
                tabControl1.SelectedTab = tabPageAuto;
            }
            else
            {
                tabControl1.SelectedTab = tabPageManual;
            }
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
            progressBar1.Value = 0;
            
            // Background thread to not freeze UI
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunAllSteps(paragraphs =>
                    {
                        bool confirmed = false;
                        this.Invoke(new Action(() =>
                        {
                            using (var confirmForm = new OutlineConfirmForm())
                            {
                                confirmForm.LoadData(paragraphs);
                                var result = confirmForm.ShowDialog(this);
                                if (result == DialogResult.OK)
                                {
                                    confirmed = true;
                                }
                            }
                        }));
                        return confirmed;
                    });
                }
                finally
                {
                    this.Invoke(new Action(() => btnRunAll.Enabled = true));
                }
            });
        }

        private void btnClean_Click(object sender, EventArgs e)
        {
            btnClean.Enabled = false;
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunCleanData();
                }
                finally
                {
                    this.Invoke(new Action(() => btnClean.Enabled = true));
                }
            });
        }

        private void btnStyle_Click(object sender, EventArgs e)
        {
            btnStyle.Enabled = false;
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunStyleRebuild();
                }
                finally
                {
                    this.Invoke(new Action(() => btnStyle.Enabled = true));
                }
            });
        }

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            if (_workflow == null) return;
            btnAnalyze.Enabled = false;
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    var paragraphs = _workflow.RunAnalysis();
                    bool confirmed = false;
                    this.Invoke(new Action(() =>
                    {
                        using (var confirmForm = new OutlineConfirmForm())
                        {
                            confirmForm.LoadData(paragraphs);
                            var result = confirmForm.ShowDialog(this);
                            if (result == DialogResult.OK)
                            {
                                confirmed = true;
                            }
                        }
                    }));

                    if (confirmed)
                    {
                        _workflow.RunFormatting(paragraphs);
                    }
                    else
                    {
                        LogMessage("【取消】已取消正文套用样式。");
                    }
                }
                catch (Exception ex)
                {
                    LogMessage($"【错误】分步解析排版失败: {ex.Message}");
                }
                finally
                {
                    this.Invoke(new Action(() => btnAnalyze.Enabled = true));
                }
            });
        }

        private void btnMedia_Click(object sender, EventArgs e)
        {
            btnMedia.Enabled = false;
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunMediaFormatting();
                }
                finally
                {
                    this.Invoke(new Action(() => btnMedia.Enabled = true));
                }
            });
        }

        private void btnLayout_Click(object sender, EventArgs e)
        {
            btnLayout.Enabled = false;
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunLayout();
                }
                finally
                {
                    this.Invoke(new Action(() => btnLayout.Enabled = true));
                }
            });
        }
    }
}
