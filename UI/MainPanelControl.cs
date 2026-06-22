using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Xml.Serialization;
using System.Collections.Generic;
using System.Text;
using WordTool.Models;

namespace WordTool.UI
{
    public partial class MainPanelControl : UserControl
    {
        private Workflows.FormatWorkflow _workflow;
        private List<FormattingTemplate> _templates = new List<FormattingTemplate>();
        private string _configPath;
        private readonly object _logLock = new object();
        private readonly Queue<string> _pendingLogs = new Queue<string>();
        private Timer _logFlushTimer;
        private readonly Color _wordBlue = Color.FromArgb(43, 87, 154);
        private readonly Color _wordBlueDark = Color.FromArgb(31, 72, 136);
        private readonly Color _surface = Color.FromArgb(248, 250, 252);
        private readonly Color _border = Color.FromArgb(218, 225, 233);
        private readonly Color _text = Color.FromArgb(31, 41, 55);
        private readonly Color _mutedText = Color.FromArgb(99, 115, 129);
        
        public MainPanelControl()
        {
            InitializeComponent();
            ApplyWordTheme();
            InitializeTemplates();
            this.cmbTemplate.SelectedIndexChanged += CmbTemplate_SelectedIndexChanged;
            InitializeLogBuffer();
        }

        public void BindWorkflow(Workflows.FormatWorkflow workflow)
        {
            _workflow = workflow;
            UpdateWorkflowTemplate();
        }

        private void InitializeTemplates()
        {
            try
            {
                string appDataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "WordTool");
                if (!Directory.Exists(appDataFolder))
                {
                    Directory.CreateDirectory(appDataFolder);
                }

                _configPath = Path.Combine(appDataFolder, "templates.xml");

                bool needReset = false;
                if (!File.Exists(_configPath))
                {
                    needReset = true;
                }
                else
                {
                    try
                    {
                        LoadTemplatesFromFile();
                        if (_templates == null || _templates.Count == 0)
                        {
                            needReset = true;
                        }
                        else
                        {
                            NormalizeLoadedTemplates();
                        }
                    }
                    catch (Exception loadEx)
                    {
                        // 遇到序列化异常或格式不匹配时，自动重置为默认值，避免程序卡死或失效
                        System.Diagnostics.Debug.WriteLine($"加载排版模板异常: {loadEx.Message}");
                        needReset = true;
                    }
                }

                if (needReset)
                {
                    _templates = new List<FormattingTemplate>
                    {
                        FormattingTemplate.GetDefaultThesisTemplate(),
                        FormattingTemplate.GetDefaultOfficialDocumentTemplate()
                    };
                    SaveTemplatesToFile();
                }

                RefreshTemplateCombo();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"加载排版模板失败: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void LoadTemplatesFromFile()
        {
            var serializer = new XmlSerializer(typeof(List<FormattingTemplate>));
            using (var reader = new StreamReader(_configPath))
            {
                _templates = (List<FormattingTemplate>)serializer.Deserialize(reader);
            }
        }

        private void NormalizeLoadedTemplates()
        {
            foreach (var template in _templates)
            {
                if (template.TableWidthPercent <= 0)
                {
                    template.TableWidthPercent = 100.0f;
                }

                if (template.TableTopBorderWidth <= 0)
                {
                    template.TableTopBorderWidth = template.TableTopBottomBorderWidth > 0 ? template.TableTopBottomBorderWidth : 1.5f;
                }

                if (template.TableHeaderBottomBorderWidth <= 0)
                {
                    template.TableHeaderBottomBorderWidth = 0.75f;
                }

                if (template.TableBottomBorderWidth <= 0)
                {
                    template.TableBottomBorderWidth = template.TableTopBottomBorderWidth > 0 ? template.TableTopBottomBorderWidth : 1.5f;
                }
            }
        }

        private void SaveTemplatesToFile()
        {
            var serializer = new XmlSerializer(typeof(List<FormattingTemplate>));
            using (var writer = new StreamWriter(_configPath))
            {
                serializer.Serialize(writer, _templates);
            }
        }

        private void RefreshTemplateCombo()
        {
            cmbTemplate.Items.Clear();
            foreach (var t in _templates)
            {
                cmbTemplate.Items.Add(t.Name);
            }

            if (cmbTemplate.Items.Count > 0)
            {
                cmbTemplate.SelectedIndex = 0;
            }
        }

        private void CmbTemplate_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateWorkflowTemplate();
        }

        private void UpdateWorkflowTemplate()
        {
            if (_workflow == null) return;
            int index = cmbTemplate.SelectedIndex;
            if (index >= 0 && index < _templates.Count)
            {
                _workflow.Template = _templates[index];
            }
        }

        private void ApplyWordTheme()
        {
            this.BackColor = _surface;
            this.Font = new Font("Microsoft YaHei UI", 9.0f);

            panelTemplate.BackColor = Color.White;
            panelTemplate.Padding = new Padding(12, 8, 12, 6);
            panelTemplate.Height = 44;

            lblTemplate.Text = "模板";
            lblTemplate.Width = 46;
            lblTemplate.ForeColor = _mutedText;
            lblTemplate.Font = new Font("Microsoft YaHei UI", 9.0f, FontStyle.Regular);

            cmbTemplate.FlatStyle = FlatStyle.Flat;
            cmbTemplate.Font = new Font("Microsoft YaHei UI", 9.0f);
            cmbTemplate.ForeColor = _text;
            cmbTemplate.BackColor = Color.White;

            StyleSecondaryButton(btnEditTemplate);
            btnEditTemplate.Width = 64;

            tabControl1.Font = new Font("Microsoft YaHei UI", 9.0f);
            tabControl1.Height = 275;
            tabPageAuto.BackColor = _surface;
            tabPageManual.BackColor = _surface;
            tabPageAuto.Padding = new Padding(12);
            tabPageManual.Padding = new Padding(12);

            StylePrimaryButton(btnRunAll);
            btnRunAll.Height = 48;

            StyleStepButton(btnClean);
            StyleStepButton(btnStyle);
            StyleStepButton(btnAnalyze);
            StyleStepButton(btnImage);
            StyleStepButton(btnTable);
            StyleStepButton(btnLayout);

            panelExecutionControl.BackColor = Color.White;
            panelExecutionControl.Height = 46;
            panelExecutionControl.Padding = new Padding(12, 8, 12, 8);
            StyleSecondaryButton(btnPause);
            StyleDangerButton(btnStop);

            progressBar1.Height = 8;

            txtLog.BorderStyle = BorderStyle.FixedSingle;
            txtLog.BackColor = Color.White;
            txtLog.ForeColor = Color.FromArgb(38, 50, 56);
            txtLog.Font = new Font("Consolas", 9.0f);
            txtLog.Margin = new Padding(12);
        }

        private void StylePrimaryButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.BackColor = _wordBlue;
            button.ForeColor = Color.White;
            button.Font = new Font("Microsoft YaHei UI", 9.5f, FontStyle.Bold);
            button.Cursor = Cursors.Hand;
            button.TextAlign = ContentAlignment.MiddleCenter;
        }

        private void StyleStepButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderColor = _border;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(232, 240, 254);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(212, 227, 252);
            button.BackColor = Color.White;
            button.ForeColor = _text;
            button.Font = new Font("Microsoft YaHei UI", 9.0f, FontStyle.Regular);
            button.Height = 38;
            button.Cursor = Cursors.Hand;
            button.TextAlign = ContentAlignment.MiddleLeft;
            button.Padding = new Padding(14, 0, 10, 0);
        }

        private void StyleSecondaryButton(Button button)
        {
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderColor = _border;
            button.FlatAppearance.BorderSize = 1;
            button.FlatAppearance.MouseOverBackColor = Color.FromArgb(238, 244, 252);
            button.FlatAppearance.MouseDownBackColor = Color.FromArgb(222, 235, 250);
            button.BackColor = Color.White;
            button.ForeColor = _wordBlueDark;
            button.Font = new Font("Microsoft YaHei UI", 9.0f, FontStyle.Regular);
            button.Cursor = Cursors.Hand;
        }

        private void StyleDangerButton(Button button)
        {
            StyleSecondaryButton(button);
            button.ForeColor = Color.FromArgb(176, 48, 48);
        }

        private void btnEditTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                int index = cmbTemplate.SelectedIndex;
                if (index < 0 || index >= _templates.Count)
                {
                    MessageBox.Show($"未选择有效的模板或模板列表为空 (当前选择索引: {index}, 模板数量: {_templates.Count})", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                var selectedTemplate = _templates[index];
                using (var editor = new TemplateEditorForm(selectedTemplate))
                {
                    if (editor.ShowDialog(this) == DialogResult.OK)
                    {
                        _templates[index] = editor.EditingTemplate;
                        SaveTemplatesToFile();
                        LogMessage($"【模板】模板“{editor.EditingTemplate.Name}”配置已成功修改并保存。");
                        UpdateWorkflowTemplate();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"打开模板编辑器失败:\n异常信息: {ex.Message}\n\n堆栈信息:\n{ex.StackTrace}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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
            string line = $"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}";
            lock (_logLock)
            {
                _pendingLogs.Enqueue(line);
            }

            if (!IsHandleCreated || IsDisposed)
            {
                return;
            }

            if (!this.InvokeRequired)
            {
                FlushLogBuffer();
            }
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

        private void SetExecutionState(bool running)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<bool>(SetExecutionState), running);
                return;
            }

            // 运行排版时，禁用标签页和模板切换，以防重入和并发错误
            tabControl1.Enabled = !running;
            panelTemplate.Enabled = !running;
            
            btnPause.Enabled = running;
            btnStop.Enabled = running;

            if (!running)
            {
                btnPause.Text = "暂停";
            }
        }

        private void InitializeLogBuffer()
        {
            _logFlushTimer = new Timer();
            _logFlushTimer.Interval = 200;
            _logFlushTimer.Tick += (sender, args) => FlushLogBuffer();
            _logFlushTimer.Start();
        }

        private void FlushLogBuffer()
        {
            if (!IsHandleCreated || this.IsDisposed || txtLog.IsDisposed) return;

            if (this.InvokeRequired)
            {
                BeginInvoke(new Action(FlushLogBuffer));
                return;
            }

            StringBuilder builder = null;
            lock (_logLock)
            {
                if (_pendingLogs.Count == 0) return;
                builder = new StringBuilder();
                while (_pendingLogs.Count > 0)
                {
                    builder.Append(_pendingLogs.Dequeue());
                }
            }

            txtLog.AppendText(builder.ToString());
            txtLog.ScrollToCaret();
        }

        private void btnPause_Click(object sender, EventArgs e)
        {
            if (_workflow == null) return;

            if (_workflow.IsPaused)
            {
                _workflow.Resume();
                btnPause.Text = "暂停";
            }
            else
            {
                _workflow.Pause();
                btnPause.Text = "继续";
            }
        }

        private void btnStop_Click(object sender, EventArgs e)
        {
            if (_workflow == null) return;

            var result = MessageBox.Show("确定要中止当前正在执行的排版操作吗？", "中止确认", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                _workflow.Cancel();
                btnStop.Enabled = false; // 防止重复点击
            }
        }

        private void btnRunAll_Click(object sender, EventArgs e)
        {
            ClearLogBuffer();
            txtLog.Clear();
            SetExecutionState(true);
            _workflow?.ResetControlStates();
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
                    SetExecutionState(false);
                }
            });
        }

        private void ClearLogBuffer()
        {
            lock (_logLock)
            {
                _pendingLogs.Clear();
            }
        }

        private void btnClean_Click(object sender, EventArgs e)
        {
            SetExecutionState(true);
            _workflow?.ResetControlStates();
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunCleanData();
                }
                finally
                {
                    SetExecutionState(false);
                }
            });
        }

        private void btnStyle_Click(object sender, EventArgs e)
        {
            SetExecutionState(true);
            _workflow?.ResetControlStates();
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunStyleRebuild();
                }
                finally
                {
                    SetExecutionState(false);
                }
            });
        }

        private void btnAnalyze_Click(object sender, EventArgs e)
        {
            if (_workflow == null) return;
            SetExecutionState(true);
            _workflow.ResetControlStates();
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
                catch (OperationCanceledException)
                {
                    LogMessage("【终止】分步解析排版被用户中止。");
                }
                catch (Exception ex)
                {
                    LogMessage($"【错误】分步解析排版失败: {ex.Message}");
                }
                finally
                {
                    SetExecutionState(false);
                }
            });
        }

        private void btnImage_Click(object sender, EventArgs e)
        {
            SetExecutionState(true);
            _workflow?.ResetControlStates();
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunImageFormatting();
                }
                finally
                {
                    SetExecutionState(false);
                }
            });
        }

        private void btnTable_Click(object sender, EventArgs e)
        {
            SetExecutionState(true);
            _workflow?.ResetControlStates();
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunTableFormatting();
                }
                finally
                {
                    SetExecutionState(false);
                }
            });
        }

        private void btnLayout_Click(object sender, EventArgs e)
        {
            SetExecutionState(true);
            _workflow?.ResetControlStates();
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunLayout();
                }
                finally
                {
                    SetExecutionState(false);
                }
            });
        }
    }
}
