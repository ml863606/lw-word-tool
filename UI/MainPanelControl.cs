using System;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Xml.Serialization;
using System.Collections.Generic;
using WordTool.Models;

namespace WordTool.UI
{
    public partial class MainPanelControl : UserControl
    {
        private Workflows.FormatWorkflow _workflow;
        private List<FormattingTemplate> _templates = new List<FormattingTemplate>();
        private string _configPath;
        
        public MainPanelControl()
        {
            InitializeComponent();
            InitializeTemplates();
            this.cmbTemplate.SelectedIndexChanged += CmbTemplate_SelectedIndexChanged;
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
                btnPause.Text = "⏸️ 暂停";
            }
        }

        private void btnPause_Click(object sender, EventArgs e)
        {
            if (_workflow == null) return;

            if (_workflow.IsPaused)
            {
                _workflow.Resume();
                btnPause.Text = "⏸️ 暂停";
            }
            else
            {
                _workflow.Pause();
                btnPause.Text = "▶️ 继续";
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

        private void btnMedia_Click(object sender, EventArgs e)
        {
            SetExecutionState(true);
            _workflow?.ResetControlStates();
            System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    _workflow?.RunMediaFormatting();
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
