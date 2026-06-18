using System;
using System.Collections.Generic;
using System.Windows.Forms;
using WordTool.Models;

namespace WordTool.UI
{
    public class OutlineConfirmForm : Form
    {
        private OutlineControl _outlineControl;

        public OutlineConfirmForm()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this._outlineControl = new OutlineControl();
            this.SuspendLayout();
            
            // 
            // _outlineControl
            // 
            this._outlineControl.Dock = DockStyle.Fill;
            this._outlineControl.Location = new System.Drawing.Point(0, 0);
            this._outlineControl.Name = "_outlineControl";
            this._outlineControl.Size = new System.Drawing.Size(480, 600);
            this._outlineControl.TabIndex = 0;
            this._outlineControl.OnConfirmFormatting += OutlineControl_OnConfirmFormatting;
            
            // 
            // OutlineConfirmForm
            // 
            this.ClientSize = new System.Drawing.Size(480, 600);
            this.Controls.Add(this._outlineControl);
            this.Name = "OutlineConfirmForm";
            this.Text = "确认文档大纲与角色识别结果";
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            
            this.ResumeLayout(false);
        }

        public void LoadData(List<AnalyzedParagraph> paragraphs)
        {
            _outlineControl.LoadData(paragraphs);
        }

        private void OutlineControl_OnConfirmFormatting(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
