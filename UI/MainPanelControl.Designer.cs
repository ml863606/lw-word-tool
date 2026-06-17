namespace WordTool.UI
{
    partial class MainPanelControl
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPageAuto;
        private System.Windows.Forms.TabPage tabPageManual;
        private System.Windows.Forms.Button btnRunAll;
        private System.Windows.Forms.Button btnClean;
        private System.Windows.Forms.Button btnStyle;
        private System.Windows.Forms.Button btnAnalyze;
        private System.Windows.Forms.Button btnMedia;
        private System.Windows.Forms.Button btnLayout;
        private System.Windows.Forms.RichTextBox txtLog;
        private System.Windows.Forms.ProgressBar progressBar1;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPageAuto = new System.Windows.Forms.TabPage();
            this.tabPageManual = new System.Windows.Forms.TabPage();
            this.btnRunAll = new System.Windows.Forms.Button();
            this.btnClean = new System.Windows.Forms.Button();
            this.btnStyle = new System.Windows.Forms.Button();
            this.btnAnalyze = new System.Windows.Forms.Button();
            this.btnMedia = new System.Windows.Forms.Button();
            this.btnLayout = new System.Windows.Forms.Button();
            this.txtLog = new System.Windows.Forms.RichTextBox();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            
            this.tabControl1.SuspendLayout();
            this.tabPageAuto.SuspendLayout();
            this.tabPageManual.SuspendLayout();
            this.SuspendLayout();

            // tabControl1
            this.tabControl1.Controls.Add(this.tabPageAuto);
            this.tabControl1.Controls.Add(this.tabPageManual);
            this.tabControl1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tabControl1.Height = 250;

            // tabPageAuto
            this.tabPageAuto.Controls.Add(this.btnRunAll);
            this.tabPageAuto.Text = "全自动巡航";
            this.tabPageAuto.Padding = new System.Windows.Forms.Padding(10);

            // btnRunAll
            this.btnRunAll.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnRunAll.Height = 60;
            this.btnRunAll.Text = "🚀 开始一键全自动排版";
            this.btnRunAll.Click += new System.EventHandler(this.btnRunAll_Click);

            // tabPageManual
            this.tabPageManual.Controls.Add(this.btnLayout);
            this.tabPageManual.Controls.Add(this.btnMedia);
            this.tabPageManual.Controls.Add(this.btnAnalyze);
            this.tabPageManual.Controls.Add(this.btnStyle);
            this.tabPageManual.Controls.Add(this.btnClean);
            this.tabPageManual.Text = "分步精准打击";
            this.tabPageManual.Padding = new System.Windows.Forms.Padding(10);

            // Buttons in Manual Tab (Dock Top adds them bottom-up in code, so reverse order)
            this.btnLayout.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnLayout.Height = 40;
            this.btnLayout.Text = "5. 生成目录与页面设置";
            this.btnLayout.Click += new System.EventHandler(this.btnLayout_Click);

            this.btnMedia.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnMedia.Height = 40;
            this.btnMedia.Text = "4. 图表居中与自动调整";
            this.btnMedia.Click += new System.EventHandler(this.btnMedia_Click);

            this.btnAnalyze.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnAnalyze.Height = 40;
            this.btnAnalyze.Text = "3. 解析并套用正文样式";
            this.btnAnalyze.Click += new System.EventHandler(this.btnAnalyze_Click);

            this.btnStyle.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnStyle.Height = 40;
            this.btnStyle.Text = "2. 重建底层标准样式";
            this.btnStyle.Click += new System.EventHandler(this.btnStyle_Click);

            this.btnClean.Dock = System.Windows.Forms.DockStyle.Top;
            this.btnClean.Height = 40;
            this.btnClean.Text = "1. 清理脏数据 (空行等)";
            this.btnClean.Click += new System.EventHandler(this.btnClean_Click);

            // txtLog
            this.txtLog.Dock = System.Windows.Forms.DockStyle.Fill;
            this.txtLog.BackColor = System.Drawing.SystemColors.Window;
            this.txtLog.ForeColor = System.Drawing.SystemColors.WindowText;
            this.txtLog.ReadOnly = true;
            this.txtLog.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;

            // progressBar1
            this.progressBar1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.progressBar1.Height = 20;

            // MainPanelControl
            this.Controls.Add(this.txtLog);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.tabControl1);
            this.Size = new System.Drawing.Size(300, 600);

            this.tabControl1.ResumeLayout(false);
            this.tabPageAuto.ResumeLayout(false);
            this.tabPageManual.ResumeLayout(false);
            this.ResumeLayout(false);
        }
    }
}
