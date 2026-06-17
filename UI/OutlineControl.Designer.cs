namespace WordTool.UI
{
    partial class OutlineControl
    {
        private System.ComponentModel.IContainer components = null;
        private System.Windows.Forms.TreeView treeViewOutline;
        private System.Windows.Forms.Button btnConfirm;
        private System.Windows.Forms.Panel panelTop;

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
            this.treeViewOutline = new System.Windows.Forms.TreeView();
            this.btnConfirm = new System.Windows.Forms.Button();
            this.panelTop = new System.Windows.Forms.Panel();
            this.panelTop.SuspendLayout();
            this.SuspendLayout();
            // 
            // treeViewOutline
            // 
            this.treeViewOutline.Dock = System.Windows.Forms.DockStyle.Fill;
            this.treeViewOutline.Location = new System.Drawing.Point(0, 40);
            this.treeViewOutline.Name = "treeViewOutline";
            this.treeViewOutline.Size = new System.Drawing.Size(250, 410);
            this.treeViewOutline.TabIndex = 0;
            // 
            // btnConfirm
            // 
            this.btnConfirm.Dock = System.Windows.Forms.DockStyle.Fill;
            this.btnConfirm.Location = new System.Drawing.Point(5, 5);
            this.btnConfirm.Name = "btnConfirm";
            this.btnConfirm.Size = new System.Drawing.Size(240, 30);
            this.btnConfirm.TabIndex = 1;
            this.btnConfirm.Text = "确认无误，执行一键排版";
            this.btnConfirm.UseVisualStyleBackColor = true;
            this.btnConfirm.Click += new System.EventHandler(this.btnConfirm_Click);
            // 
            // panelTop
            // 
            this.panelTop.Controls.Add(this.btnConfirm);
            this.panelTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.panelTop.Location = new System.Drawing.Point(0, 0);
            this.panelTop.Name = "panelTop";
            this.panelTop.Padding = new System.Windows.Forms.Padding(5);
            this.panelTop.Size = new System.Drawing.Size(250, 40);
            this.panelTop.TabIndex = 2;
            // 
            // OutlineControl
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.treeViewOutline);
            this.Controls.Add(this.panelTop);
            this.Name = "OutlineControl";
            this.Size = new System.Drawing.Size(250, 450);
            this.panelTop.ResumeLayout(false);
            this.ResumeLayout(false);
        }
    }
}
