using System;
using System.Drawing;
using System.Windows.Forms;
using WordTool.Models;

namespace WordTool.UI
{
    public class TemplateEditorForm : Form
    {
        private PropertyGrid _propertyGrid;
        private Button _btnSave;
        private Button _btnCancel;
        private Panel _btnPanel;
        
        public FormattingTemplate EditingTemplate { get; private set; }

        public TemplateEditorForm(FormattingTemplate templateToEdit)
        {
            EditingTemplate = CloneTemplate(templateToEdit);
            InitializeComponent();
            _propertyGrid.SelectedObject = EditingTemplate;
        }

        private void InitializeComponent()
        {
            this._propertyGrid = new PropertyGrid();
            this._btnSave = new Button();
            this._btnCancel = new Button();
            this._btnPanel = new Panel();
            
            this.SuspendLayout();

            // 统一应用微软雅黑字体，大小设为 10.5 (小四) 解决字体太小难以看清的问题
            var mainFont = new Font("Microsoft YaHei", 10.5f);
            var boldFont = new Font("Microsoft YaHei", 10.5f, FontStyle.Bold);

            // 
            // _propertyGrid
            // 
            this._propertyGrid.Dock = DockStyle.Fill;
            this._propertyGrid.Location = new Point(0, 0);
            this._propertyGrid.Name = "_propertyGrid";
            this._propertyGrid.Size = new Size(580, 640);
            this._propertyGrid.TabIndex = 0;
            this._propertyGrid.ToolbarVisible = false; // 隐藏工具栏，使布局更大气
            
            // 美化 PropertyGrid 视觉外观
            this._propertyGrid.Font = new Font("Microsoft YaHei", 10.0f); // 网格列表字体大小
            this._propertyGrid.ViewBackColor = Color.White;
            this._propertyGrid.ViewForeColor = Color.FromArgb(50, 50, 50);
            this._propertyGrid.CategoryForeColor = Color.FromArgb(0, 102, 204); // 分类名称颜色设为科技蓝
            this._propertyGrid.CategorySplitterColor = Color.FromArgb(235, 235, 235);
            this._propertyGrid.LineColor = Color.FromArgb(240, 240, 240); // 虚线设为更浅淡的灰
            this._propertyGrid.HelpBackColor = Color.FromArgb(245, 247, 250); // 底部说明栏背景色
            this._propertyGrid.HelpForeColor = Color.FromArgb(80, 80, 80);

            // 
            // _btnSave
            // 
            this._btnSave.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this._btnSave.Font = boldFont;
            this._btnSave.Location = new Point(350, 9);
            this._btnSave.Name = "_btnSave";
            this._btnSave.Size = new Size(100, 32);
            this._btnSave.TabIndex = 0;
            this._btnSave.Text = "保存配置";
            this._btnSave.UseVisualStyleBackColor = true;
            this._btnSave.Click += new EventHandler(this.btnSave_Click);

            // 
            // _btnCancel
            // 
            this._btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this._btnCancel.Font = mainFont;
            this._btnCancel.Location = new Point(460, 9);
            this._btnCancel.Name = "_btnCancel";
            this._btnCancel.Size = new Size(100, 32);
            this._btnCancel.TabIndex = 1;
            this._btnCancel.Text = "取消";
            this._btnCancel.UseVisualStyleBackColor = true;
            this._btnCancel.Click += new EventHandler(this.btnCancel_Click);

            // 
            // _btnPanel
            // 
            this._btnPanel.Controls.Add(this._btnSave);
            this._btnPanel.Controls.Add(this._btnCancel);
            this._btnPanel.Dock = DockStyle.Bottom;
            this._btnPanel.Height = 50;
            this._btnPanel.BackColor = Color.FromArgb(245, 245, 245);
            this._btnPanel.Location = new Point(0, 640);
            this._btnPanel.Name = "_btnPanel";
            this._btnPanel.Padding = new Padding(5);

            // 
            // TemplateEditorForm
            // 
            this.ClientSize = new Size(580, 690);
            this.Controls.Add(this._propertyGrid);
            this.Controls.Add(this._btnPanel);
            this.Name = "TemplateEditorForm";
            this.Text = $"配置排版模板 - {EditingTemplate.Name}";
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            this.Font = mainFont; // 窗体默认字体
            
            this.ResumeLayout(false);
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private FormattingTemplate CloneTemplate(FormattingTemplate original)
        {
            var clone = new FormattingTemplate
            {
                Name = original.Name,
                PaperSize = original.PaperSize,
                TopMarginCm = original.TopMarginCm,
                BottomMarginCm = original.BottomMarginCm,
                LeftMarginCm = original.LeftMarginCm,
                RightMarginCm = original.RightMarginCm,
                NormalFontNameFarEast = original.NormalFontNameFarEast,
                NormalFontNameAscii = original.NormalFontNameAscii,
                NormalFontSize = original.NormalFontSize,
                NormalLineSpacingRule = original.NormalLineSpacingRule,
                NormalFirstLineIndentChars = original.NormalFirstLineIndentChars,
                H1FontNameFarEast = original.H1FontNameFarEast,
                H1FontSize = original.H1FontSize,
                H1Bold = original.H1Bold,
                H1Alignment = original.H1Alignment,
                H1SpaceBefore = original.H1SpaceBefore,
                H1SpaceAfter = original.H1SpaceAfter,
                H2FontNameFarEast = original.H2FontNameFarEast,
                H2FontSize = original.H2FontSize,
                H2Bold = original.H2Bold,
                H2Alignment = original.H2Alignment,
                H2SpaceBefore = original.H2SpaceBefore,
                H2SpaceAfter = original.H2SpaceAfter,
                H3FontNameFarEast = original.H3FontNameFarEast,
                H3FontSize = original.H3FontSize,
                H3Bold = original.H3Bold,
                H3Alignment = original.H3Alignment,
                RefFontSize = original.RefFontSize,
                RefFirstLineIndentCm = original.RefFirstLineIndentCm,
                RefLeftIndentCm = original.RefLeftIndentCm,
                CaptionFontSize = original.CaptionFontSize,
                CaptionBold = original.CaptionBold,
                CaptionAlignment = original.CaptionAlignment,
                NoteFontNameFarEast = original.NoteFontNameFarEast,
                NoteFontSize = original.NoteFontSize,
                NoteAlignment = original.NoteAlignment,
                ImageAlignment = original.ImageAlignment,
                ImageSpaceBefore = original.ImageSpaceBefore,
                ImageSpaceAfter = original.ImageSpaceAfter,
                TableAlignment = original.TableAlignment,
                TableAutoFit = original.TableAutoFit,
                TableThreeLine = original.TableThreeLine,
                TableTopBottomBorderWidth = original.TableTopBottomBorderWidth,
                TableHeaderBottomBorderWidth = original.TableHeaderBottomBorderWidth,
                TableHeaderBold = original.TableHeaderBold,
                TableTextFontName = original.TableTextFontName,
                TableTextSize = original.TableTextSize,
                TableTextLineSpacingRule = original.TableTextLineSpacingRule,
                TableTextFirstLineIndentCm = original.TableTextFirstLineIndentCm,
                TocLevels = original.TocLevels,
                TocAutoUpdate = original.TocAutoUpdate
            };
            return clone;
        }
    }
}
