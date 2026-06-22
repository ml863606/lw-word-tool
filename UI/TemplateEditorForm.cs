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
        private Panel _headerPanel;
        private Label _titleLabel;
        private Label _subtitleLabel;
        private readonly Color _wordBlue = Color.FromArgb(43, 87, 154);
        private readonly Color _wordBlueDark = Color.FromArgb(31, 72, 136);
        private readonly Color _surface = Color.FromArgb(248, 250, 252);
        private readonly Color _border = Color.FromArgb(218, 225, 233);
        private readonly Color _text = Color.FromArgb(31, 41, 55);
        private readonly Color _mutedText = Color.FromArgb(99, 115, 129);
        
        public FormattingTemplate EditingTemplate { get; private set; }

        public TemplateEditorForm(FormattingTemplate templateToEdit)
        {
            EditingTemplate = CloneTemplate(templateToEdit);
            InitializeComponent();
            _propertyGrid.SelectedObject = EditingTemplate;
            _propertyGrid.ExpandAllGridItems();
        }

        private void InitializeComponent()
        {
            this._propertyGrid = new PropertyGrid();
            this._btnSave = new Button();
            this._btnCancel = new Button();
            this._btnPanel = new Panel();
            this._headerPanel = new Panel();
            this._titleLabel = new Label();
            this._subtitleLabel = new Label();
            
            this.SuspendLayout();

            var mainFont = new Font("Microsoft YaHei UI", 9.0f);
            var titleFont = new Font("Microsoft YaHei UI", 11.0f, FontStyle.Bold);
            var boldFont = new Font("Microsoft YaHei UI", 9.0f, FontStyle.Bold);

            //
            // _headerPanel
            //
            this._headerPanel.Dock = DockStyle.Top;
            this._headerPanel.Height = 66;
            this._headerPanel.BackColor = Color.White;
            this._headerPanel.Padding = new Padding(18, 12, 18, 10);
            this._headerPanel.Controls.Add(this._subtitleLabel);
            this._headerPanel.Controls.Add(this._titleLabel);

            //
            // _titleLabel
            //
            this._titleLabel.Dock = DockStyle.Top;
            this._titleLabel.Height = 24;
            this._titleLabel.Text = "配置排版模板";
            this._titleLabel.Font = titleFont;
            this._titleLabel.ForeColor = _text;
            this._titleLabel.TextAlign = ContentAlignment.MiddleLeft;

            //
            // _subtitleLabel
            //
            this._subtitleLabel.Dock = DockStyle.Top;
            this._subtitleLabel.Height = 20;
            this._subtitleLabel.Text = EditingTemplate.Name;
            this._subtitleLabel.Font = mainFont;
            this._subtitleLabel.ForeColor = _mutedText;
            this._subtitleLabel.TextAlign = ContentAlignment.MiddleLeft;

            // 
            // _propertyGrid
            // 
            this._propertyGrid.Dock = DockStyle.Fill;
            this._propertyGrid.Location = new Point(0, 66);
            this._propertyGrid.Name = "_propertyGrid";
            this._propertyGrid.Size = new Size(580, 640);
            this._propertyGrid.TabIndex = 0;
            this._propertyGrid.ToolbarVisible = false;
            
            this._propertyGrid.Font = new Font("Microsoft YaHei UI", 9.0f);
            this._propertyGrid.ViewBackColor = Color.White;
            this._propertyGrid.ViewForeColor = _text;
            this._propertyGrid.CategoryForeColor = _wordBlue;
            this._propertyGrid.CategorySplitterColor = _border;
            this._propertyGrid.LineColor = Color.FromArgb(235, 239, 244);
            this._propertyGrid.HelpBackColor = Color.FromArgb(243, 247, 252);
            this._propertyGrid.HelpForeColor = _mutedText;
            this._propertyGrid.BackColor = Color.White;

            // 
            // _btnSave
            // 
            this._btnSave.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this._btnSave.Font = boldFont;
            this._btnSave.Location = new Point(350, 12);
            this._btnSave.Name = "_btnSave";
            this._btnSave.Size = new Size(100, 34);
            this._btnSave.TabIndex = 0;
            this._btnSave.Text = "保存配置";
            this._btnSave.FlatStyle = FlatStyle.Flat;
            this._btnSave.FlatAppearance.BorderSize = 0;
            this._btnSave.BackColor = _wordBlue;
            this._btnSave.ForeColor = Color.White;
            this._btnSave.Cursor = Cursors.Hand;
            this._btnSave.Click += new EventHandler(this.btnSave_Click);

            // 
            // _btnCancel
            // 
            this._btnCancel.Anchor = AnchorStyles.Bottom | AnchorStyles.Right;
            this._btnCancel.Font = mainFont;
            this._btnCancel.Location = new Point(460, 12);
            this._btnCancel.Name = "_btnCancel";
            this._btnCancel.Size = new Size(100, 34);
            this._btnCancel.TabIndex = 1;
            this._btnCancel.Text = "取消";
            this._btnCancel.FlatStyle = FlatStyle.Flat;
            this._btnCancel.FlatAppearance.BorderColor = _border;
            this._btnCancel.FlatAppearance.BorderSize = 1;
            this._btnCancel.BackColor = Color.White;
            this._btnCancel.ForeColor = _wordBlueDark;
            this._btnCancel.Cursor = Cursors.Hand;
            this._btnCancel.Click += new EventHandler(this.btnCancel_Click);

            // 
            // _btnPanel
            // 
            this._btnPanel.Controls.Add(this._btnSave);
            this._btnPanel.Controls.Add(this._btnCancel);
            this._btnPanel.Dock = DockStyle.Bottom;
            this._btnPanel.Height = 58;
            this._btnPanel.BackColor = _surface;
            this._btnPanel.Location = new Point(0, 640);
            this._btnPanel.Name = "_btnPanel";
            this._btnPanel.Padding = new Padding(5);

            // 
            // TemplateEditorForm
            // 
            this.ClientSize = new Size(600, 720);
            this.Controls.Add(this._propertyGrid);
            this.Controls.Add(this._headerPanel);
            this.Controls.Add(this._btnPanel);
            this.Name = "TemplateEditorForm";
            this.Text = $"配置排版模板 - {EditingTemplate.Name}";
            this.StartPosition = FormStartPosition.CenterParent;
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.MinimizeBox = false;
            this.MaximizeBox = false;
            this.FormBorderStyle = FormBorderStyle.SizableToolWindow;
            this.BackColor = _surface;
            this.Font = mainFont;
            
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
                TableWidthPercent = original.TableWidthPercent > 0 ? original.TableWidthPercent : 100.0f,
                TableAlignment = original.TableAlignment,
                TableAutoFit = original.TableAutoFit,
                TableThreeLine = original.TableThreeLine,
                TableTopBottomBorderWidth = original.TableTopBottomBorderWidth,
                TableTopBorderWidth = original.TableTopBorderWidth > 0 ? original.TableTopBorderWidth : (original.TableTopBottomBorderWidth > 0 ? original.TableTopBottomBorderWidth : 1.5f),
                TableHeaderBottomBorderWidth = original.TableHeaderBottomBorderWidth,
                TableBottomBorderWidth = original.TableBottomBorderWidth > 0 ? original.TableBottomBorderWidth : (original.TableTopBottomBorderWidth > 0 ? original.TableTopBottomBorderWidth : 1.5f),
                TableHeaderBold = original.TableHeaderBold,
                TableTextFontName = original.TableTextFontName,
                TableTextSize = original.TableTextSize,
                TableTextAlignment = original.TableTextAlignment,
                TableTextLineSpacingRule = original.TableTextLineSpacingRule,
                TableTextFirstLineIndentCm = original.TableTextFirstLineIndentCm,
                TocLevels = original.TocLevels,
                TocAutoUpdate = original.TocAutoUpdate
            };
            return clone;
        }
    }
}
