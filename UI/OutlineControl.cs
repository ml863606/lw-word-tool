using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using WordTool.Models;

namespace WordTool.UI
{
    public partial class OutlineControl : UserControl
    {
        private List<AnalyzedParagraph> _analyzedParagraphs;
        private ContextMenuStrip _contextMenu;
        public event EventHandler OnConfirmFormatting;

        public OutlineControl()
        {
            InitializeComponent();
            InitializeContextMenu();
            
            this.treeViewOutline.NodeMouseClick += treeViewOutline_NodeMouseClick;
            this.chkShowAll.CheckedChanged += chkShowAll_CheckedChanged;
        }

        private void InitializeContextMenu()
        {
            _contextMenu = new ContextMenuStrip();
            
            // Add roles options
            AddContextMenuItem("设为 一级标题 (Heading 1)", ParagraphRole.Heading1);
            AddContextMenuItem("设为 二级标题 (Heading 2)", ParagraphRole.Heading2);
            AddContextMenuItem("设为 三级标题 (Heading 3)", ParagraphRole.Heading3);
            AddContextMenuItem("设为 参考文献 (Reference)", ParagraphRole.Reference);
            AddContextMenuItem("设为 图表题注 (Caption)", ParagraphRole.Caption);
            AddContextMenuItem("设为 表注 (TableNote)", ParagraphRole.TableNote);
            _contextMenu.Items.Add(new ToolStripSeparator());
            AddContextMenuItem("设为 普通正文 (Normal)", ParagraphRole.Normal);

            this.treeViewOutline.ContextMenuStrip = _contextMenu;
        }

        private void AddContextMenuItem(string text, ParagraphRole role)
        {
            var item = new ToolStripMenuItem(text);
            item.Tag = role;
            item.Click += ContextMenuItem_Click;
            _contextMenu.Items.Add(item);
        }

        private void ContextMenuItem_Click(object sender, EventArgs e)
        {
            var menuItem = sender as ToolStripMenuItem;
            if (menuItem == null) return;

            var selectedNode = treeViewOutline.SelectedNode;
            if (selectedNode == null) return;

            var paragraph = selectedNode.Tag as AnalyzedParagraph;
            if (paragraph == null) return;

            var newRole = (ParagraphRole)menuItem.Tag;
            paragraph.Role = newRole;

            // Update display
            if (!chkShowAll.Checked && newRole == ParagraphRole.Normal)
            {
                // If we are not showing all, remove the normal paragraph from the view
                treeViewOutline.Nodes.Remove(selectedNode);
            }
            else
            {
                selectedNode.Text = $"[{paragraph.Role}] {paragraph.TextContent}";
                UpdateNodeStyle(selectedNode, paragraph);
            }
        }

        private void treeViewOutline_NodeMouseClick(object sender, TreeNodeMouseClickEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                treeViewOutline.SelectedNode = e.Node;
            }
        }

        private void chkShowAll_CheckedChanged(object sender, EventArgs e)
        {
            if (_analyzedParagraphs != null)
            {
                LoadData(_analyzedParagraphs);
            }
        }

        public void LoadData(List<AnalyzedParagraph> paragraphs)
        {
            _analyzedParagraphs = paragraphs;
            treeViewOutline.Nodes.Clear();

            foreach (var p in paragraphs)
            {
                bool shouldShow = chkShowAll.Checked || p.Role != ParagraphRole.Normal;
                if (shouldShow)
                {
                    TreeNode node = new TreeNode($"[{p.Role}] {p.TextContent}");
                    node.Tag = p;
                    UpdateNodeStyle(node, p);
                    treeViewOutline.Nodes.Add(node);
                }
            }
        }

        private void UpdateNodeStyle(TreeNode node, AnalyzedParagraph p)
        {
            switch (p.Role)
            {
                case ParagraphRole.Heading1:
                    node.ForeColor = Color.DarkBlue;
                    node.NodeFont = new Font(treeViewOutline.Font, FontStyle.Bold);
                    break;
                case ParagraphRole.Heading2:
                    node.ForeColor = Color.Blue;
                    node.NodeFont = new Font(treeViewOutline.Font, FontStyle.Bold);
                    break;
                case ParagraphRole.Heading3:
                    node.ForeColor = Color.SteelBlue;
                    node.NodeFont = new Font(treeViewOutline.Font, FontStyle.Bold);
                    break;
                case ParagraphRole.Reference:
                    node.ForeColor = Color.DarkGreen;
                    break;
                case ParagraphRole.Caption:
                    node.ForeColor = Color.DarkMagenta;
                    break;
                case ParagraphRole.TableNote:
                    node.ForeColor = Color.DarkOrange;
                    break;
                case ParagraphRole.Normal:
                    node.ForeColor = Color.Gray;
                    node.NodeFont = new Font(treeViewOutline.Font, FontStyle.Regular);
                    break;
                default:
                    node.ForeColor = Color.Black;
                    break;
            }
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            OnConfirmFormatting?.Invoke(this, EventArgs.Empty);
        }
    }
}
