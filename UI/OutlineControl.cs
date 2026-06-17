using System;
using System.Collections.Generic;
using System.Windows.Forms;
using WordTool.Models;

namespace WordTool.UI
{
    public partial class OutlineControl : UserControl
    {
        private List<AnalyzedParagraph> _analyzedParagraphs;
        public event EventHandler OnConfirmFormatting;

        public OutlineControl()
        {
            InitializeComponent();
        }

        public void LoadData(List<AnalyzedParagraph> paragraphs)
        {
            _analyzedParagraphs = paragraphs;
            treeViewOutline.Nodes.Clear();

            // 简单展示，将所有非Normal的内容列出来供用户确认
            foreach (var p in paragraphs)
            {
                if (p.Role != ParagraphRole.Normal)
                {
                    TreeNode node = new TreeNode($"[{p.Role}] {p.TextContent}");
                    node.Tag = p;
                    treeViewOutline.Nodes.Add(node);
                }
            }
        }

        private void btnConfirm_Click(object sender, EventArgs e)
        {
            OnConfirmFormatting?.Invoke(this, EventArgs.Empty);
        }

        // 可以添加右键菜单让用户修改 TreeNode 对应的 ParagraphRole
    }
}
