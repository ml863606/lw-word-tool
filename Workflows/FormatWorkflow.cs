using System;
using System.Collections.Generic;
using WordTool.Models;
using Word = Microsoft.Office.Interop.Word;

namespace WordTool.Workflows
{
    public class FormatWorkflow
    {
        private Word.Document _doc;
        private Action<string> _logger;
        private Action<int> _progress;

        public FormatWorkflow(Word.Document doc, Action<string> logger = null, Action<int> progress = null)
        {
            _doc = doc;
            _logger = logger ?? (msg => { });
            _progress = progress ?? (pct => { });
        }

        public void RunAllSteps(Func<List<AnalyzedParagraph>, bool> confirmCallback = null)
        {
            try
            {
                _doc.Application.UndoRecord.StartCustomRecord("一键全自动排版");

                _progress(10);
                RunCleanData();

                _progress(30);
                RunStyleRebuild();

                _progress(50);
                var paragraphs = RunAnalysis();

                if (confirmCallback != null)
                {
                    _logger("正在等待大纲确认...");
                    bool confirmed = confirmCallback(paragraphs);
                    if (!confirmed)
                    {
                        _logger("【取消】用户取消了排版操作。");
                        _progress(0);
                        return;
                    }
                }

                _progress(70);
                RunFormatting(paragraphs);

                _progress(85);
                RunMediaFormatting();

                _progress(95);
                RunLayout();

                _progress(100);
                _logger("【成功】所有排版步骤已顺利完成！");
            }
            catch (Exception ex)
            {
                _logger($"【错误】排版过程中发生异常: {ex.Message}");
            }
            finally
            {
                _doc.Application.UndoRecord.EndCustomRecord();
            }
        }

        public void RunCleanData()
        {
            _logger("开始执行 [清理脏数据]...");
            var cleaner = new Cleaners.DocumentCleaner();
            cleaner.Clean(_doc, _logger);
            _logger("[清理脏数据] 完成。");
        }

        public void RunStyleRebuild()
        {
            _logger("开始执行 [重置标准样式]...");
            var styleManager = new Styles.StyleManager();
            styleManager.RebuildStyles(_doc, _logger);
            _logger("[重置标准样式] 完成。");
        }

        public List<AnalyzedParagraph> RunAnalysis()
        {
            _logger("开始执行 [解析段落角色]...");
            var analyzer = new Analyzers.ParagraphAnalyzer();
            var result = analyzer.Analyze(_doc);
            _logger($"[解析段落角色] 完成，共识别到 {result.Count} 个段落。");
            return result;
        }

        public void RunFormatting(List<AnalyzedParagraph> paragraphs)
        {
            if (paragraphs == null || paragraphs.Count == 0)
            {
                _logger("【警告】解析结果为空，跳过正文排版。");
                return;
            }
            _logger("开始执行 [应用正文排版]...");
            var formatter = new Formatters.DocumentFormatter();
            formatter.ApplyFormatting(paragraphs, _logger);
            _logger("[应用正文排版] 完成。");
        }

        public void RunMediaFormatting()
        {
            _logger("开始执行 [排版图片与表格]...");
            var mediaFormatter = new Formatters.MediaFormatter();
            mediaFormatter.FormatImages(_doc, _logger);
            mediaFormatter.FormatTables(_doc, _logger);
            _logger("[排版图片与表格] 完成。");
        }

        public void RunLayout()
        {
            _logger("开始执行 [生成或更新目录]...");
            var layoutManager = new Layout.LayoutManager();
            layoutManager.UpdateOrInsertTOC(_doc, _logger);
            _logger("[生成或更新目录] 完成。");
        }
    }
}
