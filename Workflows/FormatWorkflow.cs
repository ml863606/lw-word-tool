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
        private FormattingTemplate _template;

        // 线程同步与控制信号
        private System.Threading.CancellationTokenSource _cts;
        private System.Threading.ManualResetEventSlim _pauseEvent = new System.Threading.ManualResetEventSlim(true);
        private bool _isPaused = false;

        public bool IsPaused => _isPaused;
        public Action YieldUICallback { get; set; }

        public FormattingTemplate Template
        {
            get { return _template ?? (_template = FormattingTemplate.GetDefaultThesisTemplate()); }
            set { _template = value; }
        }

        public FormatWorkflow(Word.Document doc, Action<string> logger = null, Action<int> progress = null)
        {
            _doc = doc;
            _logger = logger ?? (msg => { });
            _progress = progress ?? (pct => { });
        }

        public void Pause()
        {
            _isPaused = true;
            _pauseEvent.Reset();
            _logger("【暂停】正在挂起排版进程...");
        }

        public void Resume()
        {
            _isPaused = false;
            _pauseEvent.Set();
            _logger("【恢复】排版进程已恢复。");
        }

        public void Cancel()
        {
            _cts?.Cancel();
            _pauseEvent.Set(); // 如果处于暂停状态，先唤醒线程以触发取消异常
            _logger("【中止】正在发送中止信号...");
        }

        public void ResetControlStates()
        {
            _cts = new System.Threading.CancellationTokenSource();
            _pauseEvent.Set();
            _isPaused = false;
        }

        public void CheckPauseAndCancel()
        {
            if (_cts != null && _cts.Token.IsCancellationRequested)
            {
                _cts.Token.ThrowIfCancellationRequested();
            }
            
            // 泵送消息至 UI 线程，以确保暂停/中止按钮点击能被处理
            try
            {
                YieldUICallback?.Invoke();
            }
            catch { }

            _pauseEvent.Wait();
        }

        private void OptimizeWordUI(bool optimize)
        {
            try
            {
                _doc.Application.ScreenUpdating = !optimize;
                try
                {
                    _doc.ShowSpellingErrors = !optimize;
                    _doc.ShowGrammaticalErrors = !optimize;
                }
                catch { }
            }
            catch { }
        }

        public void RunAllSteps(Func<List<AnalyzedParagraph>, bool> confirmCallback = null)
        {
            try
            {
                _doc.Application.UndoRecord.StartCustomRecord("一键全自动排版");
                ResetControlStates();

                _progress(5);
                RunCleanData();

                _progress(25);
                RunStyleRebuild();

                _progress(45);
                var paragraphs = RunAnalysis();

                if (confirmCallback != null)
                {
                    CheckPauseAndCancel();
                    _logger("正在等待大纲确认...");
                    bool confirmed = confirmCallback(paragraphs);
                    if (!confirmed)
                    {
                        _logger("【取消】用户取消了排版操作。");
                        _progress(0);
                        return;
                    }
                }

                _progress(65);
                RunFormatting(paragraphs);

                _progress(80);
                RunMediaFormatting();

                _progress(95);
                RunLayout();

                _progress(100);
                _logger("【成功】所有排版步骤已顺利完成！");
            }
            catch (OperationCanceledException)
            {
                _logger("【终止】一键排版任务已被用户终止。");
                _progress(0);
            }
            catch (Exception ex)
            {
                _logger($"【错误】排版过程中发生异常: {ex.Message}");
            }
            finally
            {
                _doc.Application.UndoRecord.EndCustomRecord();
                OptimizeWordUI(false); // 确保恢复屏幕刷新
            }
        }

        public void RunCleanData()
        {
            _logger("开始执行 [清理脏数据]...");
            OptimizeWordUI(true);
            try
            {
                CheckPauseAndCancel();
                var cleaner = new Cleaners.DocumentCleaner();
                cleaner.Clean(_doc, _logger, checkStatus: CheckPauseAndCancel);
                _logger("[清理脏数据] 完成。");
            }
            catch (OperationCanceledException)
            {
                _logger("【终止】[清理脏数据] 被用户中止。");
                throw;
            }
            finally
            {
                OptimizeWordUI(false);
            }
        }

        public void RunStyleRebuild()
        {
            _logger("开始执行 [重置标准样式]...");
            OptimizeWordUI(true);
            try
            {
                CheckPauseAndCancel();
                var styleManager = new Styles.StyleManager();
                styleManager.RebuildStyles(_doc, Template, _logger, checkStatus: CheckPauseAndCancel);
                _logger("[重置标准样式] 完成。");
            }
            catch (OperationCanceledException)
            {
                _logger("【终止】[重置标准样式] 被用户中止。");
                throw;
            }
            finally
            {
                OptimizeWordUI(false);
            }
        }

        public List<AnalyzedParagraph> RunAnalysis()
        {
            _logger("开始执行 [解析段落角色]...");
            OptimizeWordUI(true);
            try
            {
                CheckPauseAndCancel();
                var analyzer = new Analyzers.ParagraphAnalyzer();
                var result = analyzer.Analyze(_doc, checkStatus: CheckPauseAndCancel);
                _logger($"[解析段落角色] 完成，共识别到 {result.Count} 个段落。");
                return result;
            }
            catch (OperationCanceledException)
            {
                _logger("【终止】[解析段落角色] 被用户中止。");
                throw;
            }
            finally
            {
                OptimizeWordUI(false);
            }
        }

        public void RunFormatting(List<AnalyzedParagraph> paragraphs)
        {
            if (paragraphs == null || paragraphs.Count == 0)
            {
                _logger("【警告】解析结果为空，跳过正文排版。");
                return;
            }
            _logger("开始执行 [应用正文排版]...");
            OptimizeWordUI(true);
            try
            {
                CheckPauseAndCancel();
                var formatter = new Formatters.DocumentFormatter();
                formatter.ApplyFormatting(paragraphs, _logger, checkStatus: CheckPauseAndCancel);
                _logger("[应用正文排版] 完成。");
            }
            catch (OperationCanceledException)
            {
                _logger("【终止】[应用正文排版] 被用户中止。");
                throw;
            }
            finally
            {
                OptimizeWordUI(false);
            }
        }

        public void RunMediaFormatting()
        {
            _logger("开始执行 [排版图片与表格]...");
            OptimizeWordUI(true);
            try
            {
                CheckPauseAndCancel();
                var mediaFormatter = new Formatters.MediaFormatter();
                mediaFormatter.FormatImages(_doc, Template, _logger, checkStatus: CheckPauseAndCancel);
                mediaFormatter.FormatTables(_doc, Template, _logger, checkStatus: CheckPauseAndCancel);
                _logger("[排版图片与表格] 完成。");
            }
            catch (OperationCanceledException)
            {
                _logger("【终止】[排版图片与表格] 被用户中止。");
                throw;
            }
            finally
            {
                OptimizeWordUI(false);
            }
        }

        public void RunLayout()
        {
            _logger("开始执行 [生成或更新目录]...");
            OptimizeWordUI(true);
            try
            {
                CheckPauseAndCancel();
                var layoutManager = new Layout.LayoutManager();
                layoutManager.UpdateOrInsertTOC(_doc, Template, _logger, checkStatus: CheckPauseAndCancel);
                _logger("[生成或更新目录] 完成。");
            }
            catch (OperationCanceledException)
            {
                _logger("【终止】[生成或更新目录] 被用户中止。");
                throw;
            }
            finally
            {
                OptimizeWordUI(false);
            }
        }
    }
}
