// --------------------------------------------------------------------------------------------------------------------
// <copyright file="MainForm.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using DeliveryDocumentsCleaner.Disposables;
using DeliveryDocumentsCleaner.Interop;

namespace DeliveryDocumentsCleaner.App;

/// <summary>
/// メインフォームを表します。
/// </summary>
public partial class MainForm : Form
{
    /// <summary>
    /// <see cref="MainForm" /> クラスの新しいインスタンスを生成します。
    /// </summary>
    public MainForm()
    {
        this.InitializeComponent();
        this.InitializeStatusBar();
    }

    /// <summary>
    ///  Clean up any resources being used.
    /// </summary>
    /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            this.components?.Dispose();
        }

        base.Dispose(disposing);
    }

    /// <summary>
    /// ステータスバーを初期化します。
    /// </summary>
    /// <param name="totalCount">最大件数。</param>
    private void InitializeStatusBar(int? totalCount = null)
    {
        this.toolStripProgressBar1.Minimum = 0;
        this.toolStripProgressBar1.Maximum = totalCount ?? 100;
        this.toolStripProgressBar1.Visible = false;
        this.toolStripStatusLabel1.Text = null;
    }

    /// <summary>
    /// ステータスバーに情報を設定します。
    /// </summary>
    /// <param name="message">メッセージ。</param>
    /// <param name="completedCount">完了した件数。</param>
    private void SetStatusBar(string message, int completedCount)
    {
        var value = Math.Min(this.toolStripProgressBar1.Maximum, completedCount);
        this.toolStripProgressBar1.Value = value;
        this.toolStripProgressBar1.Visible = value > 0;
        this.toolStripStatusLabel1.Text = message;
    }

    /// <summary>
    /// 非活性化状態の範囲を表します。
    /// </summary>
    /// <returns>非活性化状態の範囲。</returns>
    private IDisposable UseBusyScope()
    {
        this.SetEnabled(false);
        return Disposable.Create(() => this.SetEnabled(true));
    }

    /// <summary>
    /// 活性化状態を設定します。
    /// </summary>
    /// <param name="isEnabled">活性化とする場合は <c>true</c>。それ以外は <c>false</c>。</param>
    private void SetEnabled(bool isEnabled) => this.tableLayoutPanel1.Enabled = isEnabled;

    /// <summary>
    /// ドキュメント一覧へ設定します。
    /// </summary>
    /// <param name="folder">選択されたフォルダパス。</param>
    private void SetFilesListFrom(string? folder = null)
    {
        this.textBoxOfFolder.Text = folder ?? string.Empty;
        this.listBoxOfFiles.Items.Clear();
        if (Directory.Exists(folder))
        {
            var files = Directory.EnumerateFiles(folder, "*.*", SearchOption.TopDirectoryOnly)
                .Select(path => new DisplayItem(path))
                .Where(item => !item.FileName.StartsWith('~'))
                .Where(item => item.DocKind.HasValue)
                .ToList();
            files.ForEach(item => this.listBoxOfFiles.Items.Add(item));
        }
    }

    private async ValueTask ExecuteToClearAsync(IReadOnlyCollection<IDocumentItem> documentItems)
    {
        var setting = new DocModifierSetting
        {
            ExcelHighlightColor = this.xlColorPanel1.XlColor,
            WordColorIds = this.wdColorPanel1.WordColorIds,
        };
        using var modifier = new DocModifier(new Progress<IDocumentItem>());

        using var completedSignal = new ManualResetEventSlim();
        var closureSignal = completedSignal;

        var totalCount = documentItems.Count;
        var countOfCalled = 0;
        var gate = new object();
        modifier.Progress.ProgressChanged += (_, item) =>
        {
            var completedCount = Interlocked.Increment(ref countOfCalled);
            lock (gate)
            {
                this.SetStatusBar($"{completedCount} {item.FileName}", completedCount);
                if (completedCount >= totalCount)
                {
                    closureSignal.Set();
                }
            }
        };

        await modifier.ExecuteToClearAsync(setting, documentItems).ConfigureAwait(true);
        completedSignal.Wait(TimeSpan.FromSeconds(totalCount * 3));
    }

    /// <summary>
    /// […] ボタンがクリックされた場合に発生するイベントのイベントハンドラです。
    /// </summary>
    /// <param name="sender">イベントのソースを表す <see cref="object" />。</param>
    /// <param name="e">イベントデータを格納している <see cref="EventArgs" />。</param>
    private void ButtonToSelectFolder_Click(object sender, EventArgs e)
    {
        using var dialog = new FolderBrowserDialog();
        dialog.ShowNewFolderButton = false;
        this.SetFilesListFrom(dialog.ShowDialog(this) == DialogResult.OK ? dialog.SelectedPath : null);
    }

    /// <summary>
    /// [クリア実行] ボタンがクリックされた場合に発生するイベントのイベントハンドラです。
    /// </summary>
    /// <param name="sender">イベントのソースを表す <see cref="object" />。</param>
    /// <param name="e">イベントデータを格納している <see cref="EventArgs" />。</param>
    private async void ButtonToExecuteClear_Click(object sender, EventArgs e)
    {
        using var scope = this.UseBusyScope();

        var files = this.listBoxOfFiles.SelectedItems.OfType<DisplayItem>().ToArray();
        this.InitializeStatusBar(files.Length);
        await this.ExecuteToClearAsync(files).ConfigureAwait(true);
    }

    /// <summary>
    /// 一覧項目を表します。
    /// </summary>
    private class DisplayItem : IDocumentItem
    {
        /// <summary>
        /// <see cref="DisplayItem" /> クラスの新しいインスタンスを初期化します。
        /// </summary>
        /// <param name="filepath">フルパス。</param>
        public DisplayItem(string filepath)
        {
            this.FilePath = filepath;
            this.FileName = Path.GetFileName(filepath);
            this.DocKind = DisplayItem.ResolveDocKind(filepath);
        }

        /// <inheritdoc />
        public DocumentKind? DocKind { get; }

        /// <inheritdoc />
        public string FilePath { get; }

        /// <inheritdoc />
        public string FileName { get; }

        /// <inheritdoc />
        public override string ToString() => this.FileName;

        /// <summary>
        /// ドキュメント種別を解決します。
        /// </summary>
        /// <param name="path">ファイルパス。</param>
        /// <returns>ドキュメント種別。</returns>
        private static DocumentKind? ResolveDocKind(string path)
        {
            var ext = Path.GetExtension(path).ToUpperInvariant();
            return ext switch
            {
                @".DOCX" => DocumentKind.Word,
                @".DOCM" => DocumentKind.Word,
                @".XLSX" => DocumentKind.Excel,
                @".XLSM" => DocumentKind.Excel,
                _ => null,
            };
        }
    }
}
