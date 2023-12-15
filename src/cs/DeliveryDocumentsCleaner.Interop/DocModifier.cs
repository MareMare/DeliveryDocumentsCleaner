// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DocModifier.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.Diagnostics;
using DeliveryDocumentsCleaner.Disposables;

namespace DeliveryDocumentsCleaner.Interop;

/// <summary>
/// ドキュメントの情報を更新する機能を提供します。
/// </summary>
public class DocModifier : DisposableBase, IDocumentModifier
{
    /// <summary>Dispose 可能なインスタンスのコンテナを表します。</summary>
    private readonly CompositeDisposable _disposables;

    /// <summary>
    /// <see cref="DocModifier" /> クラスの新しいインスタンスを初期化します。
    /// </summary>
    /// <param name="progress">進行状況の報告者。</param>
    public DocModifier(Progress<IDocumentItem> progress)
    {
        this._disposables = new CompositeDisposable();
        this.Progress = progress;
    }

    /// <summary>
    /// 進行状況の報告者を取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="Progress{IDocumentItem}" /> 型。
    /// <para>進行状況の報告者。既定値は <see langword="not null" /> です。</para>
    /// </value>
    public Progress<IDocumentItem> Progress { get; }

    /// <inheritdoc />
    public async ValueTask ExecuteToClearAsync(
        DocModifierSetting setting,
        IReadOnlyCollection<IDocumentItem>? documentItems)
    {
        if (documentItems == null || documentItems.Count == 0)
        {
            return;
        }

        var completedCount = 0;
        var totalCount = documentItems.Count;
        var chunkSize = (int)Math.Ceiling(totalCount / 5d);

        var chunkedItems = documentItems.Chunk(chunkSize);
        var tasks = chunkedItems
            .Select(items => Task.Run(() => ExecuteToClearItems(items)))
            .ToArray();
        await Task.WhenAll(tasks).ConfigureAwait(false);
        return;

        void ExecuteToClearItems(IEnumerable<IDocumentItem> items)
        {
            foreach (var item in items)
            {
                DocModifier.ExecuteToClearItem(setting, item);
                Interlocked.Increment(ref completedCount);

                this.Progress?.Report(item);
            }
        }
    }

    /// <inheritdoc />
    protected override void DisposeManagedInstances()
    {
    }

    /// <inheritdoc />
    protected override void DisposeUnmanagedInstances() => this._disposables.Dispose();

    private static void ExecuteToClearItem(DocModifierSetting setting, IDocumentItem documentItem)
    {
        Debug.Print($"{documentItem.FileName}");
        if (documentItem.DocKind == DocumentKind.Excel)
        {
            DocModifier.ExecuteToClearExcel(setting, documentItem);
        }
        else if (documentItem.DocKind == DocumentKind.Word)
        {
            DocModifier.ExecuteToClearWord(setting, documentItem);
        }
    }

    private static void ExecuteToClearExcel(DocModifierSetting setting, IDocumentItem documentItem)
    {
        using var modifier = new XlDocModifier();
        modifier.FilePath = documentItem.FilePath;
        modifier.OpenBook();
        modifier.ClearHighlightColor(setting.ExcelHighlightColor);
        modifier.ClearPersonalInfo();
        modifier.CloseBook();
    }

    private static void ExecuteToClearWord(DocModifierSetting setting, IDocumentItem documentItem)
    {
        using var modifier = new WdDocModifier();
        modifier.FilePath = documentItem.FilePath;
        modifier.OpenDoc();
        modifier.ClearHighlightColor(setting.WordColorIds?.ToArray());
        modifier.ClearPersonalInfo();
        modifier.CloseDoc();
    }
}
