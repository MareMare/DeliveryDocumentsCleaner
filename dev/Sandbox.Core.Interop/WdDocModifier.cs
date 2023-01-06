// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WdDocModifier.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using Microsoft.Office.Interop.Word;
using Sandbox.Core.Disposables;

namespace Sandbox.Core.Interop;

/// <summary>
/// Word ドキュメントの情報を更新する機能を提供します。
/// </summary>
public class WdDocModifier : DisposableBase
{
    /// <summary>VBA での True 値を表します。</summary>
    private const int VbaTrueValue = -1;

    /// <summary>蛍光ペン色のインデックスのコレクションを表します。</summary>
    private readonly IReadOnlyDictionary<WordColorIdentity, WdColorIndex> _wdColorIndexes
        = new Dictionary<WordColorIdentity, WdColorIndex>
        {
            { WordColorIdentity.Yellow, WdColorIndex.wdYellow },
            { WordColorIdentity.BrightGreen, WdColorIndex.wdBrightGreen },
            { WordColorIdentity.Turquoise, WdColorIndex.wdTurquoise },
            { WordColorIdentity.Pink, WdColorIndex.wdPink },
            { WordColorIdentity.Blue, WdColorIndex.wdBlue },
            { WordColorIdentity.Red, WdColorIndex.wdRed },
            { WordColorIdentity.DarkBlue, WdColorIndex.wdDarkBlue },
            { WordColorIdentity.Teal, WdColorIndex.wdTeal },
            { WordColorIdentity.Green, WdColorIndex.wdGreen },
            { WordColorIdentity.Violet, WdColorIndex.wdViolet },
            { WordColorIdentity.DarkRed, WdColorIndex.wdDarkRed },
            { WordColorIdentity.DarkYellow, WdColorIndex.wdDarkYellow },
            { WordColorIdentity.Gray50, WdColorIndex.wdGray50 },
            { WordColorIdentity.Gray25, WdColorIndex.wdGray25 },
            { WordColorIdentity.Black, WdColorIndex.wdBlack },
        };

    /// <summary>Dispose 可能なインスタンスのコンテナを表します。</summary>
    private readonly CompositeDisposable _disposables;

    /// <summary><see cref="IComWrapper{Application}" /> インスタンスを表します。</summary>
    private readonly IComWrapper<Application> _wdApp;

    /// <summary><see cref="IComWrapper{Document}" /> インスタンスを表します。</summary>
    private IComWrapper<Document>? _wdDoc;

    /// <summary>Word ファイルパスを表します。</summary>
    private string _filePath = default!;

    /// <summary>変更された場合は <c>true</c>。それ以外は <c>false</c>。</summary>
    private bool _isDirty;

    /// <summary>
    /// <see cref="WdDocModifier" /> クラスの新しいインスタンスを初期化します。
    /// </summary>
    public WdDocModifier()
    {
        this._disposables = new CompositeDisposable();
        this._wdApp = new Application().ToComWrap();
        this._wdApp.SetBusy();
    }

    /// <summary>
    /// Word ファイルのフルパスを取得または設定します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="string" /> 型。
    /// <para>Word ファイルのフルパス。既定値は <see langword="null" /> です。</para>
    /// </value>
    public string FilePath
    {
        get => this._filePath;
        set
        {
            this.ThrowIfAlreadyOpen();
            this._filePath = value;
        }
    }

    /// <summary>
    /// Word ファイルをオープンします。
    /// </summary>
    public void OpenDoc()
    {
        this.ThrowIfAlreadyOpen();

        this._isDirty = false;
        using var documents = this._wdApp.Target.Documents.ToComWrap();
        this._wdDoc = documents.Target.Open(this.FilePath).ToComWrap();
    }

    /// <summary>
    /// Word ファイルをクローズします。
    /// </summary>
    public void CloseDoc()
    {
        this.SaveDocIfNecessary();

        this._wdDoc?.Target.Close(WdSaveOptions.wdDoNotSaveChanges);
        this._wdDoc?.Dispose();
        this._wdDoc = null;
    }

    /// <summary>
    /// 必要に応じて Word ファイルを上書き保存します。
    /// </summary>
    /// <param name="force">強制的に保存する場合は <c>true</c>。それ以外は <c>false</c>。</param>
    public void SaveDocIfNecessary(bool force = false)
    {
        if (this._isDirty || force)
        {
            this._wdDoc?.Target.Save();
        }

        this._isDirty = false;
    }

    /// <summary>
    /// 指定された蛍光ペン識別子に対応する蛍光ペンをクリアします。
    /// </summary>
    /// <param name="wdColorIdsToClear">蛍光ペン識別子のコレクション。</param>
    public void ClearHighlightColor(IEnumerable<WordColorIdentity>? wdColorIdsToClear = null)
    {
        if (this._wdDoc is null)
        {
            return;
        }

        var colorIndexesToClear = wdColorIdsToClear?
                                      .Join(
                                          this._wdColorIndexes,
                                          outer => outer,
                                          inner => inner.Key,
                                          (_, inner) => inner.Value)
                                  ?? this._wdColorIndexes.Values;
        var mapping = colorIndexesToClear.ToDictionary(_ => _);

        using var wdRange = this._wdDoc.Target.Range(0, 0).ToComWrap();
        using var wdFind = wdRange.Target.Find.ToComWrap();
        using var wdReplacement = wdFind.Target.Replacement.ToComWrap();

        wdFind.Target.ClearFormatting();
        wdReplacement.Target.ClearFormatting();

        wdFind.Target.Highlight = WdDocModifier.VbaTrueValue;

        wdFind.Target.Text = string.Empty;
        wdReplacement.Target.Text = string.Empty;

        wdFind.Target.Forward = true;
        wdFind.Target.Wrap = WdFindWrap.wdFindStop;
        wdFind.Target.Format = true;
        wdFind.Target.MatchCase = false;
        wdFind.Target.MatchWholeWord = false;
        wdFind.Target.MatchByte = false;
        wdFind.Target.MatchAllWordForms = false;
        wdFind.Target.MatchSoundsLike = false;
        wdFind.Target.MatchWildcards = false;
        wdFind.Target.MatchFuzzy = false;

        while (true)
        {
            if (wdFind.Target.Execute(Replace: WdReplace.wdReplaceNone) != true)
            {
                break;
            }

            var targetColorIndex = wdRange.Target.HighlightColorIndex;
            if (mapping.ContainsKey(targetColorIndex))
            {
                wdRange.Target.HighlightColorIndex = WdColorIndex.wdNoHighlight;
            }

            wdRange.Target.Collapse(WdCollapseDirection.wdCollapseEnd);
        }
    }

    /// <summary>
    /// 個人情報をクリアします。
    /// </summary>
    public void ClearPersonalInfo()
    {
        if (this._wdDoc is not null)
        {
            this._wdDoc.Target.RemoveDocumentInformation(WdRemoveDocInfoType.wdRDIAll);

            // RemovePersonalInformation = True としてから保存することで個人情報を削除します。
            this._wdDoc.Target.RemovePersonalInformation = true;
            this.SaveDocIfNecessary(true);
        }
    }

    /// <inheritdoc />
    protected override void DisposeManagedInstances()
    {
        this.SaveDocIfNecessary();

        this._wdDoc?.Target.Close(WdSaveOptions.wdDoNotSaveChanges);
        this._wdApp.Target.Quit(WdSaveOptions.wdDoNotSaveChanges);
    }

    /// <inheritdoc />
    protected override void DisposeUnmanagedInstances()
    {
        this._disposables.Dispose();
        this._wdDoc?.Dispose();
        this._wdDoc = null;
        this._wdApp.Dispose();
    }

    /// <summary>
    /// 既にオープン済みであれば例外を発生させます。
    /// </summary>
    private void ThrowIfAlreadyOpen()
    {
        if (this._wdDoc is not null)
        {
            throw new InvalidOperationException($"既にオープン済みです。Path=\"{this._filePath}\"");
        }
    }
}
