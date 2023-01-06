// --------------------------------------------------------------------------------------------------------------------
// <copyright file="XlDocModifier.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.Drawing;
using Microsoft.Office.Interop.Excel;
using Sandbox.Core.Disposables;

namespace Sandbox.Core;

/// <summary>
/// Excel ドキュメントの情報を更新する機能を提供します。
/// </summary>
public class XlDocModifier : DisposableBase
{
    /// <summary>Dispose 可能なインスタンスのコンテナを表します。</summary>
    private readonly CompositeDisposable _disposables;

    /// <summary><see cref="IComWrapper{Application}" /> インスタンスを表します。</summary>
    private readonly IComWrapper<Application> _xlApp;

    /// <summary><see cref="IComWrapper{Workbook}" /> インスタンスを表します。</summary>
    private IComWrapper<Workbook>? _xlBook;

    /// <summary>Excel ファイルパスを表します。</summary>
    private string _filePath = default!;

    /// <summary>変更された場合は <c>true</c>。それ以外は <c>false</c>。</summary>
    private bool _isDirty;

    /// <summary>
    /// <see cref="XlDocModifier" /> クラスの新しいインスタンスを初期化します。
    /// </summary>
    public XlDocModifier()
    {
        this._disposables = new CompositeDisposable();
        this._xlApp = new Application().ToComWrap();
        this._xlApp.SetBusy();
    }

    /// <summary>
    /// Excel ファイルのフルパスを取得または設定します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="string" /> 型。
    /// <para>Excel ファイルのフルパス。既定値は <see langword="null" /> です。</para>
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
    /// Excel ファイルをオープンします。
    /// </summary>
    public void OpenBook()
    {
        this.ThrowIfAlreadyOpen();

        this._isDirty = false;
        using var workbooks = this._xlApp.Target.Workbooks.ToComWrap();
        this._xlBook = workbooks.Target.Open(Filename: this.FilePath).ToComWrap();
    }

    /// <summary>
    /// Excel ファイルをクローズします。
    /// </summary>
    public void CloseBook()
    {
        this.SaveBookIfNecessary();

        this._xlBook?.Target.Close(SaveChanges: false);
        this._xlBook?.Dispose();
        this._xlBook = null;
    }

    /// <summary>
    /// 必要に応じて Excel ファイルを上書き保存します。
    /// </summary>
    /// <param name="force">強制的に保存する場合は <c>true</c>。それ以外は <c>false</c>。</param>
    public void SaveBookIfNecessary(bool force = false)
    {
        if (this._isDirty || force)
        {
            this._xlBook?.Target.Save();
        }

        this._isDirty = false;
    }

    /// <summary>
    /// 指定された色のハイライトをクリアします。
    /// </summary>
    /// <param name="colorToClear">クリアするハイライトの色。</param>
    public void ClearHighlightColor(Color colorToClear) => this.ClearHighlightColor(ColorTranslator.ToOle(colorToClear));

    /// <summary>
    /// 指定された色コードのハイライトをクリアします。
    /// </summary>
    /// <param name="fullColorCode">クリアするハイライトの色コード。</param>
    public void ClearHighlightColor(long fullColorCode)
    {
        using var xlFindFormat = this._xlApp.Target.FindFormat.ToComWrap();
        using var xlReplaceFormat = this._xlApp.Target.ReplaceFormat.ToComWrap();
        using var xlFindFormatInterior = xlFindFormat.Target.Interior.ToComWrap();
        using var xlReplaceFormatInterior = xlReplaceFormat.Target.Interior.ToComWrap();
        using var xlCells = this._xlApp.Target.Cells.ToComWrap();

        xlFindFormatInterior.Target.PatternColorIndex = Constants.xlAutomatic;
        xlFindFormatInterior.Target.Color = fullColorCode;
        xlFindFormatInterior.Target.TintAndShade = 0;
        xlFindFormatInterior.Target.PatternTintAndShade = 0;

        xlReplaceFormatInterior.Target.Pattern = Constants.xlNone;
        xlReplaceFormatInterior.Target.TintAndShade = 0;
        xlReplaceFormatInterior.Target.PatternTintAndShade = 0;

        xlCells.Target.Replace2(
            What: string.Empty,
            Replacement: string.Empty,
            LookAt: XlLookAt.xlPart,
            SearchOrder: XlSearchOrder.xlByRows,
            MatchCase: false,
            SearchFormat: true,
            ReplaceFormat: true,
            FormulaVersion: XlFormulaVersion.xlReplaceFormula2);
        this._isDirty = true;
    }

    /// <summary>
    /// 個人情報をクリアします。
    /// </summary>
    public void ClearPersonalInfo()
    {
        if (this._xlBook is not null)
        {
            this._xlBook.Target.RemoveDocumentInformation(XlRemoveDocInfoType.xlRDIAll);

            // RemovePersonalInformation = True としてから保存することで個人情報を削除します。
            this._xlBook.Target.RemovePersonalInformation = true;
            this.SaveBookIfNecessary(force: true);
        }
    }

    /// <inheritdoc />
    protected override void DisposeManagedInstances()
    {
        this.SaveBookIfNecessary();

        this._xlBook?.Target.Close(SaveChanges: false);

        this._xlApp.SetBusy(false);
        this._xlApp.Target.Quit();
    }

    /// <inheritdoc />
    protected override void DisposeUnmanagedInstances()
    {
        this._disposables.Dispose();
        this._xlBook?.Dispose();
        this._xlBook = null;
        this._xlApp.Dispose();
    }

    /// <summary>
    /// 既にオープン済みであれば例外を発生させます。
    /// </summary>
    private void ThrowIfAlreadyOpen()
    {
        if (this._xlBook is not null)
        {
            throw new InvalidOperationException($"既にオープン済みです。Path=\"{this._filePath}\"");
        }
    }
}
