// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WdColorControl.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.ComponentModel;
using DeliveryDocumentsCleaner.Interop;

namespace DeliveryDocumentsCleaner.App;

/// <summary>
/// Word 用の蛍光ペンを表します。
/// </summary>
internal partial class WdColorControl : UserControl
{
    /// <summary>蛍光ペン識別子に対応する色のマッピングを表します。</summary>
    private readonly Dictionary<WordColorIdentity, Color> _mapping = new()
    {
        // NOTE: https://learn.microsoft.com/en-us/office/vba/api/word.wdcolor
        { WordColorIdentity.Yellow, Color.FromArgb(255, 255, 0) },
        { WordColorIdentity.BrightGreen, Color.FromArgb(0, 255, 0) },
        { WordColorIdentity.Turquoise, Color.FromArgb(0, 255, 255) },
        { WordColorIdentity.Pink, Color.FromArgb(255, 0, 255) },
        { WordColorIdentity.Blue, Color.FromArgb(0, 0, 255) },
        { WordColorIdentity.Red, Color.FromArgb(255, 0, 0) },
        { WordColorIdentity.DarkBlue, Color.FromArgb(0, 0, 128) },
        { WordColorIdentity.Teal, Color.FromArgb(0, 128, 128) },
        { WordColorIdentity.Green, Color.FromArgb(0, 128, 0) },
        { WordColorIdentity.Violet, Color.FromArgb(128, 0, 128) },
        { WordColorIdentity.DarkRed, Color.FromArgb(128, 0, 0) },
        { WordColorIdentity.DarkYellow, Color.FromArgb(128, 128, 0) },
        { WordColorIdentity.Gray50, Color.FromArgb(128, 128, 128) },
        { WordColorIdentity.Gray25, Color.FromArgb(192, 192, 192) },
        { WordColorIdentity.Black, Color.FromArgb(0, 0, 0) },
    };

    /// <summary>蛍光ペン識別子を表します。</summary>
    private WordColorIdentity _wordColorId;

    /// <summary>
    /// <see cref="WdColorControl" /> クラスの新しいインスタンスを初期化します。
    /// </summary>
    public WdColorControl()
    {
        this.InitializeComponent();
        this.WordColorId = WordColorIdentity.Yellow;
    }

    /// <summary>
    /// 蛍光ペンを取得または設定します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="bool" /> 型。
    /// <para>蛍光ペン。既定値は <see langword="false" /> です。</para>
    /// </value>
    public bool IsChecked
    {
        get => this.checkBox1.Checked;
        set => this.checkBox1.Checked = value;
    }

    /// <summary>
    /// 蛍光ペン識別子を取得または設定します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="WordColorIdentity" /> 型。
    /// <para>蛍光ペン識別子。既定値は <see langword="WordColorIdentity.Yellow" /> です。</para>
    /// </value>
    public WordColorIdentity WordColorId
    {
        get => this._wordColorId;
        set
        {
            this.BackColor = this._mapping.TryGetValue(value, out var color) ? color : Color.White;
            this._wordColorId = value;
        }
    }

    /// <summary>
    /// 使用中のリソースをすべてクリーンアップします。
    /// </summary>
    /// <param name="disposing">マネージド リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            if (this.components is not null)
            {
                this.components.Dispose();
            }
        }

        base.Dispose(disposing);
    }
}
