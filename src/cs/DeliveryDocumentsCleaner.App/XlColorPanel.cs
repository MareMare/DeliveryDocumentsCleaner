// --------------------------------------------------------------------------------------------------------------------
// <copyright file="XlColorPanel.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using Cyotek.Windows.Forms;

namespace DeliveryDocumentsCleaner.App;

/// <summary>
/// Excel 用の塗りつぶし色パネルを表します。
/// </summary>
public partial class XlColorPanel : UserControl
{
    /// <summary>塗りつぶし色を表します。</summary>
    private Color? _xlColor;

    /// <summary>
    /// <see cref="XlColorPanel" /> クラスの新しいインスタンスを初期化します。
    /// </summary>
    public XlColorPanel()
    {
        this.InitializeComponent();
    }

    /// <summary>
    /// 塗りつぶし色を取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="Color" /> 型。
    /// <para>塗りつぶし色。既定値は <see langword="null" /> です。</para>
    /// </value>
    public Color? XlColor
    {
        get => this._xlColor;
        private set
        {
            this.panelOfSelectedXlColor.BackColor = value ?? SystemColors.Control;
            this.labelOfSelectedXlColor.Text = value == null ? "---" : ColorTranslator.ToHtml(value.Value);
            this._xlColor = value;
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

    /// <summary>
    /// […] ボタンがクリックされた場合に発生するイベントのイベントハンドラです。
    /// </summary>
    /// <param name="sender">イベントのソースを表す <see cref="object" />。</param>
    /// <param name="e">イベントデータを格納している <see cref="EventArgs" />。</param>
    private void ButtonToSelectXlColor_Click(object sender, EventArgs e)
    {
        using var dialog = new ColorPickerDialog();
        if (this.XlColor.HasValue)
        {
            dialog.Color = this.XlColor.Value;
        }

        dialog.ShowLoad = false;
        dialog.ShowSave = false;
        this.XlColor = dialog.ShowDialog(this.ParentForm) == DialogResult.OK ? dialog.Color : null;
    }
}
