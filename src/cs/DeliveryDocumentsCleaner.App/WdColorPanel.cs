// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WdColorPanel.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner.App;

/// <summary>
/// Word 用の蛍光ペンパネルを表します。
/// </summary>
public partial class WdColorPanel : UserControl
{
    /// <summary><see cref="WdColorControl"/> のコレクションを表します。</summary>
    private readonly List<WdColorControl> _wdColorControls;

    /// <summary>
    /// <see cref="WdColorPanel" /> クラスの新しいインスタンスを初期化します。
    /// </summary>
    public WdColorPanel()
    {
        this.InitializeComponent();
        this._wdColorControls = new()
        {
            this.wdColorControl1,
            this.wdColorControl2,
            this.wdColorControl3,
            this.wdColorControl4,
            this.wdColorControl5,
            this.wdColorControl6,
            this.wdColorControl7,
            this.wdColorControl8,
            this.wdColorControl9,
            this.wdColorControl10,
            this.wdColorControl11,
            this.wdColorControl12,
            this.wdColorControl13,
            this.wdColorControl14,
            this.wdColorControl15,
        };
    }

    /// <summary>
    /// 選択済みの蛍光ペン識別子のコレクションを取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="WordColorIdentity" /> 型。
    /// <para>選択済みの蛍光ペン識別子のコレクション。既定値は要素数 0 です。</para>
    /// </value>
    public IReadOnlyCollection<WordColorIdentity> WordColorIds
    {
        get => this._wdColorControls.Where(ctrl => ctrl.IsChecked).Select(ctrl => ctrl.WordColorId).ToArray();
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
    /// [すべてチェックON] がクリックされた場合に発生するイベントのイベントハンドラです。
    /// </summary>
    /// <param name="sender">イベントのソースを表す <see cref="object" />。</param>
    /// <param name="e">イベントデータを格納している <see cref="EventArgs" />。</param>
    private void CheckBoxToCheckAll_CheckedChanged(object sender, EventArgs e)
    {
        var isChecked = this.checkBoxToCheckAll.Checked;
        this._wdColorControls.ForEach(ctrl => ctrl.IsChecked = isChecked);
    }
}
