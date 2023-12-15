// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DocModifierSetting.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.Drawing;

namespace DeliveryDocumentsCleaner;

/// <summary>
/// ドキュメントの情報を更新する設定を表します。
/// </summary>
public class DocModifierSetting
{
    /// <summary>
    /// 塗りつぶし色を取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="Color" /> 型。
    /// <para>蛍光ペン識別子のコレクション。既定値は <see langword="null" /> です。</para>
    /// </value>
    public Color? ExcelHighlightColor { get; init; }

    /// <summary>
    /// 蛍光ペン識別子のコレクションを取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="WordColorIdentity" /> 型。
    /// <para>蛍光ペン識別子のコレクション。既定値は <see langword="null" /> です。</para>
    /// </value>
    public IReadOnlyCollection<WordColorIdentity>? WordColorIds { get; init; }
}
