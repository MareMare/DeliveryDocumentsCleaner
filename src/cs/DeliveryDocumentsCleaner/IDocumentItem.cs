// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IDocumentItem.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner;

/// <summary>
/// ドキュメント項目のインターフェイスを表します。
/// </summary>
public interface IDocumentItem
{
    /// <summary>
    /// ドキュメント種別を取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="DocumentKind" /> 型。
    /// <para>ドキュメント種別。既定値は <see langword="null" /> です。</para>
    /// </value>
    DocumentKind? DocKind { get; }

    /// <summary>
    /// フルパスを取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="string" /> 型。
    /// <para>フルパス。既定値は <see langword="null" /> です。</para>
    /// </value>
    string FilePath { get; }

    /// <summary>
    /// ファイル名を取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="string" /> 型。
    /// <para>ファイル名。既定値は <see langword="null" /> です。</para>
    /// </value>
    string FileName { get; }
}
