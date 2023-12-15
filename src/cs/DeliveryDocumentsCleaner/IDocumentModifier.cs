// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IDocumentModifier.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner;

/// <summary>
/// ドキュメントの情報を更新するインターフェイスを表します。
/// </summary>
public interface IDocumentModifier : IDisposable
{
    /// <summary>
    /// 非同期操作として、クリアを実行します。
    /// </summary>
    /// <param name="setting">ドキュメントの情報を更新する設定。</param>
    /// <param name="documentItems">ドキュメント項目のコレクション。</param>
    /// <returns>完了を表す <see cref="ValueTask" />。</returns>
    ValueTask ExecuteToClearAsync(DocModifierSetting setting, IReadOnlyCollection<IDocumentItem>? documentItems);
}
