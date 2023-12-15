// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DisposableExtensions.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner.Disposables;

/// <summary>
/// <see cref="IDisposable" /> インターフェイスの拡張メソッドを提供します。
/// </summary>
public static class DisposableExtensions
{
    /// <summary>
    /// コンテナへ追加します。
    /// </summary>
    /// <typeparam name="T"><see cref="IDisposable" /> インターフェイスを実装した型。</typeparam>
    /// <param name="self"><see cref="IDisposable" /> インターフェイスを実装したしたインスタンス。</param>
    /// <param name="container">コンテナ。</param>
    /// <returns><paramref name="self" />。</returns>
    public static T AddTo<T>(this T self, CompositeDisposable container)
        where T : IDisposable
    {
        ArgumentNullException.ThrowIfNull(container);
        container.Add(self);
        return self;
    }
}
