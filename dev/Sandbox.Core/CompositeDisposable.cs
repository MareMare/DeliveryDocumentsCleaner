// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CompositeDisposable.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.Collections.Concurrent;

namespace Sandbox.Core;

/// <summary>
/// Dispose 可能なインスタンスのコンテナを表します。
/// </summary>
public class CompositeDisposable : IDisposable
{
    /// <summary>コンテナを表します。</summary>
    private readonly ConcurrentBag<IDisposable> _disposables;

    /// <summary>
    /// <see cref="CompositeDisposable" /> クラスの新しいインスタンスを初期化します。
    /// </summary>
    /// <param name="disposables">新しいリストに要素がコピーされるコレクション。</param>
    public CompositeDisposable(params IDisposable[] disposables)
    {
        this._disposables = new ConcurrentBag<IDisposable>(disposables) ??
                            throw new ArgumentNullException(nameof(disposables));
    }

    /// <summary>
    /// Dispose されたかを取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="bool" /> 型。
    /// <para>Dispose された場合 true。既定値は false です。</para>
    /// </value>
    public bool IsDisposed { get; private set; }

    /// <summary>
    /// コンテナに追加します。
    /// </summary>
    /// <param name="disposable">追加するインスタンス。</param>
    public void Add(IDisposable disposable)
    {
        if (this.IsDisposed)
        {
            disposable.Dispose();
        }
        else
        {
            this._disposables.Add(disposable);
        }
    }

    /// <inheritdoc />
    public void Dispose()
    {
        if (this.IsDisposed)
        {
            return;
        }

        this.IsDisposed = true;
        while (this._disposables.TryTake(out var disposable))
        {
            disposable.Dispose();
        }

        GC.SuppressFinalize(this);
    }
}
