// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CompositeDisposable.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.Collections.Concurrent;

namespace DeliveryDocumentsCleaner.Disposables;

/// <summary>
/// Dispose 可能なインスタンスのコンテナを表します。
/// </summary>
public class CompositeDisposable : IDisposable
{
    /// <summary>コンテナを表します。</summary>
    private readonly ConcurrentBag<IDisposable> _disposables;

    /// <summary><see cref="IDisposable.Dispose" /> メソッドが呼び出されたかをスレッドセーフで管理する値を表します。</summary>
    private long _disposableState;

    /// <summary>
    /// <see cref="CompositeDisposable" /> クラスの新しいインスタンスを初期化します。
    /// </summary>
    /// <param name="disposables">新しいリストに要素がコピーされるコレクション。</param>
    public CompositeDisposable(params IDisposable[] disposables)
    {
        ArgumentNullException.ThrowIfNull(disposables);
        this._disposables = new ConcurrentBag<IDisposable>(disposables);
    }

    /// <summary>
    /// <see cref="CompositeDisposable" /> クラスのインスタンスが GC に回収される時に呼び出されます。
    /// </summary>
    ~CompositeDisposable()
    {
        this.Dispose(false);
    }

    /// <summary>
    /// Dispose されたかを取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="bool" /> 型。
    /// <para>Dispose された場合 true。既定値は false です。</para>
    /// </value>
    public bool IsDisposed => Interlocked.Read(ref this._disposableState) == 1L;

    /// <summary>
    /// コンテナに追加します。
    /// </summary>
    /// <param name="disposable">追加するインスタンス。</param>
    public void Add(IDisposable disposable)
    {
        ArgumentNullException.ThrowIfNull(disposable);

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
        this.Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// <see cref="CompositeDisposable" /> クラスのインスタンスによって使用されているアンマネージ リソースを解放し、オプションでマネージ リソースも解放します。
    /// </summary>
    /// <param name="disposing">マネージ リソースとアンマネージ リソースの両方を解放する場合は true。アンマネージ リソースだけを解放する場合は false。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (Interlocked.CompareExchange(ref this._disposableState, 1L, 0L) == 0L)
        {
            if (disposing)
            {
                while (this._disposables.TryTake(out var disposable))
                {
                    disposable.Dispose();
                }
            }
        }
    }
}
