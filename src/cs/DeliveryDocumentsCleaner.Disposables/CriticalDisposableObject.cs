// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CriticalDisposableObject.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.Runtime.ConstrainedExecution;

namespace DeliveryDocumentsCleaner.Disposables;

/// <summary>
/// <see cref="IDisposable" /> インターフェイスを実装した <see cref="CriticalFinalizerObject" /> クラスを表します。
/// </summary>
public abstract class CriticalDisposableObject : CriticalFinalizerObject, IDisposable
{
    /// <summary>Dispose メソッドが呼び出されたかをスレッドセーフで管理する値を表します。</summary>
    private long _disposableState;

    /// <summary>
    /// <see cref="CriticalDisposableObject" /> クラスのインスタンスが GC に回収される時に呼び出されます。
    /// </summary>
    ~CriticalDisposableObject() => this.Dispose(false);

    /// <summary>
    /// Dispose されたかを取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="bool" /> 型。
    /// <para>Dispose された場合 true。既定値は false です。</para>
    /// </value>
    public bool IsDisposed => Interlocked.Read(ref this._disposableState) == 1L;

    /// <inheritdoc />
    public void Dispose()
    {
        this.Dispose(true);
        GC.SuppressFinalize(this);
    }

    /// <summary>
    /// <see cref="CriticalDisposableObject" /> クラスのインスタンスによって使用されているアンマネージ リソースを解放し、オプションでマネージ リソースも解放します。
    /// </summary>
    /// <param name="disposing">マネージ リソースとアンマネージ リソースの両方を解放する場合は true。アンマネージ リソースだけを解放する場合は false。</param>
    protected virtual void Dispose(bool disposing)
    {
        if (Interlocked.CompareExchange(ref this._disposableState, 1, 0) == 0)
        {
            this.ReleaseCore();
        }
    }

    /// <summary>
    /// リソースの解放およびリセットに関連付けられているアプリケーション定義のタスクを実行します。
    /// </summary>
    protected abstract void ReleaseCore();
}
