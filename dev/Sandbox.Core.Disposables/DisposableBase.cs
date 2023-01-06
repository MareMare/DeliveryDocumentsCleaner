// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DisposableBase.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.Diagnostics;

namespace Sandbox.Core.Disposables
{
    /// <summary>
    /// Dispose パターンを実装した基本となるオブジェクトを表します。
    /// </summary>
    public abstract class DisposableBase : IDisposable
    {
        /// <summary><see cref="IDisposable.Dispose" /> メソッドが呼び出されたかをスレッドセーフで管理する値を表します。</summary>
        private long _disposableState;

        /// <summary>
        /// <see cref="DisposableBase" /> クラスの新しいインスタンスを初期化します。
        /// </summary>
        protected DisposableBase()
        {
        }

        /// <summary>
        /// <see cref="DisposableBase" /> クラスのインスタンスが GC に回収される時に呼び出されます。
        /// </summary>
        ~DisposableBase()
        {
            this.Dispose(false);
        }

        /// <summary>
        /// <see cref="IDisposable.Dispose" /> されたかを取得します。
        /// </summary>
        /// <value>
        /// 値を表す <see cref="bool" /> 型。
        /// <para><see cref="IDisposable.Dispose" /> された場合 true。既定値は false です。</para>
        /// </value>
        public bool IsDisposed => Interlocked.Read(ref this._disposableState) == 1L;

        /// <summary>
        /// アンマネージ リソースの解放およびリセットに関連付けられているアプリケーション定義のタスクを実行します。
        /// </summary>
        public void Dispose()
        {
            this.Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// <see cref="IDisposable.Dispose" /> するときにログを出力します。
        /// </summary>
        /// <param name="disposing">マネージ リソースとアンマネージ リソースの両方を解放する場合は true。アンマネージ リソースだけを解放する場合は false。</param>
        protected virtual void WriteLogAtDisposing(bool disposing) =>
            Debug.Print(FormattableString.Invariant($"■ {this.GetType().FullName}.Disposing({disposing})..."));

        /// <summary>
        /// <see cref="IDisposable.Dispose" /> したときにログを出力します。
        /// </summary>
        protected virtual void WriteLogAtDisposed() =>
            Debug.Print(FormattableString.Invariant($"■ {this.GetType().FullName}.Disposed..."));

        /// <summary>
        /// マネージ リソースを解放します。
        /// </summary>
        protected virtual void DisposeManagedInstances()
        {
        }

        /// <summary>
        /// アンマネージ リソースを解放します。
        /// </summary>
        protected virtual void DisposeUnmanagedInstances()
        {
        }

        /// <summary>
        /// Dispose されていれば例外を発生させます。
        /// </summary>
        protected void ThrowIfDisposed()
        {
            if (this.IsDisposed)
            {
                throw new ObjectDisposedException(this.GetType().FullName);
            }
        }

        /// <summary>
        /// <see cref="DisposableBase" /> クラスのインスタンスによって使用されているアンマネージ リソースを解放し、オプションでマネージ リソースも解放します。
        /// </summary>
        /// <param name="disposing">マネージ リソースとアンマネージ リソースの両方を解放する場合は true。アンマネージ リソースだけを解放する場合は false。</param>
        private void Dispose(bool disposing)
        {
            if (Interlocked.CompareExchange(ref this._disposableState, 1L, 0L) == 0L)
            {
                this.WriteLogAtDisposing(disposing);
                if (disposing)
                {
                    // マネージ リソース (IDisposable の実装インスタンス) の解放処理をこの位置に記述します。
                    this.DisposeManagedInstances();
                }

                // アンマネージ リソース (IDisposable の非実装インスタンス) の解放処理をこの位置に記述します。
                this.DisposeUnmanagedInstances();

                this.WriteLogAtDisposed();
            }
        }
    }
}
