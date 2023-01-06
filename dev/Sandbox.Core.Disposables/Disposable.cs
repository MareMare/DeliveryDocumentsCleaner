// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Disposable.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace Sandbox.Core.Disposables
{
    /// <summary>
    /// Dispose 可能なインスタンスを提供します。
    /// </summary>
    public static class Disposable
    {
        /// <summary>
        /// 何もしない既定の Dispose 可能なインスタンスを取得します。
        /// </summary>
        /// <value>
        /// 値を表す <see cref="IDisposable" /> 型。
        /// <para>既定値は <seealso cref="EmptyDisposable" /> です。</para>
        /// </value>
        public static IDisposable Empty { get; } = new EmptyDisposable();

        /// <summary>
        /// <see cref="IDisposable.Dispose" /> が呼ばれた時の処理を行うメソッドのデリゲートを指定して <see cref="IDisposable" /> を実装したインスタンスを生成します。
        /// </summary>
        /// <param name="dispose"><see cref="IDisposable.Dispose" /> が呼ばれた時の処理を行うメソッドのデリゲート。</param>
        /// <returns>生成された <see cref="IDisposable" /> を実装したインスタンス。</returns>
        public static IDisposable Create(Action dispose) =>
            new AnonymousDisposable(dispose ?? throw new ArgumentNullException(nameof(dispose)));

        /// <summary>
        /// 何もしない既定の Dispose 可能なインスタンスを表します。
        /// </summary>
        private sealed class EmptyDisposable : IDisposable
        {
            /// <summary>
            /// <see cref="EmptyDisposable" /> クラスのインスタンスが GC に回収される時に呼び出されます。
            /// </summary>
            ~EmptyDisposable()
            {
                EmptyDisposable.Nop();
            }

            /// <inheritdoc />
            public void Dispose()
            {
                EmptyDisposable.Nop();
                GC.SuppressFinalize(this);
            }

            /// <summary>
            /// 何もしません。
            /// </summary>
            private static void Nop()
            {
                // no op
            }
        }

        /// <summary>
        /// 匿名の <see cref="IDisposable" /> インターフェイスの実装を表します。
        /// </summary>
        private sealed class AnonymousDisposable : IDisposable
        {
            /// <summary><see cref="IDisposable.Dispose" /> が呼ばれた時の処理を行うメソッドのデリゲートを表します。</summary>
            private volatile Action? _dispose;

            /// <summary><see cref="IDisposable.Dispose" /> メソッドが呼び出されたかをスレッドセーフで管理する値を表します。</summary>
            private long _disposableState;

            /// <summary>
            /// <see cref="IDisposable.Dispose" /> が呼ばれた時の処理を行うメソッドのデリゲートを指定して <see cref="IDisposable" /> を実装したインスタンスを生成します。
            /// </summary>
            /// <param name="dispose"><see cref="IDisposable.Dispose" /> が呼ばれた時の処理を行うメソッドのデリゲート。</param>
            internal AnonymousDisposable(Action dispose)
            {
                this._dispose = dispose;
            }

            /// <summary>
            /// <see cref="AnonymousDisposable" /> クラスのインスタンスが GC に回収される時に呼び出されます。
            /// </summary>
            ~AnonymousDisposable()
            {
                this.DisposeCore();
            }

            /// <inheritdoc />
            public void Dispose()
            {
                this.DisposeCore();
                GC.SuppressFinalize(this);
            }

            /// <summary>
            /// <see cref="DisposableBase" /> クラスのインスタンスによって使用されているアンマネージ リソースを解放し、オプションでマネージ リソースも解放します。
            /// </summary>
            private void DisposeCore()
            {
                if (Interlocked.CompareExchange(ref this._disposableState, 1L, 0L) == 0L)
                {
                    Interlocked.Exchange(ref this._dispose, null)?.Invoke();
                }
            }
        }
    }
}
