// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ComWrapper.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.Runtime.InteropServices;
using System.Runtime.Versioning;

namespace Sandbox.Core;

/// <summary>
/// ランタイム呼び出し可能ラッパー (RCW: Runtime Callable Wrappers) のファクトリを提供します。
/// </summary>
public static class ComWrapper
{
    /// <summary>
    /// <see cref="IComWrapper{TComObject}" /> を実装したインスタンスを生成します。
    /// </summary>
    /// <typeparam name="TComObject">Runtime Callable Wrappers の対象となる COM 型。</typeparam>
    /// <param name="target">Runtime Callable Wrappers の対象となる COM オブジェクト。</param>
    /// <returns><see cref="IComWrapper{TComObject}" /> を実装したインスタンス。</returns>
    public static IComWrapper<TComObject> Create<TComObject>(TComObject target)
        where TComObject : class => new Rcw<TComObject>(target);

    /// <summary>
    /// ランタイム呼び出し可能ラッパー (RCW: Runtime Callable Wrappers) のオブジェクトを表します。
    /// </summary>
    /// <typeparam name="TComObject">Runtime Callable Wrappers の対象となる COM 型。</typeparam>
    private sealed class Rcw<TComObject> : CriticalDisposableObject, IComWrapper<TComObject>
        where TComObject : class
    {
        /// <summary>アンマネージ メモリからマネージ オブジェクトにアクセスできるようにする機能を表します。</summary>
        private GCHandle _targetHandle;

        /// <summary>
        /// <see cref="Rcw{TComObject}" /> クラスの新しいインスタンスを初期化します。
        /// </summary>
        /// <param name="target">Runtime Callable Wrappers の対象となる COM オブジェクト。</param>
        internal Rcw(TComObject target)
        {
            var comObject = target ?? throw new ArgumentNullException(nameof(target));
            if (!Marshal.IsComObject(comObject))
            {
                throw new ArgumentException("COM オブジェクトではありません。", nameof(target));
            }

            this._targetHandle = GCHandle.Alloc(target);
        }

        /// <inheritdoc />
        public TComObject Target => this.GetTarget();

        /// <inheritdoc />
        [SupportedOSPlatform("windows")]
        protected override void ReleaseCore()
        {
            if (this._targetHandle.IsAllocated)
            {
                if (this._targetHandle.Target != null)
                {
                    _ = Marshal.FinalReleaseComObject(this._targetHandle.Target);
                }

                this._targetHandle.Target = null;
                this._targetHandle.Free();
            }
        }

        /// <summary>
        /// Runtime Callable Wrappers の対象となる COM オブジェクトを取得します。
        /// </summary>
        /// <returns>Runtime Callable Wrappers の対象となる COM オブジェクト。既定値は null です。</returns>
        private TComObject GetTarget()
        {
            if (this._targetHandle.IsAllocated)
            {
                if (this._targetHandle.Target != null)
                {
                    return (TComObject)this._targetHandle.Target;
                }
            }

            throw new ObjectDisposedException(typeof(TComObject).Name, "既にこの COM オブジェクトは破棄されています。");
        }
    }
}
