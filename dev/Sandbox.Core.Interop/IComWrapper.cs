// --------------------------------------------------------------------------------------------------------------------
// <copyright file="IComWrapper.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace Sandbox.Core.Interop;

/// <summary>
/// ランタイム呼び出し可能ラッパー (RCW: Runtime Callable Wrappers) のインターフェイスを表します。
/// </summary>
/// <typeparam name="TComObject">Runtime Callable Wrappers の対象となる COM 型。</typeparam>
public interface IComWrapper<out TComObject> : IDisposable
    where TComObject : class
{
    /// <summary>
    /// Dispose されたかを取得します。
    /// </summary>
    /// <value>
    /// 値を表す <see cref="bool" /> 型。
    /// <para>Dispose された場合 true。既定値は false です。</para>
    /// </value>
    bool IsDisposed { get; }

    /// <summary>
    /// Runtime Callable Wrappers の対象となる COM オブジェクトを取得します。
    /// </summary>
    /// <value>
    /// <para>Runtime Callable Wrappers の対象となる COM オブジェクト。既定値は null です。</para>
    /// </value>
    TComObject Target { get; }
}
