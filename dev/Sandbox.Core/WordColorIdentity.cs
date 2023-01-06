// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WordColorIdentity.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace Sandbox.Core;

/// <summary>
/// <see cref="Microsoft.Office.Interop.Word.WdColorIndex" /> と等価な蛍光ペン識別子の列挙体を表します。
/// </summary>
public enum WordColorIdentity
{
    /// <summary>[wdYellow] を表します。</summary>
    Yellow,

    /// <summary>[wdBrightGreen] を表します。</summary>
    BrightGreen,

    /// <summary>[wdTurquoise] を表します。</summary>
    Turquoise,

    /// <summary>[wdPink] を表します。</summary>
    Pink,

    /// <summary>[wdBlue] を表します。</summary>
    Blue,

    /// <summary>[wdRed] を表します。</summary>
    Red,

    /// <summary>[wdDarkBlue] を表します。</summary>
    DarkBlue,

    /// <summary>[wdTeal] を表します。</summary>
    Teal,

    /// <summary>[wdGreen] を表します。</summary>
    Green,

    /// <summary>[wdViolet] を表します。</summary>
    Violet,

    /// <summary>[wdDarkRed] を表します。</summary>
    DarkRed,

    /// <summary>[wdDarkYellow] を表します。</summary>
    DarkYellow,

    /// <summary>[wdGray50] を表します。</summary>
    Gray50,

    /// <summary>[wdGray25] を表します。</summary>
    Gray25,

    /// <summary>[wdBlack] を表します。</summary>
    Black,
}
