// --------------------------------------------------------------------------------------------------------------------
// <copyright file="XlApplicationExtensions.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using Microsoft.Office.Interop.Excel;

namespace Sandbox.Core.Interop;

/// <summary>
/// <see cref="IComWrapper{Application}" /> インターフェイスの拡張メソッドを提供します。
/// </summary>
internal static class XlApplicationExtensions
{
    /// <summary>
    /// 処理中かどうかを設定します。
    /// </summary>
    /// <param name="xlApp"><see cref="IComWrapper{Application}" /> インスタンス。</param>
    /// <param name="isBusy">処理中としてマークする場合は <see langword="true" />。それ以外は <see langword="false" />。</param>
    internal static void SetBusy(this IComWrapper<Application> xlApp, bool isBusy = true)
    {
        xlApp.Target.DisplayAlerts = !isBusy;
        xlApp.Target.DisplayStatusBar = !isBusy;
        xlApp.Target.EnableEvents = !isBusy;
        xlApp.Target.PrintCommunication = !isBusy;
        xlApp.Target.ScreenUpdating = !isBusy;
    }
}
