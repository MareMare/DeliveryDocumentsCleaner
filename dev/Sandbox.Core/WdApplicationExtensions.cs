// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WdApplicationExtensions.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using Microsoft.Office.Interop.Word;

namespace Sandbox.Core;

/// <summary>
/// <see cref="IComWrapper{Application}" /> インターフェイスの拡張メソッドを提供します。
/// </summary>
internal static class WdApplicationExtensions
{
    /// <summary>
    /// 処理中かどうかを設定します。
    /// </summary>
    /// <param name="wdApp"><see cref="IComWrapper{Application}" /> インスタンス。</param>
    /// <param name="isBusy">処理中としてマークする場合は <see langword="true" />。それ以外は <see langword="false" />。</param>
    internal static void SetBusy(this IComWrapper<Application> wdApp, bool isBusy = true) =>
        wdApp.Target.DisplayAlerts = isBusy ? WdAlertLevel.wdAlertsNone : WdAlertLevel.wdAlertsAll;
}
