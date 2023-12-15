// --------------------------------------------------------------------------------------------------------------------
// <copyright file="Program.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner.App;

/// <summary>
/// アプリケーションのエントリポイントを提供します。
/// </summary>
internal static class Program
{
    /// <summary>
    /// アプリケーションのメインエントリポイントです。
    /// </summary>
    [STAThread]
    internal static void Main()
    {
        // To customize application configuration such as set high DPI settings or default font,
        // see https://aka.ms/applicationconfiguration.
        ApplicationConfiguration.Initialize();

        using var form = new MainForm();
        Application.Run(form);
    }
}
