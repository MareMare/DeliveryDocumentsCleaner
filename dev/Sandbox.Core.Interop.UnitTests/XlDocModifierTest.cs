// --------------------------------------------------------------------------------------------------------------------
// <copyright file="XlDocModifierTest.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System.Drawing;

namespace Sandbox.Core.Interop.UnitTests;

public class XlDocModifierTest : IDisposable
{
    private static readonly string TargetExcelFilePath = Path.Combine(Environment.CurrentDirectory, "01.おためしブック.xlsx");

    [Fact]
    public void XlDocModifier_Sandbox()
    {
        using var modifier = new XlDocModifier();

        modifier.FilePath = XlDocModifierTest.TargetExcelFilePath;
        modifier.OpenBook();
        modifier.ClearHighlightColor(Color.Yellow);
        modifier.ClearPersonalInfo();
        modifier.CloseBook();
    }

    [Fact]
    public void XlDocModifier_FilePath_IfAlreadyOpened_ShouldThrow()
    {
        using var modifier = new XlDocModifier();

        modifier.FilePath = XlDocModifierTest.TargetExcelFilePath;
        modifier.OpenBook();

        Assert.Throws<InvalidOperationException>(() => modifier.FilePath = XlDocModifierTest.TargetExcelFilePath);
    }

    /// <inheritdoc />
    public void Dispose()
    {
        GC.WaitForPendingFinalizers();
        GC.Collect(0, GCCollectionMode.Forced);
        GC.WaitForPendingFinalizers();

        GC.SuppressFinalize(this);
    }
}
