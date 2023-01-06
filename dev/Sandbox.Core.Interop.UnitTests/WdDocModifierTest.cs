// --------------------------------------------------------------------------------------------------------------------
// <copyright file="WdDocModifierTest.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace Sandbox.Core.Interop.UnitTests;

public class WdDocModifierTest : IDisposable
{
    private static readonly string TargetWordFilePath = Path.Combine(Environment.CurrentDirectory, "02.おためし文書.docx");

    [Fact]
    public void WdDocModifier_Sandbox()
    {
        using var modifier = new WdDocModifier();

        modifier.FilePath = WdDocModifierTest.TargetWordFilePath;
        modifier.OpenDoc();
        modifier.ClearHighlightColor();
        modifier.ClearPersonalInfo();
        modifier.CloseDoc();
    }

    [Fact]
    public void WdDocModifier_FilePath_IfAlreadyOpened_ShouldThrow()
    {
        using var modifier = new WdDocModifier();

        modifier.FilePath = WdDocModifierTest.TargetWordFilePath;
        modifier.OpenDoc();

        Assert.Throws<InvalidOperationException>(() => modifier.FilePath = WdDocModifierTest.TargetWordFilePath);
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
