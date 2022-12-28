// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DisposableTest.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace Sandbox.Core.UnitTests;

public class DisposableTest
{
    [Fact]
    public void WhenDispose_DisposableCreated_ShouldCall()
    {
        var called = false;
        var disposable = Disposable.Create(() => called = true);
        Assert.NotNull(disposable);
        Assert.False(called);

        disposable.Dispose();
        Assert.True(called);
    }

    [Fact]
    public void WhenDispose_DisposableEmpty_DoseNotThrow()
    {
        var disposable = Disposable.Empty;
        Assert.NotNull(disposable);
        disposable.Dispose();
    }

    [Fact]
    public void WhenExpressionIsNull_ThrowException()
        => Assert.Throws<ArgumentNullException>(() => Disposable.Create(null!));
}
