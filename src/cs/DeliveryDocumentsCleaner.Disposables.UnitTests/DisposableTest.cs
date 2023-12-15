// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DisposableTest.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner.Disposables.UnitTests;

public class DisposableTest
{
    [Fact]
    public void Disposable_DisposableCreate_Dispose_CallByGc()
    {
        var called = false;
        WeakReference<IDisposable>? weak;
        void DisposeByScope()
        {
            // This will go out of scope after dispose() is executed.
            var disposable = Disposable.Create(() => called = true);
            weak = new WeakReference<IDisposable>(disposable, trackResurrection: true);
        }

        DisposeByScope();
        GC.WaitForPendingFinalizers();
        GC.Collect(0, GCCollectionMode.Forced);
        GC.WaitForPendingFinalizers();

        Assert.True(called);
        
        Assert.NotNull(weak);
        if (!weak.TryGetTarget(out var resurrection))
        {
            Assert.Fail("Failed");
        }
    }

    [Fact]
    public void Disposable_DisposableCreate_Dispose()
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
