// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CompositeDisposableTest.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner.Disposables.UnitTests;

public class CompositeDisposableTest
{
    [Fact]
    public void CompositeDisposable_Dispose()
    {
        var disposable = new CompositeDisposable();
        Assert.NotNull(disposable);
        Assert.False(disposable.IsDisposed);

        disposable.Dispose();
        Assert.True(disposable.IsDisposed);
    }

    [Fact]
    public void CompositeDisposable_Dispose_CallByGc()
    {
        static WeakReference<CompositeDisposable> DisposeByScope()
        {
            // This will go out of scope after dispose() is executed.
            var disposable = new CompositeDisposable();
            return new WeakReference<CompositeDisposable>(disposable, trackResurrection: true);
        }

        var weak = DisposeByScope();

        GC.WaitForPendingFinalizers();
        GC.Collect(0, GCCollectionMode.Forced);
        GC.WaitForPendingFinalizers();

        if (weak.TryGetTarget(out var resurrection))
        {
            Assert.True(resurrection.IsDisposed);
        }
        else
        {
            Assert.True(false, "Failed");
        }
    }

    [Fact]
    public void CompositeDisposable_AddDisposable()
    {
        var called = false;
        var disposableToAdd = Disposable.Create(() => called = true);
        var disposable = new CompositeDisposable();

        disposable.Add(disposableToAdd);
        Assert.False(disposable.IsDisposed);
        Assert.False(called);

        disposable.Dispose();
        Assert.True(disposable.IsDisposed);
        Assert.True(called);
    }

    [Fact]
    public void CompositeDisposable_AddDisposable_AlreadyDisposed()
    {
        var called = false;
        var disposableToAdd = Disposable.Create(() => called = true);
        var disposable = new CompositeDisposable();

        disposable.Dispose();
        Assert.True(disposable.IsDisposed);
        Assert.False(called);

        disposable.Add(disposableToAdd);
        Assert.True(disposable.IsDisposed);
        Assert.True(called);
    }

    [Fact]
    public void CompositeDisposable_Ctor_WithEnumerable_DoseNotThrow()
    {
        var called1 = false;
        var disposableToAdd1 = Disposable.Create(() => called1 = true);
        var called2 = false;
        var disposableToAdd2 = Disposable.Create(() => called2 = true);

        var disposable = new CompositeDisposable(disposableToAdd1, disposableToAdd2);

        Assert.False(disposable.IsDisposed);
        Assert.False(called1);
        Assert.False(called2);

        disposable.Dispose();
        Assert.True(disposable.IsDisposed);
        Assert.True(called1);
        Assert.True(called2);
    }
}
