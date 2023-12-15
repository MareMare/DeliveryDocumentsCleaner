// --------------------------------------------------------------------------------------------------------------------
// <copyright file="UnitTest1.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner.Disposables.UnitTests;

public class MyDisposableBaseTest
{
    [Fact]
    public void Dispose_CallsExplicitAndImplicitDisposal()
    {
        // Arrange
        var @explicit = false;
        var @implicit = false;
        var disposable = new MyDisposable(
            onExplicitDispose: () => @explicit = true,
            onImplicitDispose: () => @implicit = true);

        // Act
        disposable.Dispose();

        // Assert
        Assert.True(@explicit);
        Assert.True(@implicit);
        Assert.True(disposable.Disposed);
    }

    [Fact]
    public void Dispose_CallsImplicitOnlyOnFinalization()
    {
        // Arrange
        var @explicit = false;
        var @implicit = false;
        WeakReference<MyDisposable>? weak;
        void DisposeByScope()
        {
            // This will go out of scope after dispose() is executed
            var disposable = new MyDisposable(
                onExplicitDispose: () => @explicit = true,
                onImplicitDispose: () => @implicit = true);
            weak = new WeakReference<MyDisposable>(disposable, true);
        }

        // Act
        DisposeByScope();
        GC.Collect(0, GCCollectionMode.Forced);
        GC.WaitForPendingFinalizers();

        // Assert
        Assert.False(@explicit); // Not called through finalizer
        Assert.True(@implicit);

        Assert.NotNull(weak);
        if (weak.TryGetTarget(out var resurrection))
        {
            Assert.True(resurrection.Disposed);
        }
        else
        {
            Assert.Fail("Failed");
        }
    }

    [Fact]
    public void Dispose_CallsImplicitOnlyOnFinalization2()
    {
        // Arrange
        WeakReference<MyDisposable>? weak;
        void DisposeByScope()
        {
            // This will go out of scope after dispose() is executed
            var disposable = new MyDisposable();
            weak = new WeakReference<MyDisposable>(disposable, true);
        }

        // Act
        DisposeByScope();
        GC.Collect(0, GCCollectionMode.Forced);
        GC.WaitForPendingFinalizers();

        // Assert
        // Assert another here...

        Assert.NotNull(weak);
        if (weak.TryGetTarget(out var resurrection))
        {
            Assert.True(resurrection.Disposed);
        }
        else
        {
            Assert.Fail("Failed");
        }
    }


    private class MyDisposable : MyDisposableBase
    {
        private readonly Action? _onExplicitDispose;
        private readonly Action? _onImplicitDispose;

        public MyDisposable(Action? onExplicitDispose = null, Action? onImplicitDispose = null)
        {
            this._onExplicitDispose = onExplicitDispose;
            this._onImplicitDispose = onImplicitDispose;
        }

        protected override void DisposeExplicit() =>
            this._onExplicitDispose?.DynamicInvoke();

        protected override void DisposeImplicit() =>
            this._onImplicitDispose?.DynamicInvoke();
    }
}

public abstract class MyDisposableBase : IDisposable
{
    ~MyDisposableBase() => this.Dispose(false);

    public bool Disposed { get; private set; }

    /// <inheritdoc />
    public void Dispose()
    {
        this.Dispose(true);
        GC.SuppressFinalize(this);
    }

    protected virtual void Dispose(bool disposing)
    {
        if (!this.Disposed)
        {
            this.Disposed = true;

            if (disposing)
            {
                this.DisposeExplicit();
            }

            this.DisposeImplicit();
        }
    }

    protected virtual void DisposeExplicit() { }
    protected virtual void DisposeImplicit() { }
}
