// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DisposableBaseTest.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner.Disposables.UnitTests;

public class DisposableBaseTest
{
    [Fact]
    public void WhenDispose_ShouldCallMethods()
    {
        var disposable = new WrappedDisposable();
        Assert.NotNull(disposable);
        Assert.False(disposable.IsDisposed);
        Assert.False(disposable.IsCalledToDisposeManagedInstances);
        Assert.False(disposable.IsCalledToDisposeUnmanagedInstances);
        Assert.False(disposable.IsCalledToWriteLogAtDisposing);
        Assert.False(disposable.IsCalledToWriteLogAtDisposed);
        Assert.Null(disposable.ReasonForDisposing);

        disposable.Dispose();
        Assert.True(disposable.IsDisposed);
        Assert.True(disposable.IsCalledToDisposeManagedInstances);
        Assert.True(disposable.IsCalledToDisposeUnmanagedInstances);
        Assert.True(disposable.IsCalledToWriteLogAtDisposing);
        Assert.True(disposable.IsCalledToWriteLogAtDisposed);
        Assert.True(disposable.ReasonForDisposing);
    }

    [Fact]
    public void Dispose_CallsByGc()
    {
        WeakReference<WrappedDisposable>? weak;
        void DisposeByScope()
        {
            // This will go out of scope after dispose() is executed.
            var disposable = new WrappedDisposable();
            weak = new WeakReference<WrappedDisposable>(disposable, true);
        }

        DisposeByScope();
        GC.WaitForPendingFinalizers();
        GC.Collect(0, GCCollectionMode.Forced);
        GC.WaitForPendingFinalizers();

        Assert.NotNull(weak);
        if (weak.TryGetTarget(out var resurrection))
        {
            Assert.True(resurrection.IsDisposed);
            Assert.False(resurrection.IsCalledToDisposeManagedInstances); // Not called through finalizer.
            Assert.True(resurrection.IsCalledToDisposeUnmanagedInstances);
            Assert.True(resurrection.IsCalledToWriteLogAtDisposing);
            Assert.True(resurrection.IsCalledToWriteLogAtDisposed);
            Assert.False(resurrection.ReasonForDisposing); // Not called through finalizer.
        }
    }

    [Fact]
    public void WhenDispose_ThrowException()
    {
        var disposable = new WrappedDisposable();
        disposable.CallToThrowIfDisposed();
        disposable.Dispose();

        Assert.Throws<ObjectDisposedException>(() => disposable.CallToThrowIfDisposed());
    }

    private class WrappedDisposable : DisposableBase
    {
        public bool IsCalledToDisposeManagedInstances { get; private set; }
        public bool IsCalledToDisposeUnmanagedInstances { get; private set; }
        public bool IsCalledToWriteLogAtDisposing { get; private set; }
        public bool IsCalledToWriteLogAtDisposed { get; private set; }
        public bool? ReasonForDisposing { get; private set; }

        public void CallToThrowIfDisposed()
            => this.ThrowIfDisposed();

        protected override void DisposeManagedInstances()
        {
            this.IsCalledToDisposeManagedInstances = true;
            base.DisposeManagedInstances();
        }

        protected override void DisposeUnmanagedInstances()
        {
            this.IsCalledToDisposeUnmanagedInstances = true;
            base.DisposeUnmanagedInstances();
        }

        protected override void WriteLogAtDisposing(bool disposing)
        {
            this.ReasonForDisposing = disposing;
            this.IsCalledToWriteLogAtDisposing = true;
            base.WriteLogAtDisposing(disposing);
        }

        protected override void WriteLogAtDisposed()
        {
            this.IsCalledToWriteLogAtDisposed = true;
            base.WriteLogAtDisposed();
        }
    }
}
