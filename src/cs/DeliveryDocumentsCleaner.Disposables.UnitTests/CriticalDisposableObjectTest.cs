// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CriticalDisposableObjectTest.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner.Disposables.UnitTests;

public class CriticalDisposableObjectTest
{
    [Fact]
    public void CriticalDisposableObject_Dispose()
    {
        var disposable = new WrappedCriticalDisposableObject();

        Assert.False(disposable.IsDisposed);
        Assert.Equal(0, disposable.NumberOfReleaseCoreMethodCalls);

        disposable.Dispose();

        Assert.True(disposable.IsDisposed);
        Assert.Equal(1, disposable.NumberOfReleaseCoreMethodCalls);

        disposable.Dispose();

        Assert.True(disposable.IsDisposed);
        Assert.Equal(1, disposable.NumberOfReleaseCoreMethodCalls);
    }

    [Fact]
    public void CriticalDisposableObject_Dispose_CallByGc()
    {
        WeakReference<WrappedCriticalDisposableObject>? weak;
        void DisposeByScope()
        {
            // This will go out of scope after dispose() is executed.
            var disposable = new WrappedCriticalDisposableObject();
            weak = new WeakReference<WrappedCriticalDisposableObject>(disposable, trackResurrection: true);
        }

        DisposeByScope();
        GC.WaitForPendingFinalizers();
        GC.Collect(0, GCCollectionMode.Forced);
        GC.WaitForPendingFinalizers();

        Assert.NotNull(weak);
        if (weak.TryGetTarget(out var resurrection))
        {
            Assert.True(resurrection.IsDisposed);
            Assert.Equal(1, resurrection.NumberOfReleaseCoreMethodCalls);
        }
        else
        {
            Assert.Fail("Failed");
        }
    }

    private class WrappedCriticalDisposableObject : CriticalDisposableObject
    {
        public int NumberOfReleaseCoreMethodCalls { get; private set; }

        /// <inheritdoc />
        protected override void ReleaseCore() => this.NumberOfReleaseCoreMethodCalls++;
    }
}
