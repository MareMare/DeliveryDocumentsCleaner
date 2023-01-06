// --------------------------------------------------------------------------------------------------------------------
// <copyright file="CriticalDisposableObjectTest.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using System;

namespace Sandbox.Core.UnitTests;

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
        static WeakReference<WrappedCriticalDisposableObject> DisposeByScope()
        {
            // This will go out of scope after dispose() is executed.
            var disposable = new WrappedCriticalDisposableObject();
            return new WeakReference<WrappedCriticalDisposableObject>(disposable, trackResurrection: true);
        }

        var weak = DisposeByScope();

        GC.WaitForPendingFinalizers();
        GC.Collect(0, GCCollectionMode.Forced);
        GC.WaitForPendingFinalizers();

        if (weak.TryGetTarget(out var resurrection))
        {
            Assert.True(resurrection.IsDisposed);
            Assert.Equal(1, resurrection.NumberOfReleaseCoreMethodCalls);
        }
        else
        {
            Assert.True(false, "Failed");
        }
    }

    private class WrappedCriticalDisposableObject : CriticalDisposableObject
    {
        public int NumberOfReleaseCoreMethodCalls { get; private set; } = 0;

        /// <inheritdoc />
        protected override void ReleaseCore() => this.NumberOfReleaseCoreMethodCalls++;
    }
}
