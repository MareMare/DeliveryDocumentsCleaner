// --------------------------------------------------------------------------------------------------------------------
// <copyright file="ProgressExtensions.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

namespace DeliveryDocumentsCleaner;

public static class ProgressExtensions
{
    public static void Report<T>(this IProgress<T> progress, T value)
    {
        ArgumentNullException.ThrowIfNull(progress);
        progress.Report(value);
    }
}
