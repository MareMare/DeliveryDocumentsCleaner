// --------------------------------------------------------------------------------------------------------------------
// <copyright file="DocModifierTest.cs" company="MareMare">
// Copyright © 2022 MareMare. All rights reserved.
// Licensed under the MIT license. See LICENSE file in the project root for full license information.
// </copyright>
// --------------------------------------------------------------------------------------------------------------------

using Xunit.Abstractions;

namespace DeliveryDocumentsCleaner.Interop.UnitTests;

public class DocModifierTest
{
    private readonly ITestOutputHelper _testOutputHelper;

    public DocModifierTest(ITestOutputHelper testOutputHelper) =>
        this._testOutputHelper = testOutputHelper;

    //[Fact]
    public async Task Test2Async()
    {
        static string BuildDummyPath(int index)
        {
            return Path.Combine(Environment.CurrentDirectory, $"Dummy{index}.{(index % 2 == 0 ? "docx" : "xlsx")}");
        }

        var setting = new DocModifierSetting();
        var modifier = new DocModifier(new Progress<IDocumentItem>());

        using var completedSignal = new ManualResetEventSlim();
        var closureSignal = completedSignal;

        var totalCount = 30;
        var countOfCalled = 0L;
        var gate = new object();
        modifier.Progress.ProgressChanged += (_, item) =>
        {
            var next = Interlocked.Increment(ref countOfCalled);
            lock (gate)
            {
                this._testOutputHelper.WriteLine($"{next} {item.FileName}");
                if (next >= totalCount)
                {
                    closureSignal.Set();
                }
            }
        };

        var documentItems = Enumerable.Range(1, totalCount).Select(index => new DocItem(BuildDummyPath(index))).ToArray();
        await modifier.ExecuteToClearAsync(setting, documentItems).ConfigureAwait(true);

        var waited = completedSignal.Wait(TimeSpan.FromSeconds(totalCount));

        Assert.True(waited);
        Assert.Equal(totalCount, countOfCalled);
    }

    private class DocItem : IDocumentItem
    {
        public DocItem(string filepath)
        {
            this.FilePath = filepath;
            this.FileName = Path.GetFileName(filepath);
            this.DocKind = DocItem.ResolveDocKind(filepath);
        }

        public DocumentKind? DocKind { get; }

        public string FilePath { get; }

        public string FileName { get; }

        public override string ToString() => this.FileName;

        private static DocumentKind? ResolveDocKind(string path)
        {
            var ext = Path.GetExtension(path).ToLowerInvariant();
            return ext switch
            {
                @".docx" => DocumentKind.Word,
                @".docm" => DocumentKind.Word,
                @".xlsx" => DocumentKind.Excel,
                @".xlsm" => DocumentKind.Excel,
                _ => null,
            };
        }
    }
}
