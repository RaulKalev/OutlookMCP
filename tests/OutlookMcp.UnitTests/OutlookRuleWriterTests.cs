using System.Runtime.InteropServices;
using Microsoft.Extensions.Logging.Abstractions;
using OutlookMcp.Contracts;
using OutlookMcp.OutlookInterop;

namespace OutlookMcp.UnitTests;

public sealed class OutlookRuleWriterTests
{
    private const ushort VariantTypeArray = 0x2000;
    private const ushort VariantTypeVariant = 12;

    [Fact]
    public void ObjectArrayMarshalsAsSafeArrayOfVariants()
    {
        var nativeVariant = Marshal.AllocCoTaskMem(24);
        try
        {
            Marshal.GetNativeVariantForObject(new object[] { "Hexest", "PRC-26-015" }, nativeVariant);

            Assert.Equal(VariantTypeArray | VariantTypeVariant, (ushort)Marshal.ReadInt16(nativeVariant));
        }
        finally
        {
            _ = VariantClear(nativeVariant);
            Marshal.FreeCoTaskMem(nativeVariant);
        }
    }

    [Fact]
    public void OneSubjectTermUsesVariantArrayAndSavesEnabledRule()
    {
        var rules = new FakeRules();
        var destination = new FakeFolder("folder-1");
        var request = Request(["Hexest"]);

        var created = (FakeRule)new OutlookRuleWriter(NullLogger.Instance).CreateAndSave(rules, destination, request);

        var values = Assert.IsType<object[]>(created.Conditions.Subject.Text);
        Assert.Equal(["Hexest"], values.Cast<string>());
        Assert.True(created.Conditions.Subject.Enabled);
        Assert.True(created.Enabled);
        Assert.Equal(1, rules.SaveCalls);
    }

    [Fact]
    public void MultipleSubjectTermsShareOneConditionWithOrSemantics()
    {
        var rules = new FakeRules();
        var request = Request(["Hexest", "Lõhkeainetehas", "Lõhkeinetehas", "PRC-26-015"]);

        var created = (FakeRule)new OutlookRuleWriter(NullLogger.Instance).CreateAndSave(rules, new FakeFolder("folder-1"), request);

        var values = Assert.IsType<object[]>(created.Conditions.Subject.Text);
        Assert.Equal(request.SubjectContains, values.Cast<string>());
        Assert.Single(rules.Items);
    }

    [Fact]
    public void MoveActionUsesTheExactDestinationFolder()
    {
        var rules = new FakeRules();
        var destination = new FakeFolder("exact-folder-id");

        var created = (FakeRule)new OutlookRuleWriter(NullLogger.Instance).CreateAndSave(rules, destination, Request(["Hexest"]));

        Assert.True(created.Actions.MoveToFolder.Enabled);
        Assert.Same(destination, created.Actions.MoveToFolder.Folder);
    }

    [Fact]
    public void StopProcessingActionIsEnabledOnlyWhenRequested()
    {
        var writer = new OutlookRuleWriter(NullLogger.Instance);
        var enabledRules = new FakeRules();
        var disabledRules = new FakeRules();

        var enabled = (FakeRule)writer.CreateAndSave(enabledRules, new FakeFolder("folder-1"), Request(["Hexest"], stop: true));
        var disabled = (FakeRule)writer.CreateAndSave(disabledRules, new FakeFolder("folder-1"), Request(["Hexest"], stop: false));

        Assert.True(enabled.Actions.Stop.Enabled);
        Assert.False(disabled.Actions.Stop.Enabled);
    }

    [Fact]
    public void SaveFailureRemovesOnlyCreatedRuleAndPersistsCleanup()
    {
        var existing = new FakeRule("Existing rule");
        var rules = new FakeRules(existing) { FailFirstSave = true };
        var writer = new OutlookRuleWriter(NullLogger.Instance);

        var exception = Assert.Throws<COMException>(() => writer.CreateAndSave(rules, new FakeFolder("folder-1"), Request(["Hexest"])));

        Assert.Equal(unchecked((int)0x80020009), exception.HResult);
        Assert.Single(rules.Items);
        Assert.Same(existing, rules.Items[0]);
        Assert.Equal(1, rules.RemoveCalls);
        Assert.Equal(2, rules.SaveCalls);
    }

    private static CreateFolderRuleRequest Request(IReadOnlyList<string> subjects, bool stop = false) =>
        new("store-1", "folder-1", "Test rule", SubjectContains: subjects, StopProcessingMoreRules: stop, DryRun: false);

    public sealed class FakeRules(params FakeRule[] existing)
    {
        public List<FakeRule> Items { get; } = [.. existing];
        public bool FailFirstSave { get; set; }
        public int SaveCalls { get; private set; }
        public int RemoveCalls { get; private set; }
        public int Count => Items.Count;
        public FakeRule this[int index] => Items[index - 1];

        public FakeRule Create(string name, int ruleType)
        {
            Assert.Equal(0, ruleType);
            var rule = new FakeRule(name);
            Items.Insert(0, rule);
            return rule;
        }

        public void Remove(int index)
        {
            RemoveCalls++;
            Items.RemoveAt(index - 1);
        }

        public void Save(bool showProgress)
        {
            Assert.False(showProgress);
            SaveCalls++;
            if (!FailFirstSave || SaveCalls != 1) return;
            Marshal.ThrowExceptionForHR(unchecked((int)0x80020009));
        }
    }

    public sealed class FakeRule(string name)
    {
        public string Name { get; } = name;
        public FakeConditions Conditions { get; } = new();
        public FakeActions Actions { get; } = new();
        public bool Enabled { get; set; }
    }

    public sealed class FakeConditions
    {
        public FakeAddressCondition SenderAddress { get; } = new();
        public FakeTextCondition Subject { get; } = new();
        public FakeTextCondition Body { get; } = new();
        public FakeTextCondition BodyOrSubject { get; } = new();
    }

    public sealed class FakeTextCondition
    {
        public object Text { get; set; } = Array.Empty<object>();
        public bool Enabled { get; set; }
    }

    public sealed class FakeAddressCondition
    {
        public object Address { get; set; } = Array.Empty<object>();
        public bool Enabled { get; set; }
    }

    public sealed class FakeActions
    {
        public FakeMoveAction MoveToFolder { get; } = new();
        public FakeStopAction Stop { get; } = new();
    }

    public sealed class FakeMoveAction
    {
        public bool Enabled { get; set; }
        public object? Folder { get; set; }
    }

    public sealed class FakeStopAction
    {
        public bool Enabled { get; set; }
    }

    public sealed class FakeFolder(string entryId)
    {
        public string EntryID { get; } = entryId;
    }

    [DllImport("oleaut32.dll")]
    private static extern int VariantClear(IntPtr variant);
}
