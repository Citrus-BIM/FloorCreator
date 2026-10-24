#requires -Version 7.0
<#
Compiles both real WPF code-behind files and settings classes against small UI/API
stand-ins. Only settings path expressions are redirected to unique temporary XML
files; the production XML read/write logic runs without touching user settings.
#>
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$namespace = 'PhaseFilterSettingsRegression_' + [Guid]::NewGuid().ToString('N')
$sources = [System.Collections.Generic.List[string]]::new()
$cases = [System.Collections.Generic.List[string]]::new()
foreach ($variant in @('FloorCreator', 'FloorCreatorSpectrum')) {
    foreach ($file in @('FloorCreatorWPF.xaml.cs', 'FloorCreatorSettings.cs', 'PhaseFilterSelectionItem.cs', 'ElementIdCompat.cs')) {
        $source = Get-Content -LiteralPath (Join-Path $root "$variant/$file") -Raw
        $source = [regex]::Replace($source, '(?m)^using .*?;\r?\n', '')
        if ($file -eq 'FloorCreatorSettings.cs') {
            if ($variant -eq 'FloorCreator') {
                $pathExpression = 'assemblyPathAll.Replace("FloorCreator.dll", fileName)'
                if (-not $source.Contains($pathExpression)) { throw 'Main settings path seam changed.' }
                $source = $source.Replace($pathExpression, 'TestStorage.FilePath')
            } else {
                $directory = [regex]::Match($source, 'private static string SettingsDirectory\s*=>[\s\S]*?;')
                if (-not $directory.Success) { throw 'Spectrum settings directory seam changed.' }
                $source = $source.Replace($directory.Value, 'private static string SettingsDirectory => TestStorage.Directory;')
                $source = $source.Replace('Path.Combine(SettingsDirectory, FileName)', 'TestStorage.FilePath')
            }
        }
        $sources.Add($source.Replace('namespace FloorCreator', "namespace $namespace.$variant"))
    }
    $sources.Add(@'
namespace __NAMESPACE__.__VARIANT__
{
    public partial class FloorCreatorWPF
    {
        public ComboBox comboBox_FloorType = new ComboBox(), comboBox_DoorPhaseFilter = new ComboBox();
        public RadioButton rbt_ManualCreation = new RadioButton(), rbt_CreateFromParameter = new RadioButton(), rbt_InSelected = new RadioButton(), rbt_InWholeProject = new RadioButton();
        public CheckBox chk_FillDoorPatches = new CheckBox(), chk_FilterDoorsByPhase = new CheckBox(), chk_DeleteOldFloors = new CheckBox();
        public TextBox textBox_FloorLevelOffset = new TextBox();
        public GroupBox groupBox_FloorCreationOption = new GroupBox(), groupBox_InRooms = new GroupBox();
        private void InitializeComponent() {}
        public void Accept() => btn_Ok_Click(null, null);
        public void Cancel() => btn_Cancel_Click(null, null);
        public string SelectedName => (comboBox_DoorPhaseFilter.SelectedItem as PhaseFilterSelectionItem)?.Name;
        public long SelectedId => (comboBox_DoorPhaseFilter.SelectedItem as PhaseFilterSelectionItem)?.PhaseFilterId.Value ?? -1;
    }
}
'@.Replace('__NAMESPACE__', $namespace).Replace('__VARIANT__', $variant))
    $cases.Add(@'
{
    TestStorage.FilePath = prefix + "__VARIANT__.xml";
    var filters = new List<__VARIANT__.PhaseFilterSelectionItem> {
        new __VARIANT__.PhaseFilterSelectionItem(ElementId.InvalidElementId, "Нет"),
        new __VARIANT__.PhaseFilterSelectionItem(new ElementId(10), "Active view"),
        new __VARIANT__.PhaseFilterSelectionItem(new ElementId(20), "User choice") };
    var floorTypes = new List<FloorType> { new FloorType { Name = "Floor" } };
    Func<__VARIANT__.FloorCreatorWPF> open = () => new __VARIANT__.FloorCreatorWPF(floorTypes, filters, new ElementId(10));
    Seed(null);
    var first = open();
    first.chk_FillDoorPatches.IsChecked = true;
    first.chk_FilterDoorsByPhase.IsChecked = true;
    first.comboBox_DoorPhaseFilter.SelectedItem = filters[2];
    first.Accept();
    Check("__VARIANT__: active command receives selected ID", first.FilterDoorsByPhase && first.SelectedDoorPhaseFilterId.Value == 20);
    Check("__VARIANT__: actual XML persists selected name", File.ReadAllText(TestStorage.FilePath).Contains("<DoorPhaseFilterName>User choice</DoorPhaseFilterName>"));
    var reopened = open();
    Check("__VARIANT__: chosen filter overrides active-view default after reopen", reopened.SelectedName == "User choice");
    var otherDocumentFilters = new List<__VARIANT__.PhaseFilterSelectionItem> {
        new __VARIANT__.PhaseFilterSelectionItem(ElementId.InvalidElementId, "Нет"),
        new __VARIANT__.PhaseFilterSelectionItem(new ElementId(1010), "Active view"),
        new __VARIANT__.PhaseFilterSelectionItem(new ElementId(2020), "User choice") };
    var otherDocument = new __VARIANT__.FloorCreatorWPF(floorTypes, otherDocumentFilters, new ElementId(1010));
    Check("__VARIANT__: name resolves to current document ID", otherDocument.SelectedName == "User choice" && otherDocument.SelectedId == 2020);

    reopened.comboBox_DoorPhaseFilter.SelectedItem = filters[2];
    reopened.chk_FilterDoorsByPhase.IsChecked = false;
    reopened.Accept();
    var disabled = open();
    Check("__VARIANT__: disabled filter preserves preference but command gets invalid ID", !reopened.FilterDoorsByPhase && reopened.SelectedDoorPhaseFilterId.Value == -1 && disabled.SelectedName == "User choice" && disabled.chk_FilterDoorsByPhase.IsChecked == false);
    disabled.comboBox_DoorPhaseFilter.SelectedItem = filters[2];
    disabled.chk_FillDoorPatches.IsChecked = false;
    disabled.chk_FilterDoorsByPhase.IsChecked = true;
    disabled.Accept();
    Check("__VARIANT__: disabled door patches preserve preference", !disabled.FilterDoorsByPhase && disabled.SelectedDoorPhaseFilterId.Value == -1 && open().SelectedName == "User choice");
    disabled.comboBox_DoorPhaseFilter.SelectedItem = filters[0];
    disabled.chk_FillDoorPatches.IsChecked = true;
    disabled.chk_FilterDoorsByPhase.IsChecked = true;
    disabled.Accept();
    Check("__VARIANT__: explicit None survives active-view fallback", open().SelectedName == "Нет" && disabled.SelectedDoorPhaseFilterId.Value == -1);

    Seed(null);
    Check("__VARIANT__: legacy XML without new field uses active view", open().SelectedName == "Active view");
    Seed("Missing filter");
    Check("__VARIANT__: absent saved filter uses active view", open().SelectedName == "Active view");
    var noActiveMatch = new __VARIANT__.FloorCreatorWPF(floorTypes, filters, new ElementId(999));
    Check("__VARIANT__: missing saved and active filters use first item", noActiveMatch.SelectedName == "Нет");
    Seed("User choice");
    string beforeCancel = File.ReadAllText(TestStorage.FilePath);
    var cancel = open();
    cancel.comboBox_DoorPhaseFilter.SelectedItem = filters[1];
    cancel.Cancel();
    Check("__VARIANT__: cancel leaves XML unchanged", cancel.DialogResult == false && File.ReadAllText(TestStorage.FilePath) == beforeCancel);
}
'@.Replace('__VARIANT__', $variant))
}
$harness = @'
#pragma warning disable CS0219
using System;
using System.IO;
using System.Xml.Serialization;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using __NAMESPACE__.Stubs;
namespace __NAMESPACE__.Stubs
{
    public static class TestStorage { public static string FilePath; public static string Directory => Path.GetDirectoryName(FilePath); }
    public sealed class ElementId
    {
        public static readonly ElementId InvalidElementId = new ElementId(-1);
        public ElementId(long value) { Value = value; }
        public long Value { get; }
        public int IntegerValue => (int)Value;
    }
    public class FloorType { public string Name { get; set; } }
    public class RoutedEventArgs : EventArgs {}
    public enum Key { Enter, Space, Escape }
    public class KeyEventArgs : RoutedEventArgs { public Key Key; }
    public class Window { public bool? DialogResult; public void Close() {} }
    public class CheckBox { public bool? IsChecked; public bool IsEnabled; }
    public class RadioButton : CheckBox { public string Name; }
    public class TextBox { public string Text; }
    public class StackPanel { public List<object> Children = new List<object>(); }
    public class Grid : StackPanel {}
    public class GroupBox { public object Content; public bool IsEnabled; }
    public class ComboBox
    {
        public object ItemsSource;
        public string DisplayMemberPath;
        public object SelectedItem;
        public bool IsEnabled;
        public IList Items => ((IEnumerable)ItemsSource)?.Cast<object>().ToList() ?? new List<object>();
    }
}
__SOURCES__
namespace __NAMESPACE__
{
    public static class Regression
    {
        private static readonly List<string> Failures = new List<string>();
        private static int _passed;
        private static void Check(string label, bool condition) { if (!condition) Failures.Add(label); else _passed++; }
        private static void Seed(string selected)
        {
            File.WriteAllText(TestStorage.FilePath, "<FloorCreatorSettings><FillDoorPatches>true</FillDoorPatches><FilterDoorsByPhase>true</FilterDoorsByPhase>" + (selected == null ? "" : "<DoorPhaseFilterName>" + selected + "</DoorPhaseFilterName>") + "</FloorCreatorSettings>");
        }
        public static string Run(string prefix)
        {
            __CASES__
            if (Failures.Count != 0) throw new Exception(string.Join(Environment.NewLine, Failures));
            return "PASS: " + _passed + " phase filter settings assertions across both variants.";
        }
    }
}
'@.Replace('__NAMESPACE__', $namespace).Replace('__SOURCES__', ($sources -join "`n")).Replace('__CASES__', ($cases -join "`n"))
Add-Type -TypeDefinition $harness -Language CSharp
$testType = "$namespace.Regression" -as [type]
$prefix = Join-Path ([IO.Path]::GetTempPath()) ('FloorCreatorPhaseSettings-' + [Guid]::NewGuid().ToString('N') + '-')
try {
    $testType::Run($prefix)
} finally {
    foreach ($variant in @('FloorCreator', 'FloorCreatorSpectrum')) {
        foreach ($suffix in @('.xml', '.xml.tmp')) {
            $testFile = $prefix + $variant + $suffix
            if (Test-Path -LiteralPath $testFile) { Remove-Item -LiteralPath $testFile -Force }
        }
    }
}
