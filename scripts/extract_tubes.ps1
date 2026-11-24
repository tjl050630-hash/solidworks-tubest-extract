param(
    [Parameter(Mandatory = $true)]
    [string]$InputPath,

    [Parameter(Mandatory = $false)]
    [string]$ProjectName = "",

    [Parameter(Mandatory = $false)]
    [string]$OutputRoot = "",

    [Parameter(Mandatory = $false)]
    [string]$SolidWorksDir = "",

    [Parameter(Mandatory = $false)]
    [double]$MinTubeCrossSectionMm = 12.0,

    [Parameter(Mandatory = $false)]
    [double]$MinTubeLengthMm = 50.0,

    [Parameter(Mandatory = $false)]
    [switch]$KeepSolidWorksOpen
)

$ErrorActionPreference = "Stop"

if (!(Test-Path $InputPath)) {
    throw "Input file not found: $InputPath"
}

$InputPath = (Resolve-Path $InputPath).Path
$baseName = [System.IO.Path]::GetFileNameWithoutExtension($InputPath)

if ([string]::IsNullOrWhiteSpace($ProjectName)) {
    $ProjectName = $baseName
}

if ([string]::IsNullOrWhiteSpace($OutputRoot)) {
    $desktop = [Environment]::GetFolderPath('Desktop')
    $OutputRoot = Join-Path $desktop "拆图"
}

function Find-SolidWorksDir {
    param([string]$ExplicitDir)
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($ExplicitDir)) {
        $candidates += $ExplicitDir
    }
    # Check registry for installed SolidWorks path
    $regPaths = @(
        "HKLM:\SOFTWARE\SolidWorks\Applications\PDMWorks Enterprise",
        "HKLM:\SOFTWARE\SolidWorks\IM",
        "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\SLDWORKS.exe"
    )
    foreach ($rp in $regPaths) {
        try {
            $item = Get-ItemProperty -Path $rp -ErrorAction SilentlyContinue
            if ($item -and $item.'(default)') {
                $candidates += (Split-Path $item.'(default)' -Parent)
            }
            if ($item -and $item.Path) {
                $candidates += $item.Path
            }
        } catch {}
    }
    # Common installation directories across drives
    foreach ($drive in @("C:", "D:", "E:")) {
        $candidates += "$drive\Program Files\SOLIDWORKS Corp\SOLIDWORKS"
        $candidates += "$drive\SOLIDWORKS Corp\SOLIDWORKS"
    }
    foreach ($dir in ($candidates | Select-Object -Unique)) {
        if ($dir -and (Test-Path (Join-Path $dir "SolidWorks.Interop.sldworks.dll")) -and (Test-Path (Join-Path $dir "SolidWorks.Interop.swconst.dll"))) {
            return $dir
        }
    }
    throw "Could not locate SolidWorks Interop DLLs. Please pass -SolidWorksDir '<path to SOLIDWORKS folder>'."
}

$swDir = Find-SolidWorksDir -ExplicitDir $SolidWorksDir
$dll1 = Join-Path $swDir "SolidWorks.Interop.sldworks.dll"
$dll2 = Join-Path $swDir "SolidWorks.Interop.swconst.dll"

$outDir = Join-Path $OutputRoot $ProjectName
New-Item -ItemType Directory -Force -Path $outDir | Out-Null

# Record existing SolidWorks PIDs before launching so we only clean up what we spawned
$existingSwPids = @(Get-Process -Name "SLDWORKS" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)

[System.Reflection.Assembly]::LoadFrom($dll1) | Out-Null
[System.Reflection.Assembly]::LoadFrom($dll2) | Out-Null

$code = @"
using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;

public class SolidWorksTubeExtractor
{
    static string Sanitize(string s)
    {
        if (string.IsNullOrEmpty(s)) return "Tube";
        foreach (char c in Path.GetInvalidFileNameChars()) s = s.Replace(c, '_');
        return s.Replace(" ", "_").Replace("/", "_").Trim('_');
    }

    static bool IsTubeBody(Body2 b, double minCrossMm, double minLenMm)
    {
        double[] box = (double[])b.GetBodyBox();
        if (box == null || box.Length < 6) return false;
        double[] d = new double[] {
            (box[3] - box[0]) * 1000.0,
            (box[4] - box[1]) * 1000.0,
            (box[5] - box[2]) * 1000.0
        };
        Array.Sort(d);
        int faces = b.GetFaceCount();
        string name = (b.Name ?? "").ToLower();

        if (name.Contains("gusset") || name.Contains("end cap") || name.Contains("boss-extrude") || name.Contains("mirror") || name.Contains("plate"))
        {
            if (d[0] < minCrossMm && faces <= 12) return false;
        }
        if (d[0] >= minCrossMm && d[2] >= minLenMm && faces >= 4)
        {
            if (faces == 6) return false;
            return true;
        }
        return false;
    }

    static double EstimateCutLengthMm(Body2 rep, string cutListLengthStr, double maxBoxDimMm, ModelDoc2 doc)
    {
        double parsed;
        if (!string.IsNullOrWhiteSpace(cutListLengthStr) && double.TryParse(cutListLengthStr, out parsed) && parsed > 0)
        {
            int lenUnit = doc.GetUserPreferenceIntegerValue((int)swUserPreferenceIntegerValue_e.swUnitsLinear);
            if (lenUnit == (int)swLengthUnit_e.swINCHES) return Math.Round(parsed * 25.4, 1);
            if (lenUnit == (int)swLengthUnit_e.swMETER) return Math.Round(parsed * 1000.0, 1);
            if (lenUnit == (int)swLengthUnit_e.swCM) return Math.Round(parsed * 10.0, 1);
            return Math.Round(parsed, 1);
        }
        return Math.Round(maxBoxDimMm, 1);
    }

    public static void Run(string inputPath, string outDir, string prefix, double minCrossMm, double minLenMm)
    {
        ISldWorks swApp = null;
        try { swApp = (ISldWorks)Marshal.GetActiveObject("SldWorks.Application"); }
        catch { swApp = (ISldWorks)Activator.CreateInstance(Type.GetTypeFromProgID("SldWorks.Application")); }
        swApp.Visible = true;
        Console.WriteLine("SolidWorks Revision: " + swApp.RevisionNumber());

        string ext = Path.GetExtension(inputPath).ToLower();
        int err = 0, warn = 0;
        ModelDoc2 doc = null;

        if (ext == ".sldprt")
        {
            doc = swApp.OpenDoc6(inputPath, (int)swDocumentTypes_e.swDocPART, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
        }
        else if (ext == ".sldasm")
        {
            doc = swApp.OpenDoc6(inputPath, (int)swDocumentTypes_e.swDocASSEMBLY, (int)swOpenDocOptions_e.swOpenDocOptions_Silent, "", ref err, ref warn);
        }
        else if (ext == ".step" || ext == ".stp" || ext == ".igs" || ext == ".iges" || ext == ".sat")
        {
            ImportStepData stepData = (ImportStepData)swApp.GetImportFileData(inputPath);
            doc = (ModelDoc2)swApp.LoadFile4(inputPath, "r", stepData, ref err);
        }
        else
        {
            throw new Exception("Unsupported file extension: " + ext);
        }

        if (doc == null)
        {
            throw new Exception("Failed to open model in SolidWorks: " + inputPath + " (err=" + err + ", warn=" + warn + ")");
        }

        try
        {
            if (doc.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                ExtractPart(doc, outDir, Sanitize(prefix), minCrossMm, minLenMm);
            }
            else if (doc.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                ExtractAssembly(doc, outDir, Sanitize(prefix), minCrossMm, minLenMm);
            }
        }
        finally
        {
            swApp.CloseDoc(doc.GetTitle());
        }
    }

    static void ExtractPart(ModelDoc2 doc, string outDir, string prefix, double minCrossMm, double minLenMm)
    {
        int err = 0, warn = 0;
        doc.ShowNamedView2("*Isometric", (int)swStandardViews_e.swIsometricView);
        doc.ViewZoomtofit2();
        doc.SaveAs(Path.Combine(outDir, prefix + "_00_Original.png"));

        PartDoc part = (PartDoc)doc;
        object vAll = part.GetBodies2((int)swBodyType_e.swSolidBody, false);
        if (vAll == null) { Console.WriteLine("No solid bodies found."); return; }
        object[] allBodies = (object[])vAll;

        var groups = new List<Tuple<string, List<Body2>, string, string>>();
        Feature feat = (Feature)doc.FirstFeature();
        while (feat != null)
        {
            if (feat.GetTypeName2() == "CutListFolder")
            {
                BodyFolder bf = (BodyFolder)feat.GetSpecificFeature2();
                if (bf != null)
                {
                    object vBodies = bf.GetBodies();
                    if (vBodies != null)
                    {
                        var tubeBodies = new List<Body2>();
                        foreach (object ob in (object[])vBodies)
                        {
                            Body2 b = (Body2)ob;
                            if (IsTubeBody(b, minCrossMm, minLenMm)) tubeBodies.Add(b);
                        }
                        if (tubeBodies.Count > 0)
                        {
                            string desc = "", lenStr = "";
                            CustomPropertyManager cpm = feat.CustomPropertyManager;
                            if (cpm != null)
                            {
                                string vOut = "", rOut = "";
                                cpm.Get4("DESCRIPTION", false, out vOut, out rOut);
                                desc = rOut;
                                cpm.Get4("LENGTH", false, out vOut, out rOut);
                                lenStr = rOut;
                            }
                            groups.Add(Tuple.Create(feat.Name, tubeBodies, desc, lenStr));
                        }
                    }
                }
            }
            feat = (Feature)feat.GetNextFeature();
        }

        if (groups.Count == 0)
        {
            foreach (object ob in allBodies)
            {
                Body2 b = (Body2)ob;
                if (IsTubeBody(b, minCrossMm, minLenMm))
                {
                    groups.Add(Tuple.Create(b.Name, new List<Body2> { b }, "", ""));
                }
            }
        }

        foreach (object ob in allBodies)
        {
            Body2 b = (Body2)ob;
            b.HideBody(!IsTubeBody(b, minCrossMm, minLenMm));
        }
        doc.GraphicsRedraw2();
        doc.ViewZoomtofit2();
        doc.SaveAs(Path.Combine(outDir, prefix + "_00_TubesOnly.png"));

        string satAll = Path.Combine(outDir, prefix + "_AllTubesOnly.sat");
        doc.Extension.SaveAs(satAll, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref err, ref warn);
        Console.WriteLine("Saved combined tubes SAT: " + satAll);

        var csv = new StringBuilder();
        csv.AppendLine("序号,STEP文件名,切割清单项,规格描述,长度_mm,包围盒_mm,数量");

        int idx = 0;
        foreach (var grp in groups)
        {
            idx++;
            Body2 rep = grp.Item2[0];
            int qty = grp.Item2.Count;
            double[] box = (double[])rep.GetBodyBox();
            double[] d = new double[] {
                Math.Round((box[3] - box[0]) * 1000.0, 1),
                Math.Round((box[4] - box[1]) * 1000.0, 1),
                Math.Round((box[5] - box[2]) * 1000.0, 1)
            };
            Array.Sort(d);
            double lenMm = EstimateCutLengthMm(rep, grp.Item4, d[2], doc);
            string bbox = string.Format("{0}x{1}x{2}", d[0], d[1], d[2]);

            foreach (object ob in allBodies)
            {
                Body2 b = (Body2)ob;
                b.HideBody(b.Name != rep.Name);
            }

            string stepName = string.Format("{0}_T{1:D2}_{2}x{3}_L{4}mm_{5}pcs.step", prefix, idx, d[0], d[1], lenMm, qty);
            string stepPath = Path.Combine(outDir, stepName);
            bool ok = doc.Extension.SaveAs(stepPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref err, ref warn);
            Console.WriteLine(string.Format("  [{0}] {1} -> {2} (qty={3})", idx, ok ? "OK" : "FAIL", stepName, qty));
            csv.AppendLine(string.Format("{0},\"{1}\",\"{2}\",\"{3}\",{4},\"{5}\",{6}",
                idx, stepName, grp.Item1, grp.Item3, lenMm, bbox, qty));
        }

        foreach (object ob in allBodies) { ((Body2)ob).HideBody(false); }
        File.WriteAllText(Path.Combine(outDir, prefix + "_BOM.csv"), csv.ToString(), new UTF8Encoding(true));
    }

    static void ExtractAssembly(ModelDoc2 doc, string outDir, string prefix, double minCrossMm, double minLenMm)
    {
        int err = 0, warn = 0;
        doc.ShowNamedView2("*Isometric", (int)swStandardViews_e.swIsometricView);
        doc.ViewZoomtofit2();
        doc.SaveAs(Path.Combine(outDir, prefix + "_00_Original.png"));

        AssemblyDoc assy = (AssemblyDoc)doc;
        assy.ResolveAllLightWeightComponents(true);

        var counts = new Dictionary<string, int>();
        var models = new Dictionary<string, ModelDoc2>();
        var cfgs = new Dictionary<string, string>();
        var dimsMap = new Dictionary<string, double[]>();

        Configuration activeCfg = (Configuration)doc.ConfigurationManager.ActiveConfiguration;
        Component2 root = activeCfg.GetRootComponent3(true);
        if (root != null)
        {
            TraverseAssy(root, counts, models, cfgs, dimsMap, minCrossMm, minLenMm);
        }

        doc.GraphicsRedraw2();
        doc.ViewZoomtofit2();
        doc.SaveAs(Path.Combine(outDir, prefix + "_00_TubesOnly.png"));

        string satAll = Path.Combine(outDir, prefix + "_AllTubesOnly.sat");
        doc.Extension.SaveAs(satAll, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref err, ref warn);
        Console.WriteLine("Saved combined tubes SAT: " + satAll);

        var csv = new StringBuilder();
        csv.AppendLine("序号,STEP文件名,原零件名,配置名,长度_mm,包围盒_mm,数量");

        int idx = 0;
        foreach (var kv in counts)
        {
            idx++;
            ModelDoc2 cm = models[kv.Key];
            string cfgName = cfgs[kv.Key];
            double[] d = dimsMap[kv.Key];

            string origCfg = cm.ConfigurationManager.ActiveConfiguration.Name;
            if (!string.IsNullOrEmpty(cfgName) && origCfg != cfgName)
            {
                cm.ShowConfiguration2(cfgName);
            }

            string title = cm.GetTitle();
            if (title.Contains(".")) title = title.Substring(0, title.LastIndexOf('.'));
            string stepName = string.Format("{0}_T{1:D2}_{2}_L{3}mm_{4}pcs.step", prefix, idx, Sanitize(title), d[2], kv.Value);
            string stepPath = Path.Combine(outDir, stepName);
            bool ok = cm.Extension.SaveAs(stepPath, (int)swSaveAsVersion_e.swSaveAsCurrentVersion, (int)swSaveAsOptions_e.swSaveAsOptions_Silent, null, ref err, ref warn);

            if (!string.IsNullOrEmpty(cfgName) && origCfg != cfgName)
            {
                cm.ShowConfiguration2(origCfg);
            }

            string bbox = string.Format("{0}x{1}x{2}", d[0], d[1], d[2]);
            Console.WriteLine(string.Format("  [{0}] {1} -> {2} (qty={3})", idx, ok ? "OK" : "FAIL", stepName, kv.Value));
            csv.AppendLine(string.Format("{0},\"{1}\",\"{2}\",\"{3}\",{4},\"{5}\",{6}",
                idx, stepName, title, cfgName, d[2], bbox, kv.Value));
        }

        File.WriteAllText(Path.Combine(outDir, prefix + "_BOM.csv"), csv.ToString(), new UTF8Encoding(true));
    }

    static void TraverseAssy(
        Component2 comp,
        Dictionary<string, int> counts,
        Dictionary<string, ModelDoc2> models,
        Dictionary<string, string> cfgs,
        Dictionary<string, double[]> dimsMap,
        double minCrossMm,
        double minLenMm)
    {
        object vChildren = comp.GetChildren();
        if (vChildren == null) return;
        object[] children = (object[])vChildren;
        foreach (object oc in children)
        {
            Component2 c = (Component2)oc;
            if (c.IsHidden(true)) continue;
            ModelDoc2 cm = (ModelDoc2)c.GetModelDoc2();
            if (cm == null) continue;

            if (cm.GetType() == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                TraverseAssy(c, counts, models, cfgs, dimsMap, minCrossMm, minLenMm);
            }
            else if (cm.GetType() == (int)swDocumentTypes_e.swDocPART)
            {
                double[] box = (double[])c.GetBox(false, false);
                if (box == null || box.Length < 6) { c.Visible = (int)swComponentVisibilityState_e.swComponentHidden; continue; }
                double[] d = new double[] {
                    Math.Round((box[3] - box[0]) * 1000.0, 1),
                    Math.Round((box[4] - box[1]) * 1000.0, 1),
                    Math.Round((box[5] - box[2]) * 1000.0, 1)
                };
                Array.Sort(d);

                string fnLower = Path.GetFileNameWithoutExtension(cm.GetPathName() ?? cm.GetTitle()).ToLower();
                bool isFastenerOrPanel = fnLower.Contains("screw") || fnLower.Contains("bolt") || fnLower.Contains("nut")
                    || fnLower.Contains("washer") || fnLower.Contains("hinge") || fnLower.Contains("door") || fnLower.Contains("panel")
                    || fnLower.Contains("螺") || fnLower.Contains("角件") || fnLower.Contains("角铝") || fnLower.Contains("门板") || fnLower.Contains("合页");

                if (isFastenerOrPanel || d[0] < minCrossMm || d[2] < minLenMm)
                {
                    c.Visible = (int)swComponentVisibilityState_e.swComponentHidden;
                    continue;
                }

                string pPath = cm.GetPathName();
                if (string.IsNullOrEmpty(pPath)) pPath = cm.GetTitle();
                string cfg = c.ReferencedConfiguration ?? "";
                string key = pPath.ToLower() + "|" + cfg.ToLower();

                if (!counts.ContainsKey(key))
                {
                    counts[key] = 0;
                    models[key] = cm;
                    cfgs[key] = cfg;
                    dimsMap[key] = d;
                }
                counts[key]++;
            }
        }
    }
}
"@

Add-Type -TypeDefinition $code -ReferencedAssemblies @($dll1, $dll2) -Language CSharp

try {
    [SolidWorksTubeExtractor]::Run($InputPath, $outDir, $ProjectName, $MinTubeCrossSectionMm, $MinTubeLengthMm)

    Write-Host "`n=== Post-Export STEP Validation ==="
    $stepFiles = Get-ChildItem -Path $outDir -Filter "*.step"
    $allValid = $true
    foreach ($sf in $stepFiles) {
        $cnt = (Select-String -Path $sf.FullName -Pattern "MANIFOLD_SOLID_BREP" -AllMatches).Matches.Count
        $status = if ($cnt -eq 1) { "OK (1 solid)" } else { $allValid = $false; "WARN ($cnt solids)" }
        Write-Host ("  {0,-48} {1,7:N1} KB  {2}" -f $sf.Name, ($sf.Length / 1KB), $status)
    }
    Write-Host "`nOutput folder: $outDir"
    if (-not $allValid) {
        Write-Warning "One or more exported STEP files did not contain exactly 1 solid!"
    }
}
finally {
    if (-not $KeepSolidWorksOpen) {
        $nowSwPids = @(Get-Process -Name "SLDWORKS" -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
        foreach ($pidNew in $nowSwPids) {
            if ($existingSwPids -notcontains $pidNew) {
                Stop-Process -Id $pidNew -Force -ErrorAction SilentlyContinue
                Write-Host "Closed spawned SolidWorks PID=$pidNew"
            }
        }
    }
}
