---
name: powerpoint-fixture-authoring
description: Use when creating, replacing, verifying, or documenting real Microsoft PowerPoint-authored .pptx fixtures in this ts-pptx repository, especially for read-model or OOXML bugs that need desktop PowerPoint output rather than ts-pptx-generated packages.
metadata:
  internal: true
---

# PowerPoint fixture authoring

Reference fixtures are evidence for how PowerPoint writes OOXML, so never generate them with
ts-pptx. `pnpm ppt:run` sends a recipe to the worker when `TSPPTX_POWERPOINT_URL` is set (the
VM in `tools/powerpoint-vm/`; `pnpm ppt:health` confirms it answers) and otherwise runs it on
this Windows machine's PowerPoint. With neither, stop and open the fixture issue AGENTS.md
describes.

## Workflow

1. Fixtures live in `test/read/fixtures/`. Replace an existing one only when asked, and delete
   only that path. Keep exploration decks out of the repo.
2. Write the recipe as `test/read/fixtures/authoring/author-<name>.ps1` from the start; it is
   the fixture's provenance and is committed with it. That directory's `README.md` has the
   path convention, `--with`, and the sidecar formatter step.

   ```sh
   pnpm ppt:run test/read/fixtures/authoring/author-<name>.ps1 [-- -Param value]
   ```

   Do not use `powershell.exe -ExecutionPolicy Bypass`; the sandbox denies it.
3. Keep the deck minimal: deliberate slide size, stable shape names, deterministic geometry and
   colors, no external assets unless the bug needs them.
4. Verify with the helpers below: no repair prompt, `docProps/app.xml` names PowerPoint, the
   slide XML holds the construct being pinned, and a SHA-256.
5. Record provenance, hash, purpose and check date in `test/read/fixtures/README.md`. Take the
   PowerPoint build from the line `ppt:run` prints.

## COM authoring pattern

A recipe has this shape:

```powershell
$ErrorActionPreference = 'Stop'
$REPO = (Resolve-Path (Join-Path $PSScriptRoot '..\..\..\..')).Path
$out  = Join-Path $REPO 'test\read\fixtures\example.pptx'
# Snapshot pre-existing PIDs so the reap at the end only kills the server we
# spawn — never a user's interactive PowerPoint with unsaved work.
$preexistingIds = @(Get-Process POWERPNT -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Id)
$pp = $null
$pres = $null
try {
  $pp = New-Object -ComObject PowerPoint.Application
  $pp.DisplayAlerts = 1
  $pres = $pp.Presentations.Add(1)
  $pres.PageSetup.SlideWidth = 960
  $pres.PageSetup.SlideHeight = 540
  $slide = $pres.Slides.Add(1, 12) # ppLayoutBlank

  # Add named shapes/groups that exercise the target PowerPoint behavior.

  $pres.SaveAs($out)
  $pres.Saved = $true
  $pres.Close()
  $pp.Quit()
}
finally {
  if ($pres -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pres) }
  if ($pp -ne $null) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($pp) }
  [GC]::Collect()
  [GC]::WaitForPendingFinalizers()
  # Quit() can leave the automation server lingering; reap only PIDs we created.
  Get-Process POWERPNT -ErrorAction SilentlyContinue |
    Where-Object { $preexistingIds -notcontains $_.Id } |
    Stop-Process -Force -ErrorAction SilentlyContinue
}
```

For grouped shape fixtures, create the child shapes first, group by shape names,
then set transforms on the returned group. PowerPoint writes group transforms
under `p:grpSpPr/a:xfrm`, for example `rot`, `flipH`, and `flipV`.

### Freeform / boolean (`custGeom`) geometry

Build a single freeform with `BuildFreeform`/`AddNodes`/`ConvertToShape` (each
`AddNodes` appends one segment; `msoSegmentLine`=0, `msoSegmentCurve`=1 takes
control1/control2/end). Return-to-start closes the path → `a:close`.

For a hole or a boolean, `ShapeRange.MergeShapes(...)` fails under PowerShell late binding.
Open the deck with a window (`Presentations.Open(path, 0, 0, -1)`), select the shapes, and run
the ribbon command:

```powershell
$base.Select($true)     # msoTrue: replace selection
$hole.Select($false)    # msoFalse: extend selection
$pp.CommandBars.ExecuteMso('ShapesSubtract')   # or ShapesUnion / ShapesCombine / ShapesFragment
$merged = $pp.ActiveWindow.Selection.ShapeRange.Item(1)
$merged.Name = '...'    # name the merged result
```

The same `ExecuteMso`-on-a-selection pattern is the fallback for any COM method whose
enum arguments refuse to late-bind.

`BuildFreeform` emits one `a:path`, and Merge Shapes (even of disjoint shapes) folds
everything into one `a:path` with several contours. PowerPoint never writes a multi-`a:path`
`custGeom`, so that construct is not authorable here.

## Autofit bake-on-save (`normAutofit` / `spAutoFit`)

PowerPoint bakes autofit into the saved XML only with this sequence:

- Set `TextFrame2.AutoSize = msoAutoSizeNone`, pin `Width`/`Height` (`AddTextbox` ignores its
  height), add the text, and set `AutoSize` **last**. Shrink then bakes
  `<a:normAutofit fontScale lnSpcReduction/>`; resize bakes `<a:spAutoFit/>` with a fitted
  `ext.cy`. A box that has already grown bakes a bare `<a:normAutofit/>`.
- A trailing empty paragraph is not enumerable until all text exists, so insert text in one
  pass and apply paragraph formatting in a second.

`test/read/fixtures/authoring/author-deck.ps1` is the full parameterized engine.

## Font-presence guard

PowerPoint writes `latin@typeface="X"` even when `X` is not installed and substitutes only at
render time, so the XML proves nothing about metrics. Before any font-sensitive fixture, check
in a fresh process that GDI resolves each face to itself:

```powershell
Add-Type -AssemblyName System.Drawing
foreach ($face in 'Aptos','Aptos SemiBold','Calibri','Tahoma','Arial') {
  $f = New-Object System.Drawing.Font($face, 18); $resolved = $f.Name; $f.Dispose()
  if ($resolved -ne $face) { throw "FONT SUBSTITUTED: $face -> $resolved" }
}
```

`test/read/fixtures/authoring/readiness-guard.ps1` is the worked example.

The worker VM installs the needed fonts (`tools/powerpoint-vm/oem/install-fonts.ps1`). On
another Windows machine, install a font per user (copy to `%LOCALAPPDATA%\Microsoft\Windows\Fonts`
and register under `HKCU:\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Fonts`). Office cloud
fonts are visible only to Office, not GDI. `test/read/fixtures/authoring/measure-lo.py` reads
LibreOffice's fitted sizes over UNO as a second opinion.

## Helpers

Both run through `pnpm ppt:run` like a recipe, with their parameters after `--`.
The job carries the fixtures tree, so any committed fixture can be named.

**Verify** — opens the deck through PowerPoint COM, reads package metadata,
computes SHA-256, optionally lists group `a:xfrm` attributes, and reaps any
automation-server process it spawned (reported as `reapedProcessIds`):

```sh
pnpm ppt:run .agents/skills/powerpoint-fixture-authoring/scripts/verify-powerpoint-fixture.ps1 -- -Path test/read/fixtures/<fixture>.pptx -InspectGroups
```

**Dump slide XML** — prints a slide's XML so you can confirm the exact OOXML
construct PowerPoint emitted (indented by default; `-Raw` for the stored form):

```sh
pnpm ppt:run .agents/skills/powerpoint-fixture-authoring/scripts/dump-slide-xml.ps1 -- -Path test/read/fixtures/<fixture>.pptx -Slide 1
```

If the helper output does not show the expected OOXML construct, fix the
PowerPoint authoring script and regenerate the fixture rather than patching the
OOXML by hand.
