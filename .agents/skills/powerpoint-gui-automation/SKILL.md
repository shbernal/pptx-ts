---
name: powerpoint-gui-automation
description: Use in an interactive Windows session with desktop Microsoft PowerPoint (a Windows machine, or the worker VM's desktop over its web viewer or RDP) when a feature has NO COM/VBA surface at all (e.g. Insert > Zoom - Slide/Section/Summary Zoom - has no Shapes.AddZoom) and must be authored by driving the real GUI, or when you need to visually observe how a file renders/behaves in the actual desktop app rather than infer it from OOXML. Escalation path from powerpoint-fixture-authoring, not a replacement for it - try COM (and its ExecuteMso fallback) first.
metadata:
  internal: true
---

# PowerPoint GUI automation

Drives the visible PowerPoint window: foreground control, ribbon KeyTips over SendKeys, and UI
Automation (`InvokePattern`/`TogglePattern`) for what keyboard and mouse cannot reach. The Zoom
fixture (`pslz:sldZm`, `psez:sectionZm`, `psuz:summaryZm`) was authored this way.

**Try COM first** (`powerpoint-fixture-authoring`, including its `ExecuteMso` fallback). Confirm
a feature has no COM surface by checking the real object model, not by assuming.

## Prerequisites

- An interactive console session (`query session` shows `>console ... Active`): a Windows
  machine, or the worker VM's desktop over its web viewer or RDP
  (`tools/powerpoint-vm/README.md`). Session 0 or a disconnected session captures black
  frames; stop and say so. `pnpm ppt:run` is for COM recipes, not this.
- PowerPoint with a visible window (`Presentations.Open(path, 0, 0, -1)`).
- Delete `HKCU:\Software\Microsoft\Office\16.0\PowerPoint\Resiliency\{DocumentRecovery,StartupItems}`
  first; a recovery pane swallows keyboard input.
- Run scripts with `& '...\script.ps1'`, never `-ExecutionPolicy Bypass`. `uia-lib.ps1` needs an
  STA apartment (`-STA` on a raw `pwsh` call). Keep string literals ASCII-only; a non-ASCII
  character in a `.ps1` caused misleading parse errors.

## Workflow

1. **Sanity check.** Run `scripts\foreground-and-shoot.ps1 -Shot .tmp\shot0.png` and read the
   PNG to confirm PowerPoint is in front (it works around the foreground lock).
2. **Discover KeyTips** from a screenshot rather than guessing, one level at a time:
   ```
   & drive-ribbon.ps1 -KeyTips '{ESC};{ESC};%' -Shot .tmp\shot-keytips.png
   ```
3. **Open the menu item and dump the dialog.** `-Dump` prints each control's accessible
   `Name`, which can differ from its visible label ("2. Alpha 1" was `Slide 2 Alpha 1`). Never
   toggle a name you have not seen in a dump; a miss no-ops silently.
   ```
   & drive-ribbon.ps1 -KeyTips '{ESC};{ESC};%;N;Y' -UiaMenuItem 'Slide Zoom' `
       -DialogTitleLike 'Insert *Zoom*' -Dump -Shot .tmp\shot-dlg.png
   ```
4. **Toggle and submit** with the dumped names:
   ```
   & drive-ribbon.ps1 -DialogTitleLike 'Insert *Zoom*' -Toggle 'Slide 2 Alpha 1' `
       -InvokeButton 'Insert' -Shot .tmp\shot-done.png
   ```
5. **Verify the OOXML**, not the screenshot: `& save-and-extract.ps1 -NameLike 'my-fixture*'
   -DestDir .tmp\gui-extract`, then read the slide XML and its `_rels`.
6. **Scrub before committing.** `docProps/core.xml` carries the signed-in user's name and may
   carry add-in residue. Reap only `POWERPNT` processes you started.

## Why UI Automation

Ribbon dropdowns ignore synthetic input: chorded `SendKeys` fires the old Alt accelerators,
arrow keys and synthetic mouse clicks on popup items register but do not activate them. KeyTips
work only as Alt pressed and released, then one `SendWait` per letter (`Send-KeyTipSequence`).
Popup items need UIA `InvokePattern`. Do not retry the input-queue approaches.

## Other gotchas

- Unscoped screenshots and UIA desktop searches leak the user's other windows. Always pass
  `-DialogTitleLike` or a scope; `Save-WindowScreenshot` crops to the target window.
- Starting delays: ~700ms after foregrounding, ~700-900ms between KeyTips, ~1200-1500ms after a
  menu-opening invoke. Re-tune when a screenshot shows a step has not settled.
- Loading `UIAutomationClient` makes the process DPI-aware, so screenshot resolution can jump.
  Take a pixel-comparable "before" shot before dot-sourcing `uia-lib.ps1`.
- `save-and-extract.ps1` extracts to a fresh directory because the sandbox can block a
  `Remove-Item` near XML-looking text.

## Scripts

- `scripts/ppt-window-lib.ps1` - dot-source library: `Get-PptHwnd`,
  `Set-PptForeground` (the `AttachThreadInput` foreground-lock workaround),
  `Save-WindowScreenshot` (window-rect-scoped capture), `Send-KeyTipSequence`.
- `scripts/uia-lib.ps1` - dot-source library: `Find-UiaDialog` (scope by
  window title), `Find-UiaElementByName`, `Get-UiaControlDump`,
  `Invoke-UiaElement`, `Set-UiaToggleOn`.
- `scripts/foreground-and-shoot.ps1` - standalone sanity check / "just look
  at current state" tool.
- `scripts/drive-ribbon.ps1` - the combined driver: KeyTips -> UIA menu
  invoke -> optional dump/toggle/button-invoke -> screenshot. See its
  comment-based help (`Get-Help ...\drive-ribbon.ps1 -Full`) for every
  parameter and worked examples.
- `scripts/save-and-extract.ps1` - saves the presentation via its already-
  running COM instance and extracts the package so you can read the real
  OOXML the GUI action produced, instead of trusting the screenshot alone.
