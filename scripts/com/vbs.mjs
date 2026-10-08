/**
 * The VBScripts that drive desktop PowerPoint, one per deck.
 *
 * Every builder returns a complete `.vbs` source string: {@link vbsOpenHeader} opens the deck and
 * prints `OPEN_OK` or `OPEN_ERR`, the middle reads the construct under test back out over COM, and
 * {@link vbsFooter} closes and quits. The entry point ships the string and the deck as one
 * PowerPoint job, runs it under `cscript`, and hands the stdout lines to a verifier.
 *
 * No absolute path is spliced in. A builder takes the deck's file name, the script finds the deck
 * beside itself through `WScript.ScriptFullName`, and every PNG it exports is written beside it
 * too and named relative to it on stdout. So the same script runs in whatever workspace the job
 * lands in, on this machine or in a VM, and the PNGs come back as the job's output files.
 *
 * String building, not string escaping: every value spliced in here is a repo constant or a file
 * name checked by {@link deckFileName}, never caller input.
 */

import { MODEL3D_EXPORT, PRSTGEOM_CASES, PRSTGEOM_EXPORT } from './contract.mjs'

// --- 3. the VBScripts that drive PowerPoint ---------------------------------
/**
 * Shared header: create the app, open the deck, emit OPEN_OK/OPEN_ERR.
 *
 * Default is headless + read-only. Pass `withWindow` for checks that read OLE objects back:
 * a windowless PowerPoint does not instantiate embedded objects, so `Shapes` comes back without
 * them — indistinguishable from PowerPoint having discarded them. (A deck PowerPoint authored
 * itself enumerates zero shapes under a headless open too, which is how that was pinned down.)
 */
/**
 * A deck's file name, checked to be one a VBScript string literal can carry as is.
 * @param {string} name
 * @returns {string}
 */
export function deckFileName(name) {
	if (!/^[\w.-]+\.pptx$/i.test(name)) throw new Error(`not a plain .pptx file name: ${JSON.stringify(name)}`)
	return name
}

/**
 * @param {string} deckName the deck's file name, in the script's own folder
 * @param {boolean} [withWindow] open with a window, for the features a headless open will not instantiate
 * @returns {string}
 */
export function vbsOpenHeader(deckName, withWindow = false) {
	// WithWindow:=msoFalse (0) keeps it headless; ReadOnly avoids touching the file.
	const openArgs = withWindow ? '0, 0, -1' : '-1, 0, 0'
	return `Option Explicit
Dim ppt, pres, sld, shp, fso, here
Set fso = CreateObject("Scripting.FileSystemObject")
here = fso.GetParentFolderName(WScript.ScriptFullName)
On Error Resume Next
Set ppt = CreateObject("PowerPoint.Application")
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "NO_POWERPOINT\t" & Err.Description
  WScript.Quit 3
End If
Err.Clear
Set pres = ppt.Presentations.Open(fso.BuildPath(here, "${deckFileName(deckName)}"), ${openArgs})
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "OPEN_ERR\t" & Hex(Err.Number) & "\t" & Err.Description
  ppt.Quit
  WScript.Quit 2
End If
WScript.StdOut.WriteLine "OPEN_OK\t" & pres.Slides.Count
`
}

export function vbsFooter() {
	return `pres.Close
ppt.Quit
WScript.StdOut.WriteLine "DONE"
WScript.Quit 0
`
}

/** @param {string} deckName @returns {string} */
export function buildNavVbs(deckName) {
	// Emits tab-separated `ACTION` lines (slideIdx, objectName, resolvedActionNum).
	return (
		vbsOpenHeader(deckName) +
		`Dim act
For Each sld In pres.Slides
  For Each shp In sld.Shapes
    Set act = shp.ActionSettings(1) ' ppMouseClick = 1
    If Not act Is Nothing Then
      If act.Action <> 0 Then
        WScript.StdOut.WriteLine "ACTION\t" & sld.SlideIndex & "\t" & shp.Name & "\t" & act.Action
      End If
    End If
  Next
Next
` +
		vbsFooter()
	)
}

/** @param {string} deckName @returns {string} */
export function buildGeomVbs(deckName) {
	// Emits one tab-separated `CONN` line per connector shape:
	//   CONN <name> <beginConnected(-1/0)> <beginConnectedShapeName> <beginConnectionSite>
	return (
		vbsOpenHeader(deckName) +
		`Dim cf, tgt
For Each sld In pres.Slides
  For Each shp In sld.Shapes
    If shp.Connector Then
      Set cf = shp.ConnectorFormat
      tgt = ""
      If cf.BeginConnected Then tgt = cf.BeginConnectedShape.Name
      WScript.StdOut.WriteLine "CONN\t" & shp.Name & "\t" & cf.BeginConnected & "\t" & tgt & "\t" & cf.BeginConnectionSite
    End If
  Next
Next
` +
		vbsFooter()
	)
}

/** @param {string} deckName @returns {string} */
export function buildOleVbs(deckName) {
	// Emits one tab-separated `OLE` line per embedded-object shape: name, resolved ProgID.
	// msoEmbeddedOLEObject = 7, msoLinkedOLEObject = 10.
	return (
		vbsOpenHeader(deckName, true) +
		`Dim i, j
For i = 1 To pres.Slides.Count
  Set sld = pres.Slides(i)
  For j = 1 To sld.Shapes.Count
    Set shp = sld.Shapes(j)
    If shp.Type = 7 Or shp.Type = 10 Then
      WScript.StdOut.WriteLine "OLE\t" & shp.Name & "\t" & shp.OLEFormat.ProgID
    End If
  Next
Next
` +
		vbsFooter()
	)
}

/** @param {string} deckName @returns {string} */
export function buildModel3dVbs(deckName) {
	// Emits `M3D <name> <Shape.Type> <Model3D.CameraPositionZ>` per shape, then exports slide 1 to
	// PNG so the caller can prove the model actually rasterized. A window is needed here for the
	// same reason as OLE: a headless PowerPoint does not instantiate the 3D renderer.
	const png = deckFileName(deckName).replace(/\.pptx$/i, '.png')
	return (
		vbsOpenHeader(deckName, true) +
		`Dim j, camZ
Set sld = pres.Slides(1)
For j = 1 To sld.Shapes.Count
  Set shp = sld.Shapes(j)
  camZ = ""
  On Error Resume Next
  camZ = shp.Model3D.CameraPositionZ
  On Error Goto 0
  WScript.StdOut.WriteLine "M3D" & vbTab & shp.Name & vbTab & shp.Type & vbTab & camZ
Next
Err.Clear
sld.Export fso.BuildPath(here, "${png}"), "PNG", ${MODEL3D_EXPORT.w}, ${MODEL3D_EXPORT.h}
If Err.Number <> 0 Then
  WScript.StdOut.WriteLine "EXPORT_ERR" & vbTab & Hex(Err.Number) & vbTab & Err.Description
Else
  WScript.StdOut.WriteLine "EXPORT" & vbTab & "${png}"
End If
` +
		vbsFooter()
	)
}

/** @param {string} deckName @returns {string} */
export function buildPresetGeomVbs(deckName) {
	// Exports every slide to `<deck>-<n>.png` and emits one `PNG <n> <name>` line each. Nothing is
	// read back over COM on purpose: `Shape.Adjustments` reports the *stored* guide, out-of-range
	// value and all, so it cannot answer what PowerPoint paints. Only the pixels can.
	const base = deckFileName(deckName).replace(/\.pptx$/i, '')
	return (
		vbsOpenHeader(deckName) +
		`Dim i, png
For i = 1 To pres.Slides.Count
  png = "${base}-" & i & ".png"
  Err.Clear
  On Error Resume Next
  pres.Slides(i).Export fso.BuildPath(here, png), "PNG", ${PRSTGEOM_EXPORT.w}, ${PRSTGEOM_EXPORT.h}
  If Err.Number <> 0 Then
    WScript.StdOut.WriteLine "EXPORT_ERR" & vbTab & i & vbTab & Hex(Err.Number) & vbTab & Err.Description
  Else
    WScript.StdOut.WriteLine "PNG" & vbTab & i & vbTab & png
  End If
  On Error Goto 0
Next
WScript.StdOut.WriteLine "SLIDES" & vbTab & ${PRSTGEOM_CASES.length}
` +
		vbsFooter()
	)
}
