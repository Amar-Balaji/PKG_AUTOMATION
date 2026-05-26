# PKG Automation — Developer Guide

3ds Max MAXScript that batch-creates Ashley furniture package layers from an Excel mastersheet. Each row in the sheet becomes one layer containing instanced furniture, arranged in a stylised top-down composition, with its own VRay Physical Camera + target aimed at the package.

Single file: [pkg_automation.ms](pkg_automation.ms). One rollout (`pkgRollout`), one preview dialog (`pkgPreviewDlg`).

---

## 1. End-user workflow

1. Type the **SERIES NAME** (e.g. `B779`) — strips the series prefix from each kit ID in the Excel rows.
2. For each kit row in the grid: type the **Kit ID**, pick a **TYPE** (`AUTO` / `BED` / `HEADBOARD` / `NIGHTSTAND` / `CHEST` / `DRESSER` / `MIRROR` / `OTHER`), pick the matching scene object from the dropdown.
3. Browse to the **Excel file**.
4. Set **FRONT AXIS** (which world axis non-bed models face at identity), **BED HEAD** (which way the bed's headboard points — leave on `Auto`), and **GAP** (corner-to-corner spacing in inches; backend converts to cm).
5. **PICK VRAY CAMERA** — point at a template VRayPhysicalCamera. Its focal length / film gate / ISO / exposure all carry through to every cloned camera.
6. **RUN** → preview dialog lists every package that will be created → second RUN to commit.

All panel state is auto-saved to `C:\PKG SCRIPT\pkg_automation_settings.ini` and restored next time the rollout opens.

---

## 2. Excel format

Each row, anywhere in sheet 1, looks like:

```
PKG026197 (B779B16,1;B779B1,1;B779-46,1;B779-92,1)
```

Parser ([parsePackageLine](pkg_automation.ms#L207)):
- Package name = everything before the `(`.
- Inside the parens: entries split by `;`.
- Each entry: `kitId,quantity`.
- The series prefix (`B779`) is stripped from each kit ID, then a leading dash too — so `B779-92` becomes the lookup key `92`.

Quantity `2` means the same source object is cloned twice. `92(2)` in the sheet → `92,2` after the parser → two clones of the kit-92 source object → counts as **2 pieces** in the kit-count formula (see camera section).

---

## 3. Scene-graph plumbing

### Cloning
[cloneAsInstance](pkg_automation.ms#L283) clones the source as `#instance` — instances stay linked to the original, so material/mesh edits propagate. Critically it clones **the entire hierarchy** (`collectHierarchy`) and re-links parents inside the cloned set, otherwise grouped beds drop their children silently.

### Layer per package
[safeCreateHiddenLayer](pkg_automation.ms#L262) creates (or fetches) a layer named exactly after the package code, marked hidden. Every cloned node + the wrapping group are added to that layer so toggling the eye icon hides the whole package at once.

### Group head
Each package's clones are wrapped in a `group … name:pkg.packageName` so users can move/scale the whole package as one unit.

---

## 4. Layout algorithm (the interesting bit)

[arrangeBedMode](pkg_automation.ms#L530) handles **both** BED and HEADBOARD anchored packages. There's no separate headboard-mode function anymore; the only difference is how the anchor's base rotation is derived.

### 4.1 Anchor placement

- The anchor (BED or HEADBOARD) goes at world origin.
- Its **artistic rotation** is **35°** (hard-coded, was a UI option before).
- Its **base rotation** = whatever's needed to orient the model correctly:
  - **BED + BED HEAD = Auto**: temporarily reset rotation, scan the bed's children, find the geometry with the highest `max.z` (that's the headboard panel), compute the angle from bed center to headboard center → `baseRot = 90° − headAngle` so the headboard ends up at world +Y.
  - **BED + explicit axis (+Y/-Y/+X/-X)**: user says where the headboard is at identity; `bedHeadOffsetDeg` returns the rotation to put it at +Y.
  - **HEADBOARD anchor**: uses `FRONT AXIS` (`frontAxisOffsetDeg`) instead, because a standalone headboard has no separate "headboard" geometry to detect.

### 4.2 Side-piece angles ([sideAngle](pkg_automation.ms#L519))

Mirror-symmetric across the bed:

| Side  | 1 piece | 2+ pieces |
|-------|---------|-----------|
| Left  | +35°    | closest = +32.5°, rest = +35° |
| Right | −35°    | closest = −32.5°, rest = −35° |

This produces the "V opening toward the camera" silhouette in the Ashley reference shots.

### 4.3 Corner-touching placement

Not center-to-center — *corner-to-corner with `gap`*. For each side piece, in order from anchor outward:

1. Compute the piece's **rotated rectangle corners** with [cornersOfRotatedRect](pkg_automation.ms#L443), assuming the piece is centred at origin with its full `baseRot + sideAngle`.
2. The corner **facing the anchor** is the one we care about:
   - Left side → piece's max-X corner ([cornerMaxX](pkg_automation.ms#L455))
   - Right side → piece's min-X corner ([cornerMinX](pkg_automation.ms#L462))
3. Compute the world-space target for that corner:
   - Left side → `[lastLeftCorner.x − gap, lastLeftCorner.y]`
   - Right side → `[lastRightCorner.x + gap, lastRightCorner.y]`
   - `lastLeftCorner` / `lastRightCorner` start from the anchor's leftmost/rightmost corner.
4. Translate the piece's centre so its facing corner lands at that target. **Y follows from the corner alignment** — pieces sit at slightly different Y from the anchor, which is what makes the diagonal "kiss" look in the reference image.
5. Update `lastLeftCorner` / `lastRightCorner` to this piece's far-side corner for the next iteration.

### 4.4 [placeAndRotate](pkg_automation.ms#L405)

The workhorse called for every piece. Steps:

1. Reset `node.rotation` to identity.
2. Compute `groupUnionBBox`, translate node so bbox centre is at origin (decouples placement from the model's pivot — pivots can be anywhere).
3. Apply the combined rotation `(baseRotZ + rotAngle)` around Z.
4. Recompute bbox after rotation, then translate so:
   - **XY**: bbox centre lands at `targetPos.xy`.
   - **Z**: bbox **MIN** (the base of the model) lands at `targetPos.z`. This is why pieces sit on the floor (Z=0) instead of being half-buried — `targetPos.z` is always passed as `0` for ground pieces.

[stackMirrorOnDresser](pkg_automation.ms#L431) passes the dresser's `bbox.max.z` as `targetPos.z` so the mirror's base lands on the dresser top.

### 4.5 Mirrors and others

- **Mirrors**: stacked on top of the matching dresser, taking the dresser's exact angle so the two pieces line up.
- **OTHER** pieces: dropped in a back row behind the anchor, no artistic angle.

---

## 5. Camera

[cloneCameraToLayer](pkg_automation.ms#L646), called from [buildPackage](pkg_automation.ms#L750) once the layout is done.

### Cloning
- `maxOps.cloneNodes` with `cloneType:#copy` on `[camera, target]`. This preserves every camera attribute (focal length, film gate, ISO, F-number, white balance, exposure, etc.) — *do not change camera properties, only its position*.
- Renamed `<PKG>_Cam` and `<PKG>_Cam_Target`.
- Both added to the package layer so they're hidden with the rest of the package.

### Positioning

Hard rules from the spec:

| Property      | Value                              | Note                                     |
|---------------|------------------------------------|------------------------------------------|
| Camera X      | `0`                                | locked                                   |
| Camera Z      | `100 + kitCount × 20` cm           | 1=120, 2=140, 3=160, 4=180, 5=200…       |
| Camera Y      | `−distance` (auto, see below)      | "front/back can move" per user spec      |
| Target X      | `0`                                | locked                                   |
| Target Y      | `0`                                | locked                                   |
| Target Z      | package bbox vertical mid          | auto-centres package in viewport         |

**Distance calculation** ([cameraHFOVDeg](pkg_automation.ms#L477) reads the FOV from the template camera):

```
distance = (max(pkgWidth, pkgHeight) / 2) / tan(hFOV / 2)
distance = (distance + pkgDepth / 2) × 1.15
```

- `max(W, H)` because we need the larger of the two horizontal/vertical extents to fit.
- `+ pkgDepth/2` pushes back so the near face of the package fits too (perspective makes the near face appear larger).
- `× 1.15` is a 15 % safety margin so the frame doesn't graze the silhouette.

**FOV resolution** — tries property names in order:
1. `focal_length` + `film_gate` → `2 × atan(film_gate / (2 × focal_length))` (VRay Physical)
2. `focal_length` + `film_width` (Autodesk Physical Camera)
3. `.fov` (standard target/free camera)
4. Falls back to 45° if none of those resolve.

**Kit count** = `typedRoots.count` = total clones in this package (so `92(2)` adds 2 to the count, not 1).

---

## 6. Settings persistence

INI file at `C:\PKG SCRIPT\pkg_automation_settings.ini` via Max's built-in `setINISetting` / `getINISetting`.

`[main]` section: `Series`, `ExcelPath`, `FrontAxis`, `BedHead`, `GapInch`, `CameraName`.
`[kits]` section: `Count`, then `Kit<i>_ID`, `Kit<i>_Type`, `Kit<i>_Obj` per row.

Saved automatically on:
- RUN
- Camera pick
- Excel browse
- Dialog close

Loaded at the end of `on pkgRollout open` — *after* the grid columns are set up and the two default rows are added, so the saved row count can replace them via `dgvKits.Rows.Clear()`.

Camera name → if the saved name doesn't resolve to a scene node, the field stays empty (silent fallback). Same for kit-grid object names — invalid entries silently fall back to blank cells.

---

## 7. DataGridView gotchas

### Enter-key cell commit
Default dotNet behaviour bubbles Enter up to the dialog, which fires RUN, which fires before the cell value has been committed — the typed value vanishes.

Fix ([editorKeyDown](pkg_automation.ms#L778) + [`on dgvKits EditingControlShowing`](pkg_automation.ms#L884)):
- Every time a cell enters edit mode, hook the editor's `KeyDown` event.
- On Enter (`KeyValue == 13`, **not** `KeyCode as integer` — that throws because the `Keys` enum can't auto-cast): call `dgvKits.EndEdit()` explicitly, then set `SuppressKeyPress = true` and `Handled = true` so Enter never reaches the rollout.

The `removeAllEventHandlers ctrl "KeyDown"` call before `addEventHandler` prevents duplicate hooks when the same editor instance is reused across edits.

### `dotNet.addEventHandler` signature
Takes **exactly 3 positional args**: `(object, eventName, handler)`. **No `id:` keyword arg** — that's a `removeAllEventHandlers`-only thing. Passing `id:#foo` to `addEventHandler` throws "wanted 3, got 6".

### DataError
`on dgvKits DataError s e do (e.ThrowException = false)` swallows errors when a saved combo value isn't in the current dropdown items list — without this, the rollout would popup an exception every time a stale scene-object name is restored from the settings file.

---

## 8. Unit convention

**Everything is in centimetres** (the user's 3ds Max scene is configured for cm). The only inch values are the GAP dropdown labels (`2"`, `3"`, `4"`, `5"`); they're converted by [gapToCm](pkg_automation.ms#L466):

```
gap_cm = round(inch × 2.54)
```

So 2"→5, 3"→8, 4"→10, 5"→13. Integer, no decimals.

Camera Z values (120/140/160/180/200…) are cm.

---

## 9. Quick map: where things live

| Area                          | Function / lines                                       |
|-------------------------------|--------------------------------------------------------|
| Rollout / UI                  | [pkgRollout](pkg_automation.ms#L57) — top of file       |
| Settings (save/load)          | [saveSettings](pkg_automation.ms#L308) / [loadSettings](pkg_automation.ms#L361) |
| Excel parsing                 | [parsePackageLine](pkg_automation.ms#L207), [parseKitEntry](pkg_automation.ms#L194) |
| Excel COM read                | [readExcelLines](pkg_automation.ms#L226)               |
| Bed-head auto-detect          | [detectBedHeadAngleDeg](pkg_automation.ms#L342)        |
| Layout                        | [arrangeBedMode](pkg_automation.ms#L530), [arrangeBedroom](pkg_automation.ms#L675) |
| Corner math                   | [cornersOfRotatedRect](pkg_automation.ms#L443), [cornerMaxX](pkg_automation.ms#L455), [cornerMinX](pkg_automation.ms#L462) |
| Place + rotate                | [placeAndRotate](pkg_automation.ms#L405)               |
| Camera                        | [cloneCameraToLayer](pkg_automation.ms#L646), [cameraHFOVDeg](pkg_automation.ms#L477) |
| Build one package             | [buildPackage](pkg_automation.ms#L735)                 |
| Run all                       | [executePending](pkg_automation.ms#L778)               |
| Preview dialog                | [pkgPreviewDlg](pkg_automation.ms#L24)                 |

---

## 10. Common modification recipes

**Change angle preset (e.g. left side should be 30° / 27.5° instead of 35° / 32.5°)**
Edit [sideAngle](pkg_automation.ms#L519). One function, four magic numbers.

**Change camera height formula**
Edit the `camZ = …` line in [cloneCameraToLayer](pkg_automation.ms#L646).

**Add a new TYPE bucket (e.g. BENCH)**
1. Add `"BENCH"` to `typeOptions` in `on pkgRollout open`.
2. Add `case "BENCH": append benches t` in the bucketing loop in `arrangeBedMode`.
3. Add a placement loop for `benches` (likely cloned from the right-side loop).

**Change GAP options to centimetres directly**
1. Change `dropdownList ddlGap items:#(…)` to your cm values.
2. Replace `gapToCm` body with `(ddlGap.selected as integer)`.
3. Update the label to `"GAP (cm)"`.

**Stop the camera from auto-positioning** (keep template's pos)
Comment out the `camNode.pos = …` / `tgtNode.pos = …` lines at the bottom of [cloneCameraToLayer](pkg_automation.ms#L646).

---

## 11. Known limitations

- Bed-head auto-detect picks the geometry with the highest top — works for beds with a tall headboard, fails for platform beds with no headboard (falls back to "headboard at +Y"). Override with an explicit BED HEAD axis in that case.
- All packages share one camera template — if you need a different focal length per package, pick a new camera between runs.
- The corner-touching layout assumes pieces have approximately rectangular footprints. Very irregular shapes (e.g. an L-shaped dresser) will gap unevenly because we compute the local bbox rather than the actual convex hull.
- Single-mesh beds (no group hierarchy) can't be auto-detected — `detectBedHeadAngleDeg` returns the fallback 90°. Use an explicit BED HEAD axis.
- Camera distance calc uses horizontal FOV only. Render aspect ratio isn't read — if you render very tall (portrait) frames, the top/bottom may clip; pad the GAP or widen the camera FOV.
