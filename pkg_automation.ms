/*
PKG Package Script - 3ds Max

Workflow:
1. Enter SERIES NAME (e.g. B779).
2. For each kit row: type Kit ID, pick TYPE
   (BED/HEADBOARD/NIGHTSTAND/CHEST/DRESSER/MIRROR/OTHER or AUTO)
   and choose the matching scene object from the dropdown.
3. Click "+ ADD MORE KIT" to add more rows.
4. Click "REFRESH SCENE OBJECTS" anytime to rescan the scene.
5. Browse to the Excel file containing the package rows.
6. Set FRONT AXIS to the world axis the models' fronts currently face
   in your scene (most furniture libraries use +Y). The script rotates
   each piece so its front faces the camera (-Y world) before applying
   the artistic angle.
7. Set GAP (scene units) between pieces, and the default ANGLES for
   BED / NIGHTSTAND / CHEST / DRESSER.
8. Optional: click "PICK VRAY CAMERA" and pick a template
   VRayPhysicalCamera in the viewport. A copy of the camera+target is
   cloned into every package layer. The cloned camera's X is forced to
   0 to keep the bed centred on screen.
9. Click RUN. A preview dialog lists every package that will be created.
10. Click RUN in the preview dialog to actually run the script.

Bedroom mode (when at least one BED kit is mapped):
   - BED at the center, rotated by ANGLE BED
   - NIGHTSTANDs to the left of the bed. First nightstand gets ANGLE NS,
     each subsequent one alternates with ANGLE NS - 2.5 (the rule from
     the spec: paired nightstands differ by 2.5 degrees so both show
     the same amount of side in screen space).
   - CHESTs then DRESSERs to the right
   - MIRRORs stacked on top of the matching dresser (or appended to
     the right row if no dresser exists)
   - OTHERs placed behind the bed

Headboard mode (no BED, but HEADBOARD is mapped, or auto-detected):
   - All pieces laid out in a single back row, centred on origin
   - Order left to right: NIGHTSTANDs, HEADBOARD, CHESTs, DRESSERs
   - MIRRORs stacked on top of the matching dresser

TYPE = AUTO auto-classifies by bounding-box shape: the single
largest item becomes the bed; remaining are classified by their
aspect ratio (mirror = tall+thin, dresser = wide, chest = tall,
else nightstand).

Excel package format:
   PKG026197 (B779B16,1;B779B1,1;B779-46,1;B779-92,1)
   - Series prefix (B779) is stripped, leading dash also stripped.
   - Number after the comma is the instance quantity.
   - Entries are split by ;
*/

struct PKGItem (kitName, qty)
struct PKGPackage (packageName, items)

rollout pkgPreviewDlg "Confirm Packages" width:460 height:530
(
    label   lblCount   ""       pos:[10,10]  width:440 height:18
    listbox lstPkgs    ""       pos:[10,33]  width:440 height:22
    button  btnRun     "RUN"    pos:[10,488] width:215 height:32
    button  btnCancel  "Cancel" pos:[235,488] width:215 height:32

    on pkgPreviewDlg open do
    (
        local pkgs = pkgRollout.getPendingPackages()
        local items = #()
        for pkg in pkgs do
        (
            local ic = pkg.items.count
            local suffix = if ic == 1 then " kit" else " kits"
            append items (pkg.packageName + "   (" + ic as string + suffix + ")")
        )
        lstPkgs.items = items
        lblCount.text = pkgs.count as string + " package(s) will be created. Click RUN to confirm."
    )

    on btnRun pressed do
    (
        destroyDialog pkgPreviewDlg
        pkgRollout.executePending()
    )

    on btnCancel pressed do
    (
        destroyDialog pkgPreviewDlg
    )
)

rollout pkgRollout "Package Script" width:560 height:650
(
    local noneTag         = "(none)"
    local pendingSeries   = ""
    local pendingPairs    = #()
    local pendingPackages = #()
    local pickedCamera    = undefined

    label lblTitle "PACKAGE SCRIPT" pos:[200,12] width:200 height:18

    label    lblSeries  "SERIES NAME"  pos:[20,44]  width:110
    edittext edtSeries  ""             pos:[140,41] width:400 height:22

    label  lblKitHeader "KIT DETAILS"            pos:[20,74]  width:120
    button btnRefresh   "REFRESH SCENE OBJECTS"  pos:[360,70] width:180 height:24

    dotNetControl dgvKits "System.Windows.Forms.DataGridView" pos:[20,100] width:520 height:380

    dotNetControl btnDelKit "Button" pos:[160,490] width:130 height:30
    dotNetControl btnAddKit "Button" pos:[300,490] width:160 height:30

    label    lblExcel  "EXCEL FILE" pos:[20,530]  width:100
    edittext edtExcel  ""           pos:[125,527] width:320 height:22 readonly:true
    button   btnBrowse "..."        pos:[450,527] width:90  height:22

    label        lblFront   "FRONT AXIS" pos:[20,562]  width:90
    dropdownList ddlFront   items:#("+Y","-Y","+X","-X") pos:[115,558] width:55 selection:2
    label        lblBedHead "BED HEAD"   pos:[180,562] width:60
    dropdownList ddlBedHead items:#("Auto","+Y","-Y","+X","-X") pos:[245,558] width:55 selection:1
    label        lblGap "GAP (in)" pos:[315,562] width:55
    dropdownList ddlGap items:#("1","2","3","4","5") pos:[375,558] width:55 selection:2
    label        lblChestX "CHEST x" pos:[440,562] width:55
    dropdownList ddlChestX items:#("-5","-4.75","-4.5","-4.25","-4","-3.75","-3.5","-3.25","-3","-2.5","-2","-1.5","-1","-0.75","-0.5","-0.25","0","0.25","0.5","0.75","1","2","3","4","5") pos:[500,558] width:55 selection:21

    button btnPickCam   "PICK VRAY CAMERA"    pos:[20,600]  width:160 height:32
    label  lblPickedCam "(no camera picked)"  pos:[190,607] width:255 height:18
    button btnRun       "RUN"                 pos:[455,600] width:85  height:32

    fn trimAll s = trimLeft (trimRight s)

    fn collectSceneObjectNames =
    (
        local raw = #()
        for n in objects do append raw n.name
        sort raw
        local result = #(noneTag)
        for s in raw do append result s
        result
    )

    fn refreshSceneObjects =
    (
        local names = collectSceneObjectNames()
        local savedValues = #()
        local rowCount = dgvKits.RowCount
        for i = 0 to (rowCount - 1) do
        (
            local cell = dgvKits.Rows.Item[i].Cells.Item[2]
            local v = cell.Value
            local s = ""
            try (if v != undefined then s = v as string) catch()
            append savedValues s
        )
        for i = 0 to (rowCount - 1) do
            dgvKits.Rows.Item[i].Cells.Item[2].Value = undefined
        local colObj = dgvKits.Columns.Item[2]
        colObj.Items.Clear()
        for s in names do colObj.Items.Add s
        for i = 0 to (rowCount - 1) do
        (
            local prev = savedValues[i + 1]
            if prev != "" and findItem names prev > 0 then
                dgvKits.Rows.Item[i].Cells.Item[2].Value = prev
        )
    )

    fn buildLookup =
    (
        local pairs = #()
        local rowCount = dgvKits.RowCount
        for i = 0 to (rowCount - 1) do
        (
            local row = dgvKits.Rows.Item[i]
            local kitCell  = row.Cells.Item[0]
            local typeCell = row.Cells.Item[1]
            local objCell  = row.Cells.Item[2]
            local kitStr  = ""
            local typeStr = "AUTO"
            local objStr  = ""
            try (if kitCell.Value  != undefined then kitStr  = trimAll (kitCell.Value as string)) catch()
            try (if typeCell.Value != undefined then typeStr = (typeCell.Value as string))        catch()
            try (if objCell.Value  != undefined then objStr  = objCell.Value as string)           catch()
            if typeStr == "" then typeStr = "AUTO"
            if kitStr != "" and objStr != "" and objStr != noneTag then
            (
                local n = getNodeByName objStr
                if n != undefined and isValidNode n then
                    append pairs #(kitStr, n, typeStr)
            )
        )
        pairs
    )

    fn lookupNode pairs key =
    (
        for p in pairs where p[1] == key do return p[2]
        undefined
    )

    fn lookupType pairs key =
    (
        for p in pairs where p[1] == key do return p[3]
        "AUTO"
    )

    fn parseKitEntry raw =
    (
        local commaPos = findString raw ","
        if commaPos == 0 then return (PKGItem kitName:raw qty:1)
        local nm = trimAll (substring raw 1 (commaPos - 1))
        local qtyText = trimAll (substring raw (commaPos + 1) (raw.count - commaPos))
        local qty = 1
        try (qty = qtyText as integer) catch(qty = 1)
        if qty == undefined or qty < 1 then qty = 1
        PKGItem kitName:nm qty:qty
    )

    fn parsePackageLine line =
    (
        local txt = trimAll line
        if txt == "" then return undefined
        local openPos = findString txt "("
        local closePos = findString txt ")"
        if openPos == 0 or closePos == 0 or closePos <= openPos then return undefined
        local pkgName = trimAll (substring txt 1 (openPos - 1))
        local itemsRaw = substring txt (openPos + 1) (closePos - openPos - 1)
        if pkgName == "" or itemsRaw == "" then return undefined
        local cleaned = substituteString itemsRaw " " ""
        local parts = filterString cleaned ";"
        local items = #()
        for part in parts where part != "" do
            append items (parseKitEntry part)
        if items.count == 0 then return undefined
        PKGPackage packageName:pkgName items:items
    )

    fn stripSeriesPrefix kitStr series =
    (
        local rest = kitStr
        if series != "" and matchPattern kitStr pattern:(series + "*") ignoreCase:false then
            rest = substring kitStr (series.count + 1) -1
        if rest.count > 0 and rest[1] == "-" then
            rest = substring rest 2 -1
        rest
    )

    fn readExcelLines filePath =
    (
        local rows = #()
        local excelApp = undefined
        local workbook = undefined
        try
        (
            excelApp = createOleObject "Excel.Application"
            excelApp.Visible = false
            excelApp.DisplayAlerts = false
            workbook = excelApp.Workbooks.Open filePath
            local sheet = workbook.Worksheets.Item 1
            local used = sheet.UsedRange
            local rc = used.Rows.Count
            local cc = used.Columns.Count
            for r = 1 to rc do
            (
                local line = ""
                for c = 1 to cc do
                (
                    local cell = sheet.Cells.Item r c
                    local v = cell.Value2
                    if v != undefined then
                    (
                        local s = v as string
                        if s != "" then
                        (
                            if line != "" then line += " "
                            line += s
                        )
                    )
                )
                if line != "" then append rows line
            )
        )
        catch
        (
            local errInfo = getCurrentException()
            messageBox ("Failed to read Excel file.\n\nDetails:\n" + (errInfo as string) + "\n\nIf Excel.exe is hung, end it from Task Manager and try again.") title:"Excel Error"
        )
        if workbook != undefined then try (workbook.Close false) catch()
        if excelApp != undefined then try (excelApp.Quit()) catch()
        if excelApp != undefined then try (releaseOLEObject excelApp) catch()
        rows
    )

    fn safeCreateHiddenLayer layerName =
    (
        local layer = undefined
        try (layer = LayerManager.getLayerFromName layerName) catch()
        if layer == undefined then layer = LayerManager.newLayerFromName layerName
        try (layer.isHidden = true) catch()
        layer
    )

    fn collectHierarchy node =
    (
        local result = #(node)
        local idx = 1
        while idx <= result.count do
        (
            local cur = result[idx]
            for c in cur.children do append result c
            idx += 1
        )
        result
    )

    -- Clone the entire hierarchy of src and re-link parents so that
    -- the resulting clones keep the same parent/child structure as the
    -- source. This is what fixes the "missing models" bug: a group head
    -- alone does not always drag its members along on cloneNodes, and
    -- linked hierarchies definitely do not. Passing the full hierarchy
    -- and then re-parenting in the new set guarantees everything is
    -- there and attached.
    fn cloneAsInstance src =
    (
        local sources = collectHierarchy src
        local newNodes
        try (maxOps.cloneNodes sources cloneType:#instance newNodes:&newNodes) catch(newNodes = undefined)
        if newNodes == undefined or newNodes.count == 0 then return #()
        if newNodes.count == sources.count then
        (
            for i = 1 to sources.count do
            (
                local oldNode = sources[i]
                if oldNode.parent != undefined then
                (
                    local pIdx = findItem sources oldNode.parent
                    if pIdx > 0 then
                        try (newNodes[i].parent = newNodes[pIdx]) catch()
                )
            )
        )
        newNodes
    )

    -- In the returned newNodes find the cloned root: the node whose
    -- parent is not in the cloned set. Falls back to first element.
    fn findCloneRoot nodes =
    (
        for n in nodes do
        (
            if n.parent == undefined then return n
            local pInSet = false
            for other in nodes do
                if other == n.parent then ( pInSet = true ; exit )
            if not pInSet then return n
        )
        nodes[1]
    )

    fn groupUnionBBox topNode =
    (
        if not isGroupHead topNode then return #(topNode.min, topNode.max)
        local mn = [1e30, 1e30, 1e30]
        local mx = [-1e30, -1e30, -1e30]
        local found = false
        local visit = #(topNode)
        while visit.count > 0 do
        (
            local cur = visit[1]
            deleteItem visit 1
            for c in cur.children do
            (
                if isGroupHead c then append visit c
                else if isValidNode c then
                (
                    local dmn = c.min
                    local dmx = c.max
                    if dmn.x < mn.x then mn.x = dmn.x
                    if dmn.y < mn.y then mn.y = dmn.y
                    if dmn.z < mn.z then mn.z = dmn.z
                    if dmx.x > mx.x then mx.x = dmx.x
                    if dmx.y > mx.y then mx.y = dmx.y
                    if dmx.z > mx.z then mx.z = dmx.z
                    found = true
                )
            )
        )
        if not found then return #(topNode.min, topNode.max)
        #(mn, mx)
    )

    fn bboxVolume n =
    (
        local bb = groupUnionBBox n
        local d = bb[2] - bb[1]
        (abs d.x) * (abs d.y) * (abs d.z)
    )

    -- Z rotation that brings a non-bed model's front face to -Y world
    -- (camera position assumed to be at -Y, looking +Y, per spec).
    fn frontAxisOffsetDeg axisStr =
    (
        local v = 0.0
        case axisStr of
        (
            "+Y": v = 180.0
            "-Y": v = 0.0
            "+X": v = -90.0
            "-X": v = 90.0
            default: v = 0.0
        )
        v
    )

    -- Z rotation that brings the bed's headboard to +Y world (and
    -- therefore the foot to -Y world, facing the camera). The dropdown
    -- value says where the headboard currently sits in the model at
    -- identity rotation; the script then rotates the bed so that
    -- direction ends up at +Y world.
    fn bedHeadOffsetDeg headStr =
    (
        local v = 0.0
        case headStr of
        (
            "+Y": v = 0.0
            "-Y": v = 180.0
            "+X": v = 90.0
            "-X": v = -90.0
            default: v = 0.0
        )
        v
    )

    -- Auto-detect the angle (CCW from +X) at which the headboard sits
    -- relative to the bed's center, while the bed is at identity
    -- rotation. The headboard is taken to be the geometry node with
    -- the highest top (max Z) - in most bed models that is the
    -- vertical headboard panel.
    fn detectBedHeadAngleDeg bedNode =
    (
        local bb = groupUnionBBox bedNode
        local center = (bb[1] + bb[2]) / 2.0
        local geomNodes = #()
        local visit = #(bedNode)
        while visit.count > 0 do
        (
            local cur = visit[1]
            deleteItem visit 1
            for c in cur.children do
            (
                if isGroupHead c then append visit c
                else if isValidNode c then append geomNodes c
            )
        )
        if geomNodes.count == 0 then return 90.0
        local maxZ = -1e30
        local hbCenter = center
        for c in geomNodes do
        (
            local cmax = c.max
            if cmax.z > maxZ then
            (
                maxZ = cmax.z
                hbCenter = (c.min + c.max) / 2.0
            )
        )
        local dx = hbCenter.x - center.x
        local dy = hbCenter.y - center.y
        if (abs dx) < 0.001 and (abs dy) < 0.001 then return 90.0
        atan2 dy dx
    )

    -- Reset node's rotation to identity, recenter so its bbox center is
    -- at the origin, rotate around Z by (baseRotZ + rotAngle), then
    -- move so the rotated bbox lands at targetPos. XY is aligned to
    -- the bbox center, Z is aligned to the bbox MIN (the base of the
    -- model) so the piece sits on targetPos.z (= floor level when 0)
    -- instead of floating with half the model below the floor.
    fn placeAndRotate node targetPos rotAngle baseRotZ =
    (
        try (node.rotation = (quat 0 0 0 1)) catch()
        local bb = groupUnionBBox node
        local bbCenter = (bb[1] + bb[2]) / 2.0
        node.pos = node.pos - bbCenter
        try (rotate node (eulerangles 0 0 (baseRotZ + rotAngle))) catch()
        local bb2 = groupUnionBBox node
        local cx = (bb2[1].x + bb2[2].x) / 2.0
        local cy = (bb2[1].y + bb2[2].y) / 2.0
        local baseZ = bb2[1].z
        node.pos = node.pos + [targetPos.x - cx, targetPos.y - cy, targetPos.z - baseZ]
    )

    fn classifyByBBox n =
    (
        local bb = groupUnionBBox n
        local d = bb[2] - bb[1]
        local w  = abs d.x
        local de = abs d.y
        local h  = abs d.z
        local mx = w
        if de > mx then mx = de
        local mn = w
        if de < mn then mn = de
        if mx == 0 then return "OTHER"
        if h > mx * 1.5 and mn < mx * 0.4 then return "MIRROR"
        if mx > mn * 2.0 then return "DRESSER"
        if h > mx and mn < mx * 0.8 then return "CHEST"
        "NIGHTSTAND"
    )

    -- Stack a mirror on top of a dresser. The mirror's base lands on
    -- the dresser's top (placeAndRotate aligns base Z to targetPos.z).
    fn stackMirrorOnDresser mirrorNode dresserNode baseRotZ ang =
    (
        local dbb = groupUnionBBox dresserNode
        local cx = (dbb[1].x + dbb[2].x) / 2.0
        local cy = (dbb[1].y + dbb[2].y) / 2.0
        local topZ = dbb[2].z
        placeAndRotate mirrorNode [cx, cy, topZ] ang baseRotZ
    )

    -- Width/depth/height of node's bbox at identity rotation. Used so
    -- we can compute rotated-rectangle corners regardless of the
    -- node's current world transform.
    fn getLocalSize node =
    (
        local savedRot = node.rotation
        try (node.rotation = (quat 0 0 0 1)) catch()
        local bb = groupUnionBBox node
        try (node.rotation = savedRot) catch()
        [abs (bb[2].x - bb[1].x), abs (bb[2].y - bb[1].y), abs (bb[2].z - bb[1].z)]
    )

    -- The 4 world XY corners of a rectangle of `size`, centred at
    -- `centerXY`, rotated by `angDeg` around Z. Order:
    -- 1=BL(-,-), 2=BR(+,-), 3=TR(+,+), 4=TL(-,+).
    fn cornersOfRotatedRect size centerXY angDeg =
    (
        local hw = size.x / 2.0
        local hd = size.y / 2.0
        local c = cos angDeg
        local s = sin angDeg
        #(
            [centerXY.x + (-hw)*c - (-hd)*s, centerXY.y + (-hw)*s + (-hd)*c],
            [centerXY.x + ( hw)*c - (-hd)*s, centerXY.y + ( hw)*s + (-hd)*c],
            [centerXY.x + ( hw)*c - ( hd)*s, centerXY.y + ( hw)*s + ( hd)*c],
            [centerXY.x + (-hw)*c - ( hd)*s, centerXY.y + (-hw)*s + ( hd)*c]
        )
    )

    fn cornerMaxX cs =
    (
        local r = cs[1]
        for c in cs do if c.x > r.x then r = c
        r
    )

    fn cornerMinX cs =
    (
        local r = cs[1]
        for c in cs do if c.x < r.x then r = c
        r
    )

    -- GAP dropdown value (inches as string) -> cm as integer, rounded.
    -- 2" -> 5, 3" -> 8, 4" -> 10, 5" -> 13.
    fn gapToCm =
    (
        local inchVal = 2
        try (inchVal = (ddlGap.selected as integer)) catch()
        if inchVal == undefined then inchVal = 2
        ((inchVal * 2.54) + 0.5) as integer
    )

    -- Horizontal FOV (degrees) of the picked camera. Works for VRay
    -- Physical (focal_length + film_gate), Autodesk Physical
    -- (focal_length + film_width), and standard target cameras (fov).
    fn cameraHFOVDeg cam =
    (
        local fl = undefined
        local fg = undefined
        try (fl = cam.focal_length) catch()
        try (fg = cam.film_gate) catch()
        if fg == undefined then try (fg = cam.film_width) catch()
        if fl != undefined and fg != undefined and fl > 0 then
            2.0 * (atan (fg / (2.0 * fl)))
        else
        (
            local f = 45.0
            try (f = cam.fov) catch()
            f
        )
    )

    -- Angle for the i-th piece on a given side (1-based).
    -- Single piece on a side: 35. Multiple pieces: closest to anchor
    -- = 32.5, every other = 35. Right side gets the negated values.
    fn sideAngle i totalCount isRight =
    (
        local a = 0.0
        if totalCount == 1 then a = 35.0
        else if i == 1 then a = 32.5
        else a = 35.0
        if isRight then -a else a
    )

    fn arrangeBedMode typed anchorItem baseRot bedBaseRot gap =
    (
        local angAnchor = 35.0

        -- Effective rotation for the anchor piece (BED auto-detects
        -- the headboard; HEADBOARD falls back to FRONT AXIS).
        local effAnchorRot = bedBaseRot
        local anchorSize = [200.0, 100.0, 0.0]
        local anchorTotal = 0.0
        if anchorItem != undefined then
        (
            if anchorItem[2] == "BED" and ddlBedHead.selected == "Auto" then
            (
                local savedRot = anchorItem[1].rotation
                try (anchorItem[1].rotation = (quat 0 0 0 1)) catch()
                local detected = detectBedHeadAngleDeg anchorItem[1]
                try (anchorItem[1].rotation = savedRot) catch()
                effAnchorRot = 90.0 - detected
            )
            else if anchorItem[2] == "HEADBOARD" then
                effAnchorRot = baseRot

            anchorSize = getLocalSize anchorItem[1]
            placeAndRotate anchorItem[1] [0, 0, 0] angAnchor effAnchorRot
            anchorTotal = effAnchorRot + angAnchor
        )

        -- Bucket the other pieces.
        local nightstands = #()
        local chests = #()
        local dressers = #()
        local mirrors = #()
        local others = #()
        for t in typed do
        (
            case t[2] of
            (
                "BED": ()
                "HEADBOARD": ()
                "NIGHTSTAND": append nightstands t
                "CHEST": append chests t
                "DRESSER": append dressers t
                "MIRROR": append mirrors t
                default: append others t
            )
        )

        -- Anchor's leftmost / rightmost world corner is what the
        -- first neighbour's facing corner has to meet (with a gap).
        local anchorCorners = cornersOfRotatedRect anchorSize [0, 0] anchorTotal
        local lastLeft  = cornerMinX anchorCorners
        local lastRight = cornerMaxX anchorCorners

        -- LEFT side: nightstands, closest to anchor first.
        local nsCount = nightstands.count
        for i = 1 to nsCount do
        (
            local ang = sideAngle i nsCount false
            local sz  = getLocalSize nightstands[i][1]
            local tot = baseRot + ang
            local cs0 = cornersOfRotatedRect sz [0, 0] tot
            local rt0 = cornerMaxX cs0
            local cw  = [lastLeft.x - gap - rt0.x, lastLeft.y - rt0.y]
            placeAndRotate nightstands[i][1] [cw.x, cw.y, 0] ang baseRot
            local csF = cornersOfRotatedRect sz cw tot
            lastLeft = cornerMinX csF
        )

        -- RIGHT side: chests then dressers, closest to anchor first.
        -- Build the order and the per-piece angle in lockstep so we
        -- can give the matching mirror the dresser's exact angle.
        local rightItems = #()
        for c in chests do append rightItems c
        for d in dressers do append rightItems d
        local rightCount = rightItems.count
        local rightAngles = #()
        for i = 1 to rightCount do append rightAngles (sideAngle i rightCount true)

        -- When the bed-side neighbour on the right is a CHEST and
        -- there is something past it (e.g. a dresser), align the
        -- chest's BR (front-right) corner to the bed's BR with 4 x
        -- gap of X offset. This shifts the chest UP along Y so it
        -- sits behind the bed's foot line, matching the reference
        -- layout. Single piece on the right, or a non-chest first
        -- piece, keep the normal corner-to-corner placement.
        local chestSpecialMode = (rightCount >= 2 and chests.count >= 1 \
                                  and rightItems[1][2] == "CHEST")
        for i = 1 to rightCount do
        (
            local ang = rightAngles[i]
            local sz  = getLocalSize rightItems[i][1]
            local tot = baseRot + ang
            local cs0 = cornersOfRotatedRect sz [0, 0] tot
            local cw
            if i == 1 and chestSpecialMode then
            (
                -- corners order from cornersOfRotatedRect: 1=BL 2=BR 3=TR 4=TL.
                -- Y: align chest's BR with bed's BR (lifts the chest up so
                --    its body sits behind the bed's foot line).
                -- X: keep the leftmost (BL) corner clear of the bed's right
                --    edge by (chestMult * gap) so the user can dial in the
                --    visible horizontal spacing from the panel.
                local chestMult = 1.0
                try (chestMult = (ddlChestX.selected as float)) catch()
                if chestMult == undefined then chestMult = 1.0
                local chestBL = cs0[1]
                local chestBR = cs0[2]
                cw = [lastRight.x + gap * chestMult - chestBL.x, lastRight.y - chestBR.y]
            )
            else
            (
                local lf0 = cornerMinX cs0
                cw = [lastRight.x + gap - lf0.x, lastRight.y - lf0.y]
            )
            placeAndRotate rightItems[i][1] [cw.x, cw.y, 0] ang baseRot
            local csF = cornersOfRotatedRect sz cw tot
            lastRight = cornerMaxX csF
        )

        -- Mirrors stack on top of the matching dresser, using that
        -- dresser's actual angle so the two pieces line up.
        local mIdx = 1
        for m in mirrors do
        (
            if mIdx <= dressers.count then
            (
                local rIdx = chests.count + mIdx
                local dAng = if rIdx <= rightAngles.count then rightAngles[rIdx] else -35.0
                stackMirrorOnDresser m[1] dressers[mIdx][1] baseRot dAng
            )
            mIdx += 1
        )

        -- Loose / unclassified pieces: drop them in a back row behind
        -- the anchor, no artistic angle.
        if others.count > 0 then
        (
            local behindY = (anchorSize.y / 2.0) + gap * 2.0
            local cursorO = -(anchorSize.x / 2.0)
            for o in others do
            (
                local sz = getLocalSize o[1]
                placeAndRotate o[1] [cursorO + sz.x / 2.0, behindY + sz.y / 2.0, 0] 0.0 baseRot
                cursorO += (sz.x + gap)
            )
        )
    )

    fn arrangeBedroom items =
    (
        local baseRot    = frontAxisOffsetDeg ddlFront.selected
        local bedBaseRot = bedHeadOffsetDeg ddlBedHead.selected
        local gap        = gapToCm()

        -- Split explicit types from AUTO.
        local typed = #()
        local autoItems = #()
        for it in items do
        (
            local t = toUpper it[2]
            if t == "AUTO" then append autoItems it
            else append typed #(it[1], t)
        )

        local hasBed = false
        local hasHB  = false
        for t in typed do
        (
            if t[2] == "BED" do hasBed = true
            if t[2] == "HEADBOARD" do hasHB = true
        )

        if autoItems.count > 0 then
        (
            local volSorted = #()
            for a in autoItems do append volSorted a
            for i = 1 to volSorted.count - 1 do
                for j = 1 to volSorted.count - i do
                    if (bboxVolume volSorted[j][1]) < (bboxVolume volSorted[j+1][1]) then
                    (
                        local tmp = volSorted[j]
                        volSorted[j] = volSorted[j+1]
                        volSorted[j+1] = tmp
                    )
            local startIdx = 1
            if not hasBed and not hasHB and volSorted.count > 0 then
            (
                append typed #(volSorted[1][1], "BED")
                hasBed = true
                startIdx = 2
            )
            for i = startIdx to volSorted.count do
                append typed #(volSorted[i][1], classifyByBBox volSorted[i][1])
        )

        -- Anchor: BED first, else HEADBOARD. Both go through the
        -- same corner-touching layout - the only difference is how
        -- the anchor's base rotation is derived.
        local anchor = undefined
        if hasBed then
            for t in typed do if t[2] == "BED" do anchor = t
        else if hasHB then
            for t in typed do if t[2] == "HEADBOARD" do anchor = t

        arrangeBedMode typed anchor baseRot bedBaseRot gap
    )

    -- Clone the picked template camera + its target (all camera
    -- properties — focal length, film gate, ISO, etc. — carry over
    -- because we use #copy), add both to the supplied layer, rename
    -- them after the package, and position them to frame the package:
    --
    --   Camera X      = 0 (locked)
    --   Camera Z      = 100 + kitCount * 20 (cm)
    --   Camera Y      = -distance, auto-computed from template's
    --                   horizontal FOV so the package fits in frame
    --   Target X      = 0 (locked)
    --   Target Y      = 0 (locked)
    --   Target Z      = package bbox vertical mid (free, auto-centred)
    fn cloneCameraToLayer layer pkgName pkgNodes kitCount =
    (
        if pickedCamera == undefined or not isValidNode pickedCamera then return undefined
        local sources = #(pickedCamera)
        local tgt = undefined
        try (tgt = pickedCamera.target) catch()
        if tgt != undefined and isValidNode tgt then append sources tgt
        local newNodes
        try (maxOps.cloneNodes sources cloneType:#copy newNodes:&newNodes) catch(newNodes = undefined)
        if newNodes == undefined or newNodes.count == 0 then return undefined
        for n in newNodes where isValidNode n do
            try (layer.addNode n) catch()
        try
        (
            if newNodes.count >= 1 then newNodes[1].name = uniqueName (pkgName + "_Cam")
            if newNodes.count >= 2 then newNodes[2].name = uniqueName (pkgName + "_Cam_Target")
        )
        catch()

        if newNodes.count >= 2 then
        (
            -- Package bbox in world.
            local mn = [1e30, 1e30, 1e30]
            local mx = [-1e30, -1e30, -1e30]
            local found = false
            for n in pkgNodes where isValidNode n do
            (
                local nmn = n.min
                local nmx = n.max
                if nmn.x < mn.x then mn.x = nmn.x
                if nmn.y < mn.y then mn.y = nmn.y
                if nmn.z < mn.z then mn.z = nmn.z
                if nmx.x > mx.x then mx.x = nmx.x
                if nmx.y > mx.y then mx.y = nmx.y
                if nmx.z > mx.z then mx.z = nmx.z
                found = true
            )

            if found then
            (
                local camNode = newNodes[1]
                local tgtNode = newNodes[2]

                local pkgW = mx.x - mn.x
                local pkgD = mx.y - mn.y
                local pkgH = mx.z - mn.z
                local tgtZ = (mn.z + mx.z) / 2.0

                local camZ = 100.0 + (kitCount as float) * 20.0

                -- Pick the larger of width and height as the dim that
                -- has to fit horizontally in frame.
                local fitDim = if pkgW > pkgH then pkgW else pkgH

                -- Horizontal FOV from the template camera (degrees).
                local hFOV = cameraHFOVDeg camNode
                local halfFOV = hFOV / 2.0
                local tanHalf = tan halfFOV
                local distance = 500.0
                if tanHalf > 0.0001 then
                    distance = (fitDim / 2.0) / tanHalf
                -- Push back by half the package depth so the near
                -- face of the package still fits, plus 15% margin.
                distance = (distance + pkgD / 2.0) * 1.15

                try (camNode.pos = [0, -distance, camZ]) catch()
                try (tgtNode.pos = [0, 0, tgtZ]) catch()
            )
        )

        newNodes
    )

    fn buildPackage pkg series pairs missingOut =
    (
        local typedRoots = #()
        local createdAll = #()
        for item in pkg.items do
        (
            local kitKey = stripSeriesPrefix item.kitName series
            local src = lookupNode pairs kitKey
            local typ = lookupType pairs kitKey
            if src == undefined or not isValidNode src then
                appendIfUnique missingOut kitKey
            else
            (
                for i = 1 to item.qty do
                (
                    local newNodes = #()
                    try (newNodes = cloneAsInstance src) catch(newNodes = #())
                    if newNodes.count > 0 then
                    (
                        local root = findCloneRoot newNodes
                        if root != undefined and isValidNode root then
                        (
                            append typedRoots #(root, typ)
                            for n in newNodes where isValidNode n do
                                appendIfUnique createdAll n
                        )
                    )
                )
            )
        )
        if createdAll.count == 0 then return 0
        arrangeBedroom typedRoots
        local layer = safeCreateHiddenLayer pkg.packageName
        -- Add every cloned node (root + descendants) so that hiding the
        -- layer hides everything together. Adding only the root leaves
        -- members visible because each MAX node carries its own layer.
        for n in createdAll where isValidNode n do
            try (layer.addNode n) catch()
        local roots = for tr in typedRoots collect tr[1]
        local grp = undefined
        try (grp = group roots name:pkg.packageName) catch()
        if grp != undefined then try (layer.addNode grp) catch()
        cloneCameraToLayer layer pkg.packageName createdAll typedRoots.count
        createdAll.count
    )

    fn getSettingsPath = "C:\\PKG SCRIPT\\pkg_automation_settings.ini"

    fn saveSettings =
    (
        local p = getSettingsPath()
        try
        (
            setINISetting p "main" "Series"       edtSeries.text
            setINISetting p "main" "ExcelPath"    edtExcel.text
            setINISetting p "main" "FrontAxis"    ddlFront.selected
            setINISetting p "main" "BedHead"      ddlBedHead.selected
            setINISetting p "main" "GapInch"      ddlGap.selected
            setINISetting p "main" "ChestGapMult" ddlChestX.selected
            local camName = ""
            if pickedCamera != undefined and isValidNode pickedCamera then camName = pickedCamera.name
            setINISetting p "main" "CameraName" camName

            local rowCount = dgvKits.RowCount
            -- Clear any old kit entries first so removed rows don't linger.
            try
            (
                local oldKeys = getINISetting p "kits"
                for k in oldKeys do delIniSetting p "kits" k
            )
            catch()
            setINISetting p "kits" "Count" (rowCount as string)
            for i = 0 to (rowCount - 1) do
            (
                local row = dgvKits.Rows.Item[i]
                local kit = ""
                local typ = ""
                local obj = ""
                try (if row.Cells.Item[0].Value != undefined then kit = row.Cells.Item[0].Value as string) catch()
                try (if row.Cells.Item[1].Value != undefined then typ = row.Cells.Item[1].Value as string) catch()
                try (if row.Cells.Item[2].Value != undefined then obj = row.Cells.Item[2].Value as string) catch()
                local idxStr = (i + 1) as string
                setINISetting p "kits" ("Kit" + idxStr + "_ID")   kit
                setINISetting p "kits" ("Kit" + idxStr + "_Type") typ
                setINISetting p "kits" ("Kit" + idxStr + "_Obj")  obj
            )
        )
        catch()
        OK
    )

    fn loadSettings =
    (
        local p = getSettingsPath()
        if not doesFileExist p then return undefined
        try
        (
            local s = getINISetting p "main" "Series"
            if s != "" then edtSeries.text = s
            local e = getINISetting p "main" "ExcelPath"
            if e != "" then edtExcel.text = e

            local axisItems    = #("+Y","-Y","+X","-X")
            local bedHeadItems = #("Auto","+Y","-Y","+X","-X")
            local fa = getINISetting p "main" "FrontAxis"
            local fi = findItem axisItems fa
            if fi > 0 then ddlFront.selection = fi
            local bh = getINISetting p "main" "BedHead"
            local bi = findItem bedHeadItems bh
            if bi > 0 then ddlBedHead.selection = bi

            local gapItems = #("1","2","3","4","5")
            local gi = getINISetting p "main" "GapInch"
            local gIdx = findItem gapItems gi
            if gIdx > 0 then ddlGap.selection = gIdx

            local chestMultItems = #("-5","-4.75","-4.5","-4.25","-4","-3.75","-3.5","-3.25","-3","-2.5","-2","-1.5","-1","-0.75","-0.5","-0.25","0","0.25","0.5","0.75","1","2","3","4","5")
            local cmStr = getINISetting p "main" "ChestGapMult"
            local cmIdx = findItem chestMultItems cmStr
            if cmIdx > 0 then ddlChestX.selection = cmIdx

            local cn = getINISetting p "main" "CameraName"
            if cn != "" then
            (
                local n = getNodeByName cn
                if n != undefined and isValidNode n then
                (
                    pickedCamera = n
                    lblPickedCam.text = "Camera: " + n.name
                )
            )

            local cs = getINISetting p "kits" "Count"
            local count = 0
            if cs != "" then try (count = cs as integer) catch(count = 0)
            if count > 0 then
            (
                dgvKits.Rows.Clear()
                for i = 1 to count do dgvKits.Rows.Add()
                for i = 1 to count do
                (
                    local row = dgvKits.Rows.Item[i-1]
                    local idxStr = i as string
                    local kit = getINISetting p "kits" ("Kit" + idxStr + "_ID")
                    local typ = getINISetting p "kits" ("Kit" + idxStr + "_Type")
                    local obj = getINISetting p "kits" ("Kit" + idxStr + "_Obj")
                    try (if kit != "" then row.Cells.Item[0].Value = kit) catch()
                    try (if typ != "" then row.Cells.Item[1].Value = typ) catch()
                    try (if obj != "" then row.Cells.Item[2].Value = obj) catch()
                )
            )
        )
        catch()
        OK
    )

    -- Cell-edit Enter fix: by default the Enter key inside a cell
    -- editor bubbles up to the dialog and triggers RUN, which loses
    -- whatever the user just typed. Commit explicitly on Enter and
    -- swallow the key so the rollout never sees it.
    -- Use e.KeyValue (int) rather than (e.KeyCode as integer) - the
    -- .NET Keys enum can't be cast directly to MAXScript integer.
    fn editorKeyDown s e =
    (
        local kv = 0
        try (kv = e.KeyValue) catch()
        if kv == 13 then
        (
            try (dgvKits.EndEdit()) catch()
            try (dgvKits.ClearSelection()) catch()
            try (e.SuppressKeyPress = true) catch()
            try (e.Handled = true) catch()
        )
    )

    fn getPendingPackages = pendingPackages

    fn executePending =
    (
        local series   = pendingSeries
        local pairs    = pendingPairs
        local packages = pendingPackages

        local createdCount = 0
        local skippedCount = 0
        local missingKits = #()

        undo "PKG Run" on
        (
            for pkg in packages do
            (
                local n = buildPackage pkg series pairs missingKits
                if n > 0 then createdCount += 1
                else skippedCount += 1
            )
        )

        local msg = "Created " + createdCount as string + " package layer(s)."
        if skippedCount > 0 then
            msg += "\nSkipped " + skippedCount as string + " (no mapped objects matched)."
        if missingKits.count > 0 then
        (
            msg += "\n\nUnmapped Kit IDs:"
            for m in missingKits do msg += "\n  - " + m
        )
        if pickedCamera == undefined then
            msg += "\n\nNote: No template VRayPhysicalCamera was picked, so no cameras were added to layers."
        messageBox msg title:"Run Complete"
    )

    on pkgRollout open do
    (
        dgvKits.AllowUserToAddRows = false
        dgvKits.AllowUserToDeleteRows = true
        dgvKits.AllowUserToResizeRows = false
        dgvKits.RowHeadersVisible = true
        dgvKits.RowHeadersWidth = 28
        dgvKits.MultiSelect = false
        dgvKits.AutoSizeColumnsMode = (dotNetClass "System.Windows.Forms.DataGridViewAutoSizeColumnsMode").Fill
        dgvKits.SelectionMode = (dotNetClass "System.Windows.Forms.DataGridViewSelectionMode").FullRowSelect
        dgvKits.ScrollBars = (dotNetClass "System.Windows.Forms.ScrollBars").Vertical

        local colKit = dotNetObject "System.Windows.Forms.DataGridViewTextBoxColumn"
        colKit.HeaderText = "KIT ID"
        colKit.FillWeight = 22
        dgvKits.Columns.Add colKit

        local colType = dotNetObject "System.Windows.Forms.DataGridViewComboBoxColumn"
        colType.HeaderText = "TYPE"
        colType.FillWeight = 28
        local typeOptions = #("AUTO", "BED", "HEADBOARD", "NIGHTSTAND", "CHEST", "DRESSER", "MIRROR", "OTHER")
        for s in typeOptions do colType.Items.Add s
        dgvKits.Columns.Add colType

        local colObj = dotNetObject "System.Windows.Forms.DataGridViewComboBoxColumn"
        colObj.HeaderText = "SELECT OBJECT"
        colObj.FillWeight = 50
        dgvKits.Columns.Add colObj

        local names = collectSceneObjectNames()
        for s in names do colObj.Items.Add s

        dgvKits.Rows.Add()
        dgvKits.Rows.Add()

        local colorCls = dotNetClass "System.Drawing.Color"
        local flatStyle = (dotNetClass "System.Windows.Forms.FlatStyle").Flat
        local boldFont = dotNetObject "System.Drawing.Font" "Tahoma" 9.0 \
                         ((dotNetClass "System.Drawing.FontStyle").Bold)

        btnDelKit.Text      = "DELETE KIT"
        btnDelKit.BackColor = colorCls.FromArgb 178 34 34
        btnDelKit.ForeColor = colorCls.White
        btnDelKit.FlatStyle = flatStyle
        btnDelKit.Font      = boldFont

        btnAddKit.Text      = "+ ADD MORE KIT"
        btnAddKit.BackColor = colorCls.FromArgb 60 130 60
        btnAddKit.ForeColor = colorCls.White
        btnAddKit.FlatStyle = flatStyle
        btnAddKit.Font      = boldFont

        loadSettings()
    )

    on pkgRollout close do saveSettings()

    on dgvKits DataError s e do (e.ThrowException = false)

    on dgvKits EditingControlShowing s e do
    (
        local ctrl = e.Control
        try (dotNet.removeAllEventHandlers ctrl "KeyDown") catch()
        dotNet.addEventHandler ctrl "KeyDown" editorKeyDown
    )

    on btnRefresh pressed do refreshSceneObjects()

    on btnAddKit click s e do dgvKits.Rows.Add()

    on btnDelKit click s e do
    (
        try
        (
            local rowIdx = -1
            if dgvKits.SelectedRows.Count > 0 then
                rowIdx = dgvKits.SelectedRows.Item[0].Index
            else if dgvKits.CurrentCell != undefined then
                rowIdx = dgvKits.CurrentCell.RowIndex
            if rowIdx >= 0 and rowIdx < dgvKits.RowCount then
            (
                dgvKits.Rows.RemoveAt rowIdx
                saveSettings()
            )
        )
        catch()
    )

    on btnBrowse pressed do
    (
        local fp = getOpenFileName caption:"Select Excel file" \
            types:"Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*|"
        if fp != undefined then
        (
            edtExcel.text = fp
            saveSettings()
        )
    )

    on btnPickCam pressed do
    (
        local pick = pickObject message:"Pick a VRayPhysicalCamera (its target will be copied too)"
        if pick != undefined and isValidNode pick then
        (
            pickedCamera = pick
            lblPickedCam.text = "Camera: " + pick.name
            saveSettings()
        )
    )

    on btnRun pressed do
    (
        -- Force the grid to commit any cell that is still in edit mode,
        -- otherwise the value the user just typed/picked is invisible
        -- to buildLookup and we report it as "unmapped".
        try (dgvKits.EndEdit()) catch()
        saveSettings()

        local series = trimAll edtSeries.text
        if series == "" then
        (
            messageBox "Enter the SERIES NAME first." title:"Missing Series"
            return undefined
        )
        local pairs = buildLookup()
        if pairs.count == 0 then
        (
            messageBox "Add at least one Kit ID and choose a scene object from the dropdown." title:"No Kits"
            return undefined
        )
        local excelPath = edtExcel.text
        if excelPath == "" or not doesFileExist excelPath then
        (
            messageBox "Browse to a valid Excel file." title:"Missing Excel"
            return undefined
        )
        local lines = readExcelLines excelPath
        if lines.count == 0 then
        (
            messageBox "No data was read from the Excel file." title:"Empty"
            return undefined
        )

        local packages = #()
        for line in lines do
        (
            local pkg = parsePackageLine line
            if pkg != undefined then append packages pkg
        )
        if packages.count == 0 then
        (
            messageBox "No valid package rows found in Excel." title:"No Packages"
            return undefined
        )

        pendingSeries   = series
        pendingPairs    = pairs
        pendingPackages = packages

        try (destroyDialog pkgPreviewDlg) catch()
        createDialog pkgPreviewDlg style:#(#dialog, #minimize, #sysmenu)
    )
)

createDialog pkgRollout
