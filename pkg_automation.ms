/*
PKG Package Script - 3ds Max

Workflow:
1. Enter SERIES NAME (e.g. B779).
2. Type a Kit ID in each row (e.g. B16, B1, 46, 92) and choose the
   matching scene object from the dropdown.
3. Click "+ ADD MORE KIT" to add more rows.
4. Click "REFRESH SCENE OBJECTS" anytime to rescan the scene.
5. Browse to the Excel file containing the package rows.
6. Click RUN. A preview dialog lists every package that will be created.
7. Click RUN in the preview dialog to actually run the script.

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

rollout pkgRollout "Package Script" width:560 height:615
(
    local noneTag         = "(none)"
    local pendingSeries   = ""
    local pendingPairs    = #()
    local pendingPackages = #()

    label lblTitle "PACKAGE SCRIPT" pos:[200,12] width:200 height:18

    label    lblSeries  "SERIES NAME"  pos:[20,44]  width:110
    edittext edtSeries  ""             pos:[140,41] width:400 height:22

    label  lblKitHeader "KIT DETAILS"            pos:[20,74]  width:120
    button btnRefresh   "REFRESH SCENE OBJECTS"  pos:[360,70] width:180 height:24

    dotNetControl dgvKits "System.Windows.Forms.DataGridView" pos:[20,100] width:520 height:380

    button btnAddKit "+ ADD MORE KIT" pos:[200,490] width:160 height:30

    label    lblExcel  "EXCEL FILE" pos:[20,530]  width:100
    edittext edtExcel  ""           pos:[125,527] width:320 height:22 readonly:true
    button   btnBrowse "..."        pos:[450,527] width:90  height:22

    button btnRun "RUN" pos:[455,568] width:85 height:32

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
            local cell = dgvKits.Rows.Item[i].Cells.Item[1]
            local v = cell.Value
            local s = ""
            try (if v != undefined then s = v as string) catch()
            append savedValues s
        )
        for i = 0 to (rowCount - 1) do
            dgvKits.Rows.Item[i].Cells.Item[1].Value = undefined
        local colObj = dgvKits.Columns.Item[1]
        colObj.Items.Clear()
        for s in names do colObj.Items.Add s
        for i = 0 to (rowCount - 1) do
        (
            local prev = savedValues[i + 1]
            if prev != "" and findItem names prev > 0 then
                dgvKits.Rows.Item[i].Cells.Item[1].Value = prev
        )
    )

    fn buildLookup =
    (
        local pairs = #()
        local rowCount = dgvKits.RowCount
        for i = 0 to (rowCount - 1) do
        (
            local row = dgvKits.Rows.Item[i]
            local kitCell = row.Cells.Item[0]
            local objCell = row.Cells.Item[1]
            local kitStr = ""
            local objStr = ""
            try (if kitCell.Value != undefined then kitStr = trimAll (kitCell.Value as string)) catch()
            try (if objCell.Value != undefined then objStr = objCell.Value as string) catch()
            if kitStr != "" and objStr != "" and objStr != noneTag then
            (
                local n = getNodeByName objStr
                if n != undefined and isValidNode n then
                    append pairs #(kitStr, n)
            )
        )
        pairs
    )

    fn lookupNode pairs key =
    (
        for p in pairs where p[1] == key do return p[2]
        undefined
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

    fn arrangeRow nodes =
    (
        local nextX = 0.0
        for n in nodes where isValidNode n do
        (
            local bbox = groupUnionBBox n
            local bmin = bbox[1]
            local bmax = bbox[2]
            local offset = [nextX - bmin.x, 0, 0]
            n.pos = n.pos + offset
            nextX += (bmax.x - bmin.x) + 20.0
        )
    )

    fn buildPackage pkg series pairs missingOut =
    (
        local createdRoots = #()
        local createdAll   = #()
        for item in pkg.items do
        (
            local kitKey = stripSeriesPrefix item.kitName series
            local src = lookupNode pairs kitKey
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
                            append createdRoots root
                            for n in newNodes where isValidNode n do
                                appendIfUnique createdAll n
                        )
                    )
                )
            )
        )
        if createdAll.count == 0 then return 0
        arrangeRow createdRoots
        local layer = safeCreateHiddenLayer pkg.packageName
        -- Add every cloned node (root + descendants) so that hiding the
        -- layer hides everything together. Adding only the root leaves
        -- members visible because each MAX node carries its own layer.
        for n in createdAll where isValidNode n do
            try (layer.addNode n) catch()
        local grp = undefined
        try (grp = group createdRoots name:pkg.packageName) catch()
        if grp != undefined then try (layer.addNode grp) catch()
        createdAll.count
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
        colKit.FillWeight = 30
        dgvKits.Columns.Add colKit

        local colObj = dotNetObject "System.Windows.Forms.DataGridViewComboBoxColumn"
        colObj.HeaderText = "SELECT OBJECT"
        colObj.FillWeight = 70
        dgvKits.Columns.Add colObj

        local names = collectSceneObjectNames()
        for s in names do colObj.Items.Add s

        dgvKits.Rows.Add()
        dgvKits.Rows.Add()
    )

    on dgvKits DataError s e do (e.ThrowException = false)

    on btnRefresh pressed do refreshSceneObjects()

    on btnAddKit pressed do dgvKits.Rows.Add()

    on btnBrowse pressed do
    (
        local fp = getOpenFileName caption:"Select Excel file" \
            types:"Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|All Files (*.*)|*.*|"
        if fp != undefined then edtExcel.text = fp
    )

    on btnRun pressed do
    (
        -- Force the grid to commit any cell that is still in edit mode,
        -- otherwise the value the user just typed/picked is invisible
        -- to buildLookup and we report it as "unmapped".
        try (dgvKits.EndEdit()) catch()

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
