'===========  modDiagramCore  ===================================
' REQUIRES: Tools ? References ? “Microsoft Visual Basic for
'           Applications Extensibility 5.3” (VBIDE)

Option Explicit

'--------------- CONFIG --------------------------------------
Private Const PROC_RX As String = "^(?:Sub|Function|Property)\s+(\w+)"
Private Const CALL_RX As String = "\b(?:Call\s+)?(\w+)\s*(?=\()"
'-------------------------------------------------------------
Sub Diag_GenerateCallGraph()
    '-- 1. Gather caller ? callee mapping
    Dim cg As Object                 'Scripting.Dictionary
    Set cg = GetVbaCallGraph()       'assumes the helper now returns a dictionary

    '-- 2. Spin up Visio + blank doc
    Dim visApp As Object, visDoc As Object, pg As Object
    Set visApp = PrepareVisio("CallGraph")
    Set visDoc = visApp.ActiveDocument
    Set pg = visDoc.Pages(1)

    '-- 3. Dictionaries to track containers and shapes
    Dim containers As Object: Set containers = CreateObject("Scripting.Dictionary")
    Dim shapes     As Object: Set shapes = CreateObject("Scripting.Dictionary")

    '-- 4. Walk every caller
    Dim caller As Variant, callee As Variant
    Dim modName As String, mod2 As String

        For Each caller In cg.Keys
        modName = Split(CStr(caller), ".")(0)
        If Not containers.Exists(modName) Then
            Set containers(modName) = DropModuleContainer(pg, modName)
        End If
        If Not shapes.Exists(caller) Then
            Set shapes(caller) = DropProcedureShape(containers(modName), Split(caller, ".")(1))
        End If

        '-- single, correct loop over callees
        For Each callee In cg(caller)
            Dim fullCallee As String
            If InStr(callee, ".") = 0 Then
                fullCallee = modName & "." & callee
            Else
                fullCallee = CStr(callee)
            End If
            
            Dim calleeMod As String
            calleeMod = Split(fullCallee, ".")(0)
            
            If Not containers.Exists(calleeMod) Then
                Set containers(calleeMod) = DropModuleContainer(pg, calleeMod)
            End If
            If Not shapes.Exists(fullCallee) Then
                Set shapes(fullCallee) = DropProcedureShape(containers(calleeMod), Split(fullCallee, ".")(1))
            End If
            
            ConnectShapes pg, shapes(caller), shapes(fullCallee)
        Next callee
    Next caller
    
    '-- 5. Layout + save
    HierarchicalLayout pg
    SaveAndOptionallyOpen visDoc, Environ$("USERPROFILE") & "\Desktop\invSys_CallGraph.vsdx", True
End Sub

'return Dictionary: "Module.Proc" ? Collection(calleeName, ...)
Public Function GetVbaCallGraph() As Object
    Dim cg As Object: Set cg = CreateObject("Scripting.Dictionary")
    
    ' A-1 gather declare-ed API names so we can ignore them later
    Dim apiIgnore As Object: Set apiIgnore = CollectApiDeclares()
    
    ' B-1 built-ins to skip (VBA funcs + worksheet funcs)
    Dim builtIns As Object: Set builtIns = BuiltInNameSet()
    
    Dim vbComp As VBIDE.VBComponent
    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        ScanComponent vbComp, cg, builtIns, apiIgnore
    Next vbComp
    
    Set GetVbaCallGraph = cg
End Function


'==================== IMPLEMENTATION =========================
Private Sub ScanComponent(vbComp As VBIDE.VBComponent, _
                          ByRef cg As Object, _
                          ByRef builtIns As Object, _
                          ByRef apiIgnore As Object)
    
    Dim cm As VBIDE.CodeModule: Set cm = vbComp.CodeModule
    Dim total&: total = cm.CountOfLines
    If total = 0 Then Exit Sub
    
    Dim txt As String: txt = cm.Lines(1, total)
    
    'Regex for procedure headers
    Dim reProc As Object: Set reProc = CreateObject("VBScript.RegExp")
    reProc.Global = True: reProc.IgnoreCase = True: reProc.MultiLine = True
    reProc.Pattern = PROC_RX
    
    Dim m, procName$, start&, count&, body$
    For Each m In reProc.Execute(txt)
        procName = m.SubMatches(0)
        start = cm.ProcStartLine(procName, vbext_pk_Proc)
        count = cm.ProcCountLines(procName, vbext_pk_Proc)
        body = cm.Lines(start, count)
        
        Dim callerKey$: callerKey = vbComp.name & "." & procName
        Dim callees As Object: Set callees = ParseProcedureBody(body, builtIns, apiIgnore)
        cg(callerKey) = callees.Keys   'store Keys array for iteration
    Next m
End Sub


'-- stage A: clean strings, comments, continuations; stage B/C: filter
Private Function ParseProcedureBody(ByVal body As String, _
                                     builtIns As Object, _
                                     apiIgnore As Object) As Object
    'strip comment portions
    Dim reComments As Object: Set reComments = CreateObject("VBScript.RegExp")
    reComments.Pattern = "'[^\" & vbCrLf & "]*"   'everything after '
    reComments.Global = True
    body = reComments.Replace(body, "")
    
    'join line continuations
    body = Replace(body, " _" & vbCrLf, " ")
    
    'strip string literals
    Dim reStrings As Object: Set reStrings = CreateObject("VBScript.RegExp")
    reStrings.Pattern = """(?:[^""]|"""")*"""
    reStrings.Global = True
    body = reStrings.Replace(body, "")
    
    'find calls
    Dim reCall As Object: Set reCall = CreateObject("VBScript.RegExp")
    reCall.Pattern = CALL_RX
    reCall.Global = True: reCall.IgnoreCase = True
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    Dim m, callee$
    For Each m In reCall.Execute(body)
        callee = m.SubMatches(0)
        If Not builtIns.Exists(LCase$(callee)) _
           And Not apiIgnore.Exists(LCase$(callee)) Then
            dict(callee) = True
        End If
    Next m
    
    Set ParseProcedureBody = dict
End Function

'--- stage B helper: list of built-in procs to ignore -----------
Private Function BuiltInNameSet() As Object
    Dim s As Object: Set s = CreateObject("Scripting.Dictionary")
    
    'array of lower-case intrinsic names (add more any time)
    Dim arr As Variant
    arr = Split( _
      "abs,array,asc,atn,cbool,cbyte,ccur,cdate,cdbl,chr,cint,clng,cstr," & _
      "cos,createobject,exp,filelen,fix,format,hex,inputbox,instr,int,join," & _
      "lbound,lcase,len,log,ltrim,mid,now,oct,replace,round,rtrim,scriptengine," & _
      "sgn,sin,space,sqr,strcomp,string,trim,typename,ubound,ucase,val,variance," & _
      "worksheetfunction,application", ",")

    Dim item As Variant          'loop variable **must be Variant**
    For Each item In arr
        s(item) = True
    Next item
    
    Set BuiltInNameSet = s
End Function



'--- stage C helper: collect API declares so we ignore them -----
Private Function CollectApiDeclares() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim cmp As VBIDE.VBComponent, cm As VBIDE.CodeModule
    Dim total&, txt$, re As Object, m
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "Declare\s+(?:PtrSafe\s+)?(?:Function|Sub)\s+(\w+)"
    re.IgnoreCase = True: re.Global = True
    
    For Each cmp In ThisWorkbook.VBProject.VBComponents
        Set cm = cmp.CodeModule
        total = cm.CountOfLines
        If total = 0 Then GoTo nxt
        txt = cm.Lines(1, total)
        For Each m In re.Execute(txt)
            d(LCase$(m.SubMatches(0))) = True
        Next m
nxt:
    Next cmp
    Set CollectApiDeclares = d
End Function

'================== VISIO HELPERS (unchanged) ==================
'? keep your existing PrepareVisio, DropModuleContainer, etc.
'===============================================================

'--- 1. Create or reuse a running Visio instance --------------
Public Function PrepareVisio(docTag As String) As Object
    Dim visApp As Object
    Set visApp = CreateObject("Visio.Application")   'late-bound ? always works
    visApp.Visible = True
    visApp.Documents.Add ""                          'blank drawing
    Set PrepareVisio = visApp
End Function

'--- 2. Hierarchical page layout ------------------------------
Public Sub HierarchicalLayout(pg As Object)
    On Error Resume Next
    pg.Layout 1          '1 = visLayoutHierarchical (late-bound enum)
End Sub

'--- 3. Save (and optionally close) ---------------------------
Public Sub SaveAndOptionallyOpen(visDoc As Object, fPath As String, openIt As Boolean)
    visDoc.SaveAs fPath
    If Not openIt Then visDoc.Application.Quit
End Sub
'==============================================================

'---------------------------------------------------------------------
' Robust: US & Metric stencils
'---------------------------------------------------------------------
Public Function DropModuleContainer(pg As Object, modName As String) As Object
    Const visBuiltInStencilContainers As Long = 2   'container stencil
    Const visMSUS As Long = 0                       'use current units
    Const visOpenHidden As Long = 64

    Static mMaster As Object                        'cache after first lookup
    If mMaster Is Nothing Then
        Dim stencilPath$, stn As Object, tryName As Variant
        stencilPath = pg.Application.GetBuiltInStencilFile( _
                          visBuiltInStencilContainers, visMSUS)
        Set stn = pg.Application.Documents.OpenEx(stencilPath, visOpenHidden)
        
        Dim candidates
        candidates = Array("Plain", "Rectangle", "Corners", "Classic")
        
        For Each tryName In candidates
            On Error Resume Next
            Set mMaster = stn.Masters.ItemU(CStr(tryName))
            On Error GoTo 0
            If Not mMaster Is Nothing Then Exit For
        Next tryName
        
        'fallback – use a basic rectangle as a container
        If mMaster Is Nothing Then
            Dim basic As Object
            Set basic = pg.Application.Documents.OpenEx("BASIC_U.vssx", visOpenHidden)
            Set mMaster = basic.Masters.ItemU("Rectangle")
        End If
    End If
    
    'drop the container
    Set DropModuleContainer = pg.DropContainer(mMaster, Nothing)
    DropModuleContainer.Text = modName
End Function

'--- 2. Drop a procedure rectangle **into** that container ----
Public Function DropProcedureShape(cont As Object, _
                                   ByVal procName As String) As Object
    ' Use valid built-in stencil constants for Basic Shapes and Flowchart
    Const visBuiltInStencilBasic As Long = 0
    Const visBuiltInStencilFlowchart As Long = 1
    Const visMSUS As Long = 0
    Const visOpenHidden As Long = 64

    Static mProc As Object
    If mProc Is Nothing Then
        Dim app As Object: Set app = cont.ContainingPage.Application
        Dim stencilPaths, path, stn As Object
        stencilPaths = Array( _
            app.GetBuiltInStencilFile(visBuiltInStencilBasic, visMSUS), _
            app.GetBuiltInStencilFile(visBuiltInStencilFlowchart, visMSUS), _
            "BASIC_U.vssx")
        
        Dim candidates, tryName
        candidates = Array("Process", "Action", "Plain", "Rectangle", "Step")
        
        For Each path In stencilPaths
            On Error Resume Next
            Set stn = app.Documents.OpenEx(CStr(path), visOpenHidden)
            On Error GoTo 0
            If stn Is Nothing Then GoTo nxtPath
            
            For Each tryName In candidates
                On Error Resume Next
                Set mProc = stn.Masters.ItemU(CStr(tryName))
                On Error GoTo 0
                If Not mProc Is Nothing Then Exit For
            Next tryName
            If Not mProc Is Nothing Then Exit For
nxtPath:
        Next path
        
        If mProc Is Nothing Then
            Err.Raise vbObjectError + 1024, , _
                "No suitable procedure master found in scanned stencils."
        End If
    End If

    'drop on the parent page first
    Dim pg As Object: Set pg = cont.ContainingPage
    Set DropProcedureShape = pg.Drop(mProc, 0#, 0#)
    DropProcedureShape.Text = procName

    'add shape to the container so it moves/resizes with it
    Const visMemberAddExpandContainer& = 1
    cont.ContainerProperties.AddMember DropProcedureShape, visMemberAddExpandContainer
End Function

'Glue shapes
Public Sub ConnectShapes(pg As Object, shpA As Object, shpB As Object)
    Dim conn As Object
    Set conn = pg.Drop(pg.Application.ConnectorToolDataObject, 0, 0)
    conn.CellsU("BeginX").GlueTo shpA.CellsU("PinX")
    conn.CellsU("EndX").GlueTo shpB.CellsU("PinX")
End Sub