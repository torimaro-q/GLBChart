VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GLBChartBuilder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Type Vector3
    x As Single
    y As Single
    z As Single
End Type
Private Type Geometry
    Vertices() As Vector3
    Indices() As Long
End Type
Private Type FloatBytes
    B0 As Byte
    b1 As Byte
    b2 As Byte
    B3 As Byte
End Type
Private Type LongBytes
    B0 As Byte
    b1 As Byte
    b2 As Byte
    B3 As Byte
End Type
Private Type LongType
    l As Long
End Type
Private Type SingleType
    s As Single
End Type
Private Type AxSeries
    min As Single
    max As Single
    label As String
End Type
Private Type MyAxis
    x As AxSeries
    y As AxSeries
    z As AxSeries
End Type
Private Type GLTFMaterial
    r As Single
    g As Single
    b As Single
    a As Single
    AlphaBlend As Boolean
    DoubleSided As Boolean
End Type
Private Type GLTFBufferView
    byteOffset As Long
    byteLength As Long
    target As Long
End Type
Private Type GLTFAccessor
    bufferView As Long
    componentType As Long
    count As Long
    dataType As String
End Type
Private Type GLTFPrimitive
    positionAccessor As Long
    indexAccessor As Long
    material As Long
    mode As Long
End Type
Private Type GLTFMesh
    Primitives() As GLTFPrimitive
    PrimitiveCount As Long
End Type
Private Type GLTFBuilder
    Materials() As GLTFMaterial
    MaterialCount As Long
    BufferViews() As GLTFBufferView
    BufferViewCount As Long
    Accessors() As GLTFAccessor
    AccessorCount As Long
    Meshes() As GLTFMesh
    MeshCount As Long
    binary() As Byte
    BinaryLength As Long
End Type
Private AX As MyAxis
Private Const TEMPDIR = "genglb", TEMPFNAME = "temp.glb"
Private Const OBJ_SCALE As Single = 0.001
Public TargetSheet As Worksheet
Public HasLine As Boolean, HasMarker As Boolean
Public LineWidth As Single, MarkerSize As Single
Public LineColor As Long, MarkerColor As Long
Private Sub Class_Initialize()
    Set TargetSheet = ActiveSheet
    With Me
        .HasLine = False
        .HasMarker = True
        .LineWidth = 1
        .MarkerSize = 10
        .LineColor = RGB(30, 30, 30)
        .MarkerColor = RGB(50, 100, 200)
    End With
    With AX
        .x.label = "X": .y.label = "Y": .z.label = "Z"
    End With
    If Dir(TempPath, vbDirectory) = "" Then MkDir TempPath
End Sub
Public Function CreateFromRange(ByRef target As Range) As Shape
    Dim arr As Variant: arr = target.value
    Dim vec() As Vector3: ReDim vec(LBound(arr) To UBound(arr)) As Vector3
    Dim i As Long, st As Long
    Set TargetSheet = target.Parent
    With AX
        For i = LBound(arr) To UBound(arr)
            If IsNumeric(arr(i, 1)) Then
                st = i: Exit For
            Else
                .x.label = CStr(arr(i, 1)): .y.label = CStr(arr(i, 2)): .z.label = CStr(arr(i, 3))
            End If
        Next i
        .x.min = arr(st, 1): .y.min = arr(st, 2): .z.min = arr(st, 3)
        .x.max = arr(st, 1): .y.max = arr(st, 2): .z.max = arr(st, 3)
    End With
    For i = st To UBound(arr)
        vec(i) = Vector3(CSng(arr(i, 1)), CSng(arr(i, 2)), CSng(arr(i, 3)))
        With vec(i)
            If AX.x.min > .x Then AX.x.min = .x
            If AX.y.min > .y Then AX.y.min = .y
            If AX.z.min > .z Then AX.z.min = .z
            If AX.x.max < .x Then AX.x.max = .x
            If AX.y.max < .y Then AX.y.max = .y
            If AX.z.max < .z Then AX.z.max = .z
        End With
    Next i
    Dim data() As Byte: data = CreateGLB(vec)
    Set CreateFromRange = InsertGLB(data)
End Function
Private Property Get TempPath() As String
On Error GoTo getTempErr
    TempPath = (Environ("TEMP")) & "\" & TEMPDIR
    Exit Property
getTempErr:
    TempPath = ActiveWorkbook.Path & "\" & TEMPDIR
End Property
'---------------------------------------------------------------------------------------------------
Private Function InsertGLB(ByRef data() As Byte) As Shape
    Dim tpath As String, fileNo As Integer
    tpath = TempPath & "\" & TEMPFNAME
    SaveGLB data, tpath
    Set InsertGLB = TargetSheet.Shapes.Add3DModel(tpath, msoFalse, msoTrue)
    Kill tpath
End Function
Private Sub SaveGLB(ByRef glb() As Byte, ByVal filePath As String)
    Dim f As Integer: f = FreeFile
    Open filePath For Binary Access Write As #f
    Put #f, , glb
    Close #f
End Sub
Private Function CreateGLB(ByRef points() As Vector3) As Byte()
    Dim b As GLTFBuilder: b = NewGLTF()
    Dim tCl As LongType
    tCl.l = MarkerColor
    Dim pColor As LongBytes: LSet pColor = tCl
    tCl.l = LineColor
    Dim lColor As LongBytes: LSet lColor = tCl
    With pColor
        Dim pMaterial As Long: pMaterial = AddMaterial(b, CSng(.B0 / 255), CSng(.b1 / 255), CSng(.b2 / 255), 1)
    End With
    With lColor
        Dim lMaterial As Long: lMaterial = AddMaterial(b, CSng(.B0 / 255), CSng(.b1 / 255), CSng(.b2 / 255), 1)
    End With
    Dim xMaterial As Long: xMaterial = AddMaterial(b, 1, 0.1, 0.1, 1)
    Dim yMaterial As Long: yMaterial = AddMaterial(b, 0.1, 0.8, 0.2, 1)
    Dim zMaterial As Long: zMaterial = AddMaterial(b, 0.15, 0.3, 1, 1)
    Dim fMaterial As Long: fMaterial = AddMaterial(b, 1, 1, 1, 0.05, True, True)
    ' DATA MARKERS
    If HasMarker Then AddMeshGeometry b, CreatePointGeometry(points), pMaterial
    ' DATA LINE
    If HasLine Then AddMeshGeometry b, CreateLineGeometry(points), lMaterial
    ' AXES
    AddMeshGeometry b, CreateBox(0.5, 0, 0, 1, 0.005, 0.005), xMaterial
    AddMeshGeometry b, CreateBox(0, 0.5, 0, 0.005, 1, 0.005), yMaterial
    AddMeshGeometry b, CreateBox(0, 0, 0.5, 0.005, 0.005, 1), zMaterial
    ' FLOOR
    AddMeshGeometry b, CreateBox(0.5, 0.5, 0.005, 1, 1, 0.01), fMaterial
    CreateGLB = BuildGLB(b)
End Function
Private Function NewGLTF() As GLTFBuilder
    Dim b As GLTFBuilder
    With b
        .MaterialCount = 0
        .BufferViewCount = 0
        .AccessorCount = 0
        .MeshCount = 0
        .BinaryLength = 0
    End With
    NewGLTF = b
End Function
'---------------------------------------------------------------------------------------------------
Private Function AddMaterial(ByRef b As GLTFBuilder, ByVal r As Single, _
    ByVal g As Single, ByVal blue As Single, ByVal a As Single, _
    Optional ByVal AlphaBlend As Boolean = False, Optional ByVal DoubleSided As Boolean = False) As Long
    Dim i As Long: i = b.MaterialCount
    With b
        ReDim Preserve .Materials(0 To i)
        With .Materials(i)
            .r = r
            .g = g
            .b = blue
            .a = a
            .AlphaBlend = AlphaBlend
            .DoubleSided = DoubleSided
        End With
        .MaterialCount = .MaterialCount + 1
    End With
    AddMaterial = i
End Function
Private Function AppendBinary(ByRef b As GLTFBuilder, ByRef src() As Byte) As Long
    With b
        Dim n As Long: n = ByteArrayLength(src)
        If n <= 0 Then
            AppendBinary = Align4(.BinaryLength)
            Exit Function
        End If
        Dim offset As Long: offset = Align4(.BinaryLength)
        Dim oldLength As Long: oldLength = .BinaryLength
        Dim newLength As Long: newLength = offset + n
        Dim i As Long
        ' Allocate
        If oldLength = 0 Then ReDim .binary(0 To newLength - 1) Else ReDim Preserve .binary(0 To newLength - 1)
        ' Padding
        For i = oldLength To offset - 1
            .binary(i) = 0
        Next i
        ' Copy
        For i = 0 To n - 1
            .binary(offset + i) = src(LBound(src) + i)
        Next i
        .BinaryLength = newLength
    End With
    AppendBinary = offset
End Function
Private Function AddBufferView(ByRef b As GLTFBuilder, ByVal byteOffset As Long, ByVal byteLength As Long, ByVal target As Long) As Long
    With b
        Dim i As Long: i = .BufferViewCount
        ReDim Preserve .BufferViews(0 To i)
        With .BufferViews(i)
            .byteOffset = byteOffset
            .byteLength = byteLength
            .target = target
        End With
        .BufferViewCount = .BufferViewCount + 1
    End With
    AddBufferView = i
End Function
Private Function AddAccessor(ByRef b As GLTFBuilder, ByVal bufferView As Long, ByVal componentType As Long, ByVal count As Long, ByVal dataType As String) As Long
    With b
        Dim i As Long: i = .AccessorCount
        ReDim Preserve .Accessors(0 To i)
        With .Accessors(i)
            .bufferView = bufferView
            .componentType = componentType
            .count = count
            .dataType = dataType
        End With
        .AccessorCount = .AccessorCount + 1
    End With
    AddAccessor = i
End Function
Private Function AddPositionAccessor(ByRef b As GLTFBuilder, ByVal bufferView As Long, ByVal count As Long) As Long
    AddPositionAccessor = AddAccessor(b, bufferView, 5126, count, "VEC3")
End Function
Private Function AddIndexAccessor(ByRef b As GLTFBuilder, ByVal bufferView As Long, ByVal count As Long) As Long
    AddIndexAccessor = AddAccessor(b, bufferView, 5123, count, "SCALAR")
End Function
Private Function CreateBox(ByVal cx As Single, ByVal cy As Single, ByVal cz As Single, ByVal sx As Single, ByVal sy As Single, ByVal sz As Single) As Geometry
    Dim g As Geometry
    With g
        ReDim .Vertices(0 To 7)
        ReDim .Indices(0 To 35)
        Dim x0 As Single, x1 As Single, y0 As Single, y1 As Single, z0 As Single, z1 As Single
        x0 = cx - sx * 0.5: x1 = cx + sx * 0.5
        y0 = cy - sy * 0.5: y1 = cy + sy * 0.5
        z0 = cz - sz * 0.5: z1 = cz + sz * 0.5
        .Vertices(0) = Vector3(x0, y0, z0)
        .Vertices(1) = Vector3(x1, y0, z0)
        .Vertices(2) = Vector3(x1, y1, z0)
        .Vertices(3) = Vector3(x0, y1, z0)
        .Vertices(4) = Vector3(x0, y0, z1)
        .Vertices(5) = Vector3(x1, y0, z1)
        .Vertices(6) = Vector3(x1, y1, z1)
        .Vertices(7) = Vector3(x0, y1, z1)
        '--------------------------------------------------------
        .Indices(0) = 0: .Indices(1) = 2: .Indices(2) = 1
        .Indices(3) = 0: .Indices(4) = 3: .Indices(5) = 2
        '--------------------------------------------------------
        .Indices(6) = 4: .Indices(7) = 5: .Indices(8) = 6
        .Indices(9) = 4: .Indices(10) = 6: .Indices(11) = 7
        '--------------------------------------------------------
        .Indices(12) = 0: .Indices(13) = 4: .Indices(14) = 7
        .Indices(15) = 0: .Indices(16) = 7: .Indices(17) = 3
        '--------------------------------------------------------
        .Indices(18) = 1: .Indices(19) = 2: .Indices(20) = 6
        .Indices(21) = 1: .Indices(22) = 6: .Indices(23) = 5
        '--------------------------------------------------------
        .Indices(24) = 0: .Indices(25) = 1: .Indices(26) = 5
        .Indices(27) = 0: .Indices(28) = 5: .Indices(29) = 4
        '--------------------------------------------------------
        .Indices(30) = 3: .Indices(31) = 7: .Indices(32) = 6
        .Indices(33) = 3: .Indices(34) = 6: .Indices(35) = 2
    End With
    CreateBox = g
End Function
Private Function AddPositions(ByRef b As GLTFBuilder, ByRef g As Geometry) As Long
    With g
        Dim data() As Byte: data = GeometryVerticesBytes(g)
        Dim offset As Long: offset = AppendBinary(b, data)
        Dim view As Long: view = AddBufferView(b, offset, ByteArrayLength(data), 34962)
        AddPositions = AddPositionAccessor(b, view, GetVector3Count(.Vertices))
    End With
End Function
Private Function AddIndices(ByRef b As GLTFBuilder, ByRef g As Geometry) As Long
    With g
        Dim data() As Byte: data = GeometryIndexBytes(g)
        Dim count As Long: count = UBound(.Indices) - LBound(.Indices) + 1
        Dim offset As Long: offset = AppendBinary(b, data)
        Dim view As Long: view = AddBufferView(b, offset, ByteArrayLength(data), 34963)
        AddIndices = AddIndexAccessor(b, view, count)
    End With
End Function
Private Function AddPrimitive(ByRef b As GLTFBuilder, ByVal positionAccessor As Long, ByVal indexAccessor As Long, ByVal material As Long, Optional ByVal mode As Long = 4) As Long
    With b
        If .MeshCount = 0 Then
            ReDim .Meshes(0 To 0)
            .MeshCount = 1
            .Meshes(0).PrimitiveCount = 0
        End If
        Dim meshIndex As Long
        meshIndex = 0
        With .Meshes(meshIndex)
            Dim primitiveIndex As Long
            primitiveIndex = .PrimitiveCount
            ReDim Preserve .Primitives(0 To primitiveIndex)
            With .Primitives(primitiveIndex)
                .positionAccessor = positionAccessor
                .indexAccessor = indexAccessor
                .material = material
                .mode = mode
            End With
            .PrimitiveCount = .PrimitiveCount + 1
        End With
    End With
    AddPrimitive = primitiveIndex
End Function
Private Sub AddMeshGeometry(ByRef b As GLTFBuilder, ByRef g As Geometry, ByVal material As Long, Optional ByVal mode As Long = 4)
    Dim positionAccessor As Long: positionAccessor = AddPositions(b, g)
    Dim indexAccessor As Long: indexAccessor = AddIndices(b, g)
    AddPrimitive b, positionAccessor, indexAccessor, material, mode
End Sub
Private Sub AddLineSegment(ByRef g As Geometry, ByVal baseVertex As Long, ByVal baseIndex As Long, ByRef p0 As Vector3, ByRef p1 As Vector3, ByVal width As Single)
    Dim dv As Vector3: dv = Vector3(p1.x - p0.x, p1.y - p0.y, p1.z - p0.z)
    dv = getNormalizedVector(dv)
    Dim uv As Vector3
    If Abs(dv.z) < 0.9! Then uv = Vector3(0!, 0!, 1!) Else uv = Vector3(0!, 1!, 0!)
    With uv
        Dim vv As Vector3: vv = Vector3(dv.y * .z - dv.z * .y, dv.z * .x - dv.x * .z, dv.x * .y - dv.y * .x)
    End With
    vv = getNormalizedVector(vv)
    With vv
        Dim wv As Vector3: wv = Vector3(dv.y * .z - dv.z * .y, dv.z * .x - dv.x * .z, dv.x * .y - dv.y * .x)
    End With
    Dim halfW As Single: halfW = width * 0.5!
    With vv
        .x = .x * halfW: .y = .y * halfW: .z = .z * halfW
    End With
    With wv
        .x = .x * halfW: .y = .y * halfW: .z = .z * halfW
    End With
    ' 8 vertices
    With g
        .Vertices(baseVertex + 0) = Vector3(p0.x - vv.x - wv.x, p0.y - vv.y - wv.y, p0.z - vv.z - wv.z)
        .Vertices(baseVertex + 1) = Vector3(p0.x + vv.x - wv.x, p0.y + vv.y - wv.y, p0.z + vv.z - wv.z)
        .Vertices(baseVertex + 2) = Vector3(p0.x + vv.x + wv.x, p0.y + vv.y + wv.y, p0.z + vv.z + wv.z)
        .Vertices(baseVertex + 3) = Vector3(p0.x - vv.x + wv.x, p0.y - vv.y + wv.y, p0.z - vv.z + wv.z)
        .Vertices(baseVertex + 4) = Vector3(p1.x - vv.x - wv.x, p1.y - vv.y - wv.y, p1.z - vv.z - wv.z)
        .Vertices(baseVertex + 5) = Vector3(p1.x + vv.x - wv.x, p1.y + vv.y - wv.y, p1.z + vv.z - wv.z)
        .Vertices(baseVertex + 6) = Vector3(p1.x + vv.x + wv.x, p1.y + vv.y + wv.y, p1.z + vv.z + wv.z)
        .Vertices(baseVertex + 7) = Vector3(p1.x - vv.x + wv.x, p1.y - vv.y + wv.y, p1.z - vv.z + wv.z)
        ' -direction
        .Indices(baseIndex + 0) = baseVertex + 0: .Indices(baseIndex + 1) = baseVertex + 2
        .Indices(baseIndex + 2) = baseVertex + 1: .Indices(baseIndex + 3) = baseVertex + 0
        .Indices(baseIndex + 4) = baseVertex + 3: .Indices(baseIndex + 5) = baseVertex + 2
        ' +direction
        .Indices(baseIndex + 6) = baseVertex + 4: .Indices(baseIndex + 7) = baseVertex + 5
        .Indices(baseIndex + 8) = baseVertex + 6: .Indices(baseIndex + 9) = baseVertex + 4
        .Indices(baseIndex + 10) = baseVertex + 6: .Indices(baseIndex + 11) = baseVertex + 7
        ' side
        .Indices(baseIndex + 12) = baseVertex + 0: .Indices(baseIndex + 13) = baseVertex + 4
        .Indices(baseIndex + 14) = baseVertex + 7: .Indices(baseIndex + 15) = baseVertex + 0
        .Indices(baseIndex + 16) = baseVertex + 7: .Indices(baseIndex + 17) = baseVertex + 3
        .Indices(baseIndex + 18) = baseVertex + 1: .Indices(baseIndex + 19) = baseVertex + 2
        .Indices(baseIndex + 20) = baseVertex + 6: .Indices(baseIndex + 21) = baseVertex + 1
        .Indices(baseIndex + 22) = baseVertex + 6: .Indices(baseIndex + 23) = baseVertex + 5
        .Indices(baseIndex + 24) = baseVertex + 0: .Indices(baseIndex + 25) = baseVertex + 1
        .Indices(baseIndex + 26) = baseVertex + 5: .Indices(baseIndex + 27) = baseVertex + 0
        .Indices(baseIndex + 28) = baseVertex + 5: .Indices(baseIndex + 29) = baseVertex + 4
        .Indices(baseIndex + 30) = baseVertex + 3: .Indices(baseIndex + 31) = baseVertex + 7
        .Indices(baseIndex + 32) = baseVertex + 6: .Indices(baseIndex + 33) = baseVertex + 3
        .Indices(baseIndex + 34) = baseVertex + 6: .Indices(baseIndex + 35) = baseVertex + 2
    End With
End Sub
Private Sub CopyGeometry(ByRef dest As Geometry, ByRef src As Geometry, ByVal vertexOffset As Long, ByVal indexOffset As Long)
    Dim i As Long
    With src
        For i = LBound(.Vertices) To UBound(.Vertices) ' Vertices
            dest.Vertices(vertexOffset + i) = .Vertices(i)
        Next i
        For i = LBound(.Indices) To UBound(.Indices) ' Indices
            dest.Indices(indexOffset + i) = .Indices(i) + vertexOffset
        Next i
    End With
End Sub
Private Function CreatePointGeometry(ByRef points() As Vector3) As Geometry
    Dim cube As Geometry
    Dim n As Long, i As Long, baseVertex As Long, baseIndex As Long
    Dim px As Single, py As Single, pz As Single, nv As Vector3
    n = GetVector3Count(points)
    If n <= 0 Then Err.Raise vbObjectError + 1000, "CreatePointGeometry", "Vector3 is empty"
    ' Normalize
    nv = Vector3(1! / (0.000001! + AX.x.max - AX.x.min), 1! / (0.000001! + AX.y.max - AX.y.min), 1! / (0.000001! + AX.z.max - AX.z.min))
    ' Allocate
    Dim g As Geometry
    ReDim g.Vertices(0 To n * 8 - 1)
    ReDim g.Indices(0 To n * 36 - 1)
    ' Create each marker cube
    Dim LP As Long: LP = LBound(points)
    For i = LP To UBound(points)
        With points(i)
            px = (.x - AX.x.min) * nv.x
            py = (.y - AX.y.min) * nv.y
            pz = (.z - AX.z.min) * nv.z
        End With
        cube = CreateBox(px, py, pz, MarkerSize * OBJ_SCALE, MarkerSize * OBJ_SCALE, MarkerSize * OBJ_SCALE)
        baseVertex = (i - LP) * 8
        baseIndex = (i - LP) * 36
        CopyGeometry g, cube, baseVertex, baseIndex
    Next i
    CreatePointGeometry = g
End Function
Private Function CreateLineGeometry(ByRef points() As Vector3) As Geometry
    Dim n As Long: n = GetVector3Count(points)
    If n < 2 Then
        Dim emptyG As Geometry
        CreateLineGeometry = emptyG
        Exit Function
    End If
    Dim segmentCount As Long: segmentCount = n - 1
    Dim g As Geometry
    ' 1 segment = 8 vertices + 36 indices
    ReDim g.Vertices(0 To segmentCount * 8 - 1)
    ReDim g.Indices(0 To segmentCount * 36 - 1)
    Dim i As Long, baseVertex As Long, baseIndex As Long, LP As Long
    Dim nv As Vector3, p0 As Vector3, p1 As Vector3
    With AX
        nv.x = 1! / (0.000001! + .x.max - .x.min)
        nv.y = 1! / (0.000001! + .y.max - .y.min)
        nv.z = 1! / (0.000001! + .z.max - .z.min)
        LP = LBound(points)
        For i = 0 To segmentCount - 1
            p0 = Vector3((points(LP + i).x - .x.min) * nv.x, (points(LP + i).y - .y.min) * nv.y, (points(LP + i).z - .z.min) * nv.z)
            p1 = Vector3((points(LP + i + 1).x - .x.min) * nv.x, (points(LP + i + 1).y - .y.min) * nv.y, (points(LP + i + 1).z - .z.min) * nv.z)
            baseVertex = i * 8
            baseIndex = i * 36
            AddLineSegment g, baseVertex, baseIndex, p0, p1, LineWidth * OBJ_SCALE
        Next i
    End With
    CreateLineGeometry = g
End Function
'----------------------------------------------------------------------------------------------------------------------------
Private Function BuildGLB(ByRef b As GLTFBuilder) As Byte()
    Dim json As String: json = BuildJSON(b)
    BuildGLB = SerializeGLB(json, b.binary)
End Function
Private Function BuildJSON(ByRef b As GLTFBuilder) As String
    Dim s As String, i As Long, j As Long
    ' Header
    s = "{"
    s = s & """asset"":{""version"":""2.0""},"
    s = s & """scene"":0,"
    ' Scenes
    s = s & """scenes"":[{""nodes"":[0]}],"
    ' Nodes
    s = s & """nodes"":[{""mesh"":0}],"
    ' Materials
    s = s & """materials"":["
    For i = 0 To b.MaterialCount - 1
        If i > 0 Then s = s & ","
        With b.Materials(i)
            s = s & "{"
            s = s & """pbrMetallicRoughness"":{"
            s = s & """baseColorFactor"":[" & FmtF32(.r) & "," & FmtF32(.g) & "," & FmtF32(.b) & "," & FmtF32(.a) & "]"
            s = s & "}"
            If .AlphaBlend Then s = s & ",""alphaMode"":""BLEND"""
            If .DoubleSided Then s = s & ",""doubleSided"":true"
            s = s & "}"
        End With
    Next i
    s = s & "],"
    ' Meshes
    s = s & """meshes"":["
    For i = 0 To b.MeshCount - 1
        If i > 0 Then s = s & ","
        s = s & "{""primitives"":["
        For j = 0 To b.Meshes(i).PrimitiveCount - 1
            If j > 0 Then s = s & ","
            With b.Meshes(i).Primitives(j)
                s = s & "{"
                s = s & """attributes"":{""POSITION"":" & CStr(.positionAccessor) & "},"
                s = s & """indices"":" & CStr(.indexAccessor) & ","
                s = s & """material"":" & CStr(.material) & ","
                s = s & """mode"":" & CStr(.mode)
                s = s & "}"
            End With
        Next j
        s = s & "]}"
    Next i
    s = s & "],"
    ' Buffers
    s = s & """buffers"":[{""byteLength"":" & CStr(Align4(b.BinaryLength)) & "}],"
    ' BufferViews
    s = s & """bufferViews"":["
    For i = 0 To b.BufferViewCount - 1
        If i > 0 Then s = s & ","
        With b.BufferViews(i)
            s = s & "{"
            s = s & """buffer"":0,"
            s = s & """byteOffset"":" & CStr(.byteOffset) & ","
            s = s & """byteLength"":" & CStr(.byteLength)
            If .target <> 0 Then s = s & ",""target"":" & CStr(.target)
            s = s & "}"
        End With
    Next i
    s = s & "],"
    ' Accessors
    s = s & """accessors"":["
    For i = 0 To b.AccessorCount - 1
        If i > 0 Then s = s & ","
        With b.Accessors(i)
            s = s & "{"
            s = s & """bufferView"":" & CStr(.bufferView) & ","
            s = s & """componentType"":" & CStr(.componentType) & ","
            s = s & """count"":" & CStr(.count) & ","
            s = s & """type"":""" & .dataType & """"
            s = s & "}"
        End With
    Next i
    s = s & "]"
    ' End
    s = s & "}"
    BuildJSON = s
End Function

Private Function SerializeGLB(ByVal json As String, ByRef binary() As Byte) As Byte()
    Dim jsonBytes() As Byte
    Dim jsonLength As Long
    Dim jsonPaddedLength As Long
    Dim binLength As Long
    Dim binPaddedLength As Long
    Dim totalLength As Long
    Dim glb() As Byte
    Dim pos As Long
    ' JSON
    jsonBytes = StringToAsciiBytes(json)
    jsonLength = ByteArrayLength(jsonBytes)
    jsonPaddedLength = Align4(jsonLength)
    ' BIN
    binLength = ByteArrayLength(binary)
    binPaddedLength = Align4(binLength)
    ' GLB size
    totalLength = 12 + 8 + jsonPaddedLength + 8 + binPaddedLength
    ReDim glb(0 To totalLength - 1)
    ' GLB HEADER
    pos = 0
    ' magic = "glTF"
    PutU32 glb, pos, &H46546C67
    pos = pos + 4
    ' version = 2
    PutU32 glb, pos, 2
    pos = pos + 4
    ' total length
    PutU32 glb, pos, totalLength
    pos = pos + 4
    ' JSON CHUNK
    PutU32 glb, pos, jsonPaddedLength
    pos = pos + 4
    ' "JSON"
    PutU32 glb, pos, &H4E4F534A
    pos = pos + 4
    PutBytes glb, pos, jsonBytes
    pos = pos + jsonLength
    'space (0x20)
    Do While pos < 12 + 8 + jsonPaddedLength
        glb(pos) = &H20
        pos = pos + 1
    Loop
    ' BIN CHUNK
    PutU32 glb, pos, binPaddedLength
    pos = pos + 4
    ' "BIN\0"
    PutU32 glb, pos, &H4E4942
    pos = pos + 4
    If binLength > 0 Then
        PutBytes glb, pos, binary
        pos = pos + binLength
    End If
    ' BIN padding = 0
    Do While pos < totalLength
        glb(pos) = 0
        pos = pos + 1
    Loop
    SerializeGLB = glb
End Function
Private Function GeometryVerticesBytes(ByRef g As Geometry) As Byte()
    With g
        Dim count As Long: count = GetVector3Count(.Vertices)
        Dim data() As Byte: ReDim data(0 To count * 12 - 1)
        Dim i As Long, p As Long: p = 0
        If count <= 0 Then
            GeometryVerticesBytes = data
            Exit Function
        End If
        For i = LBound(.Vertices) To UBound(.Vertices)
            With .Vertices(i)
                PutF32 data, p, .x: p = p + 4
                PutF32 data, p, .y: p = p + 4
                PutF32 data, p, .z: p = p + 4
            End With
        Next i
    End With
    GeometryVerticesBytes = data
End Function
Private Function GeometryIndexBytes(ByRef g As Geometry, Optional ByVal baseIndex As Long = 0) As Byte()
    With g
        Dim count As Long: count = UBound(.Indices) - LBound(.Indices) + 1
        Dim data() As Byte: ReDim data(0 To count * 2 - 1)
        Dim i As Long, p As Long: p = 0
        For i = LBound(.Indices) To UBound(.Indices)
            PutU16 data, p, .Indices(i) + baseIndex
            p = p + 2
        Next i
    End With
    GeometryIndexBytes = data
End Function
Private Sub PutF32(ByRef buffer() As Byte, ByVal offset As Long, ByVal value As Single)
    Dim f As FloatBytes, s As SingleType: s.s = value
    LSet f = s
    With f
        buffer(offset + 0) = .B0: buffer(offset + 1) = .b1: buffer(offset + 2) = .b2: buffer(offset + 3) = .B3
    End With
End Sub
Private Sub PutU32(ByRef buffer() As Byte, ByVal offset As Long, ByVal value As Long)
    Dim b As LongBytes, x As LongType: x.l = value
    LSet b = x
    With b
        buffer(offset + 0) = .B0: buffer(offset + 1) = .b1: buffer(offset + 2) = .b2: buffer(offset + 3) = .B3
    End With
End Sub
Private Sub PutU16(ByRef buffer() As Byte, ByVal offset As Long, ByVal value As Long)
    buffer(offset) = value And &HFF&
    buffer(offset + 1) = (value \ 256) And &HFF&
End Sub
Private Sub PutBytes(ByRef dest() As Byte, ByVal offset As Long, ByRef src() As Byte)
    Dim i As Long
    For i = LBound(src) To UBound(src)
        dest(offset + i - LBound(src)) = src(i)
    Next i
End Sub
Private Function Align4(ByVal value As Long) As Long
    Align4 = (value + 3) And Not 3
End Function
Private Function StringToAsciiBytes(ByVal s As String) As Byte()
    Dim b() As Byte, i As Long
    If Len(s) = 0 Then
        ReDim b(0 To -1)
        StringToAsciiBytes = b
        Exit Function
    End If
    ReDim b(0 To Len(s) - 1)
    For i = 1 To Len(s)
        b(i - 1) = AscW(Mid$(s, i, 1))
    Next i
    StringToAsciiBytes = b
End Function
Private Function ByteArrayLength(ByRef b() As Byte) As Long
    On Error GoTo EmptyArray
    ByteArrayLength = UBound(b) - LBound(b) + 1
    Exit Function
EmptyArray:
    ByteArrayLength = 0
End Function
Private Function GetVector3Count(ByRef points() As Vector3) As Long
    On Error GoTo EmptyArray
    GetVector3Count = UBound(points) - LBound(points) + 1
    Exit Function
EmptyArray:
    GetVector3Count = 0
End Function
Private Function Vector3(ByVal x As Single, ByVal y As Single, ByVal z As Single) As Vector3
    With Vector3
        .x = x: .y = y: .z = z
    End With
End Function
Private Function FmtF32(ByVal value As Single) As String
    Dim s As String
    s = Format$(value, "0.###############")
    s = Replace$(s, Application.DecimalSeparator, ".")
    If Right$(s, 1) = "." Then
        s = Left$(s, Len(s) - 1)
    End If
    FmtF32 = s
End Function
Private Function getNormalizedVector(ByRef vec As Vector3) As Vector3
    Dim l As Single
    With vec
        l = Sqr(.x ^ 2 + .y ^ 2 + .z ^ 2)
        If l <= 0 Then getNormalizedVector = Vector3(0, 0, 0) Else getNormalizedVector = Vector3(.x / l, .y / l, .z / l)
    End With
End Function
