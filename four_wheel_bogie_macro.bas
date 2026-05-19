Attribute VB_Name = "FourWheelBogieMacro"
Option Explicit

' SolidWorks 2018 VBA Macro
' 自动生成“千斤顶全向车四轮小轮组”简化装配模型
' 所有长度单位输入为 mm，内部统一转换为 m

Dim swApp As SldWorks.SldWorks
Dim swModel As SldWorks.ModelDoc2
Dim swPart As SldWorks.PartDoc
Dim swAssy As SldWorks.AssemblyDoc
Dim swFeat As SldWorks.Feature
Dim swSketchMgr As SldWorks.SketchManager
Dim swFeatMgr As SldWorks.FeatureManager

Const MM As Double = 0.001

Sub main()
    On Error GoTo EH

    Set swApp = Application.SldWorks
    If swApp Is Nothing Then
        MsgBox "无法连接 SolidWorks。"
        Exit Sub
    End If

    Dim outDir As String
    outDir = CurDir$ & "\\"

    Dim tplPart As String, tplAssy As String
    tplPart = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplatePart)
    tplAssy = swApp.GetUserPreferenceStringValue(swUserPreferenceStringValue_e.swDefaultTemplateAssembly)

    If tplPart = "" Or tplAssy = "" Then
        MsgBox "请先在 SolidWorks 中设置默认零件/装配模板。"
        Exit Sub
    End If

    ' 1) 先建所有零件
    CreateBogieFrame outDir, tplPart
    CreateHeavyWheel outDir, tplPart
    CreateAxle outDir, tplPart, "front_axle.SLDPRT", 120, 900
    CreateAxle outDir, tplPart, "rear_axle.SLDPRT", 120, 900
    CreateAxle outDir, tplPart, "middle_shaft.SLDPRT", 80, 700
    CreateSlewingBearing outDir, tplPart
    CreateMountingPlate outDir, tplPart
    CreateSimpleCylinder outDir, tplPart, "hydraulic_motor.SLDPRT", 220, 260
    CreateSimpleCylinder outDir, tplPart, "coupling.SLDPRT", 120, 100
    CreateSimpleBox outDir, tplPart, "reducer.SLDPRT", 300, 260, 220
    CreateThreeGearBox outDir, tplPart

    ' 2) 创建装配并插入定位
    CreateAssembly outDir, tplAssy

    MsgBox "完成：模型已生成到 " & outDir
    Exit Sub

EH:
    MsgBox "宏执行失败: " & Err.Description & " (" & Err.Number & ")"
End Sub

' -------------------------
' 基础工具
' -------------------------
Private Function NewPart(tplPart As String) As SldWorks.ModelDoc2
    Dim m As SldWorks.ModelDoc2
    Set m = swApp.NewDocument(tplPart, swDwgPaperSizes_e.swDwgPaperAsize, 0, 0)
    Set NewPart = m
End Function

Private Sub SaveAndClose(ByVal fullPath As String)
    Dim e As Long, w As Long
    swModel.SaveAs3 fullPath, 0, 2
    swApp.CloseDoc swModel.GetTitle
End Sub

Private Sub SelectFrontPlane(ByVal model As SldWorks.ModelDoc2)
    model.Extension.SelectByID2 "前视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    If model.SelectionManager.GetSelectedObjectCount2(-1) = 0 Then
        model.Extension.SelectByID2 "Front Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    End If
End Sub

Private Sub SelectTopPlane(ByVal model As SldWorks.ModelDoc2)
    model.Extension.SelectByID2 "上视基准面", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    If model.SelectionManager.GetSelectedObjectCount2(-1) = 0 Then
        model.Extension.SelectByID2 "Top Plane", "PLANE", 0, 0, 0, False, 0, Nothing, 0
    End If
End Sub

' -------------------------
' bogie_frame.SLDPRT
' 箱梁结构：外框减内腔 + 轴承座孔位
' -------------------------
Private Sub CreateBogieFrame(ByVal outDir As String, ByVal tplPart As String)
    Set swModel = NewPart(tplPart)
    Set swPart = swModel
    Set swSketchMgr = swModel.SketchManager
    Set swFeatMgr = swModel.FeatureManager

    SelectTopPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCenterRectangle 0, 0, 0, 1100 * MM / 2, 900 * MM / 2, 0
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureExtrusion2 True, False, False, 0, 0, 180 * MM, 0, False, False, False, False, _
                              0, 0, False, False, False, False, True, True, True, 0, 0, False

    ' 内部减空，表示箱梁
    SelectTopPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCenterRectangle 0, 0, 0, 1000 * MM / 2, 800 * MM / 2, 0
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureCut3 True, False, False, 0, 0, 140 * MM, 0, False, False, False, False, _
                         0, 0, False, False, False, False, False, False, False, False, False, _
                         False, True, True, True, True, False, 0, 0, False

    SaveAndClose outDir & "bogie_frame.SLDPRT"
End Sub

' -------------------------
' heavy_wheel.SLDPRT
' 轮毂+中孔+8螺栓孔
' -------------------------
Private Sub CreateHeavyWheel(ByVal outDir As String, ByVal tplPart As String)
    Set swModel = NewPart(tplPart)
    Set swPart = swModel
    Set swSketchMgr = swModel.SketchManager
    Set swFeatMgr = swModel.FeatureManager

    SelectFrontPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCircleByRadius 0, 0, 0, 650 * MM / 2
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureExtrusion2 True, False, False, 0, 0, 220 * MM, 0, False, False, False, False, _
                              0, 0, False, False, False, False, True, True, True, 0, 0, False

    ' 中孔
    SelectFrontPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCircleByRadius 0, 0, 0, 120 * MM / 2
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureCut3 True, False, False, 0, 0, 240 * MM, 0, False, False, False, False, 0, 0, _
                         False, False, False, False, False, False, False, False, False, False, True, _
                         True, True, True, False, 0, 0, False

    ' 轮毂台阶（简化）
    SelectFrontPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCircleByRadius 0, 0, 0, 220 * MM / 2
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureExtrusion2 True, False, False, 0, 0, 80 * MM, 0, False, False, False, False, _
                              0, 0, False, False, False, False, True, True, True, 0, 0, False

    ' 8 螺栓孔：先打1个，再圆周阵列
    SelectFrontPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCircleByRadius 0.14, 0, 0, 0.0125
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureCut3 True, False, False, 0, 0, 240 * MM, 0, False, False, False, False, 0, 0, _
                         False, False, False, False, False, False, False, False, False, False, True, _
                         True, True, True, False, 0, 0, False

    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 "Cut-Extrude3", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0
    swModel.Extension.SelectByID2 "临时轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0
    If swModel.SelectionManager.GetSelectedObjectCount2(-1) < 2 Then
        swModel.Extension.SelectByID2 "Temporary Axis1", "AXIS", 0, 0, 0, True, 1, Nothing, 0
    End If
    swFeatMgr.FeatureCircularPattern4 8, 6.283185307, False, "NULL", False, True

    SaveAndClose outDir & "heavy_wheel.SLDPRT"
End Sub

Private Sub CreateAxle(ByVal outDir As String, ByVal tplPart As String, ByVal fileName As String, ByVal diaMM As Double, ByVal lenMM As Double)
    CreateSimpleCylinder outDir, tplPart, fileName, diaMM, lenMM
End Sub

Private Sub CreateSlewingBearing(ByVal outDir As String, ByVal tplPart As String)
    Set swModel = NewPart(tplPart)
    Set swPart = swModel
    Set swSketchMgr = swModel.SketchManager
    Set swFeatMgr = swModel.FeatureManager

    SelectFrontPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCircleByRadius 0, 0, 0, 700 * MM / 2
    swSketchMgr.CreateCircleByRadius 0, 0, 0, 500 * MM / 2
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureExtrusion2 True, False, False, 0, 0, 80 * MM, 0, False, False, False, False, _
                              0, 0, False, False, False, False, True, True, True, 0, 0, False

    ' 24 螺栓孔：外圈12 + 内圈12
    CreateBoltCircleCut 0.31, 0.01, 12, 0.08
    CreateBoltCircleCut 0.23, 0.01, 12, 0.08

    SaveAndClose outDir & "slewing_bearing.SLDPRT"
End Sub

Private Sub CreateBoltCircleCut(ByVal radius As Double, ByVal holeR As Double, ByVal qty As Integer, ByVal depth As Double)
    swModel.ClearSelection2 True
    SelectFrontPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCircleByRadius radius, 0, 0, holeR
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureCut3 True, False, False, 0, 0, depth, 0, False, False, False, False, 0, 0, _
                         False, False, False, False, False, False, False, False, False, False, True, _
                         True, True, True, False, 0, 0, False

    swModel.ClearSelection2 True
    swModel.Extension.SelectByID2 "Cut-Extrude1", "BODYFEATURE", 0, 0, 0, True, 0, Nothing, 0
    swModel.Extension.SelectByID2 "临时轴1", "AXIS", 0, 0, 0, True, 1, Nothing, 0
    If swModel.SelectionManager.GetSelectedObjectCount2(-1) < 2 Then
        swModel.Extension.SelectByID2 "Temporary Axis1", "AXIS", 0, 0, 0, True, 1, Nothing, 0
    End If
    swFeatMgr.FeatureCircularPattern4 qty, 6.283185307, False, "NULL", False, True
End Sub

Private Sub CreateMountingPlate(ByVal outDir As String, ByVal tplPart As String)
    CreateSimpleBox outDir, tplPart, "mounting_plate.SLDPRT", 900, 900, 35
End Sub

Private Sub CreateSimpleCylinder(ByVal outDir As String, ByVal tplPart As String, ByVal fileName As String, ByVal diaMM As Double, ByVal lenMM As Double)
    Set swModel = NewPart(tplPart)
    Set swPart = swModel
    Set swSketchMgr = swModel.SketchManager
    Set swFeatMgr = swModel.FeatureManager

    SelectFrontPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCircleByRadius 0, 0, 0, diaMM * MM / 2
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureExtrusion2 True, False, False, 0, 0, lenMM * MM, 0, False, False, False, False, _
                              0, 0, False, False, False, False, True, True, True, 0, 0, False

    SaveAndClose outDir & fileName
End Sub

Private Sub CreateSimpleBox(ByVal outDir As String, ByVal tplPart As String, ByVal fileName As String, ByVal lxMM As Double, ByVal lyMM As Double, ByVal lzMM As Double)
    Set swModel = NewPart(tplPart)
    Set swPart = swModel
    Set swSketchMgr = swModel.SketchManager
    Set swFeatMgr = swModel.FeatureManager

    SelectTopPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCenterRectangle 0, 0, 0, lxMM * MM / 2, lyMM * MM / 2, 0
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureExtrusion2 True, False, False, 0, 0, lzMM * MM, 0, False, False, False, False, _
                              0, 0, False, False, False, False, True, True, True, 0, 0, False

    SaveAndClose outDir & fileName
End Sub

Private Sub CreateThreeGearBox(ByVal outDir As String, ByVal tplPart As String)
    Set swModel = NewPart(tplPart)
    Set swPart = swModel
    Set swSketchMgr = swModel.SketchManager
    Set swFeatMgr = swModel.FeatureManager

    ' 底座箱体
    SelectTopPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCenterRectangle 0, 0, 0, 240 * MM / 2, 180 * MM / 2, 0
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureExtrusion2 True, False, False, 0, 0, 140 * MM, 0, False, False, False, False, _
                              0, 0, False, False, False, False, True, True, True, 0, 0, False

    ' 三个啮合圆柱齿轮（简化为圆柱）
    CreateGearCylinder -90 * MM, 0, 0, 110, 80
    CreateGearCylinder 0, 0, 0, 110, 80
    CreateGearCylinder 90 * MM, 0, 0, 110, 80

    SaveAndClose outDir & "three_gear_box.SLDPRT"
End Sub

Private Sub CreateGearCylinder(ByVal x As Double, ByVal y As Double, ByVal z As Double, ByVal diaMM As Double, ByVal thickMM As Double)
    SelectFrontPlane swModel
    swSketchMgr.InsertSketch True
    swSketchMgr.CreateCircleByRadius x, y, z, diaMM * MM / 2
    swSketchMgr.InsertSketch True
    swFeatMgr.FeatureExtrusion2 True, False, False, 0, 0, thickMM * MM, 0, False, False, False, False, _
                              0, 0, False, False, False, False, True, True, True, 0, 0, False
End Sub

Private Sub CreateAssembly(ByVal outDir As String, ByVal tplAssy As String)
    Dim m As SldWorks.ModelDoc2
    Set m = swApp.NewDocument(tplAssy, 0, 0, 0)
    Set swModel = m
    Set swAssy = swModel

    Dim cFrame As SldWorks.Component2
    Set cFrame = swAssy.AddComponent5(outDir & "bogie_frame.SLDPRT", 0, "", False, "", 0, 0, 0)

    ' 前后轴
    swAssy.AddComponent5 outDir & "front_axle.SLDPRT", 0, "", False, "", 0, 0.425, -0.12
    swAssy.AddComponent5 outDir & "rear_axle.SLDPRT", 0, "", False, "", 0, -0.425, -0.12

    ' 四轮：左右中心距760，前后轴距850
    swAssy.AddComponent5 outDir & "heavy_wheel.SLDPRT", 0, "", False, "", 0.38, 0.425, -0.12
    swAssy.AddComponent5 outDir & "heavy_wheel.SLDPRT", 0, "", False, "", -0.38, 0.425, -0.12
    swAssy.AddComponent5 outDir & "heavy_wheel.SLDPRT", 0, "", False, "", 0.38, -0.425, -0.12
    swAssy.AddComponent5 outDir & "heavy_wheel.SLDPRT", 0, "", False, "", -0.38, -0.425, -0.12

    ' 上部结构
    swAssy.AddComponent5 outDir & "slewing_bearing.SLDPRT", 0, "", False, "", 0, 0, 0.18
    swAssy.AddComponent5 outDir & "mounting_plate.SLDPRT", 0, "", False, "", 0, 0, 0.26

    ' 一侧动力件
    swAssy.AddComponent5 outDir & "reducer.SLDPRT", 0, "", False, "", 0.62, 0, 0.05
    swAssy.AddComponent5 outDir & "coupling.SLDPRT", 0, "", False, "", 0.48, 0, 0.05
    swAssy.AddComponent5 outDir & "hydraulic_motor.SLDPRT", 0, "", False, "", 0.34, 0, 0.05
    swAssy.AddComponent5 outDir & "middle_shaft.SLDPRT", 0, "", False, "", 0.2, 0, -0.06
    swAssy.AddComponent5 outDir & "three_gear_box.SLDPRT", 0, "", False, "", 0, 0, -0.05

    ' 简化：采用坐标放置表达结构关系；如需严格配合可后续补充 AddMate5
    Dim e As Long, w As Long
    swModel.SaveAs3 outDir & "four_wheel_bogie.SLDASM", 0, 2
    swApp.CloseDoc swModel.GetTitle
End Sub
