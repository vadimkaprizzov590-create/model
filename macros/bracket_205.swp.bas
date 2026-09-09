Attribute VB_Name = "Bracket205"
'==============================================================
' Деталь: Кронштейн  ЦГНТ1801.03.01.00.00.205
' Материал: Д16Т ГОСТ 4784-2019   Покрытие: Ан.Окс.ч. 3-6 мкм
' Макрос SolidWorks (VBA). Запуск: Sub main
' Все размеры в миллиметрах, в API переводятся в метры функцией MM()
'==============================================================
Option Explicit

'---------------------- ПАРАМЕТРЫ ---------------------------
Const SAVE_DIR As String = "C:\Temp\"
Const PART_NAME As String = "CGNT1801.03.01.00.00.205"

Const L_TOT As Double = 74#      ' длина 74 h8 (-0,046)
Const H_TOT As Double = 32#      ' высота по главному виду
Const T_TOT As Double = 18#      ' толщина 18 h8 (-0,027)

Const CUT_ANG As Double = 30#    ' угол срезов, град
Const CUT_H As Double = 12#      ' высота среза (допущение, см. README)

Const D_BORE As Double = 19#     ' Ø19 H8 (+0,03)
Const BORE_X As Double = 19#     ' положение центра от левого торца (19*)
Const BORE_Y As Double = 20#     ' положение центра от низа
Const CH_BORE As Double = 1#     ' фаска 1x45, 2 шт.

Const BC_DIA As Double = 28#     ' окружность центров Ø28
Const D_M4_TAP As Double = 3.3   ' диаметр под резьбу M4 (6H)
Const N_M4 As Long = 4
Const M4_START_ANG As Double = 45#
Const CH_M4 As Double = 0.7      ' фаска 0,7x45, 4 шт.

Const D_M6_TAP As Double = 5#    ' диаметр под резьбу M6 (6H)
Const M6_PITCH As Double = 50#   ' межосевое 50
Const M6_EDGE As Double = 12#    ' от торца 12*
Const M6_Z As Double = 9#        ' по центру толщины (18/2)
Const M6_DEPTH As Double = 12#

'------------------------------------------------------------
Dim swApp As Object
Dim swModel As Object
Dim swSkMgr As Object
Dim swFeatMgr As Object
Dim swExt As Object
Dim bRet As Boolean

Function MM(ByVal v As Double) As Double
    MM = v / 1000#
End Function

Sub main()

    Set swApp = Application.SldWorks
    swApp.Visible = True

    ' --- новая деталь по шаблону по умолчанию (8 = swDefaultTemplatePart)
    Dim sTemplate As String
    sTemplate = swApp.GetUserPreferenceStringValue(8)
    Set swModel = swApp.NewDocument(sTemplate, 0, 0, 0)
    If swModel Is Nothing Then
        MsgBox "Не удалось создать деталь. Проверьте шаблон детали."
        Exit Sub
    End If

    Set swExt = swModel.Extension
    Set swSkMgr = swModel.SketchManager
    Set swFeatMgr = swModel.FeatureManager

    BuildBody
    CutMainBore
    CutM4Holes
    CutM6Holes
    AddChamfers

    swModel.ViewZoomtofit2
    swModel.ClearSelection2 True
    swModel.ForceRebuild3 False

    On Error Resume Next
    swModel.SaveAs3 SAVE_DIR & PART_NAME & ".SLDPRT", 0, 2
    On Error GoTo 0

End Sub

'------------------------------------------------------------
' 1. Основное тело: контур главного вида + вытягивание на 18
'------------------------------------------------------------
Sub BuildBody()

    Dim dx As Double
    dx = CUT_H * Tan(CUT_ANG * 3.14159265358979 / 180#)   ' горизонтальная проекция среза

    SelectPlane "Front Plane", "Спереди"
    swSkMgr.InsertSketch True

    ' контур против часовой стрелки, начало координат - левый нижний угол
    Line 0, 0, L_TOT, 0
    Line L_TOT, 0, L_TOT, H_TOT - CUT_H
    Line L_TOT, H_TOT - CUT_H, L_TOT - dx, H_TOT
    Line L_TOT - dx, H_TOT, dx, H_TOT
    Line dx, H_TOT, 0, H_TOT - CUT_H
    Line 0, H_TOT - CUT_H, 0, 0

    swSkMgr.InsertSketch True

    ' бобышка-вытягивание на T_TOT (слепое)
    swFeatMgr.FeatureExtrusion2 True, False, False, 0, 0, MM(T_TOT), 0, _
        False, False, False, False, 0, 0, False, False, False, False, _
        True, True, True, 0, 0, False

    swModel.ClearSelection2 True

End Sub

'------------------------------------------------------------
' 2. Центральное отверстие Ø19 H8 - насквозь
'------------------------------------------------------------
Sub CutMainBore()

    SelectPlane "Front Plane", "Спереди"
    swSkMgr.InsertSketch True
    swSkMgr.CreateCircleByRadius MM(BORE_X), MM(BORE_Y), 0, MM(D_BORE / 2#)
    swSkMgr.InsertSketch True

    ' 1 = swEndCondThroughAll
    swFeatMgr.FeatureCut4 True, False, False, 1, 0, MM(T_TOT), 0, _
        False, False, False, False, 0, 0, False, False, False, False, False, _
        True, True, True, True, False, 0, 0, False, False

    swModel.ClearSelection2 True

End Sub

'------------------------------------------------------------
' 3. 4 отв. M4 по окружности Ø28 - насквозь
'------------------------------------------------------------
Sub CutM4Holes()

    Dim i As Long, ang As Double, x As Double, y As Double

    SelectPlane "Front Plane", "Спереди"
    swSkMgr.InsertSketch True

    For i = 0 To N_M4 - 1
        ang = (M4_START_ANG + i * (360# / N_M4)) * 3.14159265358979 / 180#
        x = BORE_X + (BC_DIA / 2#) * Cos(ang)
        y = BORE_Y + (BC_DIA / 2#) * Sin(ang)
        swSkMgr.CreateCircleByRadius MM(x), MM(y), 0, MM(D_M4_TAP / 2#)
    Next i

    swSkMgr.InsertSketch True

    swFeatMgr.FeatureCut4 True, False, False, 1, 0, MM(T_TOT), 0, _
        False, False, False, False, 0, 0, False, False, False, False, False, _
        True, True, True, True, False, 0, 0, False, False

    swModel.ClearSelection2 True

End Sub

'------------------------------------------------------------
' 4. 2 отв. M6 в нижней грани (плоскость "Сверху" совпадает с Y=0)
'------------------------------------------------------------
Sub CutM6Holes()

    SelectPlane "Top Plane", "Сверху"
    swSkMgr.InsertSketch True

    ' в эскизе на плоскости "Сверху": локальный X = глобальный X, локальный Y = -Z
    swSkMgr.CreateCircleByRadius MM(M6_EDGE), MM(-M6_Z), 0, MM(D_M6_TAP / 2#)
    swSkMgr.CreateCircleByRadius MM(M6_EDGE + M6_PITCH), MM(-M6_Z), 0, MM(D_M6_TAP / 2#)

    swSkMgr.InsertSketch True

    ' глухое на 12 мм вверх в тело (0 = swEndCondBlind)
    swFeatMgr.FeatureCut4 True, False, True, 0, 0, MM(M6_DEPTH), 0, _
        False, False, False, False, 0, 0, False, False, False, False, False, _
        True, True, True, True, False, 0, 0, False, False

    swModel.ClearSelection2 True

End Sub

'------------------------------------------------------------
' 5. Фаски по Ø19 (1x45, 2 шт.)
'    Остальные фаски (0,8x45 - 2 шт., 0,7x45 - 4 шт.) и скругления
'    R4 (2 рад.) / R2 (4 рад.) удобнее добавить вручную -
'    они зависят от точной геометрии контура.
'------------------------------------------------------------
Sub AddChamfers()

    On Error Resume Next

    swModel.ClearSelection2 True
    swExt.SelectByRay MM(BORE_X + D_BORE / 2#), MM(BORE_Y), MM(0), _
        0, 0, -1, MM(0.4), 1, False, 0, 0
    swExt.SelectByRay MM(BORE_X + D_BORE / 2#), MM(BORE_Y), MM(T_TOT), _
        0, 0, 1, MM(0.4), 1, True, 0, 0

    ' 0 = swChamferType_AngleDistance
    swFeatMgr.InsertFeatureChamfer 4, 1, MM(CH_BORE), 45# * 3.14159265358979 / 180#, 0, 0, 0, 0

    swModel.ClearSelection2 True
    On Error GoTo 0

End Sub

'---------------------- СЛУЖЕБНЫЕ ---------------------------
Sub Line(x1 As Double, y1 As Double, x2 As Double, y2 As Double)
    swSkMgr.CreateLine MM(x1), MM(y1), 0, MM(x2), MM(y2), 0
End Sub

Sub SelectPlane(sEng As String, sRus As String)
    swModel.ClearSelection2 True
    bRet = swExt.SelectByID2(sEng, "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    If Not bRet Then
        bRet = swExt.SelectByID2(sRus, "PLANE", 0, 0, 0, False, 0, Nothing, 0)
    End If
End Sub
