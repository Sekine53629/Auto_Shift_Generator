Option Explicit
'==================================================================
'  シフト表 クリック入力マクロ  ＜標準モジュール＞
'  最終更新: 2026-08-21 (v7 : パレット動的範囲＋パターン追加対応)
'  パレット構成（横1行）:
'    27行目 = ★マーカー / 28行目 = パレット本体 / 29行目 = ラベル
'  動的範囲（ブック定義の名前付き範囲）:
'    「シフトパレット範囲」= シフト入力欄（B12の次行?薬剤師出勤数の1つ上）
'    「シフトパレット」    = パレット本体（B28?ラベル行の最終セルまで自動拡張）
'  ※名前が見つからない場合のみ、下の定数にフォールバックします
'==================================================================

'------------------------ 設定 ここから ---------------------------
' 動的範囲の名前
Public Const SHIFT_RANGE_NAME  As String = "シフトパレット範囲"
Public Const PALETTE_NAME      As String = "シフトパレット"
' フォールバック用（名前付き範囲が見つからない場合のみ使用）
Public Const SHIFT_RANGE       As String = "B13:AF22"
Public Const PALETTE_RANGE     As String = "B28:M28"
' パレットの起点セル（パレット作成マクロ用）
Public Const PALETTE_HOME      As String = "B28"

Public Const CYCLE_TRIGGER As String = "double"
Public Const STAMP_TRIGGER As String = "double"

' ★横型: マーカーは「上の行」、ラベルは「下の行」（行オフセット）
Public Const MARKER_OFFSET As Long = -1
Public Const LABEL_OFFSET  As Long = 1
Public Const MARKER_CHAR   As String = "★"

Public Const APPLY_FILL          As Boolean = True
Public Const CYCLE_RESETS_FILL   As Boolean = False
Public Const MOVE_AFTER          As String = ""
Public Const SKIP_FORMULA_CELLS  As Boolean = True

Public Const ROW_OFF       As Long = 1   ' OFF（左から1番目）
Public Const ROW_CYCLE     As Long = 2   ' 切替（左から2番目）
Public Const ROW_CLEARFILL As Long = 3   ' 色消（左から3番目）
Public Const IDX_ERASE     As Long = 4   ' 消去（左から4番目・空白スタンプ）
'------------------------ 設定 ここまで ---------------------------

'==================================================================
' 動的範囲の解決
'==================================================================
'--- シフト入力範囲 ---
Private Function ShiftRange(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set ShiftRange = ws.Parent.Names(SHIFT_RANGE_NAME).RefersToRange
    On Error GoTo 0
    If ShiftRange Is Nothing Then Set ShiftRange = ws.Range(SHIFT_RANGE)
End Function

'--- パレット本体 ---
Private Function PaletteRange(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set PaletteRange = ws.Parent.Names(PALETTE_NAME).RefersToRange
    On Error GoTo 0
    If PaletteRange Is Nothing Then Set PaletteRange = ws.Range(PALETTE_RANGE)
End Function

'--- 連続切替の巡回リスト: 空白＋パレットのスタンプ項目から自動生成 ---
Private Function CycleValues(ByVal ws As Worksheet) As Variant
    Dim pal As Range, arr() As Variant, i As Long, n As Long, v As String
    Set pal = PaletteRange(ws)
    ReDim arr(0 To pal.Cells.Count)
    arr(0) = ""            ' 先頭は空白（消去）
    n = 0
    For i = IDX_ERASE To pal.Cells.Count
        v = Trim$(CStr(pal.Cells(1, i).Value))
        If Len(v) > 0 Then
            n = n + 1
            arr(n) = v
        End If
    Next i
    ReDim Preserve arr(0 To n)
    CycleValues = arr
End Function

'==================================================================
' メインハンドラ
'==================================================================
Public Sub ShiftClick_Handle(ByVal Target As Range, _
                             ByVal EventKind As String, _
                             ByRef Handled As Boolean)
    Dim ws As Worksheet, pal As Range, area As Range
    Dim idx As Long
    Handled = False
    On Error GoTo EH

    Set ws = Target.Worksheet
    Set pal = PaletteRange(ws)

    ' --- パレット上のクリック: モード切替 ---
    If Not Application.Intersect(Target, pal) Is Nothing Then
        If EventKind = "right" Then Exit Sub
        SetStamp ws, Target.Cells(1, 1)
        Handled = True
        Exit Sub
    End If

    ' --- シフト入力範囲（動的）との交差判定 ---
    Set area = Application.Intersect(ClickRange(Target, EventKind), ShiftRange(ws))
    If area Is Nothing Then Exit Sub
    If Application.CutCopyMode <> False Then Exit Sub

    idx = CurrentIndex(ws)
    If idx = ROW_OFF Then Exit Sub

    Application.EnableEvents = False

    If idx = ROW_CYCLE Then
        If area.Cells.Count = 1 Then
            If EventKind = "right" Then
                CycleOne area.Cells(1, 1), True
                Handled = True
            ElseIf Triggered(EventKind, CYCLE_TRIGGER) Then
                CycleOne area.Cells(1, 1), False
                Handled = True
            End If
        End If
    Else
        If EventKind = "right" Or Triggered(EventKind, STAMP_TRIGGER) Then
            StampArea ws, area, idx
            Handled = True
        End If
    End If

    If Handled And Len(MOVE_AFTER) > 0 Then
        Select Case LCase$(MOVE_AFTER)
            Case "down":  MoveSel area, 1, 0
            Case "right": MoveSel area, 0, 1
        End Select
    End If

Fin:
    Application.EnableEvents = True
    Exit Sub
EH:
    Application.EnableEvents = True
End Sub

'==================================================================
' 内部処理
'==================================================================
Private Function Triggered(ByVal EventKind As String, ByVal Trig As String) As Boolean
    If EventKind = "double" Then
        Triggered = True
    ElseIf EventKind = "select" Then
        Triggered = (LCase$(Trim$(Trig)) = "single")
    Else
        Triggered = False
    End If
End Function

Private Function ClickRange(ByVal Target As Range, ByVal EventKind As String) As Range
    Dim sel As Range
    Set ClickRange = Target
    If EventKind <> "right" Then Exit Function
    If Not TypeOf Selection Is Range Then Exit Function
    Set sel = Selection
    If sel Is Nothing Then Exit Function
    If sel.Worksheet.Name <> Target.Worksheet.Name Then Exit Function
    If Not Application.Intersect(sel, Target) Is Nothing Then Set ClickRange = sel
End Function

'--- 範囲へのスタンプ適用（ダブルクリック時／選択範囲実行時 共用） ---
Private Sub StampArea(ByVal ws As Worksheet, ByVal area As Range, ByVal idx As Long)
    Dim c As Range, pal As Range
    Set pal = PaletteRange(ws)
    For Each c In area.Cells
        If idx = ROW_CLEARFILL Then
            c.Interior.Pattern = xlNone
        Else
            ApplyStamp c, pal.Cells(1, idx)
        End If
    Next c
End Sub

Private Sub CycleOne(ByVal c As Range, ByVal Reverse As Boolean)
    Dim v As Variant, i As Long, idx As Long, cur As String
    If SKIP_FORMULA_CELLS Then
        If c.HasFormula Then Exit Sub
    End If
    If c.Worksheet.ProtectContents And c.Locked Then Exit Sub
    v = CycleValues(c.Worksheet)
    cur = Trim$(CStr(c.Value))
    idx = -1
    For i = LBound(v) To UBound(v)
        If cur = CStr(v(i)) Then
            idx = i
            Exit For
        End If
    Next i
    If idx = -1 Then idx = LBound(v)
    If Reverse Then idx = idx - 1 Else idx = idx + 1
    If idx > UBound(v) Then idx = LBound(v)
    If idx < LBound(v) Then idx = UBound(v)
    If Len(CStr(v(idx))) = 0 Then
        c.ClearContents
    Else
        c.Value = v(idx)
    End If
    If CYCLE_RESETS_FILL Then c.Interior.Pattern = xlNone
End Sub

Private Sub ApplyStamp(ByVal c As Range, ByVal src As Range)
    If SKIP_FORMULA_CELLS Then
        If c.HasFormula Then Exit Sub
    End If
    If c.Worksheet.ProtectContents And c.Locked Then Exit Sub
    If Len(Trim$(CStr(src.Value))) = 0 Then
        c.ClearContents
    Else
        c.Value = src.Value
    End If
    c.Font.Color = src.Font.Color
    c.Font.Bold = src.Font.Bold
    If APPLY_FILL Then
        If src.Interior.Pattern <> xlNone Then
            c.Interior.Color = src.Interior.Color
        End If
    End If
End Sub

'--- ★横型: マーカーは各セルの「上の行」を見る ---
Private Function CurrentIndex(ByVal ws As Worksheet) As Long
    Dim pal As Range, i As Long
    Set pal = PaletteRange(ws)
    For i = 1 To pal.Cells.Count
        If StrComp(Trim$(CStr(pal.Cells(1, i).Offset(MARKER_OFFSET, 0).Value)), _
                   MARKER_CHAR, vbTextCompare) = 0 Then
            CurrentIndex = i
            Exit Function
        End If
    Next i
    CurrentIndex = ROW_CYCLE
End Function

'--- ★横型: 横方向に一致セルを探し、マーカーは上の行に書く ---
Private Sub SetStamp(ByVal ws As Worksheet, ByVal c As Range)
    Dim pal As Range, i As Long, hit As Long
    Set pal = PaletteRange(ws)
    hit = 0
    For i = 1 To pal.Cells.Count
        If pal.Cells(1, i).Address = c.Address Then hit = i
    Next i
    If hit = 0 Then Exit Sub
    If hit = CurrentIndex(ws) And hit <> ROW_CYCLE Then hit = ROW_CYCLE
    Application.EnableEvents = False
    On Error Resume Next
    pal.Offset(MARKER_OFFSET, 0).ClearContents
    pal.Cells(1, hit).Offset(MARKER_OFFSET, 0).Value = MARKER_CHAR
    On Error GoTo 0
    Application.EnableEvents = True
    ShowMode ws, hit
End Sub

Private Sub ShowMode(ByVal ws As Worksheet, ByVal idx As Long)
    Dim pal As Range, s As String, op As String
    Set pal = PaletteRange(ws)
    Select Case idx
        Case ROW_OFF
            s = "シフト入力マクロ： OFF（通常のExcel操作）"
        Case ROW_CYCLE
            If LCase$(CYCLE_TRIGGER) = "single" Then op = "クリック" Else op = "ダブルクリック"
            s = "シフト入力マクロ： 連続切替　" & op & "で次の記号／右クリックで前へ"
        Case ROW_CLEARFILL
            s = "シフト入力マクロ： 背景色クリア"
        Case Else
            If LCase$(STAMP_TRIGGER) = "single" Then op = "クリック" Else op = "ダブルクリック"
            s = "シフト入力マクロ： スタンプ [ " & _
                Trim$(CStr(pal.Cells(1, idx).Value)) & " ]　" & _
                op & "で押す／範囲を選んで右クリックでまとめて押す"
    End Select
    Application.StatusBar = s
End Sub

Private Sub MoveSel(ByVal area As Range, ByVal dr As Long, ByVal dc As Long)
    Dim t As Range
    On Error Resume Next
    Set t = area.Cells(area.Cells.Count).Offset(dr, dc)
    If Not t Is Nothing Then t.Select
    On Error GoTo 0
End Sub

'==================================================================
' 手動実行マクロ
'==================================================================
Public Sub ShiftClick_選択範囲にスタンプ()
    Dim ws As Worksheet, area As Range, idx As Long
    Set ws = ActiveSheet
    If Not TypeOf Selection Is Range Then Exit Sub
    Set area = Application.Intersect(Selection, ShiftRange(ws))
    If area Is Nothing Then
        MsgBox "シフト入力欄（" & ShiftRange(ws).Address(False, False) & _
               "）を選んでから実行してください。", vbExclamation
        Exit Sub
    End If
    idx = CurrentIndex(ws)
    If idx <= ROW_CYCLE Then
        MsgBox "パレットでスタンプする記号を選んでから実行してください。", vbExclamation
        Exit Sub
    End If
    Application.EnableEvents = False
    StampArea ws, area, idx
    Application.EnableEvents = True
End Sub

Public Sub ShiftClick_連続切替に戻す()
    SetStamp ActiveSheet, PaletteRange(ActiveSheet).Cells(1, ROW_CYCLE)
End Sub

'==================================================================
' パターン追加: パレット右端に新しいシフトパターンを追加
'   （名前付き範囲はラベル行を基準に自動拡張されます）
'==================================================================
Public Sub ShiftClick_パターン追加()
    Dim ws As Worksheet, pal As Range, lastCell As Range, newCell As Range
    Dim sym As String, lab As String, i As Long
    Set ws = ActiveSheet
    Set pal = PaletteRange(ws)

    sym = Trim$(InputBox("追加するシフト記号（セルに入力される値）:", "パターン追加"))
    If Len(sym) = 0 Then Exit Sub
    ' 重複チェック
    For i = 1 To pal.Cells.Count
        If Trim$(CStr(pal.Cells(1, i).Value)) = sym Then
            MsgBox "[ " & sym & " ] は既にパレットにあります。", vbExclamation
            Exit Sub
        End If
    Next i
    lab = Trim$(InputBox("ラベル（パレット下の説明）:", "パターン追加", sym))
    If Len(lab) = 0 Then lab = sym

    Set lastCell = pal.Cells(1, pal.Cells.Count)
    Set newCell = lastCell.Offset(0, 1)

    Application.EnableEvents = False
    ' 書式を右端のセルからコピー（マーカー行・本体・ラベル行）
    lastCell.Offset(MARKER_OFFSET, 0).Resize(3, 1).Copy
    newCell.Offset(MARKER_OFFSET, 0).PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    ' 値とラベルを書き込み → ラベルが入った時点で名前付き範囲が自動拡張
    newCell.Value = sym
    newCell.Offset(LABEL_OFFSET, 0).Value = lab
    Application.EnableEvents = True

    MsgBox "パターン [ " & sym & " ] を追加しました。" & vbCrLf & _
           "パレット範囲: " & PaletteRange(ws).Address(False, False) & vbCrLf & vbCrLf & _
           "※パレットのセルに文字色・背景色を付けると、" & vbCrLf & _
           "　スタンプ時にその書式もコピーされます。", vbInformation
End Sub

'==================================================================
' パレット再作成（横型・B28起点。列幅は変更しません）
'==================================================================
Public Sub ShiftClick_パレット作成()
    Dim ws As Worksheet, pal As Range, i As Long
    Dim vals As Variant, labs As Variant
    Set ws = ActiveSheet
    vals = Array("OFF", "切替", "色消", "", "○", "●", "▲", "公休", "希休", "夏休", "有休", "有休※")
    labs = Array("停止", "順送り", "背景消", "消去", "早番", "遅半", "遅番", "公休", "希休", "夏休", "有休", "備考付")
    Set pal = ws.Range(PALETTE_HOME).Resize(1, UBound(vals) + 1)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    With pal
        .ClearContents
        .Interior.Pattern = xlNone
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Font.size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .RowHeight = 20
    End With
    With pal.Offset(MARKER_OFFSET, 0)   ' 上=マーカー行
        .ClearContents
        .HorizontalAlignment = xlCenter
        .Font.Color = RGB(192, 0, 0)
        .Font.Bold = True
    End With
    With pal.Offset(LABEL_OFFSET, 0)    ' 下=ラベル行
        .ClearContents
        .HorizontalAlignment = xlCenter
        .Font.size = 9
        .Font.Color = RGB(100, 100, 100)
    End With
    For i = 1 To pal.Cells.Count
        If Len(CStr(vals(i - 1))) > 0 Then pal.Cells(1, i).Value = vals(i - 1)
        pal.Cells(1, i).Offset(LABEL_OFFSET, 0).Value = labs(i - 1)
    Next i
    For i = 1 To 3   ' モードセルはグレー
        With pal.Cells(1, i)
            .Interior.Color = RGB(242, 242, 242)
            .Font.Color = RGB(80, 80, 80)
        End With
    Next i
    For i = 5 To 7   ' ○●▲ は少し大きく
        pal.Cells(1, i).Font.size = 12
    Next i
    With pal.Cells(1, 1).Offset(0, -1)   ' A28 に見出し
        .Value = "シフトパレット"
        .Font.Bold = True
        .HorizontalAlignment = xlRight
    End With
    pal.Cells(1, ROW_CYCLE).Offset(MARKER_OFFSET, 0).Value = MARKER_CHAR
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    ShowMode ws, ROW_CYCLE
End Sub

Public Sub ShiftClick_セルフチェック()
    Dim ws As Worksheet, msg As String, srcS As String, srcP As String
    Set ws = ActiveSheet
    On Error GoTo EH
    srcS = "定数（フォールバック）": srcP = "定数（フォールバック）"
    On Error Resume Next
    If Not ws.Parent.Names(SHIFT_RANGE_NAME).RefersToRange Is Nothing Then srcS = "名前付き範囲"
    If Not ws.Parent.Names(PALETTE_NAME).RefersToRange Is Nothing Then srcP = "名前付き範囲"
    On Error GoTo EH
    msg = "シート名          : " & ws.Name & vbCrLf & _
          "シフト入力範囲    : " & ShiftRange(ws).Address(False, False) & "　←" & srcS & vbCrLf & _
          "パレット範囲      : " & PaletteRange(ws).Address(False, False) & "　←" & srcP & vbCrLf & _
          "パレットのセル数  : " & PaletteRange(ws).Cells.Count & vbCrLf & _
          "切替サイクル項目数: " & UBound(CycleValues(ws)) + 1 & vbCrLf & _
          "現在のモード番号  : " & CurrentIndex(ws) & vbCrLf & vbCrLf & _
          "ここまで表示されれば、標準モジュールは正しく入っています。"
    MsgBox msg, vbInformation, "ShiftClick セルフチェック"
    Exit Sub
EH:
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub

