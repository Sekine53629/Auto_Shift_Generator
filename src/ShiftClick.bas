Option Explicit

'==================================================================
'  シフト表 クリック入力マクロ  ＜標準モジュール＞
'  最終更新: 2026-08-20 (v3 : ダブルクリック方式に変更)
'
'  ・シフト範囲を ダブルクリック すると
'      空白→◯→●→▲→公休→希休→空白 と切り替わる
'  ・右クリックで 1 つ前に戻す
'  ・パレットの記号をクリックすると「スタンプモード」
'  ・パレットのセルの文字色・背景色ごとスタンプされるので
'    「緑地の◯」もパレットのセルを緑に塗るだけで使える
'
'  ※ セルを選ぶだけのクリックでは値は変わりません。
'==================================================================

'------------------------ 設定 ここから ---------------------------

' シフト入力欄（マクロが効く範囲）
Public Const SHIFT_RANGE   As String = "C5:AG40"

' パレットを置く範囲（縦1列・11行）
Public Const PALETTE_RANGE As String = "AJ4:AJ14"

' ★操作方法★
'   "double" = ダブルクリックだけで反応（誤操作なし・おすすめ）
'   "single" = シングルクリックでも反応（旧方式）
Public Const CYCLE_TRIGGER As String = "double"   ' 連続切替のとき
Public Const STAMP_TRIGGER As String = "double"   ' スタンプモードのとき

' 選択中マーク（★）を出す列。-1 = パレットの1つ左の列
Public Const MARKER_OFFSET As Long = -1
' 説明ラベルを書く列。1 = パレットの1つ右の列
Public Const LABEL_OFFSET  As Long = 1
Public Const MARKER_CHAR   As String = "★"

' True: スタンプ時にパレットの背景色も反映する（色なしのパレットは色を変えない）
Public Const APPLY_FILL          As Boolean = True
' True: 連続切替のときに背景色をクリアする（土日の色を残すなら False）
Public Const CYCLE_RESETS_FILL   As Boolean = False
' 入力後にカーソルを動かす: "" / "down" / "right"
Public Const MOVE_AFTER          As String = ""
' True: 数式の入ったセルは書き換えない
Public Const SKIP_FORMULA_CELLS  As Boolean = True

' パレットの特別な行（上から何番目か）
Public Const ROW_OFF       As Long = 1   ' OFF（マクロを止める）
Public Const ROW_CYCLE     As Long = 2   ' 連続切替モード
Public Const ROW_CLEARFILL As Long = 3   ' 背景色だけ消す
' 4 行目以降 = ふつうのスタンプ

' 連続切替の順番（自由に足し引きできます。"" は空白）
Private Function CycleValues() As Variant
    CycleValues = Array("", "◯", "●", "▲", "公休", "希休")
End Function

'------------------------ 設定 ここまで ---------------------------


'==================================================================
' シートモジュールから呼ばれる本体
'   EventKind : "select" = シングルクリック / "double" = ダブルクリック
'               "right"  = 右クリック
'==================================================================
Public Sub ShiftClick_Handle(ByVal Target As Range, _
                             ByVal EventKind As String, _
                             ByRef Handled As Boolean)
    Dim ws As Worksheet, pal As Range, area As Range, c As Range
    Dim idx As Long

    Handled = False
    On Error GoTo EH

    Set ws = Target.Worksheet
    Set pal = ws.Range(PALETTE_RANGE)

    '--- パレットの操作（ここはシングルクリックでOK。値は変わらないので）---
    If Not Application.Intersect(Target, pal) Is Nothing Then
        If EventKind = "right" Then Exit Sub   '右クリックは通常メニュー（書式変更用）
        SetStamp ws, Target.Cells(1, 1)
        Handled = True
        Exit Sub
    End If

    '--- シフト範囲外なら何もしない ---
    Set area = Application.Intersect(ClickRange(Target, EventKind), ws.Range(SHIFT_RANGE))
    If area Is Nothing Then Exit Sub

    ' コピー中はコピー操作を邪魔しない
    If Application.CutCopyMode <> False Then Exit Sub

    idx = CurrentIndex(ws)
    If idx = ROW_OFF Then Exit Sub

    Application.EnableEvents = False

    If idx = ROW_CYCLE Then
        '--- 連続切替モード（安全のため単一セルのときだけ）---
        If area.Cells.Count = 1 Then
            If EventKind = "right" Then
                CycleOne area.Cells(1, 1), True       '右クリック = 1つ戻す
                Handled = True
            ElseIf Triggered(EventKind, CYCLE_TRIGGER) Then
                CycleOne area.Cells(1, 1), False
                Handled = True
            End If
        End If
    Else
        '--- スタンプモード ---
        '    右クリック = 選択範囲にまとめて押す
        If EventKind = "right" Or Triggered(EventKind, STAMP_TRIGGER) Then
            For Each c In area.Cells
                If idx = ROW_CLEARFILL Then
                    c.Interior.Pattern = xlNone
                Else
                    ApplyStamp c, pal.Cells(idx, 1)
                End If
            Next c
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
' その操作で動かしてよいか判定
'   ダブルクリックは設定に関わらず常に有効
'   シングルクリックは設定が "single" のときだけ有効
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


'==================================================================
' 右クリック時は「選択範囲全体」を対象にする（まとめてスタンプ用）
'==================================================================
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


'==================================================================
' 1 セルを次（または前）の値に切り替える
'==================================================================
Private Sub CycleOne(ByVal c As Range, ByVal Reverse As Boolean)
    Dim v As Variant, i As Long, idx As Long, cur As String

    If SKIP_FORMULA_CELLS Then
        If c.HasFormula Then Exit Sub
    End If
    If c.Worksheet.ProtectContents And c.Locked Then Exit Sub

    v = CycleValues()
    cur = Trim$(CStr(c.Value))

    idx = -1
    For i = LBound(v) To UBound(v)
        If cur = CStr(v(i)) Then
            idx = i
            Exit For
        End If
    Next i
    If idx = -1 Then idx = LBound(v)   ' リストにない値は空白扱いにして次へ

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


'==================================================================
' パレットの内容を 1 セルに転写する
'==================================================================
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
        ' パレット側が「塗りつぶしなし」のときは、元の色（土日の色など）を残す
        If src.Interior.Pattern <> xlNone Then
            c.Interior.Color = src.Interior.Color
        End If
    End If
End Sub


'==================================================================
' 現在選ばれているパレット行（1 始まり）
'==================================================================
Private Function CurrentIndex(ByVal ws As Worksheet) As Long
    Dim pal As Range, i As Long
    Set pal = ws.Range(PALETTE_RANGE)
    For i = 1 To pal.Cells.Count
        If StrComp(Trim$(CStr(pal.Cells(i, 1).Offset(0, MARKER_OFFSET).Value)), _
                   MARKER_CHAR, vbTextCompare) = 0 Then
            CurrentIndex = i
            Exit Function
        End If
    Next i
    CurrentIndex = ROW_CYCLE     ' マークが無いときは連続切替
End Function


'==================================================================
' パレットのセルを選択中にする（同じセルをダブルクリックで連続切替に戻る）
'==================================================================
Private Sub SetStamp(ByVal ws As Worksheet, ByVal c As Range)
    Dim pal As Range, i As Long, hit As Long
    Set pal = ws.Range(PALETTE_RANGE)

    hit = 0
    For i = 1 To pal.Cells.Count
        If pal.Cells(i, 1).Address = c.Address Then hit = i
    Next i
    If hit = 0 Then Exit Sub

    If hit = CurrentIndex(ws) And hit <> ROW_CYCLE Then hit = ROW_CYCLE  ' 再選択で解除

    Application.EnableEvents = False
    On Error Resume Next
    pal.Offset(0, MARKER_OFFSET).ClearContents
    pal.Cells(hit, 1).Offset(0, MARKER_OFFSET).Value = MARKER_CHAR
    On Error GoTo 0
    Application.EnableEvents = True

    ShowMode ws, hit
End Sub


Private Sub ShowMode(ByVal ws As Worksheet, ByVal idx As Long)
    Dim pal As Range, s As String, op As String
    Set pal = ws.Range(PALETTE_RANGE)

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
                Trim$(CStr(pal.Cells(idx, 1).Offset(0, LABEL_OFFSET).Value)) & " ]　" & _
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
' 選択範囲にまとめてスタンプ（ショートカットキーに割り当てても便利）
'==================================================================
Public Sub ShiftClick_選択範囲にスタンプ()
    Dim ws As Worksheet, area As Range, c As Range, idx As Long
    Set ws = ActiveSheet
    If Not TypeOf Selection Is Range Then Exit Sub
    Set area = Application.Intersect(Selection, ws.Range(SHIFT_RANGE))
    If area Is Nothing Then
        MsgBox "シフト入力欄（" & SHIFT_RANGE & "）を選んでから実行してください。", vbExclamation
        Exit Sub
    End If

    idx = CurrentIndex(ws)
    If idx <= ROW_CYCLE Then
        MsgBox "パレットでスタンプする記号を選んでから実行してください。", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False
    For Each c In area.Cells
        If idx = ROW_CLEARFILL Then
            c.Interior.Pattern = xlNone
        Else
            ApplyStamp c, ws.Range(PALETTE_RANGE).Cells(idx, 1)
        End If
    Next c
    Application.EnableEvents = True
End Sub


'==================================================================
' 連続切替モードに戻す
'==================================================================
Public Sub ShiftClick_連続切替に戻す()
    SetStamp ActiveSheet, ActiveSheet.Range(PALETTE_RANGE).Cells(ROW_CYCLE, 1)
End Sub


'==================================================================
' ★最初に 1 回だけ実行★  パレットをシート上に作ります
'==================================================================
Public Sub ShiftClick_パレット作成()
    Dim ws As Worksheet, pal As Range, i As Long
    Dim vals As Variant, labs As Variant

    Set ws = ActiveSheet
    Set pal = ws.Range(PALETTE_RANGE)

    vals = Array("OFF", "連続切替", "色クリア", "", "◯", "●", "▲", "公休", "希休", "◯", "●")
    labs = Array("マクロを止める", "ダブルクリックで順に切替", "背景色だけ消す", "消去（空白にする）", _
                 "◯", "●", "▲", "公休", "希休", "◯（緑）", "●（緑）")

    If pal.Cells.Count <> UBound(vals) + 1 Then
        MsgBox "PALETTE_RANGE の行数（" & pal.Cells.Count & "行）と項目数（" & _
               UBound(vals) + 1 & "個）が合っていません。", vbExclamation
        Exit Sub
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    With pal
        .ClearContents
        .Interior.Pattern = xlNone
        .Font.Color = RGB(0, 0, 0)
        .Font.Bold = False
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(150, 150, 150)
        .EntireColumn.ColumnWidth = 9
    End With
    pal.Offset(0, MARKER_OFFSET).ClearContents
    pal.Offset(0, MARKER_OFFSET).EntireColumn.ColumnWidth = 3
    pal.Offset(0, MARKER_OFFSET).HorizontalAlignment = xlCenter
    pal.Offset(0, MARKER_OFFSET).Font.Color = RGB(192, 0, 0)
    pal.Offset(0, LABEL_OFFSET).ClearContents
    pal.Offset(0, LABEL_OFFSET).EntireColumn.ColumnWidth = 22
    pal.Offset(0, LABEL_OFFSET).Font.Size = 9
    pal.Offset(0, LABEL_OFFSET).Font.Color = RGB(100, 100, 100)

    For i = 1 To pal.Cells.Count
        If Len(CStr(vals(i - 1))) > 0 Then pal.Cells(i, 1).Value = vals(i - 1)
        pal.Cells(i, 1).Offset(0, LABEL_OFFSET).Value = labs(i - 1)
        pal.Cells(i, 1).RowHeight = 20
    Next i

    For i = 1 To 3
        With pal.Cells(i, 1)
            .Font.Size = 10
            .Interior.Color = RGB(242, 242, 242)
            .Font.Color = RGB(80, 80, 80)
        End With
    Next i

    ' 緑地の ◯ ● （最後の2つ）
    For i = pal.Cells.Count - 1 To pal.Cells.Count
        With pal.Cells(i, 1)
            .Interior.Color = RGB(198, 239, 206)
            .Font.Color = RGB(0, 97, 0)
            .Font.Bold = True
        End With
    Next i

    With pal.Cells(1, 1).Offset(-1, 0)
        .Value = "パレット"
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    pal.Cells(ROW_CYCLE, 1).Offset(0, MARKER_OFFSET).Value = MARKER_CHAR

    Application.ScreenUpdating = True
    Application.EnableEvents = True

    ShowMode ws, ROW_CYCLE
    MsgBox "パレットを作成しました。" & vbCrLf & vbCrLf & _
           "【連続切替】シフト欄をダブルクリック → 次の記号へ" & vbCrLf & _
           "　　　　　　右クリック → 1つ前に戻る" & vbCrLf & vbCrLf & _
           "【スタンプ】パレットの記号をクリックして選択" & vbCrLf & _
           "　　　　　　→ シフト欄をダブルクリックで押す" & vbCrLf & _
           "　　　　　　→ 範囲を選んで右クリックでまとめて押す" & vbCrLf & vbCrLf & _
           "セルを選ぶだけのクリックでは値は変わりません。", vbInformation
End Sub


'==================================================================
' ★動作確認用★  設定がそろっているか調べます（Alt+F8 から実行）
'==================================================================
Public Sub ShiftClick_セルフチェック()
    Dim ws As Worksheet, msg As String
    Set ws = ActiveSheet
    On Error GoTo EH
    msg = "シート名          : " & ws.Name & vbCrLf & _
          "SHIFT_RANGE       : " & SHIFT_RANGE & vbCrLf & _
          "PALETTE_RANGE     : " & PALETTE_RANGE & vbCrLf & _
          "CYCLE_TRIGGER     : " & CYCLE_TRIGGER & vbCrLf & _
          "STAMP_TRIGGER     : " & STAMP_TRIGGER & vbCrLf & _
          "パレットのセル数  : " & ws.Range(PALETTE_RANGE).Cells.Count & vbCrLf & _
          "現在のモード番号  : " & CurrentIndex(ws) & vbCrLf & vbCrLf & _
          "ここまで表示されれば、標準モジュールは正しく入っています。"
    MsgBox msg, vbInformation, "ShiftClick セルフチェック"
    Exit Sub
EH:
    MsgBox "エラー: " & Err.Description, vbExclamation
End Sub
