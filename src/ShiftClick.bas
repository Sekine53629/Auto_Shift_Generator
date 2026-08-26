Option Explicit
'==================================================================
'  シフト表 クリック入力マクロ ＜標準モジュール ShiftClick v9.0＞
'  2026-08-26
'
'  シート上の位置は ShiftCommon が一元管理する。
'  本モジュールは「クリックしたときの振る舞い」だけを持つ。
'
'  パレット構成(横1行・表の上):
'    パレット本体行 = 日付行の PALETTE_GAP 行上
'    本体行 + MARKER_OFFSET = ★マーカー行
'    本体行 + LABEL_OFFSET  = ラベル行
'
'  v9.0 変更:
'   ・範囲/ラベル/列/オフセットの定数と解決処理を ShiftCommon に移管
'     (モジュールごとに別々のフォールバック番地を持たなくなった)
'   ・パレット位置を「過不足行の下」から「表の上」に変更
'   ・パレット生成は ShiftSetup_パレット生成 に一本化(重複解消)
'   ・全プロシージャに ErrorLogger のエラーハンドラを追加
'==================================================================
Private Const MODULE_NAME As String = "ShiftClick"

'------------------------ 設定 ここから ---------------------------
' 操作方法: "double" = ダブルクリックで反応 / "single" = クリックでも反応
Public Const CYCLE_TRIGGER As String = "double"
Public Const STAMP_TRIGGER As String = "double"

Public Const MARKER_CHAR   As String = "★"

' True: スタンプ時にパレットの背景色も反映する
Public Const APPLY_FILL         As Boolean = True
' True: 連続切替のときに背景色をクリアする(土日の色を残すなら False)
Public Const CYCLE_RESETS_FILL  As Boolean = False
' 入力後にカーソルを動かす: "" / "down" / "right"
Public Const MOVE_AFTER         As String = ""
' True: 数式の入ったセルは書き換えない
Public Const SKIP_FORMULA_CELLS As Boolean = True
'------------------------ 設定 ここまで ---------------------------

'--- 連続切替の巡回リスト: 空白＋パレットのスタンプ項目から自動生成 ---
Private Function CycleValues(ByVal ws As Worksheet) As Variant
    Dim pal As Range, arr() As Variant, i As Long, n As Long, v As String
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ws)
20  If pal Is Nothing Then
30      CycleValues = Array("")
40      Exit Function
50  End If
60  ReDim arr(0 To pal.Cells.Count)
70  arr(0) = ""            ' 先頭は空白(消去)
80  n = 0
90  For i = IDX_ERASE To pal.Cells.Count
100     v = Trim$(CStr(pal.Cells(1, i).Value))
110     If Len(v) > 0 Then
120         n = n + 1
130         arr(n) = v
140     End If
150 Next i
160 ReDim Preserve arr(0 To n)
170 CycleValues = arr
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CycleValues", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; i=" & i & "; n=" & n
    CycleValues = Array("")
End Function

'==================================================================
' メインハンドラ(シートモジュールから呼ばれる)
'==================================================================
Public Sub ShiftClick_Handle(ByVal Target As Range, _
                             ByVal EventKind As String, _
                             ByRef Handled As Boolean)
    Dim ws As Worksheet, pal As Range, area As Range, grid As Range
    Dim idx As Long
    On Error GoTo ErrHandler

10  Handled = False
20  Set ws = Target.Worksheet
30  Set pal = PaletteRange(ws)
40  If pal Is Nothing Then Exit Sub

    '--- パレット上のクリック: モード切替 ---
50  If Not Application.Intersect(Target, pal) Is Nothing Then
60      If EventKind = "right" Then Exit Sub
70      SetStamp ws, Target.Cells(1, 1)
80      Handled = True
90      Exit Sub
100 End If

    '--- シフト入力範囲との交差判定 ---
110 Set grid = ShiftInputRange(ws)
120 If grid Is Nothing Then Exit Sub
130 Set area = Application.Intersect(ClickRange(Target, EventKind), grid)
140 If area Is Nothing Then Exit Sub
150 If Application.CutCopyMode <> False Then Exit Sub

160 idx = CurrentIndex(ws)
170 If idx = IDX_OFF Then Exit Sub

180 Application.EnableEvents = False
190 If idx = IDX_CYCLE Then
200     If area.Cells.Count = 1 Then
210         If EventKind = "right" Then
220             CycleOne area.Cells(1, 1), True
230             Handled = True
240         ElseIf Triggered(EventKind, CYCLE_TRIGGER) Then
250             CycleOne area.Cells(1, 1), False
260             Handled = True
270         End If
280     End If
290 Else
300     If EventKind = "right" Or Triggered(EventKind, STAMP_TRIGGER) Then
310         StampArea ws, area, idx
320         Handled = True
330     End If
340 End If

350 If Handled And Len(MOVE_AFTER) > 0 Then
        Select Case LCase$(MOVE_AFTER)
            Case "down":  MoveSel area, 1, 0
            Case "right": MoveSel area, 0, 1
        End Select
360 End If

CleanUp:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftClick_Handle", Err.Number, Err.Description, Erl, _
             "eventKind=" & EventKind & "; target=" & Target.Address(False, False) & _
             "; idx=" & idx
    Resume CleanUp
End Sub

'==================================================================
' 内部処理
'==================================================================
Private Function Triggered(ByVal EventKind As String, ByVal Trig As String) As Boolean
    On Error GoTo ErrHandler

10  If EventKind = "double" Then
20      Triggered = True
30  ElseIf EventKind = "select" Then
40      Triggered = (LCase$(Trim$(Trig)) = "single")
50  Else
60      Triggered = False
70  End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "Triggered", Err.Number, Err.Description, Erl, _
             "eventKind=" & EventKind & "; trigger=" & Trig
    Triggered = False
End Function

'--- 右クリック時は選択範囲全体を対象にする ---
Private Function ClickRange(ByVal Target As Range, ByVal EventKind As String) As Range
    Dim sel As Range
    On Error GoTo ErrHandler

10  Set ClickRange = Target
20  If EventKind <> "right" Then Exit Function
30  If Not TypeOf Selection Is Range Then Exit Function
40  Set sel = Selection
50  If sel Is Nothing Then Exit Function
60  If sel.Worksheet.Name <> Target.Worksheet.Name Then Exit Function
70  If Not Application.Intersect(sel, Target) Is Nothing Then Set ClickRange = sel
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClickRange", Err.Number, Err.Description, Erl, _
             "eventKind=" & EventKind & "; target=" & Target.Address(False, False)
    Set ClickRange = Target
End Function

'--- 範囲へのスタンプ適用(ダブルクリック時／選択範囲実行時 共用) ---
Private Sub StampArea(ByVal ws As Worksheet, ByVal area As Range, ByVal idx As Long)
    Dim c As Range, pal As Range
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ws)
20  If pal Is Nothing Then Exit Sub
30  For Each c In area.Cells
40      If idx = IDX_CLEARFILL Then
50          c.Interior.Pattern = xlNone
60      Else
70          ApplyStamp c, pal.Cells(1, idx)
80      End If
90  Next c
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "StampArea", Err.Number, Err.Description, Erl, _
             "area=" & area.Address(False, False) & "; idx=" & idx
End Sub

'--- 1セルを次(前)の記号へ送る ---
Private Sub CycleOne(ByVal c As Range, ByVal Reverse As Boolean)
    Dim v As Variant, i As Long, idx As Long, cur As String
    On Error GoTo ErrHandler

10  If SKIP_FORMULA_CELLS Then
20      If c.HasFormula Then Exit Sub
30  End If
40  If c.Worksheet.ProtectContents And c.Locked Then Exit Sub

50  v = CycleValues(c.Worksheet)
60  cur = Trim$(CStr(c.Value))
70  idx = -1
80  For i = LBound(v) To UBound(v)
90      If cur = CStr(v(i)) Then
100         idx = i
110         Exit For
120     End If
130 Next i
140 If idx = -1 Then idx = LBound(v)
150 If Reverse Then idx = idx - 1 Else idx = idx + 1
160 If idx > UBound(v) Then idx = LBound(v)
170 If idx < LBound(v) Then idx = UBound(v)

180 If Len(CStr(v(idx))) = 0 Then
190     c.ClearContents
200 Else
210     c.Value = v(idx)
220 End If
230 If CYCLE_RESETS_FILL Then c.Interior.Pattern = xlNone
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "CycleOne", Err.Number, Err.Description, Erl, _
             "cell=" & c.Address(False, False) & "; reverse=" & Reverse & "; idx=" & idx
End Sub

'--- パレットのセルの値と書式をコピーする ---
Private Sub ApplyStamp(ByVal c As Range, ByVal src As Range)
    On Error GoTo ErrHandler

10  If SKIP_FORMULA_CELLS Then
20      If c.HasFormula Then Exit Sub
30  End If
40  If c.Worksheet.ProtectContents And c.Locked Then Exit Sub

50  If Len(Trim$(CStr(src.Value))) = 0 Then
60      c.ClearContents
70  Else
80      c.Value = src.Value
90  End If
100 c.Font.Color = src.Font.Color
110 c.Font.Bold = src.Font.Bold
120 If APPLY_FILL Then
130     If src.Interior.Pattern <> xlNone Then
140         c.Interior.Color = src.Interior.Color
150     End If
160 End If
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "ApplyStamp", Err.Number, Err.Description, Erl, _
             "cell=" & c.Address(False, False) & "; src=" & src.Address(False, False)
End Sub

'--- 選択中のモード番号(★マーカーの位置) ---
Private Function CurrentIndex(ByVal ws As Worksheet) As Long
    Dim pal As Range, i As Long
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ws)
20  If pal Is Nothing Then
30      CurrentIndex = IDX_CYCLE
40      Exit Function
50  End If
60  For i = 1 To pal.Cells.Count
70      If StrComp(Trim$(CStr(pal.Cells(1, i).Offset(MARKER_OFFSET, 0).Value)), _
                   MARKER_CHAR, vbTextCompare) = 0 Then
80          CurrentIndex = i
90          Exit Function
100     End If
110 Next i
120 CurrentIndex = IDX_CYCLE
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CurrentIndex", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; i=" & i
    CurrentIndex = IDX_CYCLE
End Function

'--- クリックされたパレットのセルをモードに設定する ---
Private Sub SetStamp(ByVal ws As Worksheet, ByVal c As Range)
    Dim pal As Range, i As Long, hit As Long
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ws)
20  If pal Is Nothing Then Exit Sub
30  hit = 0
40  For i = 1 To pal.Cells.Count
50      If pal.Cells(1, i).Address = c.Address Then hit = i
60  Next i
70  If hit = 0 Then Exit Sub
    ' 同じモードを再度クリックしたら連続切替に戻す
80  If hit = CurrentIndex(ws) And hit <> IDX_CYCLE Then hit = IDX_CYCLE

90  Application.EnableEvents = False
100 pal.Offset(MARKER_OFFSET, 0).ClearContents
110 pal.Cells(1, hit).Offset(MARKER_OFFSET, 0).Value = MARKER_CHAR

CleanUp:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
    ShowMode ws, hit
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "SetStamp", Err.Number, Err.Description, Erl, _
             "cell=" & c.Address(False, False) & "; hit=" & hit
    Resume CleanUp
End Sub

'--- ステータスバーに現在のモードを表示 ---
Private Sub ShowMode(ByVal ws As Worksheet, ByVal idx As Long)
    Dim pal As Range, s As String, op As String
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ws)
20  If pal Is Nothing Then Exit Sub
        Select Case idx
            Case IDX_OFF
                s = "シフト入力マクロ: OFF(通常のExcel操作)"
            Case IDX_CYCLE
                If LCase$(CYCLE_TRIGGER) = "single" Then op = "クリック" Else op = "ダブルクリック"
                s = "シフト入力マクロ: 連続切替　" & op & "で次の記号／右クリックで前へ"
            Case IDX_CLEARFILL
                s = "シフト入力マクロ: 背景色クリア"
            Case Else
                If LCase$(STAMP_TRIGGER) = "single" Then op = "クリック" Else op = "ダブルクリック"
                s = "シフト入力マクロ: スタンプ [ " & _
                    Trim$(CStr(pal.Cells(1, idx).Value)) & " ]　" & _
                    op & "で押す／範囲を選んで右クリックでまとめて押す"
        End Select
30  Application.StatusBar = s
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "ShowMode", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; idx=" & idx
End Sub

'--- 入力後のカーソル移動 ---
Private Sub MoveSel(ByVal area As Range, ByVal dr As Long, ByVal dc As Long)
    Dim t As Range
    On Error GoTo ErrHandler

10  Set t = area.Cells(area.Cells.Count).Offset(dr, dc)
20  If Not t Is Nothing Then t.Select
    Exit Sub
ErrHandler:
    ' シートの端では Offset が失敗する(想定内)ため通知しない
    LogError MODULE_NAME, "MoveSel", Err.Number, Err.Description, Erl, _
             "area=" & area.Address(False, False) & "; dr=" & dr & "; dc=" & dc
End Sub

'==================================================================
' 手動実行マクロ
'==================================================================
Public Sub ShiftClick_選択範囲にスタンプ()
    Dim ws As Worksheet, area As Range, grid As Range, idx As Long
    On Error GoTo ErrHandler

10  Set ws = ActiveSheet
20  If Not TypeOf Selection Is Range Then Exit Sub
30  Set grid = ShiftInputRange(ws)
40  If grid Is Nothing Then
50      MsgBox "シフト入力欄が特定できません。" & vbCrLf & _
               "ShiftClick_セルフチェック で範囲を確認してください。", vbExclamation
60      Exit Sub
70  End If
80  Set area = Application.Intersect(Selection, grid)
90  If area Is Nothing Then
100     MsgBox "シフト入力欄(" & grid.Address(False, False) & _
               ")を選んでから実行してください。", vbExclamation
110     Exit Sub
120 End If
130 idx = CurrentIndex(ws)
140 If idx <= IDX_CYCLE Then
150     MsgBox "パレットでスタンプする記号を選んでから実行してください。", vbExclamation
160     Exit Sub
170 End If

180 Application.EnableEvents = False
190 StampArea ws, area, idx

CleanUp:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftClick_選択範囲にスタンプ", _
               "Stamped " & area.Cells.Count & " cells with palette index " & idx
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftClick_選択範囲にスタンプ", Err.Number, Err.Description, Erl, _
             "idx=" & idx
    MsgBox "スタンプでエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

Public Sub ShiftClick_連続切替に戻す()
    Dim ws As Worksheet, pal As Range
    On Error GoTo ErrHandler

10  Set ws = ActiveSheet
20  Set pal = PaletteRange(ws)
30  If pal Is Nothing Then
40      MsgBox "パレットが特定できません。" & vbCrLf & _
               "ShiftSetup_パレット生成 で作成してください。", vbExclamation
50      Exit Sub
60  End If
70  SetStamp ws, pal.Cells(1, IDX_CYCLE)

    LogSuccess MODULE_NAME, "ShiftClick_連続切替に戻す", "Reset palette mode to cycle"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "ShiftClick_連続切替に戻す", Err.Number, Err.Description, Erl, _
             "sheet=" & ActiveSheet.Name
    MsgBox "モードの切替でエラーが発生しました: " & Err.Description, vbExclamation
End Sub

'==================================================================
' パターン追加: パレット右端に新しいシフトパターンを追加
'   (名前付き範囲はラベル行を基準に自動拡張される)
'==================================================================
Public Sub ShiftClick_パターン追加()
    Dim ws As Worksheet, pal As Range, lastCell As Range, newCell As Range
    Dim sym As String, lab As String, i As Long
    On Error GoTo ErrHandler

10  Set ws = ActiveSheet
20  Set pal = PaletteRange(ws)
30  If pal Is Nothing Then
40      MsgBox "パレットが特定できません。" & vbCrLf & _
               "ShiftSetup_パレット生成 で作成してください。", vbExclamation
50      Exit Sub
60  End If

70  sym = Trim$(InputBox("追加するシフト記号(セルに入力される値):", "パターン追加"))
80  If Len(sym) = 0 Then Exit Sub

    '--- 重複チェック ---
90  For i = 1 To pal.Cells.Count
100     If Trim$(CStr(pal.Cells(1, i).Value)) = sym Then
110         MsgBox "[ " & sym & " ] は既にパレットにあります。", vbExclamation
120         Exit Sub
130     End If
140 Next i

150 lab = Trim$(InputBox("ラベル(パレット下の説明):", "パターン追加", sym))
160 If Len(lab) = 0 Then lab = sym

170 Set lastCell = pal.Cells(1, pal.Cells.Count)
180 Set newCell = lastCell.Offset(0, 1)

190 Application.EnableEvents = False
    '--- 書式を右端のセルからコピー(マーカー行・本体・ラベル行) ---
200 lastCell.Offset(MARKER_OFFSET, 0).Resize(3, 1).Copy
210 newCell.Offset(MARKER_OFFSET, 0).PasteSpecial xlPasteFormats
220 Application.CutCopyMode = False
    '--- 値とラベルを書き込み → ラベルが入った時点で名前付き範囲が自動拡張 ---
230 newCell.Value = sym
240 newCell.Offset(LABEL_OFFSET, 0).Value = lab
250 Application.EnableEvents = True

260 MsgBox "パターン [ " & sym & " ] を追加しました。" & vbCrLf & _
           "パレット範囲: " & PaletteRange(ws).Address(False, False) & vbCrLf & vbCrLf & _
           "※パレットのセルに文字色・背景色を付けると、" & vbCrLf & _
           "　スタンプ時にその書式もコピーされます。", vbInformation

CleanUp:
    On Error Resume Next
    Application.CutCopyMode = False
    Application.EnableEvents = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftClick_パターン追加", _
               "Added palette pattern: symbol=" & sym & ", label=" & lab
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftClick_パターン追加", Err.Number, Err.Description, Erl, _
             "symbol=" & sym & "; label=" & lab & "; i=" & i
    MsgBox "パターン追加でエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'==================================================================
' パレット再作成
'   実体は ShiftSetup_パレット生成 に一本化(v9.0 で重複を解消)
'   既存のボタン割当を壊さないため、入口だけ残してある。
'==================================================================
Public Sub ShiftClick_パレット作成()
    On Error GoTo ErrHandler

10  ShiftSetup_パレット生成
20  ShowMode ActiveSheet, IDX_CYCLE

    LogSuccess MODULE_NAME, "ShiftClick_パレット作成", _
               "Delegated palette build to ShiftSetup_パレット生成"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "ShiftClick_パレット作成", Err.Number, Err.Description, Erl, _
             "sheet=" & ActiveSheet.Name
    MsgBox "パレット作成でエラーが発生しました: " & Err.Description, vbExclamation
End Sub

'==================================================================
' セルフチェック
'==================================================================
Public Sub ShiftClick_セルフチェック()
    Dim ws As Worksheet, msg As String, srcS As String, srcP As String
    Dim grid As Range, pal As Range, drift As Long, docRow As Long
    Dim warn As String, palRow As Long
    On Error GoTo ErrHandler

10  Set ws = ActiveSheet
20  If NamedRangeOrNothing(NM_SHIFT) Is Nothing Then
30      srcS = "ラベルから計算"
40  Else
50      srcS = "名前付き範囲"
60  End If
70  If NamedRangeOrNothing(NM_PALETTE) Is Nothing Then
80      srcP = "ラベルから計算"
90  Else
100     srcP = "名前付き範囲"
110 End If

120 Set grid = ShiftInputRange(ws)
130 Set pal = PaletteRange(ws)
140 docRow = LabelRow(ws, LBL_DOC)
150 palRow = PaletteBodyRow(ws)

    '--- 入力範囲の妥当性(下端が医師数の DOC_GAP 行上か) ---
160 drift = ShiftRangeDrift(ws)
170 If grid Is Nothing Then
180     warn = vbCrLf & "■ シフト入力欄を特定できません" & vbCrLf & _
               "　B列の開始日の数式、またはA列の「" & LBL_NOTE & "」「" & LBL_DOC & _
               "」を確認してください" & vbCrLf
190 ElseIf drift > 0 Then
200     warn = vbCrLf & "■ 範囲の下端が " & LBL_DOC & "(" & docRow & "行)の" & _
               DOC_GAP & "行上を " & drift & " 行超えています" & vbCrLf & _
               "　正しい終端: " & (docRow - DOC_GAP) & "行" & vbCrLf
210 ElseIf drift < 0 Then
220     warn = vbCrLf & "※ 範囲の下端が " & (docRow - DOC_GAP) & "行 より " & _
               Abs(drift) & " 行上です(入力できる行が少なくなっています)" & vbCrLf
230 End If

240 msg = "シート名          : " & ws.Name & vbCrLf & _
          "シフト入力範囲    : " & _
              IIf(grid Is Nothing, "(未特定)", grid.Address(False, False)) & _
              "　←" & srcS & vbCrLf & _
          "　上端=日付行の1行下 / 下端=" & LBL_DOC & "の" & DOC_GAP & "行上" & vbCrLf & _
          "パレット範囲      : " & _
              IIf(pal Is Nothing, "(未特定)", pal.Address(False, False)) & _
              "　←" & srcP & vbCrLf & _
          "　パレット本体行  : " & palRow & " 行(日付行の" & PALETTE_GAP & "行上)" & vbCrLf & _
          "パレットのセル数  : " & IIf(pal Is Nothing, 0, pal.Cells.Count) & vbCrLf & _
          "切替サイクル項目数: " & UBound(CycleValues(ws)) + 1 & vbCrLf & _
          "現在のモード番号  : " & CurrentIndex(ws) & vbCrLf & warn & vbCrLf & _
          "ここまで表示されれば、標準モジュールは正しく入っています。"
250 MsgBox msg, vbInformation, "ShiftClick セルフチェック"

    LogSuccess MODULE_NAME, "ShiftClick_セルフチェック", _
               "Reported ranges: grid=" & IIf(grid Is Nothing, "none", grid.Address(False, False)) & _
               ", drift=" & drift
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftClick_セルフチェック", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; docRow=" & docRow & "; paletteRow=" & palRow
    MsgBox "セルフチェックでエラーが発生しました: " & Err.Description, vbExclamation
End Sub
