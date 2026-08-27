Option Explicit
'==================================================================
'  シフト表 クリック入力マクロ ＜標準モジュール ShiftClick v9.4＞
'  2026-08-27
'
'  シート上の位置は ShiftCommon が一元管理する。
'  本モジュールは「クリックしたときの振る舞い」だけを持つ。
'
'  パレット構成(横1行・表の上):
'    パレット本体行 = 日付行の PALETTE_GAP 行上
'    本体行 + MARKER_OFFSET = ★マーカー行
'    本体行 + LABEL_OFFSET  = ラベル行
'
'  ボタンの種類(ShiftCommon の IDX_* が定義):
'    IDX_OFF        OFF    … マクロ停止
'    IDX_AUTO       自動   … ダブルクリックで AutoShift を起動
'    IDX_CYCLE      切替   … 順送り
'    IDX_CLEARFILL  色消   … 背景色だけ消す
'    IDX_FILL_FIRST..IDX_FILL_LAST
'                   背景緑/橙/灰 … 値を書かず背景色だけ塗る
'    IDX_ERASE      消去   … 空白スタンプ。記号はこの次から
'
'  v9.4 変更:
'   ・スタンプが背景色と文字色を持ち込まないようにした(APPLY_FILL 廃止)
'
'  v9.2 変更:
'   ・ShiftAutoBridge を廃止し、AutoShiftPreflight / ShiftAuto_事前診断 を
'     本モジュールへ移管(位置解決3関数は ShiftCommon v2.1 へ)
'
'  v9.1 変更:
'   ・「自動」ボタンを追加。ダブルクリックで AutoShift を起動する
'     (Application.Run で呼ぶため AutoShift 未実装でもコンパイル可)
'   ・背景色ボタンを追加。StampArea が「値を書かず塗るだけ」に分岐
'   ・順送りの巡回対象から医師名を除外できるようにした
'     (CYCLE_INCLUDES_DOCTORS で切替)
'   ・ShowMode に 自動 / 背景色 のメッセージを追加
'==================================================================
Private Const MODULE_NAME As String = "ShiftClick"

'------------------------ 設定 ここから ---------------------------
' 操作方法: "double" = ダブルクリックで反応 / "single" = クリックでも反応
Public Const CYCLE_TRIGGER As String = "double"
Public Const STAMP_TRIGGER As String = "double"


' True: 連続切替のときに背景色をクリアする(土日の色を残すなら False)
Public Const CYCLE_RESETS_FILL  As Boolean = False
' 入力後にカーソルを動かす: "" / "down" / "right"
Public Const MOVE_AFTER         As String = ""
' True: 数式の入ったセルは書き換えない
Public Const SKIP_FORMULA_CELLS As Boolean = True
' True: 「自動」実行前に確認ダイアログを出す
Public Const AUTO_CONFIRM       As Boolean = True
'------------------------ 設定 ここまで ---------------------------

'--- 連続切替の巡回リスト: 空白＋パレットのスタンプ項目から自動生成 ---
Private Function CycleValues(ByVal ws As Worksheet) As Variant
    Dim pal As Range, arr() As Variant, i As Long, n As Long, v As String
    Dim lastIdx As Long
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ws)
20  If pal Is Nothing Then
30      CycleValues = Array("")
40      Exit Function
50  End If

    '--- 巡回はシフト入力欄でしか働かないため、シフト記号だけを対象にする ---
    '    医師名スタンプと備考スタンプは、それぞれ専用の場所にしか押せない
60  lastIdx = IDX_DOC_FIRST - 1
70  If lastIdx > pal.Cells.Count Then lastIdx = pal.Cells.Count
100 ReDim arr(0 To lastIdx)
110 arr(0) = ""            ' 先頭は空白(消去)
120 n = 0
130 For i = IDX_ERASE To lastIdx
140     v = Trim$(CStr(pal.Cells(1, i).Value))
150     If Len(v) > 0 Then
160         n = n + 1
170         arr(n) = v
180     End If
190 Next i
200 ReDim Preserve arr(0 To n)
210 CycleValues = arr
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
    Dim ws As Worksheet, pal As Range, area As Range
    Dim kind As Long
    Dim idx As Long, palIdx As Long
    On Error GoTo ErrHandler

10  Handled = False
20  Set ws = Target.Worksheet
30  Set pal = PaletteRange(ws)
40  If pal Is Nothing Then Exit Sub

    '--- パレット上のクリック: モード切替 / 自動実行 ---
50  If Not Application.Intersect(Target, pal) Is Nothing Then
60      If EventKind = "right" Then Exit Sub
70      palIdx = PaletteIndexOf(ws, Target.Cells(1, 1))

        '--- 動作ボタン(自動/戻す/出力)はダブルクリックで即実行 ---
80      If IsActionButton(palIdx) Then
90          If EventKind = "double" Then
100             Handled = True
110             RunAction ws, palIdx
120         End If
130         Exit Sub
140     End If

150     SetStamp ws, Target.Cells(1, 1)
160     Handled = True
170     Exit Sub
180 End If

    '--- 書き込み先の判定: シフト入力欄 / 備考行 / 医師名欄 ---
210 Set area = ClickTargetArea(ws, ClickRange(Target, EventKind), kind)
220 If area Is Nothing Then Exit Sub
230 If Application.CutCopyMode <> False Then Exit Sub

240 idx = CurrentIndex(ws)
250 If idx = IDX_OFF Then Exit Sub
    '--- 動作ボタンがモードとして残っていても入力欄では何もしない ---
260 If IsActionButton(idx) Then Exit Sub
    '--- 記号の種類と書き込み先が噛み合わない組み合わせは無視する ---
265 If Not StampAllowedHere(idx, kind) Then Exit Sub

270 Application.EnableEvents = False
280 If idx = IDX_CYCLE Then
290     If area.Cells.Count = 1 Then
300         If EventKind = "right" Then
310             CycleOne area.Cells(1, 1), True
320             Handled = True
330         ElseIf Triggered(EventKind, CYCLE_TRIGGER) Then
340             CycleOne area.Cells(1, 1), False
350             Handled = True
360         End If
370     End If
380 Else
390     If EventKind = "right" Or Triggered(EventKind, STAMP_TRIGGER) Then
400         StampArea ws, area, idx
410         Handled = True
420     End If
430 End If

440 If Handled And Len(MOVE_AFTER) > 0 Then
        Select Case LCase$(MOVE_AFTER)
            Case "down":  MoveSel area, 1, 0
            Case "right": MoveSel area, 0, 1
        End Select
450 End If

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
' 「自動」ボタン
'==================================================================
'--- クリックされたセルがパレットの何番目かを返す(0 = パレット外) ---
Private Function PaletteIndexOf(ByVal ws As Worksheet, ByVal c As Range) As Long
    Dim pal As Range, i As Long
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ws)
20  If pal Is Nothing Then Exit Function
30  For i = 1 To pal.Cells.Count
40      If pal.Cells(1, i).Address = c.Address Then
50          PaletteIndexOf = i
60          Exit Function
70      End If
80  Next i
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "PaletteIndexOf", Err.Number, Err.Description, Erl, _
             "cell=" & c.Address(False, False)
    PaletteIndexOf = 0
End Function

'--- 動作ボタンを実行する ---
Private Sub RunAction(ByVal ws As Worksheet, ByVal palIdx As Long)
    On Error GoTo ErrHandler

        Select Case palIdx
            Case IDX_AUTO
                RunAutoShift ws
            Case IDX_UNDO
                Application.Run UNDO_MACRO
            Case IDX_EXPORT
                Application.Run EXPORT_MACRO
        End Select
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "RunAction", Err.Number, Err.Description, Erl, _
             "palIdx=" & palIdx
    MsgBox "実行できませんでした: " & Err.Description, vbExclamation
End Sub

'--- AutoShift を起動する ---
'    Application.Run で呼ぶため、AutoShift が未実装でもコンパイルは通る。
Private Sub RunAutoShift(ByVal ws As Worksheet)
    Dim savedMode As XlCalculation, savedEvents As Boolean
    Dim reason As String
    On Error GoTo ErrHandler

    '--- 事前診断: 対象月セル・入力範囲・依存シートを確認 ---
5   If Not AutoShiftPreflight(reason) Then
6       MsgBox "自動作成を実行できません。" & vbCrLf & vbCrLf & reason, _
               vbExclamation, "シフト自動作成"
7       Exit Sub
8   End If

10  If AUTO_CONFIRM Then
20      If MsgBox("シフトを自動作成します。" & vbCrLf & _
                  "入力済みのセルは保持され、空白セルのみ埋まります。" & vbCrLf & vbCrLf & _
                  "実行しますか?", _
                  vbQuestion + vbYesNo, "シフト自動作成") <> vbYes Then Exit Sub
30  End If

40  Application.StatusBar = "シフト自動作成を実行中..."
50  savedMode = Application.Calculation
60  savedEvents = Application.EnableEvents
70  Application.EnableEvents = False
80  Application.Calculation = xlCalculationManual

90  Application.Run AUTOSHIFT_MACRO

100 Application.Calculation = savedMode
110 Application.CalculateFull

CleanUp:
    On Error Resume Next
    Application.Calculation = savedMode
    Application.EnableEvents = True
    Application.StatusBar = False
    On Error GoTo 0
    LogSuccess MODULE_NAME, "RunAutoShift", "Ran macro: " & AUTOSHIFT_MACRO
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "RunAutoShift", Err.Number, Err.Description, Erl, _
             "macro=" & AUTOSHIFT_MACRO & "; sheet=" & ws.Name
    MsgBox "自動作成でエラーが発生しました:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "マクロ名 [ " & AUTOSHIFT_MACRO & " ] が存在するか確認してください。" & vbCrLf & _
           "(ShiftCommon の AUTOSHIFT_MACRO で変更できます)", _
           vbExclamation, "シフト自動作成"
    Resume CleanUp
End Sub

'==================================================================
' 内部処理
'==================================================================
'--- クリック先の種別を判定して対象範囲を返す ---
'    シフト入力欄 → 備考行 → 医師名欄 の順に見る。
'    どれでもなければ Nothing (kind = TGT_NONE)。
Private Function ClickTargetArea(ByVal ws As Worksheet, ByVal clicked As Range, _
                                 ByRef kind As Long) As Range
    Dim rng As Range
    On Error GoTo ErrHandler

10  kind = TGT_NONE

20  Set rng = ShiftInputRange(ws)
30  If Not rng Is Nothing Then
40      Set ClickTargetArea = Application.Intersect(clicked, rng)
50      If Not ClickTargetArea Is Nothing Then
60          kind = TGT_SHIFT
70          Exit Function
80      End If
90  End If

100 Set rng = NoteRange(ws)
110 If Not rng Is Nothing Then
120     Set ClickTargetArea = Application.Intersect(clicked, rng)
130     If Not ClickTargetArea Is Nothing Then
140         kind = TGT_NOTE
150         Exit Function
160     End If
170 End If

180 Set rng = DoctorBlock(ws)
190 If Not rng Is Nothing Then
200     Set ClickTargetArea = Application.Intersect(clicked, rng)
210     If Not ClickTargetArea Is Nothing Then kind = TGT_DOC
220 End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClickTargetArea", Err.Number, Err.Description, Erl, _
             "clicked=" & clicked.Address(False, False)
    Set ClickTargetArea = Nothing
    kind = TGT_NONE
End Function

'--- その記号をその場所に押してよいか ---
'    シフト入力欄 : シフト記号のみ(医師名・備考スタンプは不可)
'    備考行       : 備考スタンプ + 消去 / 色消
'    医師名欄     : 医師名スタンプ + 消去 / 色消
'    それぞれの場所に別の記号が入ると集計がずれるため、専用にする。
'    消去と色消はどこでも許可する(書き間違いを直せなくなるため)。
Private Function StampAllowedHere(ByVal idx As Long, ByVal kind As Long) As Boolean
    On Error GoTo ErrHandler

10  If idx = IDX_ERASE Or idx = IDX_CLEARFILL Then
20      StampAllowedHere = (kind <> TGT_NONE)
30      Exit Function
40  End If
        Select Case kind
            Case TGT_DOC
                StampAllowedHere = IsDoctorStamp(idx)
            Case TGT_NOTE
                StampAllowedHere = IsNoteStamp(idx)
            Case TGT_SHIFT
                StampAllowedHere = Not IsDoctorStamp(idx) And Not IsNoteStamp(idx)
            Case Else
                StampAllowedHere = False
        End Select
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "StampAllowedHere", Err.Number, Err.Description, Erl, _
             "idx=" & idx & "; kind=" & kind
    StampAllowedHere = False
End Function

'--- 押せない組み合わせのときの説明文 ---
Private Function StampDeniedText(ByVal idx As Long, ByVal kind As Long) As String
    On Error GoTo ErrHandler

        Select Case kind
            Case TGT_DOC
                StampDeniedText = "医師名欄には医師名スタンプしか押せません。"
            Case TGT_NOTE
                StampDeniedText = "備考行には備考スタンプしか押せません。"
            Case TGT_SHIFT
                If IsDoctorStamp(idx) Then
                    StampDeniedText = "医師名スタンプは医師名欄にだけ押せます。"
                Else
                    StampDeniedText = "備考スタンプは備考行にだけ押せます。"
                End If
            Case Else
                StampDeniedText = "ここには押せません。"
        End Select
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "StampDeniedText", Err.Number, Err.Description, Erl, _
             "idx=" & idx & "; kind=" & kind
    StampDeniedText = "ここには押せません。"
End Function

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

'--- 範囲へのスタンプ適用 ---
'    色消          … 背景色だけ消す(値は残す)
'    背景緑/橙/灰  … 背景色だけ塗る(値は書き換えない)
'    それ以外      … 値＋書式をコピー
Private Sub StampArea(ByVal ws As Worksheet, ByVal area As Range, ByVal idx As Long)
    Dim c As Range, pal As Range, src As Range
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ws)
20  If pal Is Nothing Then Exit Sub
30  If idx < 1 Or idx > pal.Cells.Count Then Exit Sub
40  Set src = pal.Cells(1, idx)

50  For Each c In area.Cells
60      If idx = IDX_CLEARFILL Then
70          c.Interior.Pattern = xlNone
80      ElseIf idx >= IDX_FILL_FIRST And idx <= IDX_FILL_LAST Then
90          ApplyFillOnly c, src
100     Else
110         ApplyStamp c, src
120     End If
130 Next c
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "StampArea", Err.Number, Err.Description, Erl, _
             "area=" & area.Address(False, False) & "; idx=" & idx
End Sub

'--- 背景色だけを塗る(値・文字色は触らない) ---
Private Sub ApplyFillOnly(ByVal c As Range, ByVal src As Range)
    On Error GoTo ErrHandler

10  If SKIP_FORMULA_CELLS Then
20      If c.HasFormula Then Exit Sub
30  End If
40  If c.Worksheet.ProtectContents And c.Locked Then Exit Sub

50  If src.Interior.Pattern <> xlNone Then
60      c.Interior.Color = src.Interior.Color
70  Else
        ' パレット側に色が無いときは塗りを外す
80      c.Interior.Pattern = xlNone
90  End If
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "ApplyFillOnly", Err.Number, Err.Description, Erl, _
             "cell=" & c.Address(False, False) & "; src=" & src.Address(False, False)
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

'--- パレットのセルの値をコピーする(色は持ち込まない) ---
'    パレット上の色分け(医師名スタンプの背景色・文字色など)は
'    「パレットのどのボタンか」を見分けるための飾りであって、
'    シフト表側の見た目ではない。貼り付け先の色はそのまま残す。
'    背景色を塗りたいときは背景色ボタン(ApplyFillOnly)を使う。
'    太字だけは記号の強調に使うので引き継ぐ。
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
100 c.Font.Bold = src.Font.Bold
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
    Dim pal As Range, hit As Long
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ws)
20  If pal Is Nothing Then Exit Sub
30  hit = PaletteIndexOf(ws, c)
40  If hit = 0 Then Exit Sub
    ' 同じモードを再度クリックしたら連続切替に戻す
50  If hit = CurrentIndex(ws) And hit <> IDX_CYCLE Then hit = IDX_CYCLE

60  Application.EnableEvents = False
70  pal.Offset(MARKER_OFFSET, 0).ClearContents
80  pal.Cells(1, hit).Offset(MARKER_OFFSET, 0).Value = MARKER_CHAR

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
30  If LCase$(STAMP_TRIGGER) = "single" Then op = "クリック" Else op = "ダブルクリック"

        Select Case idx
            Case IDX_OFF
                s = "シフト入力マクロ: OFF(通常のExcel操作)"

            Case IDX_AUTO
                s = "シフト入力マクロ: 自動作成　ダブルクリックで実行"

            Case IDX_UNDO
                s = "シフト入力マクロ: 元に戻す　ダブルクリックで直前の変更を取り消す"

            Case IDX_EXPORT
                s = "シフト入力マクロ: 印刷出力　ダブルクリックで PDF / Excel に書き出す"

            Case IDX_CYCLE
                If LCase$(CYCLE_TRIGGER) = "single" Then op = "クリック" Else op = "ダブルクリック"
                s = "シフト入力マクロ: 連続切替　" & op & "で次の記号／右クリックで前へ"

            Case IDX_CLEARFILL
                s = "シフト入力マクロ: 背景色クリア　" & op & "で消す"

            Case IDX_FILL_FIRST To IDX_FILL_LAST
                s = "シフト入力マクロ: 背景色ペイント [ " & _
                    Trim$(CStr(pal.Cells(1, idx).Offset(LABEL_OFFSET, 0).Value)) & _
                    " ]　" & op & "で塗る／範囲を選んで右クリックでまとめて塗る"

            Case Else
                s = "シフト入力マクロ: スタンプ [ " & _
                    Trim$(CStr(pal.Cells(1, idx).Value)) & " ]　" & _
                    op & "で押す／範囲を選んで右クリックでまとめて押す"
                If IsDoctorStamp(idx) Then
                    s = s & "　(医師名欄のみ)"
                ElseIf IsNoteStamp(idx) Then
                    s = s & "　(備考行のみ)"
                End If
        End Select
40  Application.StatusBar = s
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
    ' シートの端では Offset が失敗する(想定内)
    LogError MODULE_NAME, "MoveSel", Err.Number, Err.Description, Erl, _
             "area=" & area.Address(False, False) & "; dr=" & dr & "; dc=" & dc
End Sub

'==================================================================
' 手動実行マクロ
'==================================================================
Public Sub ShiftClick_選択範囲にスタンプ()
    Dim ws As Worksheet, area As Range
    Dim kind As Long, idx As Long, stamped As Long
    On Error GoTo ErrHandler

10  Set ws = ActiveSheet
20  If Not TypeOf Selection Is Range Then Exit Sub
80  Set area = ClickTargetArea(ws, Selection, kind)
90  If area Is Nothing Then
100     MsgBox "シフト入力欄・備考行・医師名欄のいずれかを" & vbCrLf & _
               "選んでから実行してください。", vbExclamation
110     Exit Sub
120 End If
130 idx = CurrentIndex(ws)
    '--- モードボタン(OFF/自動/切替/色消)はスタンプではない ---
140 If idx <= IDX_CLEARFILL Then
150     MsgBox "パレットでスタンプする記号または背景色を選んでから" & vbCrLf & _
               "実行してください。", vbExclamation
160     Exit Sub
170 End If
    '--- 記号の種類と書き込み先が噛み合わない組み合わせは拒否する ---
175 If Not StampAllowedHere(idx, kind) Then
176     MsgBox StampDeniedText(idx, kind), vbExclamation
177     Exit Sub
178 End If

180 Application.EnableEvents = False
185 stamped = area.Cells.Count
190 StampArea ws, area, idx

CleanUp:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
    ' area が未設定のままここに来ることがあるため、件数は先に控えておく
    LogSuccess MODULE_NAME, "ShiftClick_選択範囲にスタンプ", _
               "Stamped " & stamped & " cells with palette index " & idx
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
200 lastCell.Offset(MARKER_OFFSET, 0).Resize(PALETTE_ROWS, 1).Copy
210 newCell.Offset(MARKER_OFFSET, 0).PasteSpecial xlPasteFormats
220 Application.CutCopyMode = False
    '--- 値とラベルを書き込み ---
230 newCell.Value = sym
240 newCell.Offset(LABEL_OFFSET, 0).Value = lab
250 Application.EnableEvents = True

    '--- 名前付き範囲を追随させる ---
260 ShiftSetup_名前付き範囲更新

270 MsgBox "パターン [ " & sym & " ] を追加しました。" & vbCrLf & _
           "パレット範囲: " & PaletteRange(ws).Address(False, False) & vbCrLf & vbCrLf & _
           "※パレットのセルに付けた文字色・背景色は" & vbCrLf & _
           "　シフト表には移りません_値と太字だけが移ります_。" & vbCrLf & _
           "　シフト表を塗るときは背景色ボタンを使ってください。", vbInformation

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
' パレット再作成(入口だけ。実体は ShiftSetup)
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
    Dim warn As String, palRow As Long, missing As String
    Dim noteRng As Range, docBlk As Range
    On Error GoTo ErrHandler

10  Set ws = ActiveSheet
20  If NamedRangeOrNothing(NM_SHIFT) Is Nothing Then
30      srcS = "計算で解決"
40  Else
50      srcS = "名前付き範囲"
60  End If
70  If NamedRangeOrNothing(NM_PALETTE) Is Nothing Then
80      srcP = "計算で解決"
90  Else
100     srcP = "名前付き範囲"
110 End If

120 Set grid = ShiftInputRange(ws)
130 Set pal = PaletteRange(ws)
140 docRow = ShiftDocRow(ws)
150 palRow = PaletteBodyRow(ws)
155 Set noteRng = NoteRange(ws)
156 Set docBlk = DoctorBlock(ws)

    '--- 入力範囲の妥当性 ---
160 drift = ShiftRangeDrift(ws)
170 If grid Is Nothing Then
180     warn = vbCrLf & "■ シフト入力欄を特定できません" & vbCrLf & _
               "　B列の開始日の数式、またはA列の「" & LBL_NOTE & "」を確認してください" & vbCrLf
190 ElseIf drift > 0 Then
200     warn = vbCrLf & "■ 範囲の下端が " & LBL_DOC & "(" & docRow & "行)の" & _
               DOC_GAP & "行上を " & drift & " 行超えています" & vbCrLf & _
               "　正しい終端: " & (docRow - DOC_GAP) & "行" & vbCrLf
210 ElseIf drift < 0 Then
220     warn = vbCrLf & "※ 範囲の下端が " & (docRow - DOC_GAP) & "行 より " & _
               Abs(drift) & " 行上です(入力できる行が少なくなっています)" & vbCrLf
230 End If

    '--- 依存シートの有無 ---
240 If Not SheetExists(SHT_CFG) Then missing = missing & " " & SHT_CFG
250 If Not SheetExists(SHT_HOLIDAY) Then missing = missing & " " & SHT_HOLIDAY
260 If Not SheetExists(SHT_LOG) Then missing = missing & " " & SHT_LOG
270 If Len(missing) > 0 Then
280     warn = warn & vbCrLf & "■ 不足しているシート:" & missing & vbCrLf & _
               "　ShiftSchema_不足シート生成 で作成できます" & vbCrLf
290 End If

300 msg = "シート名          : " & ws.Name & vbCrLf & _
          "シフト入力範囲    : " & _
              IIf(grid Is Nothing, "(未特定)", grid.Address(False, False)) & _
              "　←" & srcS & vbCrLf & _
          "　上端=再掲日付行の1行下 / 下端=" & LBL_DOC & "の" & DOC_GAP & "行上" & vbCrLf & _
          "備考行(書込可)    : " & _
              IIf(noteRng Is Nothing, "(未特定)", noteRng.Address(False, False)) & vbCrLf & _
          "医師名欄(書込可)  : " & _
              IIf(docBlk Is Nothing, "(未特定)", docBlk.Address(False, False)) & vbCrLf & _
          "パレット範囲      : " & _
              IIf(pal Is Nothing, "(未特定)", pal.Address(False, False)) & _
              "　←" & srcP & vbCrLf & _
          "　パレット本体行  : " & palRow & " 行(日付行の" & PALETTE_GAP & "行上)" & vbCrLf & _
          "パレットのセル数  : " & IIf(pal Is Nothing, 0, pal.Cells.Count) & vbCrLf & _
          "切替サイクル項目数: " & UBound(CycleValues(ws)) + 1 & _
              "(シフト記号のみ)" & vbCrLf & _
          "現在のモード番号  : " & CurrentIndex(ws) & vbCrLf & warn & vbCrLf & _
          "ここまで表示されれば、標準モジュールは正しく入っています。"
310 MsgBox msg, vbInformation, "ShiftClick セルフチェック"

    LogSuccess MODULE_NAME, "ShiftClick_セルフチェック", _
               "grid=" & IIf(grid Is Nothing, "none", grid.Address(False, False)) & _
               ", drift=" & drift & ", missing=[" & Trim$(missing) & "]"
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftClick_セルフチェック", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; docRow=" & docRow & "; paletteRow=" & palRow
    MsgBox "セルフチェックでエラーが発生しました: " & Err.Description, vbExclamation
End Sub


'==================================================================
' 自動作成の事前診断 (v9.2: ShiftAutoBridge から移管)
'   ShiftAuto に入る前に前提を検査する。
'   RunAutoShift から呼ばれ、問題があれば実行を止める。
'==================================================================
Public Function AutoShiftPreflight(ByRef reason As String) As Boolean
    Dim ws As Worksheet, grid As Range, mc As Range
    Dim docRow As Long, drift As Long, missing As String
    On Error GoTo ErrHandler

10  reason = ""
20  Set ws = ShiftSheet()

    '--- 対象月 ---
30  Set mc = MonthCell(ws)
40  If mc Is Nothing Then
50      reason = reason & "■ 対象月のセルが見つかりません" & vbCrLf & _
                 "　" & HeaderRow(ws) & "行のA列に年月を入れてください" & vbCrLf
60  ElseIf Not IsDate(mc.Value) Then
70      reason = reason & "■ 対象月(" & mc.Address(False, False) & _
                 ")が日付ではありません" & vbCrLf
80  End If

    '--- 入力欄 ---
90  Set grid = ShiftInputRange(ws)
100 If grid Is Nothing Then
110     reason = reason & "■ シフト入力欄を特定できません" & vbCrLf
120 Else
130     drift = ShiftRangeDrift(ws)
140     docRow = ShiftDocRow(ws)
150     If drift > 0 Then
160         reason = reason & "■ 入力欄の下端が " & LBL_DOC & "(" & docRow & _
                     "行)の" & DOC_GAP & "行上を " & drift & " 行超えています" & vbCrLf & _
                     "　正しい終端: " & (docRow - DOC_GAP) & "行" & vbCrLf
170     End If
180 End If

    '--- 依存シート ---
190 If Not SheetExists(SHT_CFG) Then missing = missing & " " & SHT_CFG
200 If Not SheetExists(SHT_HOLIDAY) Then missing = missing & " " & SHT_HOLIDAY
210 If Not SheetExists(SHT_LOG) Then missing = missing & " " & SHT_LOG
220 If Len(missing) > 0 Then
230     reason = reason & "■ 不足しているシート:" & missing & vbCrLf & _
                 "　ShiftSchema_不足シート生成 で作成できます" & vbCrLf
240 End If

250 AutoShiftPreflight = (Len(reason) = 0)
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "AutoShiftPreflight", Err.Number, Err.Description, Erl, _
             "docRow=" & docRow & "; drift=" & drift
    reason = reason & "■ 診断中にエラー: " & Err.Description & vbCrLf
    AutoShiftPreflight = False
End Function

'--- 診断結果をダイアログで見せる(単体実行用) ---
Public Sub ShiftAuto_事前診断()
    Dim reason As String, ok As Boolean, ws As Worksheet, mc As Range
    Dim grid As Range
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  ok = AutoShiftPreflight(reason)
30  Set mc = MonthCell(ws)
40  Set grid = ShiftInputRange(ws)

50  If ok Then
60      MsgBox "自動作成の前提は満たされています。" & vbCrLf & vbCrLf & _
               "対象月    : " & IIf(mc Is Nothing, "(不明)", _
                   mc.Address(False, False) & " = " & Format(mc.Value, "yyyy年m月")) & vbCrLf & _
               "入力範囲  : " & IIf(grid Is Nothing, "(未特定)", _
                   grid.Address(False, False)) & vbCrLf & _
               "日付行    : " & DateRow(ws) & " 行" & vbCrLf & _
               "医師数行  : " & ShiftDocRow(ws) & " 行" & vbCrLf & _
               "起動マクロ: " & AUTOSHIFT_MACRO, _
               vbInformation, "自動作成 事前診断"
70  Else
80      MsgBox "自動作成を実行する前に、次を解消してください。" & vbCrLf & vbCrLf & reason, _
               vbExclamation, "自動作成 事前診断"
90  End If

    LogSuccess MODULE_NAME, "ShiftAuto_事前診断", "ok=" & ok
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftAuto_事前診断", Err.Number, Err.Description, Erl, ""
    MsgBox "事前診断でエラーが発生しました: " & Err.Description, vbExclamation
End Sub
