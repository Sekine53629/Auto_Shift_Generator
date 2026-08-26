Option Explicit
'==================================================================
'  シフト表 共通モジュール ＜標準モジュール ShiftCommon v1.0＞
'  2026-08-26
'
'  シート名・A列ラベル・列・行オフセット・範囲解決をここに一元化する。
'  ShiftAuto / ShiftClick / ShiftSetup は範囲を自前で持たず、
'  必ず本モジュールの関数を経由すること。
'  (v6-v8 で各モジュールが別々のフォールバック範囲を持ち、
'   仕様がずれ続けたため v1.0 で集約した)
'
'  範囲の仕様:
'    シフト入力欄  上端 = 日付行の1行下
'                  下端 = 医師数行の DOC_GAP 行上 (間の1行は緩衝行)
'    パレット      本体行 = 日付行の PALETTE_GAP 行上
'                  本体行 + MARKER_OFFSET = ★マーカー行
'                  本体行 + LABEL_OFFSET  = ラベル行
'
'  解決順序: 名前付き範囲 → A列ラベルからの計算 → Nothing
'  (ハードコードした番地は持たない。Nothing は呼び出し側で通知する)
'==================================================================
Private Const MODULE_NAME As String = "ShiftCommon"

'--- シート名 ---
Public Const SHT_SHIFT   As String = "シフト"
Public Const SHT_CFG     As String = "自動作成設定"
Public Const SHT_HOLIDAY As String = "祝日マスタ"
Public Const SHT_LOG     As String = "シフト変更ログ"

'--- A列の基準ラベル(前方一致で検索) ---
Public Const LBL_NAME    As String = "氏名"
Public Const LBL_WEEK    As String = "曜日"
Public Const LBL_DOC     As String = "医師数"
Public Const LBL_PHARM   As String = "薬剤師出勤数"
Public Const LBL_CLERK   As String = "事務員出勤数"
Public Const LBL_SHORT   As String = "過不足"
Public Const LBL_PALETTE As String = "シフトパレット"

'--- 名前付き範囲の名前 ---
Public Const NM_SHIFT   As String = "シフトパレット範囲"
Public Const NM_PALETTE As String = "シフトパレット"

'--- 日付/シフトの列範囲 ---
Public Const COL_FIRST As String = "B"
Public Const COL_LAST  As String = "AF"

'--- 行オフセット ---
Public Const DOC_GAP       As Long = 2    ' 入力欄の下端 = 医師数行の n 行上
Public Const PALETTE_GAP   As Long = 3    ' パレット本体行 = 日付行の n 行上
Public Const MARKER_OFFSET As Long = -1   ' ★マーカー行(本体行からの相対)
Public Const LABEL_OFFSET  As Long = 1    ' ラベル行(本体行からの相対)

'--- パレットの固定位置(左からの番号) ---
Public Const IDX_OFF       As Long = 1    ' OFF(マクロ停止)
Public Const IDX_CYCLE     As Long = 2    ' 連続切替
Public Const IDX_CLEARFILL As Long = 3    ' 背景色クリア
Public Const IDX_ERASE     As Long = 4    ' 消去(空白スタンプ)

'==================================================================
' 基本ヘルパー
'==================================================================
'--- シフトシートを取得(無ければ ActiveSheet) ---
Public Function ShiftSheet() As Worksheet
    On Error GoTo ErrHandler

10  Set ShiftSheet = ThisWorkbook.Worksheets(SHT_SHIFT)
    Exit Function
ErrHandler:
    ' シート名が変更されている場合のみここに来る(想定内)
20  Set ShiftSheet = ActiveSheet
End Function

'--- 名前付き範囲を取得(未定義/参照切れなら Nothing) ---
Public Function NamedRangeOrNothing(ByVal nm As String) As Range
    On Error GoTo ErrHandler

10  Set NamedRangeOrNothing = ThisWorkbook.Names(nm).RefersToRange
    Exit Function
ErrHandler:
    ' 名前が無い/参照切れは正常系として扱い、呼び出し側でラベル計算に落とす
20  Set NamedRangeOrNothing = Nothing
End Function

'--- A列を前方一致で検索して行番号を返す(0 = 未検出) ---
Public Function LabelRow(ByVal ws As Worksheet, ByVal lbl As String) As Long
    Dim f As Range
    On Error GoTo ErrHandler

10  Set f = ws.Columns(1).Find(What:=lbl, LookIn:=xlValues, _
                              LookAt:=xlPart, MatchCase:=False)
20  If Not f Is Nothing Then LabelRow = f.Row
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "LabelRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; label=" & lbl
    LabelRow = 0
End Function

'--- 列番号 → 列文字 ---
Public Function ColLetter(ByVal colNo As Long) As String
    Dim s As String, n As Long
    On Error GoTo ErrHandler

10  n = colNo
20  Do While n > 0
30      s = Chr$(65 + ((n - 1) Mod 26)) & s
40      n = (n - 1) \ 26
50  Loop
60  ColLetter = s
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ColLetter", Err.Number, Err.Description, Erl, _
             "colNo=" & colNo
    ColLetter = ""
End Function

'==================================================================
' 行の解決(すべて A列ラベル基準)
'==================================================================
'--- 日付行 = 氏名行の1行上 ---
Public Function DateRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    On Error GoTo ErrHandler

10  r = LabelRow(ws, LBL_NAME)
20  If r > 1 Then DateRow = r - 1
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "DateRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name
    DateRow = 0
End Function

'--- 入力欄の上端 = 曜日行より下・医師数行より上で A列に値がある最初の行 ---
Public Function ShiftTopRow(ByVal ws As Worksheet) As Long
    Dim wkRow As Long, docRow As Long, r As Long
    On Error GoTo ErrHandler

10  wkRow = LabelRow(ws, LBL_WEEK)
20  docRow = LabelRow(ws, LBL_DOC)
30  If wkRow = 0 Or docRow = 0 Then Exit Function
40  For r = wkRow + 1 To docRow - 1
50      If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then
60          ShiftTopRow = r
70          Exit Function
80      End If
90  Next r
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ShiftTopRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; weekRow=" & wkRow & "; docRow=" & docRow
    ShiftTopRow = 0
End Function

'--- 入力欄の下端 = 医師数行の DOC_GAP 行上 ---
Public Function ShiftBottomRow(ByVal ws As Worksheet) As Long
    Dim docRow As Long
    On Error GoTo ErrHandler

10  docRow = LabelRow(ws, LBL_DOC)
20  If docRow > DOC_GAP Then ShiftBottomRow = docRow - DOC_GAP
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ShiftBottomRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; docRow=" & docRow
    ShiftBottomRow = 0
End Function

'--- パレット本体行 = 日付行の PALETTE_GAP 行上 ---
Public Function PaletteBodyRow(ByVal ws As Worksheet) As Long
    Dim dRow As Long
    On Error GoTo ErrHandler

10  dRow = DateRow(ws)
20  If dRow > PALETTE_GAP Then PaletteBodyRow = dRow - PALETTE_GAP
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "PaletteBodyRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; dateRow=" & dRow
    PaletteBodyRow = 0
End Function

'==================================================================
' 範囲の解決(名前付き範囲 → ラベル計算 → Nothing)
'==================================================================
'--- シフト入力欄 ---
Public Function ShiftInputRange(ByVal ws As Worksheet) As Range
    Dim topR As Long, botR As Long
    On Error GoTo ErrHandler

10  Set ShiftInputRange = NamedRangeOrNothing(NM_SHIFT)
20  If Not ShiftInputRange Is Nothing Then Exit Function

30  topR = ShiftTopRow(ws)
40  botR = ShiftBottomRow(ws)
50  If topR = 0 Or botR = 0 Or topR > botR Then Exit Function
60  Set ShiftInputRange = ws.Range(COL_FIRST & topR & ":" & COL_LAST & botR)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ShiftInputRange", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; topRow=" & topR & "; bottomRow=" & botR
    Set ShiftInputRange = Nothing
End Function

'--- パレット本体(横1行) ---
Public Function PaletteRange(ByVal ws As Worksheet) As Range
    Dim palRow As Long, lastCol As Long, firstCol As Long
    On Error GoTo ErrHandler

10  Set PaletteRange = NamedRangeOrNothing(NM_PALETTE)
20  If Not PaletteRange Is Nothing Then Exit Function

30  palRow = PaletteBodyRow(ws)
40  If palRow = 0 Then Exit Function
50  firstCol = ws.Range(COL_FIRST & "1").Column
    ' ラベル行の最終セルまでを本体の幅とみなす
60  lastCol = ws.Cells(palRow + LABEL_OFFSET, ws.Columns.Count).End(xlToLeft).Column
70  If lastCol < firstCol Then lastCol = firstCol
80  Set PaletteRange = ws.Range(ws.Cells(palRow, firstCol), ws.Cells(palRow, lastCol))
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "PaletteRange", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; paletteRow=" & palRow
    Set PaletteRange = Nothing
End Function

'--- 入力欄の下端のずれ(0=正常 / 正=下すぎ / 負=上すぎ) ---
Public Function ShiftRangeDrift(ByVal ws As Worksheet) As Long
    Dim rng As Range, docRow As Long, endRow As Long
    On Error GoTo ErrHandler

10  Set rng = ShiftInputRange(ws)
20  docRow = LabelRow(ws, LBL_DOC)
30  If rng Is Nothing Or docRow = 0 Then Exit Function
40  endRow = rng.Row + rng.Rows.Count - 1
50  ShiftRangeDrift = endRow - (docRow - DOC_GAP)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ShiftRangeDrift", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; docRow=" & docRow & "; endRow=" & endRow
    ShiftRangeDrift = 0
End Function
