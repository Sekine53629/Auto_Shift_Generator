Option Explicit
'==================================================================
'  シフト表 初期設定マクロ ＜標準モジュール ShiftSetup v1.1＞
'  2026-08-26
'  ShiftCommon / ShiftClick と併用。
'  シート上の位置はすべて ShiftCommon のラベル検索で動的に決定する。
'
'  実行するのは基本これ1本:
'      ShiftSetup_初期設定実行
'  内訳:
'      1) パレットの生成(日付行の PALETTE_GAP 行上)
'      2) 年月セル・祝日サマリーの数式セット
'      3) AH:AL 集計列(見出し＋数式)のセット
'      4) 名前付き範囲の再定義(シフトパレット範囲 / シフトパレット)
'
'  v1.1 変更:
'   ・シート名/ラベル/列/行オフセット/範囲解決を ShiftCommon に移管
'   ・全プロシージャに ErrorLogger のエラーハンドラを追加
'   ・色・フォントサイズ・記号などのリテラルを定数化
'   ・医師名スタンプを既定値(医師1..医師9)に変更
'==================================================================
Private Const MODULE_NAME As String = "ShiftSetup"

'--------------------------- 設定 ---------------------------------
' 集計列(このモジュール専用)
Private Const COL_AGG       As String = "AH"  ' 集計列の先頭
Private Const COL_AGG_END   As String = "AL"  ' 集計列の最終
Private Const COL_MONTH     As String = "AG"  ' 年月シリアルの置き場
Private Const COL_HOL_SUMM  As String = "G"   ' 祝日サマリーの置き場

' パレット内の位置(左からの番号)
Private Const IDX_SYM_FIRST As Long = 5       ' ○ の位置
Private Const IDX_SYM_LAST  As Long = 7       ' ▲ の位置
Private Const IDX_DOC_START As Long = 14      ' 医師名スタンプの開始位置

' 医師数の「忙しい日」判定に使う人数(集計列 5診出勤)
Private Const DOC_BUSY_COUNT As Long = 5

' 書式
Private Const CLR_BLACK     As Long = 0
Private Const PAL_ROW_HEIGHT As Double = 20

' フォントサイズ
Private Const FS_LABEL  As Long = 9
Private Const FS_BODY   As Long = 10
Private Const FS_SYMBOL As Long = 12
Private Const FS_TITLE  As Long = 14

'--- パレット定義(ShiftClick の並びと必ず一致させること) ---
'    1=OFF 2=切替 3=色消 4=消去(空白) 5以降=スタンプ記号
'    1..13 = シフト記号 / 14..22 = 医師名スタンプ
'  ※医師名は個人情報のため既定値をダミーにしてある。
'    運用時は下の "医師1".."医師9" を実際の医師名に書き換えて使うこと。
'    書き換えた内容はリポジトリにコミットしないこと。
Private Function SS_PalVals() As Variant
    On Error GoTo ErrHandler
    SS_PalVals = Array("OFF", "切替", "色消", "", "○", "●", "▲", _
                       "公休", "希休", "夏休", "有休", "有休※", "銀行", _
                       "医師1", "医師2", "医師3", "医師4", "医師5", _
                       "医師6", "医師7", "医師8", "医師9")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_PalVals", Err.Number, Err.Description, Erl, ""
End Function

Private Function SS_PalLabs() As Variant
    On Error GoTo ErrHandler
    SS_PalLabs = Array("停止", "順送り", "背景消", "消去", "早番", "遅半", "遅番", _
                       "公休", "希休", "夏休", "有休", "備考付", "銀行", _
                       "医師", "医師", "医師", "医師", "医師", _
                       "医師", "医師", "医師", "医師")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_PalLabs", Err.Number, Err.Description, Erl, ""
End Function

'--- 集計列の見出し ---
Private Function SS_AggHeads() As Variant
    On Error GoTo ErrHandler
    SS_AggHeads = Array("休", "○早番", "▲遅番", "●遅半", _
                        DOC_BUSY_COUNT & "診出勤")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_AggHeads", Err.Number, Err.Description, Erl, ""
End Function

'--- 集計列「休」に数える記号 ---
Private Function SS_OffSyms() As Variant
    On Error GoTo ErrHandler
    SS_OffSyms = Array("公休", "希休", "夏休", "有休", "有休※")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_OffSyms", Err.Number, Err.Description, Erl, ""
End Function

'--- 集計列「出勤」に数える記号 ---
Private Function SS_WorkSyms() As Variant
    On Error GoTo ErrHandler
    SS_WorkSyms = Array("○", "◯", "▲", "●")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_WorkSyms", Err.Number, Err.Description, Erl, ""
End Function

'==================== 色(定数として保持) =========================
Private Function ClrBorder() As Long
    On Error GoTo ErrHandler
    ClrBorder = RGB(150, 150, 150)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrBorder", Err.Number, Err.Description, Erl, ""
End Function
Private Function ClrMarker() As Long
    On Error GoTo ErrHandler
    ClrMarker = RGB(192, 0, 0)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrMarker", Err.Number, Err.Description, Erl, ""
End Function
Private Function ClrModeBg() As Long
    On Error GoTo ErrHandler
    ClrModeBg = RGB(242, 242, 242)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrModeBg", Err.Number, Err.Description, Erl, ""
End Function
Private Function ClrModeFg() As Long
    On Error GoTo ErrHandler
    ClrModeFg = RGB(80, 80, 80)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrModeFg", Err.Number, Err.Description, Erl, ""
End Function
Private Function ClrLabelFg() As Long
    On Error GoTo ErrHandler
    ClrLabelFg = RGB(100, 100, 100)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrLabelFg", Err.Number, Err.Description, Erl, ""
End Function
Private Function ClrDocBg() As Long
    On Error GoTo ErrHandler
    ClrDocBg = RGB(226, 239, 218)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrDocBg", Err.Number, Err.Description, Erl, ""
End Function
Private Function ClrDocFg() As Long
    On Error GoTo ErrHandler
    ClrDocFg = RGB(0, 61, 0)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrDocFg", Err.Number, Err.Description, Erl, ""
End Function
Private Function ClrHidden() As Long
    On Error GoTo ErrHandler
    ClrHidden = RGB(255, 255, 255)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrHidden", Err.Number, Err.Description, Erl, ""
End Function

'====================== 1) パレット生成 ==========================
Public Sub ShiftSetup_パレット生成()
    Dim ws As Worksheet, dateRowNo As Long, palRow As Long
    Dim pal As Range, vals As Variant, labs As Variant, i As Long
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  dateRowNo = DateRow(ws)
30  If dateRowNo = 0 Then
40      MsgBox "A列に「" & LBL_NAME & "」が見つかりません。", vbExclamation
50      Exit Sub
60  End If

70  palRow = PaletteBodyRow(ws)
80  If palRow < 2 Then
90      MsgBox "パレットを置く行が足りません。表の上に " & _
               (PALETTE_GAP + 1) & " 行以上必要です。", vbExclamation
100     Exit Sub
110 End If

120 vals = SS_PalVals()
130 labs = SS_PalLabs()
140 Set pal = ws.Cells(palRow, ws.Range(COL_FIRST & "1").Column) _
                .Resize(1, UBound(vals) + 1)

150 Application.EnableEvents = False
160 Application.ScreenUpdating = False

    '--- 本体行 ---
170 With pal
        .ClearContents
        .Interior.Pattern = xlNone
        .Font.Color = CLR_BLACK
        .Font.Bold = False
        .Font.Size = FS_BODY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = ClrBorder()
        .RowHeight = PAL_ROW_HEIGHT
180 End With

    '--- ★マーカー行(上) ---
190 With pal.Offset(MARKER_OFFSET, 0)
        .ClearContents
        .HorizontalAlignment = xlCenter
        .Font.Color = ClrMarker()
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = ClrBorder()
200 End With

    '--- ラベル行(下) ---
210 With pal.Offset(LABEL_OFFSET, 0)
        .ClearContents
        .HorizontalAlignment = xlCenter
        .Font.Size = FS_LABEL
        .Font.Color = ClrLabelFg()
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = ClrBorder()
220 End With

    '--- 値とラベル ---
230 For i = 1 To pal.Cells.Count
240     If Len(CStr(vals(i - 1))) > 0 Then pal.Cells(1, i).Value = vals(i - 1)
250     pal.Cells(1, i).Offset(LABEL_OFFSET, 0).Value = labs(i - 1)
260 Next i

    '--- モードセル(OFF/切替/色消)はグレー ---
270 For i = IDX_OFF To IDX_CLEARFILL
280     With pal.Cells(1, i)
            .Interior.Color = ClrModeBg()
            .Font.Color = ClrModeFg()
290     End With
300 Next i

    '--- ○●▲ を少し大きく ---
310 For i = IDX_SYM_FIRST To IDX_SYM_LAST
320     pal.Cells(1, i).Font.Size = FS_SYMBOL
330 Next i

    '--- 医師名スタンプは色分けして区別 ---
340 For i = IDX_DOC_START To pal.Cells.Count
350     With pal.Cells(1, i)
            .Interior.Color = ClrDocBg()
            .Font.Color = ClrDocFg()
            .Font.Size = FS_BODY
360     End With
370 Next i

    '--- 見出し(マーカー行の左) ---
380 With pal.Cells(1, 1).Offset(MARKER_OFFSET, -1)
        .Value = LBL_PALETTE
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = ClrBorder()
390 End With

    '--- 初期モード = 切替 ---
400 pal.Cells(1, IDX_CYCLE).Offset(MARKER_OFFSET, 0).Value = "★"

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSetup_パレット生成", _
               "Built palette row " & palRow & " with " & (UBound(SS_PalVals()) + 1) & " cells"
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSetup_パレット生成", Err.Number, Err.Description, Erl, _
             "dateRow=" & dateRowNo & "; paletteRow=" & palRow & "; i=" & i
    MsgBox "パレット生成でエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'============ 2) 年月セル・祝日サマリーの数式セット ==============
'   A(日付行-1) = =DATE(1900,AG<行>,1)   年月表示
'   G(日付行-1) = 土日公休/祝日カウントの LET 数式
Public Sub ShiftSetup_ヘッダ数式()
    Dim ws As Worksheet, dateRowNo As Long, hRow As Long
    Dim f As String
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  dateRowNo = DateRow(ws)
30  If dateRowNo = 0 Then
40      MsgBox "A列に「" & LBL_NAME & "」が見つかりません。", vbExclamation
50      Exit Sub
60  End If
70  hRow = dateRowNo - 1          ' タイトル/年月行

80  Application.EnableEvents = False

    '--- 年月セル(A列) ---
90  With ws.Cells(hRow, 1)
        .Formula = "=DATE(1900," & COL_MONTH & hRow & ",1)"
        .NumberFormatLocal = "[$-ja-JP]ge""."" m""月"""
        .Font.Bold = True
        .Font.Size = FS_TITLE
        .HorizontalAlignment = xlRight
100 End With

    '--- 年月シリアル(AG列・白文字で隠す) ---
110 With ws.Cells(hRow, ws.Range(COL_MONTH & "1").Column)
120     If Len(Trim$(CStr(.Value))) = 0 Then
130         .Value = (Year(Date) - 1900) * 12 + Month(Date)
140     End If
        .Font.Color = ClrHidden()
150 End With

    '--- 祝日サマリー(G列) ---
160 f = "=LET(d," & COL_FIRST & dateRowNo & ":" & COL_LAST & dateRowNo & _
        ",inM,--(MONTH(d)=MONTH(A" & hRow & "))" & _
        ",wk,SUMPRODUCT(inM*(WEEKDAY(d,2)>5))" & _
        ",hol,SUMPRODUCT(inM*(WEEKDAY(d,2)<6)*COUNTIF(" & SHT_HOLIDAY & "!$A:$A,d))" & _
        ",""土日公休""&wk&""回　祝日""&hol&""回　休みのトータル""&(wk+hol)&""回"")"
170 With ws.Cells(hRow, ws.Range(COL_HOL_SUMM & "1").Column)
        .Formula2 = f
        .Font.Italic = True
        .Font.Size = FS_LABEL
        .HorizontalAlignment = xlLeft
180 End With

CleanUp:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSetup_ヘッダ数式", _
               "Set month cell and holiday summary on row " & hRow
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSetup_ヘッダ数式", Err.Number, Err.Description, Erl, _
             "dateRow=" & dateRowNo & "; headerRow=" & hRow
    MsgBox "ヘッダ数式のセットでエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'============ 3) AH:AL 集計列(見出し＋数式)セット ================
Public Sub ShiftSetup_集計列数式()
    Dim ws As Worksheet, topR As Long, botR As Long, hdrRow As Long
    Dim docRow As Long, aggCol As Long, r As Long, i As Long
    Dim heads As Variant, rng As Range
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  topR = ShiftTopRow(ws)
30  botR = ShiftBottomRow(ws)
40  docRow = LabelRow(ws, LBL_DOC)
50  If topR = 0 Or botR = 0 Or docRow = 0 Then
60      MsgBox "A列の基準ラベルが見つかりません。", vbExclamation
70      Exit Sub
80  End If

90  hdrRow = topR - 1
100 aggCol = ws.Range(COL_AGG & "1").Column
110 heads = SS_AggHeads()

120 Application.EnableEvents = False
130 Application.ScreenUpdating = False

    '--- 見出し行 ---
140 For i = 0 To UBound(heads)
150     With ws.Cells(hdrRow, aggCol + i)
            .Value = heads(i)
            .Font.Bold = True
            .Font.Size = FS_LABEL
            .HorizontalAlignment = xlCenter
160     End With
170 Next i

    '--- 数式行(A列にスタッフ名がある行のみ) ---
180 For r = topR To botR
190     If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then
200         ws.Range(ws.Cells(r, aggCol), ws.Cells(r, aggCol + UBound(heads))).ClearContents
210     Else
220         ws.Cells(r, aggCol + 0).Formula = SS_CountFormula(r, SS_OffSyms())
230         ws.Cells(r, aggCol + 1).Formula = SS_CountFormula(r, Array("○", "◯"))
240         ws.Cells(r, aggCol + 2).Formula = SS_CountFormula(r, Array("▲"))
250         ws.Cells(r, aggCol + 3).Formula = SS_CountFormula(r, Array("●"))
260         ws.Cells(r, aggCol + 4).Formula = SS_BusyDayFormula(r, docRow)
270     End If
280 Next r

    '--- 書式 ---
290 Set rng = ws.Range(ws.Cells(topR, aggCol), _
                       ws.Cells(botR, aggCol + UBound(heads)))
300 With rng
        .Font.Bold = True
        .Font.Size = FS_BODY
        .HorizontalAlignment = xlCenter
310 End With

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSetup_集計列数式", _
               "Set aggregate formulas for rows " & topR & "-" & botR
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSetup_集計列数式", Err.Number, Err.Description, Erl, _
             "topRow=" & topR & "; bottomRow=" & botR & "; docRow=" & docRow & "; r=" & r
    MsgBox "集計列のセットでエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'--- 指定行の記号カウント数式を組み立てる ---
Private Function SS_CountFormula(ByVal r As Long, ByVal syms As Variant) As String
    Dim rngRef As String, s As String, i As Long
    On Error GoTo ErrHandler

10  rngRef = COL_FIRST & r & ":" & COL_LAST & r
20  For i = LBound(syms) To UBound(syms)
30      If Len(s) > 0 Then s = s & "+"
40      s = s & "COUNTIF(" & rngRef & ",""" & syms(i) & """)"
50  Next i
60  SS_CountFormula = "=" & s
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_CountFormula", Err.Number, Err.Description, Erl, _
             "r=" & r & "; i=" & i
    SS_CountFormula = ""
End Function

'--- 医師が DOC_BUSY_COUNT 名の日の出勤数を数える数式を組み立てる ---
Private Function SS_BusyDayFormula(ByVal r As Long, ByVal docRow As Long) As String
    Dim rngRef As String, docRef As String, syms As Variant
    Dim s As String, i As Long
    On Error GoTo ErrHandler

10  rngRef = COL_FIRST & r & ":" & COL_LAST & r
20  docRef = COL_FIRST & "$" & docRow & ":" & COL_LAST & "$" & docRow
30  syms = SS_WorkSyms()
40  For i = LBound(syms) To UBound(syms)
50      If Len(s) > 0 Then s = s & "+"
60      s = s & "(" & rngRef & "=""" & syms(i) & """)"
70  Next i
80  SS_BusyDayFormula = "=SUMPRODUCT((" & docRef & "=" & DOC_BUSY_COUNT & ")*(" & s & "))"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_BusyDayFormula", Err.Number, Err.Description, Erl, _
             "r=" & r & "; docRow=" & docRow & "; i=" & i
    SS_BusyDayFormula = ""
End Function

'====================== 4) 名前付き範囲 ==========================
Public Sub ShiftSetup_名前付き範囲更新()
    Dim ws As Worksheet, topR As Long, botR As Long
    Dim palRow As Long, nPal As Long, addr As String, lastCol As String
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  topR = ShiftTopRow(ws)
30  botR = ShiftBottomRow(ws)
40  palRow = PaletteBodyRow(ws)
50  If topR = 0 Or botR = 0 Or palRow = 0 Then
60      MsgBox "基準ラベル(" & LBL_NAME & " / " & LBL_WEEK & " / " & _
               LBL_DOC & ")が見つかりません。", vbExclamation
70      Exit Sub
80  End If

    '--- シフト入力範囲 ---
90  addr = "=" & ws.Name & "!$" & COL_FIRST & "$" & topR & _
           ":$" & COL_LAST & "$" & botR
100 SS_ReplaceName NM_SHIFT, addr

    '--- パレット本体 ---
110 nPal = UBound(SS_PalVals()) + 1
120 lastCol = ColLetter(ws.Range(COL_FIRST & "1").Column + nPal - 1)
130 addr = "=" & ws.Name & "!$" & COL_FIRST & "$" & palRow & _
           ":$" & lastCol & "$" & palRow
140 SS_ReplaceName NM_PALETTE, addr

    LogSuccess MODULE_NAME, "ShiftSetup_名前付き範囲更新", _
               "Redefined named ranges: shift=" & COL_FIRST & topR & ":" & COL_LAST & botR & _
               ", palette row=" & palRow
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSetup_名前付き範囲更新", Err.Number, Err.Description, Erl, _
             "topRow=" & topR & "; bottomRow=" & botR & "; paletteRow=" & palRow & _
             "; addr=" & addr
    MsgBox "名前付き範囲の更新でエラーが発生しました: " & Err.Description, vbExclamation
End Sub

'--- 名前付き範囲を貼り替える(既存があれば削除してから追加) ---
Private Sub SS_ReplaceName(ByVal nm As String, ByVal refersTo As String)
    On Error GoTo ErrHandler

    ' 既存の名前は無い場合もあるため、削除失敗は正常系として扱う
10  On Error Resume Next
20  ThisWorkbook.Names(nm).Delete
30  On Error GoTo ErrHandler

40  ThisWorkbook.Names.Add Name:=nm, RefersTo:=refersTo
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SS_ReplaceName", Err.Number, Err.Description, Erl, _
             "name=" & nm & "; refersTo=" & refersTo
End Sub

'====================== 一括実行 =================================
Public Sub ShiftSetup_初期設定実行()
    Dim ws As Worksheet, msg As String
    Dim dateRowNo As Long, topR As Long, botR As Long
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  dateRowNo = DateRow(ws)
30  topR = ShiftTopRow(ws)
40  botR = ShiftBottomRow(ws)
50  If dateRowNo = 0 Or topR = 0 Or botR = 0 Then
60      MsgBox "A列の基準ラベルが見つからないため中止しました。" & vbCrLf & _
               "必要なラベル: " & LBL_NAME & " / " & LBL_WEEK & " / " & LBL_DOC, _
               vbExclamation, "初期設定"
70      Exit Sub
80  End If

90  Application.Calculation = xlCalculationManual

100 ShiftSetup_パレット生成
110 ShiftSetup_ヘッダ数式
120 ShiftSetup_集計列数式
130 ShiftSetup_名前付き範囲更新

140 Application.Calculation = xlCalculationAutomatic
150 Application.CalculateFull

160 msg = "初期設定が完了しました。" & vbCrLf & vbCrLf & _
          "日付行          : " & dateRowNo & " 行" & vbCrLf & _
          "シフト入力範囲  : " & COL_FIRST & topR & ":" & COL_LAST & botR & vbCrLf & _
          "パレット本体行  : " & PaletteBodyRow(ws) & " 行" & vbCrLf & _
          "集計列          : " & COL_AGG & "-" & COL_AGG_END
170 MsgBox msg, vbInformation, "初期設定"

CleanUp:
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSetup_初期設定実行", _
               "Completed setup: dateRow=" & dateRowNo & ", input=" & _
               COL_FIRST & topR & ":" & COL_LAST & botR
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSetup_初期設定実行", Err.Number, Err.Description, Erl, _
             "dateRow=" & dateRowNo & "; topRow=" & topR & "; bottomRow=" & botR
    MsgBox "初期設定でエラーが発生しました: " & Err.Description, vbExclamation, "初期設定"
    Resume CleanUp
End Sub
