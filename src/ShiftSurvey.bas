Option Explicit
'==================================================================
'  シフト表 構造調査マクロ ＜標準モジュール ShiftSurvey v1.0＞
'  2026-08-26
'
'  目的:
'    ShiftSetup を「白紙のブックから生成できる」ようにするための現状把握。
'    実物のシートが実際にどういう行構成・数式・書式になっているかを
'    レポートシートに書き出す。
'
'  重要: 本モジュールは読み取り専用。
'        レポートシート以外は一切書き換えない。
'
'  個人情報:
'    MASK_NAMES = True のとき、既知の見出し以外のA列の文字列は
'    「(氏名1)」等に伏せて出力する。B:AF のセルの値は出力しない
'    (医師名・スタッフ名が出ないようにするため)。
'    レポートをそのまま共有できるようにするための既定値。
'==================================================================
Private Const MODULE_NAME As String = "ShiftSurvey"

'--- 出力先 ---
Private Const RPT_SHEET As String = "シート調査結果"

'--- 氏名・医師名を伏せる ---
Private Const MASK_NAMES As Boolean = True

'--- 走査範囲 ---
Private Const SCAN_ROWS As Long = 80    ' シフトシートを上から何行見るか
Private Const SCAN_COLS As Long = 40    ' 値・数式の有無を数える列数(A から)

'--- レポートの列 ---
Private Const RPT_COLS As Long = 11

'==================================================================
' 入口
'==================================================================
Public Sub ShiftSurvey_シート構造調査()
    Dim ws As Worksheet, rpt As Worksheet, r As Long
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  Set rpt = SV_PrepareReport()
30  r = 1

40  r = SV_WriteBookInfo(rpt, r)
50  r = SV_WriteNames(rpt, r)
60  r = SV_WriteDetected(rpt, ws, r)
70  r = SV_WriteRowDump(rpt, ws, r)
80  r = SV_WriteColWidths(rpt, ws, r)
90  r = SV_WriteConfigSheet(rpt, r)

100 rpt.Columns(1).ColumnWidth = 6
110 rpt.Columns(2).ColumnWidth = 22
120 rpt.Columns(3).Resize(1, RPT_COLS - 2).ColumnWidth = 16
130 rpt.Activate
140 rpt.Range("A1").Select

150 MsgBox "シートの構造を「" & RPT_SHEET & "」に書き出しました。" & vbCrLf & vbCrLf & _
           "氏名は " & IIf(MASK_NAMES, "伏せて", "そのまま") & "出力しています。" & vbCrLf & _
           "このシートの内容をそのまま共有すれば、" & vbCrLf & _
           "実際の行構成に合わせてセットアップを組み直せます。", _
           vbInformation, "シート構造調査"

    LogSuccess MODULE_NAME, "ShiftSurvey_シート構造調査", _
               "Wrote structure report for sheet " & ws.Name & " (" & (r - 1) & " report rows)"
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSurvey_シート構造調査", Err.Number, Err.Description, Erl, _
             "reportRow=" & r
    MsgBox "調査でエラーが発生しました: " & Err.Description, vbExclamation
End Sub

'==================================================================
' レポートシートの準備
'==================================================================
Private Function SV_PrepareReport() As Worksheet
    Dim rpt As Worksheet
    On Error GoTo ErrHandler

    ' 既存のレポートは作り直す(調査結果以外は触らない)
10  On Error Resume Next
20  Application.DisplayAlerts = False
30  ThisWorkbook.Worksheets(RPT_SHEET).Delete
40  Application.DisplayAlerts = True
50  On Error GoTo ErrHandler

60  Set rpt = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
70  rpt.Name = RPT_SHEET
80  Set SV_PrepareReport = rpt
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_PrepareReport", Err.Number, Err.Description, Erl, _
             "sheet=" & RPT_SHEET
    Set SV_PrepareReport = Nothing
End Function

'--- 見出し行を書く ---
Private Function SV_Section(ByVal rpt As Worksheet, ByVal r As Long, ByVal title As String) As Long
    On Error GoTo ErrHandler

10  rpt.Cells(r, 1).Value = "■ " & title
20  With rpt.Range(rpt.Cells(r, 1), rpt.Cells(r, RPT_COLS))
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
30  End With
40  SV_Section = r + 1
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SV_Section", Err.Number, Err.Description, Erl, _
             "r=" & r & "; title=" & title
    SV_Section = r + 1
End Function

'--- 1行分を書く ---
Private Function SV_Row(ByVal rpt As Worksheet, ByVal r As Long, ByVal vals As Variant) As Long
    Dim i As Long
    On Error GoTo ErrHandler

10  For i = LBound(vals) To UBound(vals)
20      rpt.Cells(r, i - LBound(vals) + 1).Value = vals(i)
30  Next i
40  SV_Row = r + 1
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SV_Row", Err.Number, Err.Description, Erl, _
             "r=" & r & "; i=" & i
    SV_Row = r + 1
End Function

'==================================================================
' 1) ブック概要
'==================================================================
Private Function SV_WriteBookInfo(ByVal rpt As Worksheet, ByVal r0 As Long) As Long
    Dim r As Long, sh As Worksheet
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "1. ブック概要")
30  r = SV_Row(rpt, r, Array("", "ブック名", ThisWorkbook.Name))
40  r = SV_Row(rpt, r, Array("", "Excelバージョン", Application.Version))
50  r = SV_Row(rpt, r, Array("", "調査日時", Format(Now, "yyyy/mm/dd hh:nn:ss")))
60  r = r + 1
70  r = SV_Row(rpt, r, Array("", "シート名", "使用範囲", "行数", "列数"))
80  For Each sh In ThisWorkbook.Worksheets
90      If sh.Name <> RPT_SHEET Then
100         r = SV_Row(rpt, r, Array("", sh.Name, sh.UsedRange.Address(False, False), _
                                     sh.UsedRange.Rows.Count, sh.UsedRange.Columns.Count))
110     End If
120 Next sh
130 SV_WriteBookInfo = r + 1
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_WriteBookInfo", Err.Number, Err.Description, Erl, "r=" & r
    SV_WriteBookInfo = r + 1
End Function

'==================================================================
' 2) 名前付き範囲
'==================================================================
Private Function SV_WriteNames(ByVal rpt As Worksheet, ByVal r0 As Long) As Long
    Dim r As Long, nm As Name, refs As String
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "2. 名前付き範囲")
30  r = SV_Row(rpt, r, Array("", "名前", "参照先", "解決できるか"))
40  If ThisWorkbook.Names.Count = 0 Then
50      r = SV_Row(rpt, r, Array("", "(定義なし)"))
60  Else
70      For Each nm In ThisWorkbook.Names
80          refs = SV_SafeRefersTo(nm)
90          r = SV_Row(rpt, r, Array("", nm.Name, nm.RefersTo, refs))
100     Next nm
110 End If
120 SV_WriteNames = r + 1
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_WriteNames", Err.Number, Err.Description, Erl, "r=" & r
    SV_WriteNames = r + 1
End Function

'--- 名前付き範囲が実アドレスに解決できるか ---
Private Function SV_SafeRefersTo(ByVal nm As Name) As String
    On Error GoTo ErrHandler

10  SV_SafeRefersTo = nm.RefersToRange.Address(False, False)
    Exit Function
ErrHandler:
    ' 参照切れ・数式定義など(想定内)
20  SV_SafeRefersTo = "(解決できません)"
End Function

'==================================================================
' 3) ShiftCommon が検出している位置
'==================================================================
Private Function SV_WriteDetected(ByVal rpt As Worksheet, ByVal ws As Worksheet, _
                                  ByVal r0 As Long) As Long
    Dim r As Long, c As Range, grid As Range, pal As Range
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "3. ShiftCommon が今どう判定しているか")
30  r = SV_Row(rpt, r, Array("", "項目", "結果", "備考"))

40  Set c = StartDateCell(ws)
50  If c Is Nothing Then
60      r = SV_Row(rpt, r, Array("", "開始日の数式セル", "(未検出)", _
                                 "B列に「数式かつ日付」のセットが無い"))
70  Else
80      r = SV_Row(rpt, r, Array("", "開始日の数式セル", c.Address(False, False), c.Formula))
90  End If

100 r = SV_Row(rpt, r, Array("", "日付行", DateRow(ws)))
110 r = SV_Row(rpt, r, Array("", "曜日行(ラベル)", LabelRow(ws, LBL_WEEK)))
120 r = SV_Row(rpt, r, Array("", "備考行(ラベル)", LabelRow(ws, LBL_NOTE)))
130 r = SV_Row(rpt, r, Array("", "医師数行", ShiftDocRow(ws), _
                             "備考行 + " & NOTE_TO_DOC))
140 r = SV_Row(rpt, r, Array("", "入力欄 上端", ShiftTopRow(ws)))
150 r = SV_Row(rpt, r, Array("", "入力欄 下端", ShiftBottomRow(ws)))
160 r = SV_Row(rpt, r, Array("", "パレット本体行", PaletteBodyRow(ws), _
                             "日付行 - " & PALETTE_GAP))
170 r = SV_Row(rpt, r, Array("", "下端のずれ", ShiftRangeDrift(ws), _
                             "0=正常 / 正=下すぎ / 負=上すぎ"))

180 Set grid = ShiftInputRange(ws)
190 r = SV_Row(rpt, r, Array("", "シフト入力範囲", _
                             IIf(grid Is Nothing, "(未特定)", grid.Address(False, False))))
200 Set pal = PaletteRange(ws)
210 r = SV_Row(rpt, r, Array("", "パレット範囲", _
                             IIf(pal Is Nothing, "(未特定)", pal.Address(False, False))))

220 SV_WriteDetected = r + 1
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_WriteDetected", Err.Number, Err.Description, Erl, "r=" & r
    SV_WriteDetected = r + 1
End Function

'==================================================================
' 4) 行別ダンプ(ここが本命)
'==================================================================
Private Function SV_WriteRowDump(ByVal rpt As Worksheet, ByVal ws As Worksheet, _
                                 ByVal r0 As Long) As Long
    Dim r As Long, i As Long, nameSeq As Long
    Dim aVal As String, bCell As Range
    Dim nVal As Long, nFml As Long, lastC As Long
    Dim bKind As String, bFml As String
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "4. シフトシートの行別ダンプ(上から " & SCAN_ROWS & " 行)")
30  r = SV_Row(rpt, r, Array("", "行", "A列の値", "B列の種別", "値のある列数", _
                             "数式のある列数", "最終列", "行の高さ", "塗り", "太字", "B列の数式"))

40  For i = 1 To SCAN_ROWS
50      aVal = Trim$(CStr(ws.Cells(i, 1).Value))
60      Set bCell = ws.Cells(i, 2)
70      nVal = Application.WorksheetFunction.CountA(ws.Range(ws.Cells(i, 2), ws.Cells(i, SCAN_COLS)))
80      nFml = SV_CountFormulas(ws, i)
90      lastC = SV_LastCol(ws, i)
100     bKind = SV_CellKind(bCell)
110     bFml = ""
120     If bCell.HasFormula Then bFml = bCell.Formula

        ' 空行(A列も B以降も空)は行数節約のため飛ばさず、状態だけ残す
130     If MASK_NAMES Then
140         If Len(aVal) > 0 Then
150             If Not SV_IsKnownLabel(aVal) Then
160                 nameSeq = nameSeq + 1
170                 aVal = "(氏名" & nameSeq & ")"
180             End If
190         End If
200     End If

210     r = SV_Row(rpt, r, Array("", i, aVal, bKind, nVal, nFml, _
                                 IIf(lastC = 0, "", SV_ColLetterSafe(lastC)), _
                                 Round(ws.Rows(i).RowHeight, 1), _
                                 IIf(ws.Cells(i, 2).Interior.Pattern <> xlNone, "有", ""), _
                                 IIf(ws.Cells(i, 1).Font.Bold, "有", ""), _
                                 bFml))
220 Next i

230 SV_WriteRowDump = r + 1
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_WriteRowDump", Err.Number, Err.Description, Erl, _
             "r=" & r & "; i=" & i
    SV_WriteRowDump = r + 1
End Function

'--- セルの中身の種別(値そのものは出さない) ---
Private Function SV_CellKind(ByVal c As Range) As String
    Dim s As String
    On Error GoTo ErrHandler

10  If Len(Trim$(CStr(c.Value))) = 0 Then
20      SV_CellKind = "空"
30      Exit Function
40  End If
50  If IsDate(c.Value) Then
60      s = "日付"
70  ElseIf IsNumeric(c.Value) Then
80      s = "数値"
90  Else
100     s = "文字"
110 End If
120 If c.HasFormula Then s = s & "(数式)"
130 SV_CellKind = s
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SV_CellKind", Err.Number, Err.Description, Erl, _
             "cell=" & c.Address(False, False)
    SV_CellKind = "(不明)"
End Function

'--- その行で数式が入っている列の数 ---
Private Function SV_CountFormulas(ByVal ws As Worksheet, ByVal rowNo As Long) As Long
    Dim c As Long, n As Long
    On Error GoTo ErrHandler

10  For c = 2 To SCAN_COLS
20      If ws.Cells(rowNo, c).HasFormula Then n = n + 1
30  Next c
40  SV_CountFormulas = n
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SV_CountFormulas", Err.Number, Err.Description, Erl, _
             "rowNo=" & rowNo & "; c=" & c
    SV_CountFormulas = 0
End Function

'--- その行で値が入っている最終列(0 = 空行) ---
Private Function SV_LastCol(ByVal ws As Worksheet, ByVal rowNo As Long) As Long
    Dim c As Long
    On Error GoTo ErrHandler

10  For c = SCAN_COLS To 1 Step -1
20      If Len(Trim$(CStr(ws.Cells(rowNo, c).Value))) > 0 Then
30          SV_LastCol = c
40          Exit Function
50      End If
60  Next c
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SV_LastCol", Err.Number, Err.Description, Erl, _
             "rowNo=" & rowNo & "; c=" & c
    SV_LastCol = 0
End Function

Private Function SV_ColLetterSafe(ByVal colNo As Long) As String
    On Error GoTo ErrHandler

10  SV_ColLetterSafe = ColLetter(colNo)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SV_ColLetterSafe", Err.Number, Err.Description, Erl, _
             "colNo=" & colNo
    SV_ColLetterSafe = ""
End Function

'--- 伏せなくてよい既知の見出しか ---
Private Function SV_IsKnownLabel(ByVal s As String) As Boolean
    Dim known As Variant, i As Long, t As String
    On Error GoTo ErrHandler

10  t = Trim$(s)
20  known = Array(LBL_WEEK, LBL_NOTE, LBL_DOC, LBL_PHARM, LBL_CLERK, _
                  LBL_SHORT, LBL_PALETTE, LBL_DOCTORS, _
                  "氏名", "合計", "休", "早番", "遅番", "遅半", _
                  "公休", "希休", "夏休", "有休", "銀行", "停止", "順送り", _
                  "背景消", "消去", "備考付", "医師")
30  For i = LBound(known) To UBound(known)
40      If Len(CStr(known(i))) > 0 Then
50          If InStr(1, t, CStr(known(i)), vbTextCompare) = 1 Then
60              SV_IsKnownLabel = True
70              Exit Function
80          End If
90      End If
100 Next i
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SV_IsKnownLabel", Err.Number, Err.Description, Erl, "s=" & s
    SV_IsKnownLabel = False
End Function

'==================================================================
' 5) 列幅
'==================================================================
Private Function SV_WriteColWidths(ByVal rpt As Worksheet, ByVal ws As Worksheet, _
                                   ByVal r0 As Long) As Long
    Dim r As Long, c As Long
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "5. 列幅(A から " & SCAN_COLS & " 列)")
30  r = SV_Row(rpt, r, Array("", "列", "幅", "非表示"))
40  For c = 1 To SCAN_COLS
50      r = SV_Row(rpt, r, Array("", SV_ColLetterSafe(c), _
                                 Round(ws.Columns(c).ColumnWidth, 1), _
                                 IIf(ws.Columns(c).Hidden, "非表示", "")))
60  Next c
70  SV_WriteColWidths = r + 1
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_WriteColWidths", Err.Number, Err.Description, Erl, _
             "r=" & r & "; c=" & c
    SV_WriteColWidths = r + 1
End Function

'==================================================================
' 6) 自動作成設定シート
'==================================================================
Private Function SV_WriteConfigSheet(ByVal rpt As Worksheet, ByVal r0 As Long) As Long
    Dim r As Long, cfg As Worksheet, i As Long, lbl As String, v As String
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "6. 自動作成設定シート")

30  Set cfg = SV_SheetOrNothing(SHT_CFG)
40  If cfg Is Nothing Then
50      r = SV_Row(rpt, r, Array("", "(シートがありません)"))
60      SV_WriteConfigSheet = r + 1
70      Exit Function
80  End If

    '--- メンバー表のヘッダー(4行目) ---
90  r = SV_Row(rpt, r, Array("", "メンバー表ヘッダー(4行目)"))
100 For i = 1 To 9
110     r = SV_Row(rpt, r, Array("", SV_ColLetterSafe(i) & "4", _
                                 CStr(cfg.Cells(4, i).Value)))
120 Next i
130 r = r + 1

    '--- 全体設定(K/L) ---
140 r = SV_Row(rpt, r, Array("", "全体設定(K列ラベル / L列の値)"))
150 For i = 4 To 20
160     lbl = Trim$(CStr(cfg.Cells(i, 11).Value))
170     v = Trim$(CStr(cfg.Cells(i, 12).Value))
180     If Len(lbl) > 0 Or Len(v) > 0 Then
190         r = SV_Row(rpt, r, Array("", "K" & i & " / L" & i, lbl, v))
200     End If
210 Next i

220 r = r + 1
230 r = SV_Row(rpt, r, Array("", "メンバー行数", _
                             SV_CountMembers(cfg), "A列が空欄になるまで"))

240 SV_WriteConfigSheet = r + 1
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_WriteConfigSheet", Err.Number, Err.Description, Erl, _
             "r=" & r & "; i=" & i
    SV_WriteConfigSheet = r + 1
End Function

Private Function SV_SheetOrNothing(ByVal nm As String) As Worksheet
    On Error GoTo ErrHandler

10  Set SV_SheetOrNothing = ThisWorkbook.Worksheets(nm)
    Exit Function
ErrHandler:
    ' シートが無い(想定内)
20  Set SV_SheetOrNothing = Nothing
End Function

Private Function SV_CountMembers(ByVal cfg As Worksheet) As Long
    Dim r As Long, n As Long
    On Error GoTo ErrHandler

10  r = 5
20  Do While Len(Trim$(CStr(cfg.Cells(r, 1).Value))) > 0
30      n = n + 1
40      r = r + 1
50      If n > 500 Then Exit Do
60  Loop
70  SV_CountMembers = n
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SV_CountMembers", Err.Number, Err.Description, Erl, "r=" & r
    SV_CountMembers = 0
End Function
