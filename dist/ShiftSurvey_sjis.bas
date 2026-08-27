Attribute VB_Name = "ShiftSurvey"
Option Explicit
'==================================================================
'  シフト表 構造調査マクロ ＜標準モジュール ShiftSurvey v2.2＞
'  2026-08-27
'
'  目的:
'    ShiftSetup / ShiftSchema を「白紙のブックから生成できる」ように
'    するための現状把握。実物のシートが実際にどういう行構成・数式・
'    書式になっているかをレポートシートに書き出す。
'
'  重要: 本モジュールは読み取り専用。
'        レポートシート以外は一切書き換えない。
'
'  個人情報:
'    MASK_NAMES = True のとき、既知の見出し以外のA列の文字列は
'    「(氏名1)」等に伏せて出力する。B:AF のセルの値は出力しない
'    (医師名・スタッフ名が出ないようにするため)。
'
'  v2.2 変更:
'   ・医師名ラベルの参照を LBL_DOC_STAMP に統一(リテラルを廃止)。
'
'  v2.1 変更:
'   ・パレットの定数一覧から廃止した IDX_DOC_LAST / IDX_NOTE_FIRST を外し、
'     ラベルから求めた医師名の最終位置を出すよう変更。
'
'  v2.0 変更:
'   ・パレットの内訳(インデックス/値の有無/塗り)を出力する節を追加
'     → IDX_* 定数とライブのパレットのずれを検出できる
'   ・依存シートの有無と、ShiftCommon の定数実測値の照合を追加
'   ・StartDateCell / DateFormulaRow の2個目(再掲日付行)も出力
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
30  If rpt Is Nothing Then
40      MsgBox "レポートシートを準備できませんでした。", vbExclamation
50      Exit Sub
60  End If
70  r = 1

80  Application.ScreenUpdating = False

90  r = SV_WriteBookInfo(rpt, r)
100 r = SV_WriteNames(rpt, r)
110 r = SV_WriteDetected(rpt, ws, r)
120 r = SV_WritePalette(rpt, ws, r)
130 r = SV_WriteRowDump(rpt, ws, r)
140 r = SV_WriteColWidths(rpt, ws, r)
150 r = SV_WriteConfigSheet(rpt, r)

160 rpt.Columns(1).ColumnWidth = 6
170 rpt.Columns(2).ColumnWidth = 24
180 rpt.Range(rpt.Cells(1, 3), rpt.Cells(1, RPT_COLS)).EntireColumn.ColumnWidth = 16
190 rpt.Activate
200 rpt.Range("A1").Select

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSurvey_シート構造調査", _
               "Wrote structure report for sheet " & ws.Name & " (" & (r - 1) & " rows)"
210 MsgBox "シートの構造を「" & RPT_SHEET & "」に書き出しました。" & vbCrLf & vbCrLf & _
           "氏名は " & IIf(MASK_NAMES, "伏せて", "そのまま") & "出力しています。" & vbCrLf & _
           "このシートの内容を共有すれば、" & vbCrLf & _
           "実際の行構成に合わせてセットアップを組み直せます。", _
           vbInformation, "シート構造調査"
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSurvey_シート構造調査", Err.Number, Err.Description, Erl, _
             "reportRow=" & r
    MsgBox "調査でエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
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

60  Set rpt = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
70  rpt.Name = RPT_SHEET
80  Set SV_PrepareReport = rpt
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_PrepareReport", Err.Number, Err.Description, Erl, _
             "sheet=" & RPT_SHEET
    Set SV_PrepareReport = Nothing
End Function

'--- 見出し行を書く ---
Private Function SV_Section(ByVal rpt As Worksheet, ByVal r As Long, _
                            ByVal title As String) As Long
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
Private Function SV_Row(ByVal rpt As Worksheet, ByVal r As Long, _
                        ByVal vals As Variant) As Long
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

    '--- 依存シートの有無 ---
70  r = SV_Row(rpt, r, Array("", "必要なシート", "存在するか"))
80  r = SV_Row(rpt, r, Array("", SHT_SHIFT, IIf(SheetExists(SHT_SHIFT), "有", "■無")))
90  r = SV_Row(rpt, r, Array("", SHT_CFG, IIf(SheetExists(SHT_CFG), "有", "■無")))
100 r = SV_Row(rpt, r, Array("", SHT_HOLIDAY, IIf(SheetExists(SHT_HOLIDAY), "有", "■無")))
110 r = SV_Row(rpt, r, Array("", SHT_LOG, IIf(SheetExists(SHT_LOG), "有", "■無")))
120 r = r + 1

130 r = SV_Row(rpt, r, Array("", "シート名", "使用範囲", "行数", "列数"))
140 For Each sh In ThisWorkbook.Worksheets
150     If sh.Name <> RPT_SHEET Then
160         r = SV_Row(rpt, r, Array("", sh.Name, sh.UsedRange.Address(False, False), _
                                     sh.UsedRange.Rows.Count, sh.UsedRange.Columns.Count))
170     End If
180 Next sh
190 SV_WriteBookInfo = r + 1
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
90          r = SV_Row(rpt, r, Array("", nm.Name, "'" & nm.refersTo, refs))
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
    Dim r As Long, c As Range, grid As Range, pal As Range, blk As Range
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "3. ShiftCommon が今どう判定しているか")
30  r = SV_Row(rpt, r, Array("", "項目", "結果", "備考"))

40  Set c = StartDateCell(ws)
50  If c Is Nothing Then
60      r = SV_Row(rpt, r, Array("", "開始日の数式セル", "(未検出)", _
                                 "B列に「数式かつ日付」のセルが無い"))
70  Else
80      r = SV_Row(rpt, r, Array("", "開始日の数式セル", c.Address(False, False), _
                                 "'" & c.Formula))
90  End If

100 r = SV_Row(rpt, r, Array("", "日付行", DateRow(ws)))
110 r = SV_Row(rpt, r, Array("", "再掲日付行", DateFormulaRow(ws, 2), _
                             "入力欄の上端の基準"))
120 r = SV_Row(rpt, r, Array("", "年月・タイトル行", HeaderRow(ws), "日付行 - 1"))
130 r = SV_Row(rpt, r, Array("", "曜日行(ラベル)", LabelRow(ws, LBL_WEEK), _
                             "0=A列にラベル無し(正常)"))
140 r = SV_Row(rpt, r, Array("", "備考行(ラベル)", LabelRow(ws, LBL_NOTE)))
150 r = SV_Row(rpt, r, Array("", "医師数行", ShiftDocRow(ws), _
                             "備考行 + " & NOTE_TO_DOC))
160 r = SV_Row(rpt, r, Array("", "入力欄 上端", ShiftTopRow(ws), _
                             "再掲日付行 + " & DATE_REPEAT_GAP))
170 r = SV_Row(rpt, r, Array("", "入力欄 下端", ShiftBottomRow(ws), _
                             "医師数行 - " & DOC_GAP))
180 r = SV_Row(rpt, r, Array("", "パレット本体行", PaletteBodyRow(ws), _
                             "日付行 - " & PALETTE_GAP))
190 r = SV_Row(rpt, r, Array("", "下端のずれ", ShiftRangeDrift(ws), _
                             "0=正常 / 正=下すぎ / 負=上すぎ"))

200 Set blk = DoctorBlock(ws)
210 r = SV_Row(rpt, r, Array("", "医師名欄", _
                             IIf(blk Is Nothing, "(未特定)", blk.Address(False, False)), _
                             "既定 " & DOC_BLOCK_ROWS & " 行"))
215 r = SV_Row(rpt, r, Array("", "備考行", LabelRow(ws, LBL_NOTE), _
                             "入力欄の外。クリック入力とログの対象"))
220 Set grid = ShiftInputRange(ws)
230 r = SV_Row(rpt, r, Array("", "シフト入力範囲", _
                             IIf(grid Is Nothing, "(未特定)", grid.Address(False, False))))
240 Set pal = PaletteRange(ws)
250 r = SV_Row(rpt, r, Array("", "パレット範囲", _
                             IIf(pal Is Nothing, "(未特定)", pal.Address(False, False))))

260 SV_WriteDetected = r + 1
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_WriteDetected", Err.Number, Err.Description, Erl, "r=" & r
    SV_WriteDetected = r + 1
End Function

'==================================================================
' 4) パレットの内訳(IDX_* 定数とのずれを検出する)
'==================================================================
Private Function SV_WritePalette(ByVal rpt As Worksheet, ByVal ws As Worksheet, _
                                 ByVal r0 As Long) As Long
    Dim r As Long, pal As Range, i As Long
    Dim v As String, lab As String, role As String, fill As String
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "4. パレットの内訳(ShiftCommon の IDX_* と照合)")

30  Set pal = PaletteRange(ws)
40  If pal Is Nothing Then
50      r = SV_Row(rpt, r, Array("", "(パレットを特定できません)"))
60      SV_WritePalette = r + 1
70      Exit Function
80  End If

    '--- 定数の一覧 ---
90  r = SV_Row(rpt, r, Array("", "定数", "値"))
100 r = SV_Row(rpt, r, Array("", "IDX_OFF", IDX_OFF))
110 r = SV_Row(rpt, r, Array("", "IDX_AUTO", IDX_AUTO))
120 r = SV_Row(rpt, r, Array("", "IDX_CYCLE", IDX_CYCLE))
130 r = SV_Row(rpt, r, Array("", "IDX_CLEARFILL", IDX_CLEARFILL))
140 r = SV_Row(rpt, r, Array("", "IDX_FILL_FIRST", IDX_FILL_FIRST))
150 r = SV_Row(rpt, r, Array("", "IDX_FILL_LAST", IDX_FILL_LAST))
160 r = SV_Row(rpt, r, Array("", "IDX_ERASE", IDX_ERASE))
170 r = SV_Row(rpt, r, Array("", "IDX_SYM_FIRST", IDX_SYM_FIRST))
180 r = SV_Row(rpt, r, Array("", "IDX_SYM_LAST", IDX_SYM_LAST))
185 r = SV_Row(rpt, r, Array("", "IDX_UNDO", IDX_UNDO))
186 r = SV_Row(rpt, r, Array("", "IDX_EXPORT", IDX_EXPORT))
190 r = SV_Row(rpt, r, Array("", "IDX_DOC_FIRST", IDX_DOC_FIRST))
195 r = SV_Row(rpt, r, Array("", "医師名の最終位置", LastDoctorIndex(), _
                             "ラベルが「" & LBL_DOC_STAMP & "」の最後の位置"))
196 r = SV_Row(rpt, r, Array("", "DOC_SLOTS(生成時の枠数)", DOC_SLOTS, _
                             "パレット生成で作る医師枠の数"))
200 r = SV_Row(rpt, r, Array("", "パレットのセル数", pal.Cells.Count))
210 r = r + 1

    '--- セルごとの内訳 ---
220 r = SV_Row(rpt, r, Array("", "番号", "セル", "値", "ラベル", _
                             "定数上の役割", "塗り", MARKER_CHAR))
230 For i = 1 To pal.Cells.Count
240     v = Trim$(CStr(pal.Cells(1, i).Value))
250     lab = Trim$(CStr(pal.Cells(1, i).Offset(LABEL_OFFSET, 0).Value))
260     role = SV_PaletteRole(i)
270     If pal.Cells(1, i).Interior.Pattern <> xlNone Then fill = "有" Else fill = ""
        '--- 医師名は伏せる ---
280     If MASK_NAMES And IsDoctorStamp(i) And Len(v) > 0 Then
290         If lab = LBL_DOC_STAMP Then _
                v = "(" & LBL_DOC_STAMP & (i - IDX_DOC_FIRST + 1) & ")"
300     End If
310     r = SV_Row(rpt, r, Array("", i, pal.Cells(1, i).Address(False, False), _
                                 v, lab, role, fill, _
                                 IIf(StrComp(Trim$(CStr(pal.Cells(1, i) _
                                     .Offset(MARKER_OFFSET, 0).Value)), _
                                     MARKER_CHAR, vbTextCompare) = 0, MARKER_CHAR, "")))
320 Next i

330 SV_WritePalette = r + 1
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_WritePalette", Err.Number, Err.Description, Erl, _
             "r=" & r & "; i=" & i
    SV_WritePalette = r + 1
End Function

'--- パレット番号 → 定数上の役割 ---
Private Function SV_PaletteRole(ByVal idx As Long) As String
    On Error GoTo ErrHandler
        Select Case idx
            Case IDX_OFF:       SV_PaletteRole = "OFF"
            Case IDX_AUTO:      SV_PaletteRole = "自動(AutoShift)"
            Case IDX_UNDO:      SV_PaletteRole = "戻す(変更の取消)"
            Case IDX_EXPORT:    SV_PaletteRole = "出力(PDF/Excel)"
            Case IDX_CYCLE:     SV_PaletteRole = "連続切替"
            Case IDX_CLEARFILL: SV_PaletteRole = "背景色クリア"
            Case IDX_FILL_FIRST To IDX_FILL_LAST
                                SV_PaletteRole = "背景色ペイント"
            Case IDX_ERASE:     SV_PaletteRole = "消去(空白)"
            Case Else
                If IsDoctorStamp(idx) Then
                    SV_PaletteRole = "医師名スタンプ"
                ElseIf IsNoteStamp(idx) Then
                    SV_PaletteRole = "備考スタンプ"
                ElseIf idx >= IDX_SYM_FIRST And idx <= IDX_SYM_LAST Then
                    SV_PaletteRole = "出勤記号"
                Else
                    SV_PaletteRole = "スタンプ"
                End If
        End Select
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SV_PaletteRole", Err.Number, Err.Description, Erl, "idx=" & idx
    SV_PaletteRole = "(不明)"
End Function

'==================================================================
' 5) 行別ダンプ
'==================================================================
Private Function SV_WriteRowDump(ByVal rpt As Worksheet, ByVal ws As Worksheet, _
                                 ByVal r0 As Long) As Long
    Dim r As Long, i As Long, nameSeq As Long
    Dim aVal As String, bCell As Range
    Dim nVal As Long, nFml As Long, lastC As Long
    Dim bKind As String, bFml As String
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "5. シフトシートの行別ダンプ(上から " & SCAN_ROWS & " 行)")
30  r = SV_Row(rpt, r, Array("", "行", "A列の値", "B列の種別", "値のある列数", _
                             "数式のある列数", "最終列", "行の高さ", "塗り", "太字", _
                             "B列の数式"))

40  For i = 1 To SCAN_ROWS
50      aVal = Trim$(CStr(ws.Cells(i, 1).Value))
60      Set bCell = ws.Cells(i, 2)
70      nVal = Application.WorksheetFunction.CountA( _
                   ws.Range(ws.Cells(i, 2), ws.Cells(i, SCAN_COLS)))
80      nFml = SV_CountFormulas(ws, i)
90      lastC = SV_LastCol(ws, i)
100     bKind = SV_CellKind(bCell)
110     bFml = ""
120     If bCell.HasFormula Then bFml = "'" & bCell.Formula

        '--- 氏名を伏せる ---
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
                  "背景", "消去", "備考付", "医師", "自動", "○", "●", "▲")
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
' 6) 列幅
'==================================================================
Private Function SV_WriteColWidths(ByVal rpt As Worksheet, ByVal ws As Worksheet, _
                                   ByVal r0 As Long) As Long
    Dim r As Long, c As Long
    On Error GoTo ErrHandler
10  r = r0
20  r = SV_Section(rpt, r, "6. 列幅(A から " & SCAN_COLS & " 列)")
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
' 7) 自動作成設定シート
'==================================================================
Private Function SV_WriteConfigSheet(ByVal rpt As Worksheet, ByVal r0 As Long) As Long
    Dim r As Long, cfg As Worksheet, i As Long, lbl As String, v As String
    On Error GoTo ErrHandler

10  r = r0
20  r = SV_Section(rpt, r, "7. 自動作成設定シート")

30  Set cfg = SheetOrNothing(SHT_CFG)
40  If cfg Is Nothing Then
50      r = SV_Row(rpt, r, Array("", "(シートがありません)", _
                                 "ShiftSchema_不足シート生成 で作成できます"))
60      SV_WriteConfigSheet = r + 1
70      Exit Function
80  End If

    '--- メンバー表の見出し ---
90  r = SV_Row(rpt, r, Array("", "メンバー表ヘッダー(" & CFG_HDR_ROW & "行目)"))
100 For i = CFG_COL_NAME To CFG_COL_MEMO
110     r = SV_Row(rpt, r, Array("", SV_ColLetterSafe(i) & CFG_HDR_ROW, _
                                 CStr(cfg.Cells(CFG_HDR_ROW, i).Value)))
120 Next i
130 r = r + 1

    '--- 全体設定(K/L) ---
140 r = SV_Row(rpt, r, Array("", "全体設定(K列ラベル / L列の値)"))
150 For i = CFG_SET_ROW To CFG_SET_ROW + 16
160     lbl = Trim$(CStr(cfg.Cells(i, CFG_COL_SETK).Value))
170     v = Trim$(CStr(cfg.Cells(i, CFG_COL_SETV).Value))
180     If Len(lbl) > 0 Or Len(v) > 0 Then
190         r = SV_Row(rpt, r, Array("", "K" & i & " / L" & i, lbl, v))
200     End If
210 Next i

220 r = r + 1
230 r = SV_Row(rpt, r, Array("", "メンバー行数", SV_CountMembers(cfg), _
                             "A列が空欄になるまで"))

240 SV_WriteConfigSheet = r + 1
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SV_WriteConfigSheet", Err.Number, Err.Description, Erl, _
             "r=" & r & "; i=" & i
    SV_WriteConfigSheet = r + 1
End Function

Private Function SV_CountMembers(ByVal cfg As Worksheet) As Long
    Dim r As Long, n As Long
    On Error GoTo ErrHandler
10  r = CFG_FIRST_ROW
20  Do While Len(Trim$(CStr(cfg.Cells(r, CFG_COL_NAME).Value))) > 0
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
