Attribute VB_Name = "ShiftExport"
Option Explicit
'==================================================================
'  シフト表 出力モジュール ＜標準モジュール ShiftExport v1.0＞
'  2026-08-27
'
'  目的:
'    印刷・配布用にシフト表を切り出して、PDF または Excel で保存する。
'
'  出力の仕様:
'    範囲   = 年月・タイトル行(日付行の1行上) ～ 過不足行
'             列は A ～ EXPORT_COL_LAST(集計列まで)
'             パレットの3行は範囲外なので出力されない
'    内容   = 数式は値に変換する(元ブックが無くても崩れないため)
'    ページ = 横向き / 横1ページ x 縦1ページに収める
'
'  元のシートは一切書き換えない。新しいブックに複製してから加工する。
'==================================================================
Private Const MODULE_NAME As String = "ShiftExport"

'--- 出力する最終列(集計列 AL まで含める) ---
Private Const EXPORT_COL_LAST As String = "AL"

'--- 余白(cm) ---
Private Const MARGIN_CM As Double = 0.6

'--- 形式の選択 ---
Private Const FMT_CANCEL As Long = 0
Private Const FMT_PDF    As Long = 1
Private Const FMT_XLSX   As Long = 2

'==================================================================
' 入口
'==================================================================
Public Sub ShiftExport_シフト表出力()
    Dim ws As Worksheet, src As Range
    Dim fmt As Long, path As String
    Dim newBk As Workbook, dst As Worksheet
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  Set src = XP_SourceRange(ws)
30  If src Is Nothing Then
40      MsgBox "出力する範囲を特定できません。" & vbCrLf & _
               "ShiftClick_セルフチェック で行構成を確認してください。", _
               vbExclamation, "シフト表の出力"
50      Exit Sub
60  End If

70  fmt = XP_AskFormat()
80  If fmt = FMT_CANCEL Then Exit Sub

90  path = XP_AskPath(ws, fmt)
100 If Len(path) = 0 Then Exit Sub

110 Application.ScreenUpdating = False

120 Set newBk = XP_BuildBook(ws, src, dst)
130 If newBk Is Nothing Then
140     MsgBox "出力用ブックを作成できませんでした。", vbExclamation, "シフト表の出力"
150     GoTo CleanUp
160 End If

170 XP_SetPageSetup dst
180 If Not XP_Save(newBk, dst, fmt, path) Then GoTo CleanUp

190 MsgBox "シフト表を出力しました。" & vbCrLf & vbCrLf & path, _
           vbInformation, "シフト表の出力"

CleanUp:
    On Error Resume Next
    Application.CutCopyMode = False
    If Not newBk Is Nothing Then newBk.Close SaveChanges:=False
    Application.ScreenUpdating = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftExport_シフト表出力", _
               "format=" & fmt & "; range=" & _
               IIf(src Is Nothing, "none", src.Address(False, False))
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftExport_シフト表出力", Err.Number, Err.Description, Erl, _
             "format=" & fmt & "; path=" & path
    MsgBox "出力でエラーが発生しました: " & Err.Description, vbExclamation, "シフト表の出力"
    Resume CleanUp
End Sub

'==================================================================
' 出力範囲
'==================================================================
'--- 年月・タイトル行 ～ 過不足行 / A ～ EXPORT_COL_LAST ---
Private Function XP_SourceRange(ByVal ws As Worksheet) As Range
    Dim topR As Long, botR As Long, lastCol As Long
    On Error GoTo ErrHandler

10  topR = HeaderRow(ws)
20  botR = ShortageRow(ws)
30  If topR = 0 Or botR = 0 Or topR > botR Then Exit Function
40  lastCol = ws.Range(EXPORT_COL_LAST & "1").Column
50  Set XP_SourceRange = ws.Range(ws.Cells(topR, 1), ws.Cells(botR, lastCol))
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "XP_SourceRange", Err.Number, Err.Description, Erl, _
             "topRow=" & topR & "; bottomRow=" & botR
    Set XP_SourceRange = Nothing
End Function

'==================================================================
' 形式と保存先
'==================================================================
Private Function XP_AskFormat() As Long
    Dim a As VbMsgBoxResult
    On Error GoTo ErrHandler

10  a = MsgBox("シフト表を出力します。形式を選んでください。" & vbCrLf & vbCrLf & _
               "[はい]    PDF" & vbCrLf & _
               "[いいえ]  Excel (.xlsx)" & vbCrLf & _
               "[キャンセル] やめる", _
               vbYesNoCancel + vbQuestion, "シフト表の出力")
20  If a = vbYes Then
30      XP_AskFormat = FMT_PDF
40  ElseIf a = vbNo Then
50      XP_AskFormat = FMT_XLSX
60  Else
70      XP_AskFormat = FMT_CANCEL
80  End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "XP_AskFormat", Err.Number, Err.Description, Erl, ""
    XP_AskFormat = FMT_CANCEL
End Function

'--- 保存先を聞く(空文字 = キャンセル) ---
Private Function XP_AskPath(ByVal ws As Worksheet, ByVal fmt As Long) As String
    Dim baseName As String, filt As String, ext As String, v As Variant
    On Error GoTo ErrHandler

10  baseName = XP_DefaultName(ws)
20  If fmt = FMT_PDF Then
30      filt = "PDF ファイル (*.pdf),*.pdf"
40      ext = ".pdf"
50  Else
60      filt = "Excel ブック (*.xlsx),*.xlsx"
70      ext = ".xlsx"
80  End If

90  v = Application.GetSaveAsFilename( _
            InitialFileName:=baseName & ext, _
            FileFilter:=filt, _
            Title:="シフト表の保存先")
100 If VarType(v) = vbBoolean Then Exit Function   ' キャンセル
110 XP_AskPath = CStr(v)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "XP_AskPath", Err.Number, Err.Description, Erl, _
             "format=" & fmt
    XP_AskPath = ""
End Function

'--- 既定のファイル名(対象月から作る) ---
Private Function XP_DefaultName(ByVal ws As Worksheet) As String
    Dim mc As Range
    On Error GoTo ErrHandler

10  XP_DefaultName = "シフト表"
20  Set mc = MonthCell(ws)
30  If mc Is Nothing Then Exit Function
40  If Not IsDate(mc.Value) Then Exit Function
50  XP_DefaultName = "シフト表_" & Format(mc.Value, "yyyy年m月")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "XP_DefaultName", Err.Number, Err.Description, Erl, ""
    XP_DefaultName = "シフト表"
End Function

'==================================================================
' 出力用ブックの組み立て
'==================================================================
'--- 新しいブックに、値と書式だけを複製する ---
'    元シートをコピーするとシートモジュールまで付いてくるため、
'    空のブックを作って貼り付ける方式にしている。
Private Function XP_BuildBook(ByVal ws As Worksheet, ByVal src As Range, _
                              ByRef dst As Worksheet) As Workbook
    Dim bk As Workbook, i As Long
    On Error GoTo ErrHandler

10  Set bk = Workbooks.Add(xlWBATWorksheet)
20  Set dst = bk.Worksheets(1)

30  src.Copy
40  dst.Range("A1").PasteSpecial xlPasteColumnWidths
50  dst.Range("A1").PasteSpecial xlPasteAll        ' 書式・結合セルを持ってくる
60  dst.Range("A1").PasteSpecial xlPasteValues     ' 数式は値に置き換える
70  Application.CutCopyMode = False

    '--- 行の高さは貼り付けで写らないので個別に合わせる ---
80  For i = 1 To src.Rows.Count
90      dst.Rows(i).RowHeight = src.Rows(i).RowHeight
100 Next i

110 dst.Range("A1").Select
120 Set XP_BuildBook = bk
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "XP_BuildBook", Err.Number, Err.Description, Erl, _
             "src=" & src.Address(False, False) & "; i=" & i
    On Error Resume Next
    Application.CutCopyMode = False
    If Not bk Is Nothing Then bk.Close SaveChanges:=False
    On Error GoTo 0
    Set XP_BuildBook = Nothing
End Function

'--- 横向き / 横1 x 縦1 に収める ---
Private Sub XP_SetPageSetup(ByVal dst As Worksheet)
    On Error GoTo ErrHandler

10  With dst.PageSetup
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .PrintArea = dst.UsedRange.Address
        .LeftMargin = Application.CentimetersToPoints(MARGIN_CM)
        .RightMargin = Application.CentimetersToPoints(MARGIN_CM)
        .TopMargin = Application.CentimetersToPoints(MARGIN_CM)
        .BottomMargin = Application.CentimetersToPoints(MARGIN_CM)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
        .CenterHorizontally = True
20  End With
    Exit Sub
ErrHandler:
    ' プリンタ未設定の環境では PageSetup が失敗することがある(出力自体は続行)
    LogError MODULE_NAME, "XP_SetPageSetup", Err.Number, Err.Description, Erl, _
             "sheet=" & dst.Name
End Sub

'--- 保存(成功なら True) ---
Private Function XP_Save(ByVal bk As Workbook, ByVal dst As Worksheet, _
                         ByVal fmt As Long, ByVal path As String) As Boolean
    On Error GoTo ErrHandler

10  If fmt = FMT_PDF Then
20      dst.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path, _
                                Quality:=xlQualityStandard, _
                                IncludeDocProperties:=False, _
                                IgnorePrintAreas:=False, _
                                OpenAfterPublish:=False
30  Else
40      Application.DisplayAlerts = False
50      bk.SaveAs Filename:=path, FileFormat:=xlOpenXMLWorkbook
60      Application.DisplayAlerts = True
70  End If
80  XP_Save = True
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "XP_Save", Err.Number, Err.Description, Erl, _
             "format=" & fmt & "; path=" & path
    On Error Resume Next
    Application.DisplayAlerts = True
    On Error GoTo 0
    MsgBox "保存できませんでした: " & Err.Description & vbCrLf & vbCrLf & _
           "同じ名前のファイルを開いていないか確認してください。", _
           vbExclamation, "シフト表の出力"
    XP_Save = False
End Function
