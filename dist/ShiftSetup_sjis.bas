Attribute VB_Name = "ShiftSetup"
Option Explicit
'==================================================================
'  シフト表 初期設定マクロ ＜標準モジュール ShiftSetup v2.5＞
'  2026-08-27
'  ShiftCommon / ShiftSchema / ShiftClick と併用。
'  シート上の位置はすべて ShiftCommon が解決する。
'
'  実行するのは基本これ1本:
'      ShiftSetup_初期設定実行
'  内訳:
'      0) 不足シートの生成      → ShiftSchema_不足シート生成
'      1) パレットの生成
'      2) 年月セル・祝日サマリーの数式
'      3) 集計行(医師数(診)/薬剤師出勤数/過不足)の数式
'      4) 集計列(AH:AL)の見出しと数式
'      5) 名前付き範囲の再定義(動的数式)
'
'  v2.5 変更:
'   ・医師名スタンプの装飾を位置(IDX_DOC_FIRST..IDX_DOC_LAST)から
'     ラベル判定(IsDoctorStamp)に変更。医師枠の増減に追随する。
'   ・装飾のガードを IDX_NOTE_FIRST から IDX_SYM_LAST に変更。
'     位置固定なのはモード・背景色・記号までなので、そこまでを見る。
'
'  v2.4 変更(検証報告 2026-08-27 の指摘に対応):
'   ・パレットの名前付き範囲を静的アドレスから動的数式に変更
'     実行のたびに手入力の動的定義を固定値で潰していた
'     (v2.1 で入力欄だけ直し、パレットが取り残されていた)
'   ・集計列の見出しを「空欄のときだけ書く」方式に変更
'     手で短縮した見出し(例 5診)を実行のたびに戻していた
'   ・SS_パレット装飾 の範囲外ガードを全ループに統一
'
'  v2.1 変更:
'   ・名前付き範囲を静的アドレスではなく動的数式で定義するよう変更
'     旧: "=シフト!$B$14:$AF$27" (固定。行を増減すると手動修正が必要)
'     新: "=シフト!$B$14:INDEX(シフト!$B:$AF,MATCH(\"備考*\",シフト!$A:$A,0)-NOTE_GAP,n)"
'     ※旧方式は実行するたびに手入力の動的定義を固定値で潰していた
'   ・入力欄の下端は「A列の備考ラベルの NOTE_GAP 行上」で解決する
'     (備考は1行固定なので、備考の1行上=区切りの空行、2行上=最終スタッフ行)
'   ・医師名リスト範囲 / 備考行範囲 も併せて定義する
'
'  v2.0 変更:
'   ・パレット定義をライブの26項目に一致させた
'     旧: 22項目(自動・背景色ボタンが無く、医師名がダミー)
'     ※旧定義のまま実行すると現在のパレットを壊していた
'   ・自動/背景緑/背景橙/背景灰 を追加。背景色ボタンは値を持たず
'     背景色のみを持つ(ShiftClick 側で塗り専用として扱う)
'   ・IDX_* / 色 / オフセットの重複定義を廃止し ShiftCommon を参照
'   ・重複していたセクション番号 "3)" を 3)/4) に整理
'   ・依存シートを先に作るため ShiftSchema を最初に呼ぶ
'   ・医師名は設定シートから読む方式に変更(ハードコードを廃止)
'==================================================================
Private Const MODULE_NAME As String = "ShiftSetup"

'--------------------------- 設定 ---------------------------------
' 集計列(このモジュール専用)
Private Const COL_AGG       As String = "AH"  ' 集計列の先頭
Private Const COL_AGG_END   As String = "AL"  ' 集計列の最終
Private Const COL_MONTH     As String = "AG"  ' 年月シリアルの置き場
Private Const COL_HOL_SUMM  As String = "I"   ' 祝日サマリーの置き場

' 医師数の「忙しい日」判定に使う人数(集計列 n診出勤)
Private Const DOC_BUSY_COUNT As Long = 5

' 集計行の見出し(空欄のときだけ補う)
Private Const HEAD_DOC   As String = "医師数(診)"
Private Const HEAD_PHARM As String = "薬剤師出勤数"
Private Const HEAD_SHORT As String = "過不足"


' 設定シートの「必要出勤数(医師数+n)の n」を探すラベル(前方一致)
Private Const CFG_KEY_REQ  As String = "必要出勤"
Private Const REQ_FALLBACK As Long = 1

' パレットの幅を数えるときに見る列数(ラベル行の COUNTA 範囲)
Private Const PAL_SCAN_COLS As Long = 200

' 書式
Private Const CLR_BLACK      As Long = 0
Private Const PAL_ROW_HEIGHT As Double = 20
Private Const FS_LABEL  As Long = 9
Private Const FS_BODY   As Long = 10
Private Const FS_SYMBOL As Long = 12
Private Const FS_TITLE  As Long = 14

'==================================================================
'  パレット定義
'  ShiftCommon の IDX_* と必ず一致させること。
'    1 OFF   2 自動   3 切替   4 色消
'    5 背景緑 6 背景橙 7 背景灰   ← 値は空。背景色のみ持つ
'    8 消去
'    9 ○  10 ●  11 ▲
'   12 公休 13 希休 14 夏休 15 有休 16 有休※
'   17-25 医師名   ← 設定シートから読む。無ければ "医師1".."医師9"
'   26 銀行
'
'  ※医師名は個人情報のためコードに実名を書かない。
'    シフトシートの医師名欄に既に入っている値を優先して読む。
'==================================================================

Private Function SS_PalValsBase() As Variant
    On Error GoTo ErrHandler
    '  医師名の手前まで(1..18)
    '  空文字は「値を持たないボタン」(背景色ペイント3つと消去)
    SS_PalValsBase = Array("OFF", "自動", "戻す", "出力", "切替", "色消", _
                           "", "", "", "", _
                           "○", "●", "▲", _
                           "公休", "希休", "夏休", "有休", "有休※")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_PalValsBase", Err.Number, Err.Description, Erl, ""
End Function

Private Function SS_PalLabsBase() As Variant
    On Error GoTo ErrHandler
    SS_PalLabsBase = Array("停止", "自動", "元に戻す", "印刷出力", "順送り", "背景消", _
                           "背景緑", "背景橙", "背景灰", "消去", _
                           "早番", "遅半", "遅番", _
                           "公休", "希休", "夏休", "有休", "備考付")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_PalLabsBase", Err.Number, Err.Description, Erl, ""
End Function

'--- 医師名の後ろに付く項目 ---
Private Function SS_PalValsTail() As Variant
    On Error GoTo ErrHandler
    SS_PalValsTail = Array("銀行")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_PalValsTail", Err.Number, Err.Description, Erl, ""
End Function
Private Function SS_PalLabsTail() As Variant
    On Error GoTo ErrHandler
    SS_PalLabsTail = Array("銀行")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_PalLabsTail", Err.Number, Err.Description, Erl, ""
End Function

'--- 背景色ボタンの色(IDX_FILL_FIRST から順に) ---
Private Function SS_FillColors() As Variant
    On Error GoTo ErrHandler
    SS_FillColors = Array(RGB(155, 187, 89), _
                          RGB(250, 191, 143), _
                          RGB(175, 175, 175))
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_FillColors", Err.Number, Err.Description, Erl, ""
End Function

'--- パレットの値配列を組み立てる(医師名は既存パレットから引き継ぐ) ---
Private Function SS_PalVals(ByVal ws As Worksheet) As Variant
    Dim base As Variant, tail As Variant, arr() As Variant
    Dim i As Long, n As Long, k As Long
    On Error GoTo ErrHandler

10  base = SS_PalValsBase()
20  tail = SS_PalValsTail()
30  n = (UBound(base) + 1) + DOC_SLOTS + (UBound(tail) + 1)
40  ReDim arr(0 To n - 1)

50  For i = LBound(base) To UBound(base)
60      arr(k) = base(i)
70      k = k + 1
80  Next i
    '--- 医師名: 既存パレットの値を引き継ぐ(無ければ既定名) ---
90  For i = 1 To DOC_SLOTS
100     arr(k) = SS_DoctorName(ws, i)
110     k = k + 1
120 Next i
130 For i = LBound(tail) To UBound(tail)
140     arr(k) = tail(i)
150     k = k + 1
160 Next i
170 SS_PalVals = arr
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SS_PalVals", Err.Number, Err.Description, Erl, _
             "i=" & i & "; k=" & k
End Function

Private Function SS_PalLabs() As Variant
    Dim base As Variant, tail As Variant, arr() As Variant
    Dim i As Long, n As Long, k As Long
    On Error GoTo ErrHandler

10  base = SS_PalLabsBase()
20  tail = SS_PalLabsTail()
30  n = (UBound(base) + 1) + DOC_SLOTS + (UBound(tail) + 1)
40  ReDim arr(0 To n - 1)
50  For i = LBound(base) To UBound(base)
60      arr(k) = base(i)
70      k = k + 1
80  Next i
90  For i = 1 To DOC_SLOTS
100     arr(k) = "医師"
110     k = k + 1
120 Next i
130 For i = LBound(tail) To UBound(tail)
140     arr(k) = tail(i)
150     k = k + 1
160 Next i
170 SS_PalLabs = arr
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SS_PalLabs", Err.Number, Err.Description, Erl, _
             "i=" & i & "; k=" & k
End Function

'--- 医師名スタンプの値を決める ---
'    既存パレットの該当セルに値があればそれを使う(実名を保持)。
'    無ければ "医師n"。コードに実名を持たないための仕組み。
Private Function SS_DoctorName(ByVal ws As Worksheet, ByVal slot As Long) As String
    Dim pal As Range, v As String, idx As Long
    On Error GoTo ErrHandler

10  SS_DoctorName = "医師" & slot
20  Set pal = PaletteRange(ws)
30  If pal Is Nothing Then Exit Function
40  idx = IDX_DOC_FIRST + slot - 1
50  If idx > pal.Cells.Count Then Exit Function
60  v = Trim$(CStr(pal.Cells(1, idx).Value))
70  If Len(v) > 0 Then SS_DoctorName = v
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SS_DoctorName", Err.Number, Err.Description, Erl, _
             "slot=" & slot & "; idx=" & idx
    SS_DoctorName = "医師" & slot
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
    SS_WorkSyms = Array(SYM_EARLY, SYM_EARLY_ALT, SYM_LATE, SYM_MID)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_WorkSyms", Err.Number, Err.Description, Erl, ""
End Function

'--- 色は ShiftCommon の色パレットを使う(v2.2でPrivate定義を廃止) ---


'====================== 1) パレット生成 ==========================
Public Sub ShiftSetup_パレット生成()
    Dim ws As Worksheet, dateRowNo As Long, palRow As Long
    Dim pal As Range, vals As Variant, labs As Variant
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  dateRowNo = DateRow(ws)
30  If dateRowNo = 0 Then
40      MsgBox "B列にシフト表の開始日(数式)が見つかりません。", vbExclamation
50      Exit Sub
60  End If

70  palRow = PaletteBodyRow(ws)
80  If palRow = 0 Or palRow + MARKER_OFFSET < 1 Then
90      MsgBox "パレットを置く行が足りません。" & vbCrLf & _
               "日付行(" & dateRowNo & "行)の直上に " & PALETTE_ROWS & _
               " 行必要です。", vbExclamation
100     Exit Sub
110 End If

    '--- 医師名を先に読むため、値の組み立ては書き換え前に行う ---
120 vals = SS_PalVals(ws)
130 labs = SS_PalLabs()
140 Set pal = ws.Cells(palRow, ws.Range(COL_FIRST & "1").Column) _
                .Resize(1, UBound(vals) + 1)

150 Application.EnableEvents = False
160 Application.ScreenUpdating = False

170 SS_パレット行書式 pal
180 SS_パレット値書込 pal, vals, labs
190 SS_パレット装飾 pal
200 SS_パレット見出し pal

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSetup_パレット生成", _
               "Built palette row " & palRow & " with " & (UBound(vals) + 1) & " cells"
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSetup_パレット生成", Err.Number, Err.Description, Erl, _
             "dateRow=" & dateRowNo & "; paletteRow=" & palRow
    MsgBox "パレット生成でエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'--- パレットが使う3行(マーカー/本体/ラベル)の書式を整える ---
Private Sub SS_パレット行書式(ByVal pal As Range)
    On Error GoTo ErrHandler

    '--- 本体行 ---
10  With pal
        .ClearContents
        .Interior.Pattern = xlNone
        .Font.Color = CLR_BLACK
        .Font.Bold = False
        .Font.size = FS_BODY
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Color = ClrBorder()
        .RowHeight = PAL_ROW_HEIGHT
20  End With

    '--- ★マーカー行(上) ---
30  With pal.Offset(MARKER_OFFSET, 0)
        .ClearContents
        .HorizontalAlignment = xlCenter
        .Font.Color = ClrMarker()
        .Font.Bold = True
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = ClrBorder()
40  End With

    '--- ラベル行(下) ---
50  With pal.Offset(LABEL_OFFSET, 0)
        .ClearContents
        .HorizontalAlignment = xlCenter
        .Font.size = FS_LABEL
        .Font.Color = ClrSubFg()
        .Borders(xlEdgeTop).LineStyle = xlContinuous
        .Borders(xlEdgeTop).Color = ClrBorder()
60  End With
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SS_パレット行書式", Err.Number, Err.Description, Erl, _
             "palette=" & pal.Address(False, False)
End Sub

'--- 本体行に値、ラベル行に説明を書く ---
Private Sub SS_パレット値書込(ByVal pal As Range, ByVal vals As Variant, ByVal labs As Variant)
    Dim i As Long
    On Error GoTo ErrHandler

10  For i = 1 To pal.Cells.Count
20      If Len(CStr(vals(i - 1))) > 0 Then pal.Cells(1, i).Value = vals(i - 1)
30      pal.Cells(1, i).Offset(LABEL_OFFSET, 0).Value = labs(i - 1)
40  Next i
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SS_パレット値書込", Err.Number, Err.Description, Erl, _
             "i=" & i & "; cells=" & pal.Cells.Count
End Sub

'--- ボタンの種類ごとに色と文字サイズを付ける ---
Private Sub SS_パレット装飾(ByVal pal As Range)
    Dim i As Long, fills As Variant
    On Error GoTo ErrHandler

10  fills = SS_FillColors()

    '--- 記号位置の定数がパレット幅を超えていないか ---
    '    医師名の範囲はラベルで判定するので、ここで見るのは
    '    位置固定のボタン(モード・背景色・記号)の最終位置まで。
15  If IDX_SYM_LAST > pal.Cells.Count Then
16      LogError MODULE_NAME, "SS_パレット装飾", 0, _
                 "パレットの幅が定数より狭いため装飾を中止", Erl, _
                 "cells=" & pal.Cells.Count & "; needed=" & IDX_SYM_LAST
17      Exit Sub
18  End If

    '--- モードセル(OFF/自動/戻す/出力/切替/色消)はグレー ---
20  For i = IDX_OFF To IDX_CLEARFILL
30      If i <= pal.Cells.Count Then
            With pal.Cells(1, i)
                .Interior.Color = ClrModeBg()
                .Font.Color = ClrModeFg()
            End With
40      End If
50  Next i

    '--- 背景色ボタン: 値は持たず背景色のみ ---
60  For i = IDX_FILL_FIRST To IDX_FILL_LAST
70      If i <= pal.Cells.Count Then pal.Cells(1, i).Interior.Color = fills(i - IDX_FILL_FIRST)
80  Next i

    '--- ○●▲ を少し大きく ---
90  For i = IDX_SYM_FIRST To IDX_SYM_LAST
100     If i <= pal.Cells.Count Then pal.Cells(1, i).Font.size = FS_SYMBOL
110 Next i

    '--- 医師名スタンプは色分けして区別 ---
    '    位置ではなくラベルで判定する。枠を増減しても追随する。
120 For i = 1 To pal.Cells.Count
130     If IsDoctorStamp(i) Then
140         With pal.Cells(1, i)
                .Interior.Color = ClrInputBg()
                .Font.Color = ClrDocFg()
                .Font.size = FS_BODY
150         End With
160     End If
170 Next i
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SS_パレット装飾", Err.Number, Err.Description, Erl, _
             "i=" & i & "; cells=" & pal.Cells.Count
End Sub

'--- A列の見出しと、初期モード(切替)の★を置く ---
Private Sub SS_パレット見出し(ByVal pal As Range)
    On Error GoTo ErrHandler

10  With pal.Cells(1, 1).Offset(0, -1)
        .Value = LBL_PALETTE
        .Font.Bold = True
        .HorizontalAlignment = xlRight
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = ClrBorder()
20  End With
30  pal.Cells(1, IDX_CYCLE).Offset(MARKER_OFFSET, 0).Value = MARKER_CHAR
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SS_パレット見出し", Err.Number, Err.Description, Erl, _
             "palette=" & pal.Address(False, False)
End Sub

'============ 2) 年月セル・祝日サマリーの数式セット ==============
'   A(日付行-1) = =DATE(1900,AG<行>,1)   年月表示
'   I(日付行-1) = 土日公休/祝日カウントの LET 数式
Public Sub ShiftSetup_ヘッダ数式()
    Dim ws As Worksheet, dateRowNo As Long, hRow As Long
    Dim f As String
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  dateRowNo = DateRow(ws)
30  If dateRowNo = 0 Then
40      MsgBox "B列にシフト表の開始日(数式)が見つかりません。", vbExclamation
50      Exit Sub
60  End If
70  hRow = HeaderRow(ws)
80  If hRow = 0 Then Exit Sub

    '--- パレットの3行と重ならないか確認 ---
90  If hRow >= PaletteBodyRow(ws) + MARKER_OFFSET And _
       hRow <= PaletteBodyRow(ws) + LABEL_OFFSET Then
100     MsgBox "年月行(" & hRow & "行)がパレットの3行と重なります。" & vbCrLf & _
               "上書きを避けるため、ヘッダ数式は書き込みませんでした。" & vbCrLf & vbCrLf & _
               "ShiftSurvey_シート構造調査 で行構成を確認してください。", vbExclamation
110     Exit Sub
120 End If

130 Application.EnableEvents = False

    '--- 年月セル(A列) ---
140 With ws.Cells(hRow, 1)
        .Formula = "=DATE(1900," & COL_MONTH & hRow & ",1)"
        .NumberFormatLocal = "[$-ja-JP]ge"".""  m""月"""
        .Font.Bold = True
        .Font.size = FS_TITLE
        .HorizontalAlignment = xlRight
150 End With

    '--- 年月シリアル(AG列・白文字で隠す) ---
160 With ws.Cells(hRow, ws.Range(COL_MONTH & "1").Column)
170     If Len(Trim$(CStr(.Value))) = 0 Then
180         .Value = (Year(Date) - 1900) * 12 + Month(Date)
190     End If
        .Font.Color = ClrHidden()
200 End With

    '--- 祝日サマリー ---
210 f = "=LET(d," & COL_FIRST & dateRowNo & ":" & COL_LAST & dateRowNo & _
        ",inM,--(MONTH(d)=MONTH(A" & hRow & "))" & _
        ",wk,SUMPRODUCT(inM*(WEEKDAY(d,2)>5))" & _
        ",hol,SUMPRODUCT(inM*(WEEKDAY(d,2)<6)*COUNTIF(" & SHT_HOLIDAY & "!$A:$A,d))" & _
        ",""土日公休""&wk&""回　祝日""&hol&""回　休みのトータル""&(wk+hol)&""回"")"
220 With ws.Cells(hRow, ws.Range(COL_HOL_SUMM & "1").Column)
        .Formula2 = f
        .Font.Italic = True
        .Font.size = FS_LABEL
        .HorizontalAlignment = xlLeft
230 End With

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

'========= 3) 集計行(医師数(診)/薬剤師出勤数/過不足)の数式 =========
'   位置: 備考行の NOTE_TO_DOC(2) 行下から3行
'     医師数(診)   = 医師名欄ブロックを列ごとに COUNTA
'     薬剤師出勤数 = 入力欄のうち区分が薬剤師の行の出勤記号を数える
'     過不足       = 薬剤師出勤数 - (医師数 + 必要出勤数の n)
Public Sub ShiftSetup_集計行数式()
    Dim ws As Worksheet, noteRow As Long, docRow As Long
    Dim topR As Long, botR As Long, docBlk As Range
    Dim firstCol As Long, lastCol As Long, c As Long, colL As String
    Dim blkTop As Long, blkBot As Long, noBlock As Boolean
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  noteRow = LabelRow(ws, LBL_NOTE)
30  If noteRow = 0 Then
40      MsgBox "A列に「" & LBL_NOTE & "」が見つかりません。" & vbCrLf & _
               "集計行の位置を決められないため中止しました。", vbExclamation
50      Exit Sub
60  End If
70  docRow = noteRow + NOTE_TO_DOC
80  topR = ShiftTopRow(ws)
90  botR = ShiftBottomRow(ws)
100 If topR = 0 Or botR = 0 Then
110     MsgBox "シフト入力欄を特定できないため中止しました。", vbExclamation
120     Exit Sub
130 End If

140 Set docBlk = DoctorBlock(ws)
150 noBlock = (docBlk Is Nothing)
160 If Not noBlock Then
170     blkTop = docBlk.Row
180     blkBot = docBlk.Row + docBlk.Rows.Count - 1
190 End If

200 firstCol = ws.Range(COL_FIRST & "1").Column
210 lastCol = ws.Range(COL_LAST & "1").Column

220 Application.EnableEvents = False
230 Application.ScreenUpdating = False

    '--- A列の見出し(空欄のときだけ補う) ---
240 SS_FillHeading ws, docRow, HEAD_DOC
250 SS_FillHeading ws, docRow + 1, HEAD_PHARM
260 SS_FillHeading ws, docRow + 2, HEAD_SHORT

270 For c = firstCol To lastCol
280     colL = ColLetter(c)
        '--- 医師数(診) ---
290     If Not noBlock Then
300         ws.Cells(docRow, c).Formula = _
                "=COUNTA(" & colL & blkTop & ":" & colL & blkBot & ")"
310     End If
        '--- 薬剤師出勤数 ---
320     ws.Cells(docRow + 1, c).Formula2 = SS_PharmFormula(colL, topR, botR)
        '--- 過不足 ---
330     ws.Cells(docRow + 2, c).Formula = _
            "=" & colL & (docRow + 1) & "-(" & colL & docRow & "+" & SS_RequiredN() & ")"
340 Next c

350 With ws.Range(ws.Cells(docRow, firstCol), ws.Cells(docRow + 2, lastCol))
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.size = FS_BODY
360 End With

370 If noBlock Then
380     MsgBox "医師名欄を特定できなかったため、" & vbCrLf & _
               "「" & HEAD_DOC & "」の数式だけは書き込みませんでした。" & vbCrLf & vbCrLf & _
               "ShiftSurvey_シート構造調査 で行構成を確認してください。", vbExclamation
390 End If

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSetup_集計行数式", _
               "Set aggregate rows " & docRow & "-" & (docRow + 2) & _
               "; doctorBlock=" & IIf(noBlock, "none", blkTop & "-" & blkBot)
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSetup_集計行数式", Err.Number, Err.Description, Erl, _
             "noteRow=" & noteRow & "; docRow=" & docRow & "; topRow=" & topR & _
             "; bottomRow=" & botR & "; c=" & c
    MsgBox "集計行のセットでエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'--- A列の見出しが空欄のときだけ書き込む ---
Private Sub SS_FillHeading(ByVal ws As Worksheet, ByVal r As Long, ByVal txt As String)
    On Error GoTo ErrHandler
10  If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then ws.Cells(r, 1).Value = txt
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SS_FillHeading", Err.Number, Err.Description, Erl, _
             "r=" & r & "; txt=" & txt
End Sub

'--- 必要出勤数の n を設定シートから読む数式片 ---
Private Function SS_RequiredN() As String
    On Error GoTo ErrHandler
10  SS_RequiredN = "IFERROR(INDEX(" & SHT_CFG & "!$L:$L,MATCH(""" & CFG_KEY_REQ & _
                   "*""," & SHT_CFG & "!$K:$K,0))," & REQ_FALLBACK & ")"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_RequiredN", Err.Number, Err.Description, Erl, ""
    SS_RequiredN = CStr(REQ_FALLBACK)
End Function

'--- 薬剤師出勤数の数式(区分が薬剤師の行だけ出勤記号を数える) ---
Private Function SS_PharmFormula(ByVal colL As String, _
                                 ByVal topR As Long, ByVal botR As Long) As String
    Dim syms As Variant, i As Long, symPart As String, rngRef As String
    On Error GoTo ErrHandler
10  rngRef = colL & topR & ":" & colL & botR
20  syms = SS_WorkSyms()
30  For i = LBound(syms) To UBound(syms)
40      If Len(symPart) > 0 Then symPart = symPart & "+"
50      symPart = symPart & "(" & rngRef & "=""" & syms(i) & """)"
60  Next i
70  SS_PharmFormula = "=SUMPRODUCT((IFERROR(INDEX(" & SHT_CFG & "!$B:$B,MATCH($A" & topR & _
                      ":$A" & botR & "," & SHT_CFG & "!$A:$A,0)),"""")=""" & KIND_PH & _
                      """)*(" & symPart & "))"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_PharmFormula", Err.Number, Err.Description, Erl, _
             "colL=" & colL & "; topR=" & topR & "; botR=" & botR
    SS_PharmFormula = ""
End Function

'============ 4) 集計列(AH:AL)の見出しと数式 ====================
Public Sub ShiftSetup_集計列数式()
    Dim ws As Worksheet, topR As Long, botR As Long, hdrRow As Long
    Dim docRow As Long, aggCol As Long, r As Long, i As Long
    Dim heads As Variant, rng As Range
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  topR = ShiftTopRow(ws)
30  botR = ShiftBottomRow(ws)
40  docRow = ShiftDocRow(ws)
50  If topR = 0 Or botR = 0 Or docRow = 0 Then
60      MsgBox "基準となる行が見つかりません。", vbExclamation
70      Exit Sub
80  End If

90  hdrRow = topR - 1
100 aggCol = ws.Range(COL_AGG & "1").Column
110 heads = SS_AggHeads()

120 Application.EnableEvents = False
130 Application.ScreenUpdating = False

    '--- 見出し行(空欄のときだけ書く。手で短縮した見出しを戻さない) ---
140 For i = 0 To UBound(heads)
150     With ws.Cells(hdrRow, aggCol + i)
            If Len(Trim$(CStr(.Value))) = 0 Then .Value = heads(i)
            .Font.Bold = True
            .Font.size = FS_LABEL
            .HorizontalAlignment = xlCenter
160     End With
170 Next i

    '--- 数式行(A列にスタッフ名がある行のみ) ---
180 For r = topR To botR
190     If Len(Trim$(CStr(ws.Cells(r, 1).Value))) = 0 Then
200         ws.Range(ws.Cells(r, aggCol), ws.Cells(r, aggCol + UBound(heads))).ClearContents
210     Else
220         ws.Cells(r, aggCol + 0).Formula = SS_CountFormula(r, SS_OffSyms())
230         ws.Cells(r, aggCol + 1).Formula = SS_CountFormula(r, Array(SYM_EARLY, SYM_EARLY_ALT))
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
        .Font.size = FS_BODY
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

'--- 医師が DOC_BUSY_COUNT 名の日の出勤数を数える数式 ---
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

'====================== 5) 名前付き範囲 ==========================
Public Sub ShiftSetup_名前付き範囲更新()
    Dim ws As Worksheet, topR As Long, botR As Long
    Dim palRow As Long, addr As String
    On Error GoTo ErrHandler

10  Set ws = ShiftSheet()
20  topR = ShiftTopRow(ws)
30  botR = ShiftBottomRow(ws)
40  palRow = PaletteBodyRow(ws)
50  If topR = 0 Or botR = 0 Or palRow = 0 Then
60      MsgBox "基準(B列の開始日の数式 / A列の " & LBL_NOTE & ")が見つかりません。", _
               vbExclamation
70      Exit Sub
80  End If

    '--- シフト入力範囲(動的: 下端は備考ラベルを追う) ---
    '    上端は再掲日付行の1行下で確定するため固定値でよい。
    '    下端は行の増減に追従させたいので INDEX/MATCH で解決する。
90  addr = SS_GridFormula(ws, topR)
100 SS_ReplaceName NM_SHIFT, addr

    '--- 医師名リスト範囲(静的) ---
    '    医師名欄は日付行からの相対位置で決まるが、日付行は
    '    「B列で最初に見つかる数式かつ日付のセル」であり、
    '    ワークシート数式では表せない。そのため実アドレスで固定する。
    '    行構成を変えたら、この初期設定を実行し直して貼り替えること。
102 SS_ReplaceName NM_DOCLIST, SS_BlockFormula(ws, DoctorBlock(ws))

    '--- 備考行範囲(動的: 1行固定) ---
104 SS_ReplaceName NM_NOTEROW, SS_NoteFormula(ws)

    '--- パレット本体(動的: A列の見出しを追い、幅はラベル行から数える) ---
110 SS_ReplaceName NM_PALETTE, SS_PaletteFormula(ws)

    LogSuccess MODULE_NAME, "ShiftSetup_名前付き範囲更新", _
               "shift=" & COL_FIRST & topR & ":" & COL_LAST & botR & "(dynamic)" & _
               ", paletteRow=" & palRow & "(dynamic)"
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSetup_名前付き範囲更新", Err.Number, Err.Description, Erl, _
             "topRow=" & topR & "; bottomRow=" & botR & "; paletteRow=" & palRow & _
             "; addr=" & addr
    MsgBox "名前付き範囲の更新でエラーが発生しました: " & Err.Description, vbExclamation
End Sub

'--- 入力欄の動的数式を組み立てる ---
'    =シフト!$B$14:INDEX(シフト!$B:$AF,MATCH("備考*",シフト!$A:$A,0)-NOTE_GAP,列数)
'    NOTE_GAP は ShiftCommon の DOC_GAP - NOTE_TO_DOC で導出する。
'    (下端 = 医師数行 - DOC_GAP かつ 医師数行 = 備考行 + NOTE_TO_DOC より
'     下端 = 備考行 - (DOC_GAP - NOTE_TO_DOC))
Private Function SS_GridFormula(ByVal ws As Worksheet, ByVal topR As Long) As String
    Dim nCol As Long, gap As Long   ' gap = NOTE_GAP (ShiftCommon)
    On Error GoTo ErrHandler

10  nCol = ws.Range(COL_LAST & "1").Column - ws.Range(COL_FIRST & "1").Column + 1
20  gap = NOTE_GAP
30  SS_GridFormula = "=" & SS_Q(ws.Name) & "!$" & COL_FIRST & "$" & topR & _
                     ":INDEX(" & SS_Q(ws.Name) & "!$" & COL_FIRST & ":$" & COL_LAST & _
                     ",MATCH(""" & LBL_NOTE & "*""," & SS_Q(ws.Name) & "!$A:$A,0)-" & _
                     gap & "," & nCol & ")"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_GridFormula", Err.Number, Err.Description, Erl, _
             "topRow=" & topR & "; nCol=" & nCol
    SS_GridFormula = ""
End Function

'--- パレット本体(1行)の動的数式を組み立てる ---
'    行  = A列の「シフトパレット」見出しの行(パレット本体行と同じ行)
'    幅  = ラベル行の入力済みセル数
'          本体行は背景色ボタンなどが空なので数えられない。
'          ラベル行は必ず埋まるため、こちらを幅の基準にする。
'          (ShiftCommon.PaletteRange のフォールバックと同じ考え方)
Private Function SS_PaletteFormula(ByVal ws As Worksheet) As String
    Dim q As String, m As String, colOff As Long
    On Error GoTo ErrHandler

10  q = SS_Q(ws.Name)
20  colOff = ws.Range(COL_FIRST & "1").Column - 1
30  m = "MATCH(""" & LBL_PALETTE & "*""," & q & "!$A:$A,0)-1"
40  SS_PaletteFormula = "=OFFSET(" & q & "!$A$1," & m & "," & colOff & ",1," & _
                        "COUNTA(OFFSET(" & q & "!$A$1," & m & "+" & LABEL_OFFSET & _
                        "," & colOff & ",1," & PAL_SCAN_COLS & ")))"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_PaletteFormula", Err.Number, Err.Description, Erl, _
             "colOffset=" & colOff
    SS_PaletteFormula = ""
End Function

'--- 備考行(1行)の動的数式を組み立てる ---
Private Function SS_NoteFormula(ByVal ws As Worksheet) As String
    Dim q As String, m As String
    On Error GoTo ErrHandler

10  q = SS_Q(ws.Name)
20  m = "MATCH(""" & LBL_NOTE & "*""," & q & "!$A:$A,0)"
30  SS_NoteFormula = "=INDEX(" & q & "!$" & COL_FIRST & ":$" & COL_FIRST & "," & m & ")" & _
                     ":INDEX(" & q & "!$" & COL_LAST & ":$" & COL_LAST & "," & m & ")"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_NoteFormula", Err.Number, Err.Description, Erl, ""
    SS_NoteFormula = ""
End Function

'--- 既存 Range から静的アドレス数式を組み立てる(医師名欄など固定ブロック用) ---
Private Function SS_BlockFormula(ByVal ws As Worksheet, ByVal rng As Range) As String
    On Error GoTo ErrHandler

10  If rng Is Nothing Then Exit Function
20  SS_BlockFormula = "=" & SS_Q(ws.Name) & "!" & rng.Address(True, True)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_BlockFormula", Err.Number, Err.Description, Erl, ""
    SS_BlockFormula = ""
End Function

'--- シート名を数式で使える形に整える(空白等があれば単引用符で囲む) ---
Private Function SS_Q(ByVal nm As String) As String
    On Error GoTo ErrHandler

10  If nm Like "*[!0-9A-Za-z_]*" Then
20      SS_Q = "'" & Replace(nm, "'", "''") & "'"
30  Else
40      SS_Q = nm
50  End If
    Exit Function
ErrHandler:
    SS_Q = nm
End Function

'--- 名前付き範囲を貼り替える ---
Private Sub SS_ReplaceName(ByVal nm As String, ByVal refersTo As String)
    On Error GoTo ErrHandler
    ' 既存の名前は無い場合もあるため、削除失敗は正常系として扱う
10  On Error Resume Next
20  ThisWorkbook.Names(nm).Delete
30  On Error GoTo ErrHandler
40  ThisWorkbook.Names.Add Name:=nm, refersTo:=refersTo
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SS_ReplaceName", Err.Number, Err.Description, Erl, _
             "name=" & nm & "; refersTo=" & refersTo
End Sub

'--- 完了ダイアログ用: 医師名欄ブロックの表示文字列 ---
Private Function SS_BlockText(ByVal ws As Worksheet) As String
    Dim b As Range
    On Error GoTo ErrHandler
10  Set b = DoctorBlock(ws)
20  If b Is Nothing Then
30      SS_BlockText = "(未検出)"
40  Else
50      SS_BlockText = b.Row & "-" & (b.Row + b.Rows.Count - 1) & " 行"
60  End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SS_BlockText", Err.Number, Err.Description, Erl, ""
    SS_BlockText = "(不明)"
End Function

'====================== 一括実行 =================================
Public Sub ShiftSetup_初期設定実行()
    Dim ws As Worksheet, msg As String
    Dim dateRowNo As Long, topR As Long, botR As Long
    On Error GoTo ErrHandler

    '--- 0) 依存シートを先に用意する(存在しないものだけ生成) ---
10  ShiftSchema_不足シート生成

20  Set ws = ShiftSheet()
30  dateRowNo = DateRow(ws)
40  topR = ShiftTopRow(ws)
50  botR = ShiftBottomRow(ws)
60  If dateRowNo = 0 Or topR = 0 Or botR = 0 Then
70      MsgBox "基準が見つからないため中止しました。" & vbCrLf & _
               "必要な基準: B列の開始日の数式 / A列の " & LBL_NOTE, _
               vbExclamation, "初期設定"
80      Exit Sub
90  End If

100 Application.Calculation = xlCalculationManual

110 ShiftSetup_パレット生成
120 ShiftSetup_ヘッダ数式
130 ShiftSetup_集計行数式
140 ShiftSetup_集計列数式
150 ShiftSetup_名前付き範囲更新

160 Application.Calculation = xlCalculationAutomatic
170 Application.CalculateFull

180 msg = "初期設定が完了しました。" & vbCrLf & vbCrLf & _
          "日付行(開始日)  : " & dateRowNo & " 行" & vbCrLf & _
          "パレット        : " & (PaletteBodyRow(ws) + MARKER_OFFSET) & "-" & _
                                 (PaletteBodyRow(ws) + LABEL_OFFSET) & " 行" & vbCrLf & _
          "年月・タイトル  : " & HeaderRow(ws) & " 行" & vbCrLf & _
          "医師名欄        : " & SS_BlockText(ws) & vbCrLf & _
          "シフト入力範囲  : " & COL_FIRST & topR & ":" & COL_LAST & botR & vbCrLf & _
          "集計行          : " & ShiftDocRow(ws) & "-" & (ShiftDocRow(ws) + 2) & " 行" & vbCrLf & _
          "集計列          : " & COL_AGG & "-" & COL_AGG_END
190 MsgBox msg, vbInformation, "初期設定"

CleanUp:
    On Error Resume Next
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSetup_初期設定実行", _
               "dateRow=" & dateRowNo & ", input=" & _
               COL_FIRST & topR & ":" & COL_LAST & botR
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSetup_初期設定実行", Err.Number, Err.Description, Erl, _
             "dateRow=" & dateRowNo & "; topRow=" & topR & "; bottomRow=" & botR
    MsgBox "初期設定でエラーが発生しました: " & Err.Description, vbExclamation, "初期設定"
    Resume CleanUp
End Sub
