Attribute VB_Name = "ShiftSchema"
Option Explicit
'==================================================================
'  シフト表 シート生成モジュール ＜標準モジュール ShiftSchema v1.5＞
'  2026-08-27
'
'  v1.5: 全体設定に「不足を埋めるときの連勤上限の上乗せ」を追加。
'
'  v1.4: 全体設定に「混雑日の遅番(▲) 最低人数/日」を追加。
'
'  v1.3: 全体設定に「事務員の早番(○) 人数/日」を追加。既存シートには
'        空欄の行だけが足されるため、再実行しても設定値は消えない。
'
'  目的:
'    自動作成設定 / 祝日マスタ / シフト変更ログ が存在しない場合に
'    見出し・ラベル・初期設定値まで含めて生成する。
'
'  設計方針:
'    ・冪等(べきとう)。既にあるシートは作り直さない。
'      既存シートに対しては「空欄の見出しだけ補う」動作にとどめる。
'      → 既に運用中のデータを壊さない
'    ・セルの位置は ShiftCommon のスキーマ定数のみを参照する。
'      本モジュールに番地リテラルは書かない。
'    ・ShiftSetup_初期設定実行 から最初に呼ばれることを想定。
'
'  入口:
'    ShiftSchema_不足シート生成   … 3シートまとめて(存在しないものだけ)
'    ShiftSchema_自動作成設定生成 … 個別
'    ShiftSchema_祝日マスタ生成   … 個別
'    ShiftSchema_変更ログ生成     … 個別
'    ShiftSchema_祝日マスタ取込   … 内閣府の CSV を Power Query で取り込む
'
'  v1.2 変更:
'   ・祝日マスタの取り込みを Power Query で行えるようにした。
'     手で CSV を貼る運用も従来どおり使える。
'==================================================================
Private Const MODULE_NAME As String = "ShiftSchema"

'--- 祝日マスタの取り込み(Power Query) ---
'    内閣府が公開している「国民の祝日」CSV。Shift-JIS(932)・2列。
'    URL が変わったらここだけ直す。
Private Const HOL_SOURCE_URL  As String = _
    "https://www8.cao.go.jp/chosei/shukujitsu/syukujitsu.csv"
Private Const HOL_SOURCE_CP   As Long = 932    ' CSV の文字コード(Shift-JIS)
Private Const HOL_QUERY_NAME  As String = "祝日マスタ取込"
Private Const HOL_TABLE_NAME  As String = "祝日一覧"
'    OLEDB 接続文字列のひな形。Power Query の出力を表として貼るのに使う
Private Const PQ_CONN_HEAD    As String = _
    "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location="
Private Const PQ_CONN_TAIL    As String = ";Extended Properties="""""

'--- 書式 ---
Private Const FS_TITLE  As Long = 14
Private Const FS_NOTE   As Long = 9
Private Const FS_HEAD   As Long = 10

'--- 入力規則(ドロップダウン)を設定する行数 ---
Private Const MAX_VALIDATION_ROWS As Long = 200

'--- 列幅 ---
Private Const W_NAME    As Double = 24
Private Const W_NORMAL  As Double = 14
Private Const W_MEMO    As Double = 34
Private Const W_SETKEY  As Double = 30
Private Const W_SETVAL  As Double = 22
Private Const W_DATE    As Double = 13

'==================== メンバー表の見出し =========================
Private Function SC_CfgHeads() As Variant
    On Error GoTo ErrHandler
    SC_CfgHeads = Array("氏名", "区分", "休業(○=休業中)", "勤務ルール", _
                        "固定曜日(例:月火金土)", "週勤務日数", _
                        "月間休日数(空欄=土日祝と同数)", "遅番・遅半 可否", "備考")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_CfgHeads", Err.Number, Err.Description, Erl, ""
End Function

'--- 全体設定のラベルと既定値(K列 / L列) ---
Private Function SC_SetKeys() As Variant
    On Error GoTo ErrHandler
    SC_SetKeys = Array("全体設定(自動作成ルール)", _
                       "早番(○) 人数/日(基本1・最大2)", _
                       "遅番(▲) 最低人数/日(誤差-1)", _
                       "連勤の上限(日)", _
                       "連休の上限(日)", _
                       "週の基本休日数", _
                       "必要出勤数(医師数+n)の n", _
                       "ノルマ外の休み記号(カンマ区切り)", _
                       "週の定義(固定)", _
                       "事務員の2人目以降の記号", _
                       "事務員の早番(○) 人数/日", _
                       "混雑日_医師" & DOC_BUSY_N & "名_の遅番(▲) 最低人数/日 0=通常と同じ", _
                       "不足を埋めるときの連勤上限の上乗せ(日) 0=上乗せなし")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_SetKeys", Err.Number, Err.Description, Erl, ""
End Function

Private Function SC_SetVals() As Variant
    On Error GoTo ErrHandler
    '  1件目は見出しなので値なし("値" を入れる)
    SC_SetVals = Array("値", 1, 3, 3, 3, 2, 1, "有休", _
                       "日曜始まり・土曜終わり", "●", CLERK_EARLY_DEFAULT, 0, 0)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_SetVals", Err.Number, Err.Description, Erl, ""
End Function

'--- 区分・勤務ルール・可否の選択肢(入力規則に使う) ---
Private Function SC_KindList() As String
    On Error GoTo ErrHandler
    SC_KindList = "薬剤師,事務員"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_KindList", Err.Number, Err.Description, Erl, ""
End Function
Private Function SC_RuleList() As String
    On Error GoTo ErrHandler
    SC_RuleList = "通常,固定曜日,週N日,手動"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_RuleList", Err.Number, Err.Description, Erl, ""
End Function
Private Function SC_YesNoList() As String
    On Error GoTo ErrHandler
    SC_YesNoList = "可,不可"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_YesNoList", Err.Number, Err.Description, Erl, ""
End Function

'--- 祝日マスタの見出し ---
Private Function SC_HolHeads() As Variant
    On Error GoTo ErrHandler
    SC_HolHeads = Array("日付", "名称")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_HolHeads", Err.Number, Err.Description, Erl, ""
End Function

'--- 変更ログの見出し ---
Private Function SC_LogHeads() As Variant
    On Error GoTo ErrHandler
    SC_LogHeads = Array("日時", "セル", "変更前", "変更後", "ユーザー", "備考")
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_LogHeads", Err.Number, Err.Description, Erl, ""
End Function

'--- 色は ShiftCommon の色パレットを使う(v1.1でPrivate定義を廃止) ---


'==================================================================
' 入口: 不足しているシートだけ生成する
'==================================================================
Public Sub ShiftSchema_不足シート生成()
    Dim made As String, skipped As String
    On Error GoTo ErrHandler

10  Application.ScreenUpdating = False

20  If SheetExists(SHT_CFG) Then
30      skipped = skipped & vbCrLf & "　・" & SHT_CFG
40  Else
50      ShiftSchema_自動作成設定生成
60      made = made & vbCrLf & "　・" & SHT_CFG
70  End If

80  If SheetExists(SHT_HOLIDAY) Then
90      skipped = skipped & vbCrLf & "　・" & SHT_HOLIDAY
100 Else
110     ShiftSchema_祝日マスタ生成
120     made = made & vbCrLf & "　・" & SHT_HOLIDAY
130 End If

140 If SheetExists(SHT_LOG) Then
150     skipped = skipped & vbCrLf & "　・" & SHT_LOG
160 Else
170     ShiftSchema_変更ログ生成
180     made = made & vbCrLf & "　・" & SHT_LOG
190 End If

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSchema_不足シート生成", _
               "created=[" & Replace(Replace(made, vbCrLf, ""), "　・", " ") & "]"
200 MsgBox "シートの確認が完了しました。" & vbCrLf & vbCrLf & _
           "新しく作成:" & IIf(Len(made) = 0, " (なし)", made) & vbCrLf & vbCrLf & _
           "既にあったため変更なし:" & IIf(Len(skipped) = 0, " (なし)", skipped), _
           vbInformation, "シート生成"
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSchema_不足シート生成", Err.Number, Err.Description, Erl, _
             "made=" & made
    MsgBox "シート生成でエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'==================================================================
' 1) 自動作成設定
'==================================================================
Public Sub ShiftSchema_自動作成設定生成()
    Dim ws As Worksheet, heads As Variant, keys As Variant, vals As Variant
    Dim i As Long, r As Long, isNew As Boolean
    On Error GoTo ErrHandler

10  Set ws = SC_GetOrAdd(SHT_CFG, isNew)
20  If ws Is Nothing Then Exit Sub

30  Application.ScreenUpdating = False

    '--- タイトルと説明(空欄のときだけ) ---
40  SC_SetIfBlank ws.Cells(1, 1), "シフト自動作成 設定"
50  With ws.Cells(1, 1)
        .Font.Bold = True
        .Font.size = FS_TITLE
60  End With
70  SC_SetIfBlank ws.Cells(2, 1), _
        "希望休はシフト表に「希休」スタンプで入力してください。" & _
        "実行時、入力済みのセルはすべて保持され、空白セルのみ自動で埋まります。"
80  SC_SetIfBlank ws.Cells(3, 1), _
        "派遣スタッフの行をシフト表に追加したら、この表にも氏名を追加し" & _
        "「勤務ルール=手動」を設定してください" & _
        "(自動入力の対象外・入力済みの出勤は人数計算に反映されます)。"
90  With ws.Range(ws.Cells(2, 1), ws.Cells(3, 1))
        .Font.size = FS_NOTE
        .Font.Color = ClrSubFg()
100 End With

    '--- メンバー表の見出し ---
110 heads = SC_CfgHeads()
120 For i = LBound(heads) To UBound(heads)
130     SC_SetIfBlank ws.Cells(CFG_HDR_ROW, CFG_COL_NAME + i), heads(i)
140 Next i
150 SC_StyleHeader ws.Range(ws.Cells(CFG_HDR_ROW, CFG_COL_NAME), _
                            ws.Cells(CFG_HDR_ROW, CFG_COL_MEMO))

    '--- 全体設定(K列ラベル / L列の値) ---
160 keys = SC_SetKeys()
170 vals = SC_SetVals()
180 For i = LBound(keys) To UBound(keys)
190     r = CFG_SET_ROW + i
200     SC_SetIfBlank ws.Cells(r, CFG_COL_SETK), keys(i)
210     SC_SetIfBlank ws.Cells(r, CFG_COL_SETV), vals(i)
220 Next i
230 SC_StyleHeader ws.Range(ws.Cells(CFG_SET_ROW, CFG_COL_SETK), _
                            ws.Cells(CFG_SET_ROW, CFG_COL_SETV))
240 With ws.Range(ws.Cells(CFG_SET_ROW + 1, CFG_COL_SETV), _
                  ws.Cells(CFG_SET_ROW + UBound(keys), CFG_COL_SETV))
        .Interior.Color = ClrInputBg()
        .HorizontalAlignment = xlCenter
250 End With

    '--- 入力規則(新規作成時のみ。既存シートには触らない) ---
260 If isNew Then
270     SC_AddList ws, CFG_COL_KIND, SC_KindList()
280     SC_AddList ws, CFG_COL_RULE, SC_RuleList()
290     SC_AddList ws, CFG_COL_LATE, SC_YesNoList()
300     SC_AddList ws, CFG_COL_CLOSED, "○"
310 End If

    '--- 列幅 ---
320 ws.Columns(CFG_COL_NAME).ColumnWidth = W_NAME
330 For i = CFG_COL_KIND To CFG_COL_LATE
340     ws.Columns(i).ColumnWidth = W_NORMAL
350 Next i
360 ws.Columns(CFG_COL_MEMO).ColumnWidth = W_MEMO
370 ws.Columns(CFG_COL_SETK).ColumnWidth = W_SETKEY
380 ws.Columns(CFG_COL_SETV).ColumnWidth = W_SETVAL

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSchema_自動作成設定生成", _
               "sheet=" & SHT_CFG & "; isNew=" & isNew
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSchema_自動作成設定生成", Err.Number, Err.Description, Erl, _
             "i=" & i & "; r=" & r
    MsgBox SHT_CFG & " の生成でエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'==================================================================
' 2) 祝日マスタ
'==================================================================
Public Sub ShiftSchema_祝日マスタ生成()
    Dim ws As Worksheet, heads As Variant, i As Long, isNew As Boolean
    On Error GoTo ErrHandler

10  Set ws = SC_GetOrAdd(SHT_HOLIDAY, isNew)
20  If ws Is Nothing Then Exit Sub

30  Application.ScreenUpdating = False

40  heads = SC_HolHeads()
50  For i = LBound(heads) To UBound(heads)
60      SC_SetIfBlank ws.Cells(HOL_HDR_ROW, HOL_COL_DATE + i), heads(i)
70  Next i
80  SC_StyleHeader ws.Range(ws.Cells(HOL_HDR_ROW, HOL_COL_DATE), _
                            ws.Cells(HOL_HDR_ROW, HOL_COL_NAME))

    '--- 日付列の表示形式 ---
90  ws.Columns(HOL_COL_DATE).NumberFormatLocal = "yyyy/m/d"
100 ws.Columns(HOL_COL_DATE).ColumnWidth = W_DATE
110 ws.Columns(HOL_COL_NAME).ColumnWidth = W_NAME

    '--- 使い方メモ(新規作成時のみ) ---
120 If isNew Then
130     With ws.Cells(HOL_HDR_ROW, HOL_COL_NAME + 2)
            .Value = "ShiftSchema_祝日マスタ取込 を実行すると、内閣府の" & _
                     "「国民の祝日」CSV を取り込みます。" & _
                     "手で A/B 列に貼り付けても構いません。" & _
                     "シフト表の祝日カウントは A 列を参照します。"
            .Font.size = FS_NOTE
            .Font.Color = ClrSubFg()
140     End With
150 End If

    '--- 中身が空なら取り込みを勧める(外部から取得するので必ず尋ねる) ---
160 If SC_祝日データ行数(ws) = 0 And SC_PQ使用可() Then
170     If MsgBox(SHT_HOLIDAY & " が空です。" & vbCrLf & _
                  "内閣府の「国民の祝日」CSV を取り込みますか?" & vbCrLf & vbCrLf & _
                  "取得元: " & HOL_SOURCE_URL & vbCrLf & _
                  "_インターネットに接続します_", _
                  vbYesNo + vbQuestion, SHT_HOLIDAY) = vbYes Then
180         ShiftSchema_祝日マスタ取込
190     End If
200 End If

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSchema_祝日マスタ生成", _
               "sheet=" & SHT_HOLIDAY & "; isNew=" & isNew
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSchema_祝日マスタ生成", Err.Number, Err.Description, Erl, _
             "i=" & i
    MsgBox SHT_HOLIDAY & " の生成でエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'==================================================================
' 祝日マスタの取り込み (Power Query)
'   内閣府の CSV を取り込んで A/B 列に展開する。
'   以後は表を右クリック →「更新」で最新に貼り替えられる。
'
'   Power Query は Excel 2016 以降の機能。2013 以前では使えないので、
'   その場合は従来どおり手で貼り付ける運用に戻す(判定は SC_PQ使用可)。
'==================================================================
Public Sub ShiftSchema_祝日マスタ取込()
    Dim ws As Worksheet, lo As ListObject, isNew As Boolean
    Dim n As Long
    On Error GoTo ErrHandler

10  If Not SC_PQ使用可() Then
20      MsgBox "この Excel では Power Query を使えません_2016 以降が必要_。" & vbCrLf & _
               "内閣府の CSV を " & SHT_HOLIDAY & " の A/B 列に貼り付けてください。", _
               vbExclamation, SHT_HOLIDAY
30      Exit Sub
40  End If

50  Set ws = SC_GetOrAdd(SHT_HOLIDAY, isNew)
60  If ws Is Nothing Then Exit Sub

    '--- 既に取り込み済みなら、作り直さず更新するだけ ---
70  Set lo = SC_祝日テーブル(ws)
80  If lo Is Nothing Then
        '--- 手で貼った祝日が既にある場合は、消す前に確認する ---
90      If SC_祝日データ行数(ws) > 0 Then
100         If MsgBox(SHT_HOLIDAY & " の A/B 列に既にデータがあります。" & vbCrLf & _
                      "Power Query の取り込みに切り替えると、この内容は" & vbCrLf & _
                      "取得した祝日で置き換わります。" & vbCrLf & vbCrLf & _
                      "切り替えますか?", vbYesNo + vbExclamation, SHT_HOLIDAY) <> vbYes Then
110             Exit Sub
120         End If
130     End If
140 End If

150 Application.ScreenUpdating = False
160 SC_祝日クエリを作る
170 If lo Is Nothing Then Set lo = SC_祝日テーブルを作る(ws)
180 If lo Is Nothing Then GoTo CleanUp
190 lo.QueryTable.Refresh BackgroundQuery:=False
200 n = lo.ListRows.Count
210 ws.Columns(HOL_COL_DATE).NumberFormatLocal = "yyyy/m/d"

220 MsgBox "祝日を " & n & " 件取り込みました。" & vbCrLf & vbCrLf & _
           "以後は表の中で右クリック →「更新」で最新にできます。" & vbCrLf & _
           "取得元: " & HOL_SOURCE_URL, vbInformation, SHT_HOLIDAY

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSchema_祝日マスタ取込", _
               "Loaded " & n & " holidays from " & HOL_SOURCE_URL
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSchema_祝日マスタ取込", Err.Number, Err.Description, Erl, _
             "url=" & HOL_SOURCE_URL & "; rows=" & n
    MsgBox "祝日の取り込みでエラーが発生しました:" & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "ネットワークに繋がらない場合や、外部データの取得が" & vbCrLf & _
           "許可されていない場合に起こります。" & vbCrLf & _
           "その場合は CSV を A/B 列に手で貼り付けてください。", _
           vbExclamation, SHT_HOLIDAY
    Resume CleanUp
End Sub

'--- この Excel で Power Query が使えるか ---
'    Workbook.Queries は Excel 2016 以降にしかない。参照できるかで判定する。
'    古い Excel でもコンパイルが通るよう Object 経由で呼ぶ(遅延バインド)。
Private Function SC_PQ使用可() As Boolean
    Dim wb As Object, n As Long
    On Error GoTo ErrHandler

10  Set wb = ThisWorkbook
20  n = wb.Queries.Count
30  SC_PQ使用可 = True
    Exit Function
ErrHandler:
    '--- Excel 2013 以前。想定内なので記録だけして False を返す ---
    LogError MODULE_NAME, "SC_PQ使用可", Err.Number, Err.Description, Erl, _
             "Power Query unavailable"
    SC_PQ使用可 = False
End Function

'--- 取り込みクエリを作る(同名があれば作り直す) ---
Private Sub SC_祝日クエリを作る()
    Dim wb As Object
    On Error GoTo ErrHandler

10  Set wb = ThisWorkbook
    '--- 同名クエリが無い場合の削除失敗は正常系 ---
20  On Error Resume Next
30  wb.Queries(HOL_QUERY_NAME).Delete
40  On Error GoTo ErrHandler
    '--- Name は VBA の予約語。名前付き引数を避けて位置指定で渡す ---
50  wb.Queries.Add HOL_QUERY_NAME, SC_祝日M式()

    LogSuccess MODULE_NAME, "SC_祝日クエリを作る", "query=" & HOL_QUERY_NAME
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SC_祝日クエリを作る", Err.Number, Err.Description, Erl, _
             "query=" & HOL_QUERY_NAME
End Sub

'--- 取り込みクエリの M 式 ---
'    列名は SC_HolHeads と揃える。見出しを変えたら取り込みも追随する。
Private Function SC_祝日M式() As String
    '--- 変数名に cDate は使わない。VBA の組み込み関数 CDate と衝突する ---
    Dim heads As Variant, colDate As String, colName As String
    On Error GoTo ErrHandler

10  heads = SC_HolHeads()
20  colDate = CStr(heads(LBound(heads)))
30  colName = CStr(heads(LBound(heads) + 1))

40  SC_祝日M式 = _
        "let" & vbLf & _
        "  Src = Csv.Document(Web.Contents(""" & HOL_SOURCE_URL & """)," & _
            "[Delimiter="","", Columns=2, Encoding=" & HOL_SOURCE_CP & _
            ", QuoteStyle=QuoteStyle.None])," & vbLf & _
        "  Hdr = Table.PromoteHeaders(Src, [PromoteAllScalars=true])," & vbLf & _
        "  Cols = Table.ColumnNames(Hdr)," & vbLf & _
        "  Ren = Table.RenameColumns(Hdr, {{Cols{0}, """ & colDate & """}, " & _
            "{Cols{1}, """ & colName & """}})," & vbLf & _
        "  Typed = Table.TransformColumnTypes(Ren, {{""" & colDate & """, type date}, " & _
            "{""" & colName & """, type text}})," & vbLf & _
        "  Clean = Table.SelectRows(Typed, each [" & colDate & "] <> null)" & vbLf & _
        "in" & vbLf & _
        "  Clean"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_祝日M式", Err.Number, Err.Description, Erl, ""
    SC_祝日M式 = ""
End Function

'--- クエリの出力先テーブルを作る(A1 起点) ---
Private Function SC_祝日テーブルを作る(ByVal ws As Worksheet) As ListObject
    Dim lo As ListObject, lastR As Long
    On Error GoTo ErrHandler

    '--- 貼り替え先を空にしておく(見出しはクエリが書く) ---
10  lastR = ws.Cells(ws.Rows.Count, HOL_COL_DATE).End(xlUp).Row
20  If lastR >= HOL_HDR_ROW Then
30      ws.Range(ws.Cells(HOL_HDR_ROW, HOL_COL_DATE), _
                 ws.Cells(lastR, HOL_COL_NAME)).ClearContents
40  End If

50  Set lo = ws.ListObjects.Add( _
                SourceType:=0, _
                Source:=PQ_CONN_HEAD & HOL_QUERY_NAME & PQ_CONN_TAIL, _
                Destination:=ws.Cells(HOL_HDR_ROW, HOL_COL_DATE))
60  lo.Name = HOL_TABLE_NAME
70  With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = "SELECT * FROM [" & HOL_QUERY_NAME & "]"
        .RowNumbers = False
        .PreserveFormatting = True
        .RefreshStyle = xlInsertDeleteCells
        .AdjustColumnWidth = False
        .BackgroundQuery = False
80  End With
90  Set SC_祝日テーブルを作る = lo
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_祝日テーブルを作る", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; lastRow=" & lastR
    Set SC_祝日テーブルを作る = Nothing
End Function

'--- 既にある取り込みテーブルを返す(無ければ Nothing) ---
Private Function SC_祝日テーブル(ByVal ws As Worksheet) As ListObject
    Dim lo As ListObject
    On Error GoTo ErrHandler

10  For Each lo In ws.ListObjects
20      If StrComp(lo.Name, HOL_TABLE_NAME, vbTextCompare) = 0 Then
30          Set SC_祝日テーブル = lo
40          Exit Function
50      End If
60  Next lo
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_祝日テーブル", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name
    Set SC_祝日テーブル = Nothing
End Function

'--- 手で入っている祝日の件数(見出し行を除く) ---
Private Function SC_祝日データ行数(ByVal ws As Worksheet) As Long
    Dim lastR As Long
    On Error GoTo ErrHandler

10  lastR = ws.Cells(ws.Rows.Count, HOL_COL_DATE).End(xlUp).Row
20  If lastR > HOL_HDR_ROW Then SC_祝日データ行数 = lastR - HOL_HDR_ROW
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SC_祝日データ行数", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name
    SC_祝日データ行数 = 0
End Function

'==================================================================
' 3) シフト変更ログ
'==================================================================
Public Sub ShiftSchema_変更ログ生成()
    Dim ws As Worksheet, heads As Variant, i As Long, isNew As Boolean
    On Error GoTo ErrHandler

10  Set ws = SC_GetOrAdd(SHT_LOG, isNew)
20  If ws Is Nothing Then Exit Sub

30  Application.ScreenUpdating = False

40  heads = SC_LogHeads()
50  For i = LBound(heads) To UBound(heads)
60      SC_SetIfBlank ws.Cells(LOG_HDR_ROW, LOG_COL_TIME + i), heads(i)
70  Next i
80  SC_StyleHeader ws.Range(ws.Cells(LOG_HDR_ROW, LOG_COL_TIME), _
                            ws.Cells(LOG_HDR_ROW, LOG_COL_NOTE))

90  ws.Columns(LOG_COL_TIME).NumberFormatLocal = "yyyy/m/d hh:mm:ss"
100 ws.Columns(LOG_COL_TIME).ColumnWidth = 20
110 ws.Columns(LOG_COL_ADDR).ColumnWidth = 12
120 ws.Columns(LOG_COL_BEFORE).ColumnWidth = W_NORMAL
130 ws.Columns(LOG_COL_AFTER).ColumnWidth = W_NORMAL
140 ws.Columns(LOG_COL_USER).ColumnWidth = 18
150 ws.Columns(LOG_COL_NOTE).ColumnWidth = W_MEMO

    '--- 見出し行の固定(新規作成時のみ) ---
160 If isNew Then
170     ws.Activate
180     ws.Rows(LOG_FIRST_ROW).Select
190     ActiveWindow.FreezePanes = True
200 End If

CleanUp:
    On Error Resume Next
    Application.ScreenUpdating = True
    On Error GoTo 0
    LogSuccess MODULE_NAME, "ShiftSchema_変更ログ生成", _
               "sheet=" & SHT_LOG & "; isNew=" & isNew
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "ShiftSchema_変更ログ生成", Err.Number, Err.Description, Erl, _
             "i=" & i
    MsgBox SHT_LOG & " の生成でエラーが発生しました: " & Err.Description, vbExclamation
    Resume CleanUp
End Sub

'==================================================================
' 内部ヘルパー
'==================================================================
'--- シートを取得。無ければ末尾に追加する ---
Private Function SC_GetOrAdd(ByVal nm As String, ByRef isNew As Boolean) As Worksheet
    Dim ws As Worksheet
    On Error GoTo ErrHandler

10  isNew = False
20  Set ws = SheetOrNothing(nm)
30  If Not ws Is Nothing Then
40      Set SC_GetOrAdd = ws
50      Exit Function
60  End If
70  Set ws = ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
80  ws.Name = nm
90  isNew = True
100 Set SC_GetOrAdd = ws
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "SC_GetOrAdd", Err.Number, Err.Description, Erl, _
             "name=" & nm
    Set SC_GetOrAdd = Nothing
End Function

'--- 空欄のときだけ書き込む(既存データを壊さない) ---
Private Sub SC_SetIfBlank(ByVal c As Range, ByVal v As Variant)
    On Error GoTo ErrHandler

10  If Len(Trim$(CStr(c.Value))) = 0 Then c.Value = v
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SC_SetIfBlank", Err.Number, Err.Description, Erl, _
             "cell=" & c.Address(False, False)
End Sub

'--- 見出し行の書式 ---
Private Sub SC_StyleHeader(ByVal rng As Range)
    On Error GoTo ErrHandler

10  With rng
        .Font.Bold = True
        .Font.size = FS_HEAD
        .Interior.Color = ClrHeadBg()
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlContinuous
        .Borders.Color = ClrBorder()
20  End With
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SC_StyleHeader", Err.Number, Err.Description, Erl, _
             "range=" & rng.Address(False, False)
End Sub

'--- 列に入力規則(ドロップダウン)を設定する ---
'    見出し行の1行下から下方向 MAX_VALIDATION_ROWS 行に適用
Private Sub SC_AddList(ByVal ws As Worksheet, ByVal colNo As Long, ByVal listCsv As String)
    Dim rng As Range
    On Error GoTo ErrHandler

10  Set rng = ws.Range(ws.Cells(CFG_FIRST_ROW, colNo), _
                       ws.Cells(CFG_FIRST_ROW + MAX_VALIDATION_ROWS - 1, colNo))
20  With rng.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:=listCsv
        .IgnoreBlank = True
        .InCellDropdown = True
30  End With
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SC_AddList", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; col=" & colNo & "; list=" & listCsv
End Sub
