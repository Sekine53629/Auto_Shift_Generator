Attribute VB_Name = "ShiftCommon"
Option Explicit
'==================================================================
'  シフト表 共通モジュール ＜標準モジュール ShiftCommon v3.1＞
'  2026-08-28
'
'  v3.1: ノルマ外の休み記号の既定値 PAID_OFF_DEFAULT を追加。
'        ShiftAuto と ShiftSchema に "有休" が二重に書かれていたのを
'        ここへ寄せた。既定に夏休が無く、夏休が公休ノルマを食っていた。
'
'  v3.0: 上限判定の失敗時に使う CNT_LARGE を追加。
'
'  v2.9: 事務員の早番人数の既定値 CLERK_EARLY_DEFAULT と、
'        混雑日のしきい値 DOC_BUSY_N を追加。DOC_BUSY_N は
'        集計列の「5診出勤」と自動作成の均等化で共用する。
'
'  v2.8: 貼り付け先の名前付き範囲を「シフトパレット範囲」から
'        「シフト入力範囲」に改名。中身は入力欄でパレットではない。
'        旧名を消すため NM_OBSOLETE を追加。
'
'  v2.7: 休業者のスタッフ行に使う ClrLeaveBg(#BFBFBF)を追加。
'
'  v2.6: v2.5 の判定が一度も True にならなかったのを修正。
'        判定に LBL_DOCTORS("医師名")を使っていたが、パレット3行目に
'        実際に書かれるのは "医師" で、前方一致にならなかった。
'        LBL_DOCTORS はA列(シートの医師名欄)のラベルであり別物。
'        パレット用に LBL_DOC_STAMP("医師")を分け、生成側の
'        SS_PalLabs と判定側の IsDoctorStamp が同じ定数を見るようにした。
'
'  v2.5: 医師名スタンプの判定を位置からラベルへ変更。
'        パレット3行目が「医師」で始まるかで見る。
'        医師枠を9→10に増やしたとき IDX_DOC_LAST が古いままで
'        10人目が備考スタンプと誤判定された。位置固定をやめて解消。
'        IDX_DOC_LAST / IDX_NOTE_FIRST を廃止し、代わりに
'        LastDoctorIndex() / PaletteLabel() を追加。
'        DOC_SLOTS は「生成時に作る枠数」としてのみ残す(9→10)。
'
'  v2.4: シフト記号(SYM_*)・区分(KIND_*)・★マーカー(MARKER_CHAR)を
'        本モジュールに集約。IsEarlySym で ○/◯ の入力揺れを吸収する。
'        未使用だった MonthNumber を削除。
'
'  シート名・ラベル・列・行オフセット・範囲解決をここに一元化する。
'  ShiftAuto / ShiftClick / ShiftSetup / ShiftSchema は範囲を自前で
'  持たず、必ず本モジュールの関数を経由すること。
'
'  ------------------------------------------------------------
'  実シートの行構成 (2026-08-27 実測):
'    1行  ★マーカー行
'    2行  パレット本体行      A2 = "シフトパレット"
'    3行  パレットラベル行
'    4行  年月・タイトル行    A4 = 年月 / D4 = タイトル / I4 = 祝日サマリー
'    5行  日付行(開始日の数式) ← 基準セル
'    6行  曜日行 =TEXT(B5,"aaa")
'    7-11行 医師名欄(5行)
'    12行 空行
'    13行 日付の再掲
'    14行 スタッフ入力の開始
'    29行 備考
'    31行 医師数(診)  = 備考 + 2
'    32行 薬剤師出勤数
'    33行 過不足
'    35-37行 凡例
'
'  v2.3 変更:
'   ・色定義を本モジュールに一元化(色パレット節を追加)
'     ShiftSchema と ShiftSetup が別々に Private で持っていたため、
'     ClrBorder が二重定義、ClrSetBg と ClrDocBg が同値の別名だった。
'     片方だけ色を変えると画面が食い違う状態だったのを解消。
'   ・改名: ClrSetBg/ClrDocBg → ClrInputBg、ClrNoteFg/ClrLabelFg → ClrSubFg
'
'  v2.2 変更:
'   ・名前付き範囲の定数を追加(NM_DOCLIST / NM_NOTEROW)
'     ShiftSetup が3つの名前をまとめて再定義できるようにするため。
'   ・NOTE_GAP を追加(入力欄の下端 = 備考行 - NOTE_GAP)
'     DOC_GAP - NOTE_TO_DOC を各所で手計算していたのを定数化。
'   ・MonthCell 参照漏れの是正: シートモジュール側が Range("A1") 固定
'     だったため、A4 の年月を変更しても期替わり判定が働かなかった。
'     → シートモジュールを MonthCell(Me) 経由に変更(本モジュールは変更なし)
'
'  v2.1 変更:
'   ・ShiftAutoBridge を吸収して廃止(モジュール数を削減)
'     MonthCell / MonthValue / MonthNumber / DateRowForGrid を移管。
'     位置の解決は本モジュールに一元化するという方針に反していたため。
'     ※AutoShiftPreflight は ShiftClick へ移管
'
'  v2.0 変更:
'   ・IDX_* をライブのパレット(26項目)に合わせて修正
'     旧: OFF=1 / 切替=2 / 色消=3 / 消去=4
'     新: OFF=1 / 自動=2 / 切替=3 / 色消=4 / 背景色=5-7 / 消去=8
'     ※旧定数のままだと「切替」を押すと色消モードになり、
'       CycleValues がモードボタンの文字を巡回に混ぜていた
'   ・PALETTE_GAP を 2 → 3 に修正(実シートに合わせる)
'   ・IDX_AUTO / IDX_FILL_FIRST / IDX_FILL_LAST / IDX_DOC_FIRST を追加
'   ・入力欄の上端を「2つ目の日付行の1行下」で解決するよう変更
'     (A列から「曜日」「氏名」ラベルが無くなったため)
'   ・生成対象シートのスキーマ定数を追加(ShiftSchema が使う)
'==================================================================
Private Const MODULE_NAME As String = "ShiftCommon"

'--- シート名 ---
Public Const SHT_SHIFT   As String = "シフト"
Public Const SHT_CFG     As String = "自動作成設定"
Public Const SHT_HOLIDAY As String = "祝日マスタ"
Public Const SHT_LOG     As String = "シフト変更ログ"

'--- A列の基準ラベル(前方一致で検索) ---
Public Const LBL_WEEK    As String = "曜日"
Public Const LBL_NOTE     As String = "備考"
Public Const LBL_DOC      As String = "医師数"
Public Const LBL_PHARM    As String = "薬剤師出勤数"
Public Const LBL_CLERK    As String = "事務員出勤数"
Public Const LBL_SHORT    As String = "過不足"
Public Const LBL_PALETTE  As String = "シフトパレット"
Public Const LBL_DOCTORS  As String = "医師名"   ' A列: シートの医師名欄
'--- パレット3行目(ラベル行)に入る医師名スタンプの表示 ---
'    A列の LBL_DOCTORS("医師名")とは別物。パレットには "医師" と書く。
'    生成(SS_PalLabs)と判定(IsDoctorStamp)が同じ定数を見ること。
Public Const LBL_DOC_STAMP As String = "医師"

'--- 名前付き範囲の名前 ---
'    名前と中身が対応していることを必ず保つこと。
'
'      シフト入力範囲   … スタッフの行 x 日付の列。スタンプの貼り付け先。
'      シフトパレット   … パレット本体(横1行)。押すボタンが並ぶ行。
'      医師名リスト範囲 … 医師名を書く欄。目印であり、マクロは読まない。
'      備考行範囲       … 備考の1行。
'
'    v2.8 まで、貼り付け先の名前が「シフトパレット範囲」だった。
'    パレット本体の「シフトパレット」と1文字違いで並ぶうえ、中身は
'    パレットではなく入力欄なので、どちらがどちらか分からなくなる。
'    中身に合わせて「シフト入力範囲」に改名した。
Public Const NM_SHIFT   As String = "シフト入力範囲"
Public Const NM_PALETTE As String = "シフトパレット"
'    DoctorBlock は毎回計算するので、この名前は読まない。名前ボックスから
'    医師名欄へ飛ぶための目印として置いてある。行を増減すると計算側は
'    追随するが、この名前は初期設定を実行し直すまで古いままになる。
Public Const NM_DOCLIST As String = "医師名リスト範囲"
Public Const NM_NOTEROW As String = "備考行範囲"
'--- 廃止した名前(初期設定の実行時に消す。増えたらカンマで足す) ---
'    消さないと、中身と食い違う名前がブックに残り続ける。
Public Const NM_OBSOLETE As String = "シフトパレット範囲"

'--- 日付/シフトの列範囲 ---
Public Const COL_FIRST As String = "B"
Public Const COL_LAST  As String = "AF"

'--- 行オフセット ---
'  入力欄の下端 = 医師数行の n 行上
'  実測: 医師数(診)=31行 / 最終スタッフ行=27行 → n=4
'  (医師数行の直上に「備考」行とその上の空行があるため 2 ではない)
Public Const DOC_GAP        As Long = 4
Public Const NOTE_TO_DOC    As Long = 2   ' 医師数行 = 備考行の n 行下
'  入力欄の下端 = 備考行の NOTE_GAP 行上 (= DOC_GAP - NOTE_TO_DOC)
'  備考の1行上は区切りの空行、2行上が最終スタッフ行
Public Const NOTE_GAP       As Long = 2
Public Const PALETTE_ROWS   As Long = 3   ' パレットが使う行数
Public Const PALETTE_GAP    As Long = 3   ' パレット本体行 = 日付行の n 行上
Public Const MARKER_OFFSET  As Long = -1  ' ★マーカー行(本体行からの相対)
Public Const LABEL_OFFSET   As Long = 1   ' ラベル行(本体行からの相対)
Public Const DOC_BLOCK_ROWS As Long = 5   ' 医師名欄の行数
Public Const DATE_REPEAT_GAP As Long = 1  ' 入力欄の上端 = 再掲日付行の n 行下

'--- 開始日の数式セルを探す行数の上限 ---
Private Const MAX_SCAN_ROWS As Long = 200

'==================================================================
'  パレットの固定位置(左からの番号)
'  ライブのパレット B2:AA3 = 26項目
'    1 OFF   2 自動   3 切替   4 色消
'    5 背景緑 6 背景橙 7 背景灰
'    8 消去
'    9 ○  10 ●  11 ▲
'   12 公休 13 希休 14 夏休 15 有休 16 有休※
'   17-25 医師名(9名)
'   26 銀行
'==================================================================
Public Const IDX_OFF        As Long = 1   ' OFF(マクロ停止)
Public Const IDX_AUTO       As Long = 2   ' 自動(AutoShift 起動)
Public Const IDX_UNDO       As Long = 3   ' 戻す(直前のセッションを取り消す)
Public Const IDX_EXPORT     As Long = 4   ' 出力(印刷用に PDF / Excel へ)
Public Const IDX_CYCLE      As Long = 5   ' 連続切替
Public Const IDX_CLEARFILL  As Long = 6   ' 背景色クリア
Public Const IDX_FILL_FIRST As Long = 7   ' 背景色ペイントの先頭(背景緑)
Public Const IDX_FILL_LAST  As Long = 9   ' 背景色ペイントの最終(背景灰)
Public Const IDX_ERASE      As Long = 10  ' 消去(空白スタンプ)。記号はこの次から
Public Const IDX_SYM_FIRST  As Long = 11  ' ○ の位置
Public Const IDX_SYM_LAST   As Long = 13  ' ▲ の位置
Public Const IDX_DOC_FIRST  As Long = 19  ' 医師名スタンプの開始位置
' パレット生成で作る医師枠の数。判定には使わない(ラベルで見る)
Public Const DOC_SLOTS      As Long = 10  ' 医師名スタンプの数
' 医師名の範囲は IsDoctorStamp / LastDoctorIndex がラベルから求める。
' 位置を定数で持つと枠の増減に追随できないため IDX_DOC_LAST /
' IDX_NOTE_FIRST は廃止した(v2.5)。

'--- 動作ボタン(モードにならず、ダブルクリックで即実行する) ---
'    自動 / 戻す / 出力 の3つ。IDX_AUTO..IDX_EXPORT が連続していること。
Public Const IDX_ACTION_FIRST As Long = IDX_AUTO
Public Const IDX_ACTION_LAST  As Long = IDX_EXPORT

'--- 書き込み先の種別(クリック入力の判定に使う) ---
Public Const TGT_NONE  As Long = 0   ' 書き込めない場所
Public Const TGT_SHIFT As Long = 1   ' シフト入力欄(スタッフの行)
Public Const TGT_NOTE  As Long = 2   ' 備考行
Public Const TGT_DOC   As Long = 3   ' 医師名欄

'--- ★マーカー(選択中のパレット位置を示す文字) ---
Public Const MARKER_CHAR As String = "★"

'--- シフト記号(全モジュール共通) ---
Public Const SYM_EARLY     As String = "○"
Public Const SYM_EARLY_ALT As String = "◯"   ' 全角の別字体。入力揺れとして受ける
Public Const SYM_MID       As String = "●"
Public Const SYM_LATE      As String = "▲"
Public Const SYM_OFF       As String = "公休"

'--- 区分の正規値(これ以外は設定チェックで警告) ---
' 事務員の早番(○)人数/日の既定値(設定シートに行が無いときに使う)
Public Const CLERK_EARLY_DEFAULT As Long = 1

' 公休ノルマに数えない休みの記号(設定シート L11 の既定値)。
' IsPaidOff は部分一致で見るため、「有休」は「有休※」も拾う。
' 集計列の「休」は全部の休みを足す別物なので、そちらとは一致しない。
Public Const PAID_OFF_DEFAULT As String = "有休,夏休"

' 「混雑日」とみなす医師数。集計列の「5診出勤」と自動作成の均等化で
' 同じ値を使う(片方だけ変えると表と中身がずれるため共有定数とする)
Public Const DOC_BUSY_N As Long = 5

' 上限判定が例外で答えを出せないときに返す値。
' 「上限を超えている」と見なして動かさない側に倒すために使う
Public Const CNT_LARGE As Long = 32767

Public Const KIND_PH As String = "薬剤師"
Public Const KIND_CL As String = "事務員"

'--- 起動するマクロ名(Application.Run で呼ぶため文字列) ---
'  ShiftAuto v8.0.0 の入口プロシージャ名。
'  実体は Public Sub シフト自動作成() なので日本語名を指定する。
Public Const AUTOSHIFT_MACRO As String = "シフト自動作成"

'--- 動作ボタンから Application.Run で呼ぶマクロ名 ---
Public Const UNDO_MACRO   As String = "シフト変更を戻す"
Public Const EXPORT_MACRO As String = "ShiftExport_シフト表出力"

'--- 整合性チェックのマクロ名(自動実行前の確認用) ---
Public Const AUTOCHECK_MACRO As String = "シフト設定チェック"

'==================================================================
'  生成対象シートのスキーマ(ShiftSchema が使う)
'==================================================================
'--- 自動作成設定: メンバー表 ---
Public Const CFG_HDR_ROW    As Long = 4   ' 見出し行
Public Const CFG_FIRST_ROW  As Long = 5   ' メンバー1行目
Public Const CFG_COL_NAME   As Long = 1   ' A 氏名
Public Const CFG_COL_KIND   As Long = 2   ' B 区分
Public Const CFG_COL_CLOSED As Long = 3   ' C 休業
Public Const CFG_COL_RULE   As Long = 4   ' D 勤務ルール
Public Const CFG_COL_FIXDOW As Long = 5   ' E 固定曜日
Public Const CFG_COL_WEEKN  As Long = 6   ' F 週勤務日数
Public Const CFG_COL_OFFDAY As Long = 7   ' G 月間休日数
Public Const CFG_COL_LATE   As Long = 8   ' H 遅番・遅半 可否
Public Const CFG_COL_MEMO   As Long = 9   ' I 備考

'--- 自動作成設定: 全体設定(K=ラベル / L=値) ---
Public Const CFG_SET_ROW   As Long = 4    ' 全体設定の見出し行
Public Const CFG_COL_SETK  As Long = 11   ' K ラベル
Public Const CFG_COL_SETV  As Long = 12   ' L 値

'--- 祝日マスタ ---
Public Const HOL_HDR_ROW   As Long = 1
Public Const HOL_FIRST_ROW As Long = 2
Public Const HOL_COL_DATE  As Long = 1    ' A 日付
Public Const HOL_COL_NAME  As Long = 2    ' B 名称

'--- シフト変更ログ ---
Public Const LOG_HDR_ROW   As Long = 1
Public Const LOG_FIRST_ROW As Long = 2
Public Const LOG_COL_TIME  As Long = 1    ' A 日時
Public Const LOG_COL_ADDR  As Long = 2    ' B セル
Public Const LOG_COL_BEFORE As Long = 3   ' C 変更前
Public Const LOG_COL_AFTER  As Long = 4   ' D 変更後
Public Const LOG_COL_USER   As Long = 5   ' E ユーザー
Public Const LOG_COL_NOTE   As Long = 6   ' F 備考

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

'--- シートを取得(無ければ Nothing) ---
Public Function SheetOrNothing(ByVal nm As String) As Worksheet
    On Error GoTo ErrHandler
10  Set SheetOrNothing = ThisWorkbook.Worksheets(nm)
    Exit Function
ErrHandler:
    ' シートが無いのは正常系(ShiftSchema が生成する)
20  Set SheetOrNothing = Nothing
End Function

'--- シートが存在するか ---
Public Function SheetExists(ByVal nm As String) As Boolean
    On Error GoTo ErrHandler
10  SheetExists = Not (SheetOrNothing(nm) Is Nothing)
    Exit Function
ErrHandler:
    SheetExists = False
End Function

'--- 名前付き範囲を取得(未定義/参照切れなら Nothing) ---
Public Function NamedRangeOrNothing(ByVal nm As String) As Range
    On Error GoTo ErrHandler
10  Set NamedRangeOrNothing = ThisWorkbook.Names(nm).RefersToRange
    Exit Function
ErrHandler:
    ' 名前が無い/参照切れは正常系。呼び出し側でラベル計算に落とす
20  Set NamedRangeOrNothing = Nothing
End Function

'--- そのパレット番号が医師名スタンプか ---
'    判定はラベル行(パレット本体行の LABEL_OFFSET 下)の文字が
'    LBL_DOC_STAMP("医師")で始まるかで行う。
'    A列用の LBL_DOCTORS("医師名")ではない。取り違えると
'    前方一致にならず、判定が一度も True にならない(v2.5 の不具合)。
'    位置を定数で持つと医師枠を増減したときに追随できない。
'    実際にそれが起きた: 医師を9枠から10枠に増やしたとき、
'    IDX_DOC_LAST が古いままで10人目が備考スタンプと誤判定された。
Public Function IsDoctorStamp(ByVal idx As Long) As Boolean
    Dim lab As String
    On Error GoTo ErrHandler

10  lab = PaletteLabel(idx)
20  If Len(lab) = 0 Then Exit Function
30  IsDoctorStamp = (InStr(1, lab, LBL_DOC_STAMP, vbTextCompare) = 1)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "IsDoctorStamp", Err.Number, Err.Description, Erl, _
             "idx=" & idx & "; label=" & lab
    IsDoctorStamp = False
End Function

'--- パレットの指定位置のラベル(3行目)を返す。範囲外なら空文字 ---
'    ラベル行は必ず埋まる前提。本体行は背景色ボタンなどが空なので
'    種別の判定には使えない。
Public Function PaletteLabel(ByVal idx As Long) As String
    Dim pal As Range
    On Error GoTo ErrHandler

10  If idx < 1 Then Exit Function
20  Set pal = PaletteRange(ShiftSheet())
30  If pal Is Nothing Then Exit Function
40  If idx > pal.Cells.Count Then Exit Function
50  PaletteLabel = Trim$(CStr(pal.Cells(1, idx).Offset(LABEL_OFFSET, 0).Value))
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "PaletteLabel", Err.Number, Err.Description, Erl, _
             "idx=" & idx
    PaletteLabel = ""
End Function

'--- そのパレット番号が動作ボタン(自動/戻す/出力)か ---
'    モードとして選択されず、ダブルクリックでその場で実行する。
Public Function IsActionButton(ByVal idx As Long) As Boolean
    On Error GoTo ErrHandler

10  IsActionButton = (idx >= IDX_ACTION_FIRST And idx <= IDX_ACTION_LAST)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "IsActionButton", Err.Number, Err.Description, Erl, _
             "idx=" & idx
    IsActionButton = False
End Function

'--- そのパレット番号が備考スタンプか ---
'    最後の医師名より後ろにあるスタンプ(銀行など)。備考行にしか押せない。
'    IsDoctorStamp と同じ理由でラベル基準にする。位置を定数で持たない。
Public Function IsNoteStamp(ByVal idx As Long) As Boolean
    Dim lastDoc As Long
    On Error GoTo ErrHandler

10  If idx < 1 Then Exit Function
    '--- 医師名スタンプ自身は備考スタンプではない ---
20  If IsDoctorStamp(idx) Then Exit Function
30  lastDoc = LastDoctorIndex()
40  If lastDoc = 0 Then Exit Function
50  IsNoteStamp = (idx > lastDoc)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "IsNoteStamp", Err.Number, Err.Description, Erl, _
             "idx=" & idx & "; lastDoc=" & lastDoc
    IsNoteStamp = False
End Function

'--- 医師名スタンプの最終位置をラベルから求める(0 = 医師枠が無い) ---
'    定数 IDX_DOC_LAST の代わり。枠を増減しても追随する。
Public Function LastDoctorIndex() As Long
    Dim pal As Range, i As Long
    On Error GoTo ErrHandler

10  Set pal = PaletteRange(ShiftSheet())
20  If pal Is Nothing Then Exit Function
30  For i = pal.Cells.Count To 1 Step -1
40      If IsDoctorStamp(i) Then
50          LastDoctorIndex = i
60          Exit Function
70      End If
80  Next i
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "LastDoctorIndex", Err.Number, Err.Description, Erl, _
             "i=" & i
    LastDoctorIndex = 0
End Function

'--- 早番記号か(○ と ◯ の入力揺れを吸収する) ---
Public Function IsEarlySym(ByVal v As String) As Boolean
    On Error GoTo ErrHandler

10  IsEarlySym = (v = SYM_EARLY Or v = SYM_EARLY_ALT)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "IsEarlySym", Err.Number, Err.Description, Erl, _
             "v=" & v
    IsEarlySym = False
End Function

'--- A列を前方一致で検索して行番号を返す(0 = 未検出) ---

'====================== 色パレット ================================
'  シート上の配色はすべてここで決める。
'  各モジュールで Private に持つと二重定義や同値の別名が生まれるため、
'  必ず本モジュールの関数を呼ぶこと。
'------------------------------------------------------------------
' 見出し・枠
Public Function ClrHeadBg() As Long
    On Error GoTo ErrHandler
    ClrHeadBg = RGB(217, 217, 217)   ' 見出し行の背景(薄いグレー)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrHeadBg", Err.Number, Err.Description, Erl, ""
End Function

Public Function ClrBorder() As Long
    On Error GoTo ErrHandler
    ClrBorder = RGB(150, 150, 150)   ' 罫線
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrBorder", Err.Number, Err.Description, Erl, ""
End Function

' 文字
Public Function ClrSubFg() As Long
    On Error GoTo ErrHandler
    ClrSubFg = RGB(100, 100, 100)    ' 補助文字(注記・ラベル)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrSubFg", Err.Number, Err.Description, Erl, ""
End Function

Public Function ClrModeFg() As Long
    On Error GoTo ErrHandler
    ClrModeFg = RGB(80, 80, 80)      ' モード表示の文字
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrModeFg", Err.Number, Err.Description, Erl, ""
End Function

Public Function ClrHidden() As Long
    On Error GoTo ErrHandler
    ClrHidden = RGB(255, 255, 255)   ' 背景と同色にして見せない文字
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrHidden", Err.Number, Err.Description, Erl, ""
End Function

' 入力欄・設定欄
Public Function ClrInputBg() As Long
    On Error GoTo ErrHandler
    ClrInputBg = RGB(226, 239, 218)  ' 入力/設定セルの背景(薄い緑)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrInputBg", Err.Number, Err.Description, Erl, ""
End Function

Public Function ClrDocFg() As Long
    On Error GoTo ErrHandler
    ClrDocFg = RGB(0, 61, 0)         ' 医師名欄の文字(濃い緑)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrDocFg", Err.Number, Err.Description, Erl, ""
End Function

' パレット・モード表示
Public Function ClrModeBg() As Long
    On Error GoTo ErrHandler
    ClrModeBg = RGB(242, 242, 242)   ' モード表示の背景
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrModeBg", Err.Number, Err.Description, Erl, ""
End Function

'--- 休業中のスタッフ行(#BFBFBF) ---
'    自動作成の対象外であることを一目で分かるようにする。
'    この色かどうかで「マクロが塗った行か」を判別するため、
'    背景色ペイントの色とは重ならない値にすること。
Public Function ClrLeaveBg() As Long
    On Error GoTo ErrHandler
    ClrLeaveBg = RGB(191, 191, 191)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrLeaveBg", Err.Number, Err.Description, Erl, ""
End Function

Public Function ClrMarker() As Long
    On Error GoTo ErrHandler
    ClrMarker = RGB(192, 0, 0)       ' 選択中マーカー(★)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ClrMarker", Err.Number, Err.Description, Erl, ""
End Function

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
' 行の解決
'==================================================================
'--- B列で「数式かつ日付」のセルを上から n 個目まで探す ---
'    1個目 = 日付行 / 2個目 = 再掲日付行(スタッフ入力欄の直上)
Public Function DateFormulaRow(ByVal ws As Worksheet, ByVal nth As Long) As Long
    Dim c As Range, r As Long, col As Long, hit As Long
    On Error GoTo ErrHandler
10  col = ws.Range(COL_FIRST & "1").Column
20  For r = 1 To MAX_SCAN_ROWS
30      Set c = ws.Cells(r, col)
40      If c.HasFormula Then
50          If IsDate(c.Value) Then
60              hit = hit + 1
70              If hit = nth Then
80                  DateFormulaRow = r
90                  Exit Function
100             End If
110         End If
120     End If
130 Next r
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "DateFormulaRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; nth=" & nth & "; r=" & r
    DateFormulaRow = 0
End Function

'--- シフト表の開始日セル(B列で最初の「数式かつ日付」) ---
Public Function StartDateCell(ByVal ws As Worksheet) As Range
    Dim r As Long
    On Error GoTo ErrHandler
10  r = DateFormulaRow(ws, 1)
20  If r > 0 Then Set StartDateCell = ws.Cells(r, ws.Range(COL_FIRST & "1").Column)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "StartDateCell", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name
    Set StartDateCell = Nothing
End Function

'--- 日付行 = 開始日の数式が入っている行 ---
Public Function DateRow(ByVal ws As Worksheet) As Long
    On Error GoTo ErrHandler
10  DateRow = DateFormulaRow(ws, 1)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "DateRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name
    DateRow = 0
End Function

'--- 医師数行: 備考行の NOTE_TO_DOC 行下。備考が無ければA列ラベルで探す ---
Public Function ShiftDocRow(ByVal ws As Worksheet) As Long
    Dim noteRow As Long
    On Error GoTo ErrHandler
10  noteRow = LabelRow(ws, LBL_NOTE)
20  If noteRow > 0 Then
30      ShiftDocRow = noteRow + NOTE_TO_DOC
40  Else
50      ShiftDocRow = LabelRow(ws, LBL_DOC)
60  End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ShiftDocRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; noteRow=" & noteRow
    ShiftDocRow = 0
End Function

'--- 入力欄の上端 ---
'    再掲日付行(B列2個目の日付数式)の1行下。
'    再掲が無ければ 曜日ラベル → 日付行 の順にフォールバックし、
'    A列に値がある最初の行を上端とする。
Public Function ShiftTopRow(ByVal ws As Worksheet) As Long
    Dim repeatRow As Long, startRow As Long, docRow As Long, r As Long
    On Error GoTo ErrHandler
10  docRow = ShiftDocRow(ws)
20  If docRow = 0 Then Exit Function

    '--- 本線: 再掲日付行の1行下 ---
30  repeatRow = DateFormulaRow(ws, 2)
40  If repeatRow > 0 And repeatRow + DATE_REPEAT_GAP < docRow Then
50      ShiftTopRow = repeatRow + DATE_REPEAT_GAP
60      Exit Function
70  End If

    '--- 予備: 曜日ラベル/日付行より下でA列に値がある最初の行 ---
80  startRow = LabelRow(ws, LBL_WEEK)
90  If startRow = 0 Then startRow = DateRow(ws)
100 If startRow = 0 Then Exit Function
110 For r = startRow + 1 To docRow - 1
120     If Len(Trim$(CStr(ws.Cells(r, 1).Value))) > 0 Then
130         ShiftTopRow = r
140         Exit Function
150     End If
160 Next r
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ShiftTopRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; repeatRow=" & repeatRow & "; docRow=" & docRow
    ShiftTopRow = 0
End Function

'--- 入力欄の下端 = 医師数行の DOC_GAP 行上 ---
Public Function ShiftBottomRow(ByVal ws As Worksheet) As Long
    Dim docRow As Long
    On Error GoTo ErrHandler
10  docRow = ShiftDocRow(ws)
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

'--- 年月・タイトル行 = 日付行の1行上(パレットラベル行の1行下) ---
Public Function HeaderRow(ByVal ws As Worksheet) As Long
    Dim dRow As Long
    On Error GoTo ErrHandler
10  dRow = DateRow(ws)
20  If dRow > 1 Then HeaderRow = dRow - 1
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "HeaderRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; dateRow=" & dRow
    HeaderRow = 0
End Function

'--- 医師名欄のブロック ---
'    名前付き範囲 NM_DOCLIST は読まない。あれは静的なアドレスなので、
'    行を増減すると古い位置を指したままになる。ここは毎回計算する。
'    上端 = 曜日行の1行下(曜日行が無ければ日付行+2)
'    下端 = 上端 + DOC_BLOCK_ROWS - 1
'    A列に「医師名」があればその行を上端として優先する
Public Function DoctorBlock(ByVal ws As Worksheet) As Range
    Dim topR As Long, botR As Long, firstCol As Long, lastCol As Long
    Dim dRow As Long, repeatRow As Long
    On Error GoTo ErrHandler
10  topR = LabelRow(ws, LBL_DOCTORS)
20  If topR = 0 Then
30      dRow = DateRow(ws)
40      If dRow = 0 Then Exit Function
50      topR = dRow + 2            ' 日付行 → 曜日行 → 医師名欄
60  End If
    '--- 下端: 再掲日付行の1行上を上限とする ---
70  botR = topR + DOC_BLOCK_ROWS - 1
80  repeatRow = DateFormulaRow(ws, 2)
90  If repeatRow > topR Then
100     If botR >= repeatRow Then botR = repeatRow - 1
110 End If
120 If botR < topR Then Exit Function
130 firstCol = ws.Range(COL_FIRST & "1").Column
140 lastCol = ws.Range(COL_LAST & "1").Column
150 Set DoctorBlock = ws.Range(ws.Cells(topR, firstCol), ws.Cells(botR, lastCol))
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "DoctorBlock", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; topRow=" & topR & "; bottomRow=" & botR
    Set DoctorBlock = Nothing
End Function

'==================================================================
' 範囲の解決(名前付き範囲 → 計算 → Nothing)
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

'--- 過不足行(表の最終行。出力範囲の下端に使う) ---
Public Function ShortageRow(ByVal ws As Worksheet) As Long
    Dim r As Long
    On Error GoTo ErrHandler

10  r = LabelRow(ws, LBL_SHORT)
20  If r > 0 Then
30      ShortageRow = r
40  Else
        ' ラベルが無い場合は 医師数行 + 2 (医師数 / 薬剤師出勤数 / 過不足)
50      r = ShiftDocRow(ws)
60      If r > 0 Then ShortageRow = r + 2
70  End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ShortageRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name
    ShortageRow = 0
End Function

'--- 備考行 ---
'    シフト入力欄(スタッフの行)には含めない。自動作成の対象外であり、
'    日ごとのメモを書く行だが、クリック入力とログの対象には含める。
Public Function NoteRow(ByVal ws As Worksheet) As Long
    On Error GoTo ErrHandler

10  NoteRow = LabelRow(ws, LBL_NOTE)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "NoteRow", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name
    NoteRow = 0
End Function

'--- 備考行の入力範囲(B..AF) ---
Public Function NoteRange(ByVal ws As Worksheet) As Range
    Dim r As Long
    On Error GoTo ErrHandler

10  Set NoteRange = NamedRangeOrNothing(NM_NOTEROW)
20  If Not NoteRange Is Nothing Then Exit Function

30  r = NoteRow(ws)
40  If r = 0 Then Exit Function
50  Set NoteRange = ws.Range(COL_FIRST & r & ":" & COL_LAST & r)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "NoteRange", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; noteRow=" & r
    Set NoteRange = Nothing
End Function

'--- 2つの範囲を結合する(片方が Nothing でも落ちない) ---
Public Function UnionSafe(ByVal a As Range, ByVal b As Range) As Range
    On Error GoTo ErrHandler

10  If a Is Nothing Then
20      Set UnionSafe = b
30  ElseIf b Is Nothing Then
40      Set UnionSafe = a
50  Else
60      Set UnionSafe = Application.Union(a, b)
70  End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "UnionSafe", Err.Number, Err.Description, Erl, ""
    Set UnionSafe = a
End Function

'--- 人が手で書き込める範囲(スタッフ入力欄 + 備考行) ---
'    ShiftInputRange は自動作成のアルゴリズムが使う「スタッフの行」なので
'    備考行を含めない。クリック入力と手動変更ログはこちらを使う。
Public Function EditableRange(ByVal ws As Worksheet) As Range
    On Error GoTo ErrHandler

10  Set EditableRange = UnionSafe(ShiftInputRange(ws), NoteRange(ws))
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "EditableRange", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name
    Set EditableRange = Nothing
End Function

'--- 入力欄の下端のずれ(0=正常 / 正=下すぎ / 負=上すぎ) ---
Public Function ShiftRangeDrift(ByVal ws As Worksheet) As Long
    Dim rng As Range, docRow As Long, endRow As Long
    On Error GoTo ErrHandler
10  Set rng = ShiftInputRange(ws)
20  docRow = ShiftDocRow(ws)
30  If rng Is Nothing Or docRow = 0 Then Exit Function
40  endRow = rng.Row + rng.Rows.Count - 1
50  ShiftRangeDrift = endRow - (docRow - DOC_GAP)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "ShiftRangeDrift", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; docRow=" & docRow & "; endRow=" & endRow
    ShiftRangeDrift = 0
End Function


'==================================================================
' 対象月の解決 (v2.1: ShiftAutoBridge から移管)
'   実シートは年月が「年月・タイトル行」のA列にある(A1ではない)。
'   旧 ShiftAuto は ws.Range("A1") 固定だったため必ず中断していた。
'   探索順: 年月・タイトル行のA列 → A1(旧仕様) → 開始日セル
'==================================================================
Public Function MonthCell(ByVal ws As Worksheet) As Range
    Dim hRow As Long
    On Error GoTo ErrHandler

10  hRow = HeaderRow(ws)
20  If hRow > 0 Then
30      If IsDate(ws.Cells(hRow, 1).Value) Then
40          Set MonthCell = ws.Cells(hRow, 1)
50          Exit Function
60      End If
70  End If
80  If IsDate(ws.Range("A1").Value) Then
90      Set MonthCell = ws.Range("A1")
100     Exit Function
110 End If
    '--- 予備: 日付行の先頭(開始日)そのもの ---
120 Set MonthCell = StartDateCell(ws)
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "MonthCell", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name & "; headerRow=" & hRow
    Set MonthCell = Nothing
End Function

'--- 対象月の日付を返す(取得できなければ 0) ---
Public Function MonthValue(ByVal ws As Worksheet) As Date
    Dim c As Range
    On Error GoTo ErrHandler

10  Set c = MonthCell(ws)
20  If c Is Nothing Then Exit Function
30  If IsDate(c.Value) Then MonthValue = CDate(c.Value)
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "MonthValue", Err.Number, Err.Description, Erl, _
             "sheet=" & ws.Name
End Function


'--- 入力欄に対応する日付行 (v2.1: ShiftAutoBridge から移管) ---
'    grid.Row - 1 が日付ならそれを使い、そうでなければ本来の日付行に
'    フォールバックする。→ 再掲日付行を消しても壊れない
Public Function DateRowForGrid(ByVal ws As Worksheet, ByVal grid As Range) As Long
    Dim r As Long, c As Range
    On Error GoTo ErrHandler

10  If grid Is Nothing Then Exit Function
20  r = grid.Row - 1
30  If r >= 1 Then
40      Set c = ws.Cells(r, grid.Column)
50      If IsDate(c.Value) Then
60          DateRowForGrid = r
70          Exit Function
80      End If
90  End If
100 DateRowForGrid = DateRow(ws)
    Exit Function

ErrHandler:
    LogError MODULE_NAME, "DateRowForGrid", Err.Number, Err.Description, Erl, _
             "grid=" & grid.Address(False, False) & "; r=" & r
    DateRowForGrid = 0
End Function
