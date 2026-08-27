Attribute VB_Name = "AutoShiftGenerator"
Option Explicit
'==================================================================
'  ShiftAuto v9.2.0
'  自動作成の入口と前半工程(準備～週リスト)。
'  v9.2.0: 記号・区分の定数を ShiftCommon へ移管。
'          モジュール外から呼ばれない AS_* を Private に変更。
'  ※ ShiftAuto / ShiftAutoPlace / ShiftAutoLog の3本で1組。
'    共有状態は ShiftAuto の Public 変数に置く。
'==================================================================
Private Const MODULE_NAME As String = "ShiftAuto"

' シート名・A列ラベル・範囲解決・シフト記号・区分は ShiftCommon を参照する
' 全体設定(K/L列)を何行下まで探すか
Public Const CFG_SCAN_ROWS As Long = 30

' 氏名として扱わない集計行ラベル(前方一致・カンマ区切り)
Private Const NON_NAME_LABELS As String = _
    "医師数,薬剤師出勤数,事務員出勤数,過不足,合計,シフトパレット,備考,医師名"
' 予定ステータス
Public Const ST_SKIP  As Long = -1   ' 月外・休業・空行・集計行
Public Const ST_WORK  As Long = 1    ' 自動:出勤
Public Const ST_OFF   As Long = 2    ' 自動:公休
Public Const ST_FWORK As Long = 3    ' 既存入力:出勤(○◯●▲)
Public Const ST_FOFF  As Long = 4    ' 既存入力:休み(希休・有休・公休等)
'--- モジュール内共有状態 ---
Public mPlan() As Long, mKind() As String, mRule() As String
Public mLeave() As Boolean, mCanLate() As Boolean
Public mSkipRow() As Boolean          ' 氏名が空・集計行 = 処理対象外
Public mDayIn() As Boolean, mDayWD() As Long, mDayHol() As Boolean
Public mDayDoc() As Long, mDayReq() As Long, mWkKey() As Long
Public mCov() As Long          ' 薬剤師の予定出勤数
Public mCovG() As Long         ' 事務員の予定出勤数
Public mSymb() As String, mCntE() As Long, mCntM() As Long, mCntL() As Long
Public mMaxRun As Long, mMaxOffRun As Long, mPaidSyms As String
Public mNP As Long, mND As Long
'--- v9.1.1 追加: 宣言漏れの補完 -----------------------------------
'  Option Explicit 下で未宣言だった共有状態。ShiftAutoPlace /
'  ShiftAutoLog からも参照するため Public とする。
'  シート・範囲
Public mWs As Worksheet, mCfg As Worksheet, mHolWs As Worksheet
Public mGrid As Range
Public mDateRow As Long, mDocRow As Long
'  対象月
Public mMonthDt As Date, mMonthNum As Long
Public mDayDt() As Date
'  全体設定(K/L列)から読む値
Public mEarlyN As Long, mLateMin As Long
Public mWeekBase As Long, mReqPlus As Long
Public mGSym As String                ' 事務員の2人目の記号
'  メンバー情報
Public mName() As String
Public mWD() As Boolean               ' (i, 1-7) 固定曜日
Public mWeekN() As Long               ' 週N日ルールの N
Public mQuota() As Long               ' 月間休日数(-1 = 未指定)
Public mActiveN As Long, mSkipN As Long
'  不整合の検出結果(メッセージを連結)
Public mDupName As String, mMissing As String
Public mBadKind As String, mOrphan As String
'  週リスト
Public mWkList() As Long, mNW As Long
'  配置・書き込みの集計
Public mTargetOff As Long             ' 公休ノルマ(土日祝の日数)
Public mWritten As Long               ' 書き込んだセル数
Public mUnmet As String               ' 公休ノルマ未達の一覧
Public mSess As Long, mLogged As Long ' 変更ログのセッション番号と件数
'-------------------------------------------------------------------

'==================================================================
' 氏名判定
'==================================================================
'--- 集計行のラベルか?(氏名ではない行を除外) ---

'==================================================================
' 氏名判定
'==================================================================
'--- 集計行のラベルか?(氏名ではない行を除外) ---
Public Function IsNonName(ByVal s As String) As Boolean
    On Error GoTo ErrHandler
    Dim parts() As String, p As Long, t As String
    t = Trim$(s)
    If Len(t) = 0 Then Exit Function
    parts = Split(NON_NAME_LABELS, ",")
    For p = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(p))) > 0 Then
            If InStr(1, t, Trim$(parts(p)), vbTextCompare) = 1 Then
                IsNonName = True
                Exit Function
            End If
        End If
    Next p
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "IsNonName", Err.Number, Err.Description, Erl, _
             "s=" & s
End Function


'==================================================================
'  入口: シフト自動作成
'    工程を順に呼ぶだけの司令塔。各工程は False で中断を意味する。
'    ロジックは各 AS_* に置き、ここには置かない。
'==================================================================
Public Sub シフト自動作成()
    On Error GoTo ErrHandler

    '--- 0) 前回実行の値をクリア(Public 変数は実行間で残るため) ---
5   AS_状態リセット

    '--- 1) 前提の解決(シート・設定値・入力欄) ---
10  If Not AS_準備() Then Exit Sub
20  If Not AS_日情報() Then Exit Sub

    '--- 2) メンバーの読込と不整合の検出 ---
30  If Not AS_メンバー読込() Then Exit Sub
40  If Not AS_孤児検出() Then Exit Sub
50  If Not AS_事前確認() Then Exit Sub

    '--- 3) 出勤/休みの素案づくり ---
60  If Not AS_既存分類() Then Exit Sub
70  If Not AS_ルール適用() Then Exit Sub
80  If Not AS_予定出勤数() Then Exit Sub

    '--- 4) 公休の配置(週N日 → 週の基本休 → 残ノルマ → 連勤緩和) ---
90  If Not AS_週N日ルール() Then Exit Sub
100 If Not AS_週リスト() Then Exit Sub
110 If Not AS_公休ノルマ() Then Exit Sub
120 If Not AS_残ノルマ配置() Then Exit Sub
130 If Not AS_連勤緩和() Then Exit Sub

    '--- 5) 医師5名日の出勤を均等化 ---
140 FiveBalance

    '--- 6) 記号(○●▲)の割当と均等化 ---
150 If Not AS_記号割当() Then Exit Sub
160 SymbolBalance

    '--- 7) 書き込みと結果報告 ---
170 If Not AS_書き込み() Then Exit Sub
180 If Not AS_レポート() Then Exit Sub

    LogSuccess MODULE_NAME, "シフト自動作成", _
               "np=" & mNP & "; nd=" & mND & "; written=" & mWritten
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "シフト自動作成", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
End Sub


'------------------------------------------------------------------
' AS_状態リセット
'   共有状態は Public のため、実行し終えても値が残る。
'   2回目以降の実行で書込セル数が累積したり、前回の警告文が
'   そのまま出たりするのを防ぐため、入口で毎回クリアする。
'------------------------------------------------------------------
Private Sub AS_状態リセット()
    On Error GoTo ErrHandler
10  mWritten = 0: mTargetOff = 0
20  mLogged = 0: mSess = 0
30  mActiveN = 0: mSkipN = 0: mNW = 0
40  mUnmet = "": mDupName = "": mMissing = ""
50  mBadKind = "": mOrphan = ""
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "AS_状態リセット", Err.Number, Err.Description, Erl, ""
End Sub

'------------------------------------------------------------------
' AS_準備
'   シート・設定値・入力欄を解決する。以降の全工程の前提。
'   設定シートが無い場合は生成を促し、False を返して中断する。
'------------------------------------------------------------------
Private Function AS_準備() As Boolean
    On Error GoTo ErrHandler

    '=== 準備 ===
10   Set mWs = ShiftSheet()
20   Set mCfg = SheetOrNothing(SHT_CFG)
30   Set mHolWs = SheetOrNothing(SHT_HOLIDAY)
40   If mCfg Is Nothing Then
50       If MsgBox("設定シート「" & SHT_CFG & "」がありません。" & vbCrLf & _
                  "既定値で自動生成しますか?(氏名はシフト表から取込)" & vbCrLf & vbCrLf & _
                  "※生成後、区分(事務員)・固定曜日・週N日・手動(派遣)・" & vbCrLf & _
                  "　休業・遅番不可 は手動で設定してください。", _
                  vbYesNo + vbQuestion, "設定シート自動生成") = vbYes Then
60           Set mCfg = BuildCfgSheet(mWs)
70           MsgBox "設定シート「" & SHT_CFG & "」を既定値で生成しました。" & vbCrLf & _
                   "内容を確認・修正のうえ、もう一度実行してください。", vbInformation
80       End If
90       AS_準備 = False: Exit Function
100  End If

    '--- 対象月: 年月行のA列 → A1 → 開始日 の順に探索(Bridge) ---
110  mMonthDt = MonthValue(mWs)
120  If mMonthDt = 0 Then
130      MsgBox "対象月の日付が見つかりません。" & vbCrLf & _
               HeaderRow(mWs) & "行のA列に年月を入れてください。", vbExclamation
140      AS_準備 = False: Exit Function
150  End If

    '=== 全体設定の読込(K列ラベルの部分一致検索) ===
160  mEarlyN = CLng(CfgNum(mCfg, "早番", 1))
170  mLateMin = CLng(CfgNum(mCfg, "遅番", 3))
180  mMaxRun = CLng(CfgNum(mCfg, "連勤", 3))
190  mMaxOffRun = CLng(CfgNum(mCfg, "連休", 3))
200  mWeekBase = CLng(CfgNum(mCfg, "週の基本", 2))
210  mReqPlus = CLng(CfgNum(mCfg, "必要出勤", 1))
220  mPaidSyms = CfgTxt(mCfg, "ノルマ外", "有休")
230  mGSym = CfgTxt(mCfg, "2人目", SYM_MID)
240  If mEarlyN < 0 Then mEarlyN = 0
250  If mLateMin < 0 Then mLateMin = 0
260  If mMaxRun < 1 Then mMaxRun = 1
270  If mMaxOffRun < 1 Then mMaxOffRun = 1
280  If mWeekBase < 0 Then mWeekBase = 0

290  Set mGrid = ShiftInputRange(mWs)
300  If mGrid Is Nothing Then
310      MsgBox "シフト入力欄を特定できません。" & vbCrLf & _
               "B列の開始日の数式、またはA列の「" & LBL_NOTE & "」「" & LBL_DOC & _
               "」を確認してください。", vbExclamation: AS_準備 = False: Exit Function
320  End If
330  mNP = mGrid.Rows.Count: mND = mGrid.Columns.Count
    '--- 日付行: 再掲行の有無に依存しない解決(Bridge) ---
340  mDateRow = DateRowForGrid(mWs, mGrid)
350  If mDateRow = 0 Then
360      MsgBox "日付行を特定できません。", vbExclamation: AS_準備 = False: Exit Function
370  End If
380  mDocRow = ShiftDocRow(mWs)
390  If mDocRow = 0 Then
400      MsgBox "「" & LBL_DOC & "」行が見つかりません。", vbExclamation: AS_準備 = False: Exit Function
410  End If
    '--- 入力範囲が集計行に近すぎないか確認 ---
420  If mGrid.Row + mNP - 1 > mDocRow - DOC_GAP Then
430      MsgBox "シフト入力範囲(" & mGrid.Address(False, False) & ")が" & vbCrLf & _
               "「" & LBL_DOC & "」行(" & mDocRow & "行)の" & DOC_GAP & "行上を超えています。" & vbCrLf & vbCrLf & _
               "名前付き範囲「" & NM_SHIFT & "」の終端を" & vbCrLf & _
               "「" & LBL_DOC & "」の" & DOC_GAP & "行上(" & (mDocRow - DOC_GAP) & "行)に修正してください。" & vbCrLf & _
               "(集計行は氏名として扱われないため処理は続行できます)", vbExclamation
440  End If
450  mMonthNum = Month(mMonthDt)

    AS_準備 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_準備", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_準備 = False
End Function


'------------------------------------------------------------------
' AS_日情報
'   日付・曜日・祝日・医師数・必要出勤数を列ごとに求め、
'   公休ノルマ(mTargetOff)を土日祝の数から算出する。
'------------------------------------------------------------------
Private Function AS_日情報() As Boolean
    Dim j As Long
    On Error GoTo ErrHandler

    '=== 日情報 ===
10   ReDim mDayDt(1 To mND): ReDim mDayIn(1 To mND): ReDim mDayWD(1 To mND)
20   ReDim mDayHol(1 To mND): ReDim mDayDoc(1 To mND): ReDim mDayReq(1 To mND)
30   ReDim mWkKey(1 To mND)
40   mTargetOff = 0
50   For j = 1 To mND
60       mDayIn(j) = False
70       If IsDate(mWs.Cells(mDateRow, mGrid.Column + j - 1).Value) Then
80           mDayDt(j) = mWs.Cells(mDateRow, mGrid.Column + j - 1).Value
90           mDayIn(j) = (Month(mDayDt(j)) = mMonthNum)
100      End If
110      If mDayIn(j) Then
120          mDayWD(j) = Weekday(mDayDt(j), vbSunday)
130          mWkKey(j) = CLng(mDayDt(j)) - (mDayWD(j) - 1)
140          If Not mHolWs Is Nothing Then
150              mDayHol(j) = (Application.CountIf(mHolWs.Columns(HOL_COL_DATE), mDayDt(j)) > 0)
160          End If
170          mDayDoc(j) = Val(mWs.Cells(mDocRow, mGrid.Column + j - 1).Value)
180          mDayReq(j) = mDayDoc(j) + mReqPlus
190          If mDayWD(j) = 1 Or mDayWD(j) = 7 Then
200              mTargetOff = mTargetOff + 1
210          ElseIf mDayHol(j) Then
220              mTargetOff = mTargetOff + 1
230          End If
240      End If
250  Next j

    AS_日情報 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_日情報", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_日情報 = False
End Function


'------------------------------------------------------------------
' AS_メンバー読込
'   シフト表の氏名をマスタと照合し、区分・ルール・休業等を読む。
'   氏名が空の行と集計行ラベルは処理対象外(mSkipRow)。
'   同名重複・区分の異常値はここで検出する。
'------------------------------------------------------------------
Private Function AS_メンバー読込() As Boolean
    Dim pFound() As Boolean           ' v9.1.1: 宣言漏れ(設定表で照合できたか)
    Dim i As Long
    Dim ii As Long
    Dim r As Long
    On Error GoTo ErrHandler

    '=== メンバー設定読込(氏名キー照合・空行と集計行はスキップ) ===
10   ReDim mName(1 To mNP): ReDim mKind(1 To mNP): ReDim mLeave(1 To mNP)
20   ReDim mRule(1 To mNP): ReDim mWD(1 To mNP, 1 To 7): ReDim mWeekN(1 To mNP)
30   ReDim mQuota(1 To mNP): ReDim mCanLate(1 To mNP): ReDim pFound(1 To mNP)
40   ReDim mSkipRow(1 To mNP)
50   mActiveN = 0: mSkipN = 0
60   For i = 1 To mNP
70       mName(i) = Trim$(CStr(mWs.Cells(mGrid.Row + i - 1, 1).Value))
80       mKind(i) = KIND_PH: mRule(i) = "通常": mCanLate(i) = True: mQuota(i) = -1
        '--- 氏名が空の行・集計行ラベルは処理対象外 ---
90       If Len(mName(i)) = 0 Or IsNonName(mName(i)) Then
100          mSkipRow(i) = True
110          mSkipN = mSkipN + 1
120      Else
130          mSkipRow(i) = False
140          mActiveN = mActiveN + 1
            '--- 同名の重複チェック(照合が先勝ちになり誤配置の原因) ---
150          For ii = 1 To i - 1
160              If Not mSkipRow(ii) Then
170                  If mName(ii) = mName(i) Then
180                      If InStr(mDupName, "・" & mName(i) & vbCrLf) = 0 Then
190                          mDupName = mDupName & "・" & mName(i) & vbCrLf
200                      End If
210                  End If
220              End If
230          Next ii
            '--- マスタ照合(氏名キー・行番号に依存しない) ---
240          r = CFG_FIRST_ROW
250          Do While Len(Trim$(CStr(mCfg.Cells(r, CFG_COL_NAME).Value))) > 0
260              If Trim$(CStr(mCfg.Cells(r, CFG_COL_NAME).Value)) = mName(i) Then
270                  pFound(i) = True
280                  If Len(Trim$(CStr(mCfg.Cells(r, CFG_COL_KIND).Value))) > 0 Then _
                        mKind(i) = Trim$(CStr(mCfg.Cells(r, CFG_COL_KIND).Value))
290                  mLeave(i) = (Trim$(CStr(mCfg.Cells(r, CFG_COL_CLOSED).Value)) <> "")
300                  If Len(Trim$(CStr(mCfg.Cells(r, CFG_COL_RULE).Value))) > 0 Then _
                        mRule(i) = Trim$(CStr(mCfg.Cells(r, CFG_COL_RULE).Value))
310                  ParseWD CStr(mCfg.Cells(r, CFG_COL_FIXDOW).Value), mWD, i
320                  mWeekN(i) = Val(mCfg.Cells(r, CFG_COL_WEEKN).Value)
330                  If Len(Trim$(CStr(mCfg.Cells(r, CFG_COL_OFFDAY).Value))) > 0 Then _
                        mQuota(i) = Val(mCfg.Cells(r, CFG_COL_OFFDAY).Value)
340                  mCanLate(i) = (Trim$(CStr(mCfg.Cells(r, CFG_COL_LATE).Value)) <> "不可")
350                  Exit Do
360              End If
370              r = r + 1
380          Loop
390          If Not pFound(i) Then
400              mMissing = mMissing & "・" & mName(i) & vbCrLf
410          ElseIf mKind(i) <> KIND_PH And mKind(i) <> KIND_CL Then
                '--- 区分が正規値以外だと人数計算に計上されず静かに壊れる ---
420              mBadKind = mBadKind & "・" & mName(i) & " : 区分「" & mKind(i) & "」" & vbCrLf
430          End If
440      End If
450  Next i

460  If mActiveN = 0 Then
470      MsgBox "シフト入力欄(" & mGrid.Address(False, False) & ")のA列に氏名がありません。" & vbCrLf & _
               "氏名の記入位置をご確認ください。", vbExclamation
480      AS_メンバー読込 = False: Exit Function
490  End If

    AS_メンバー読込 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_メンバー読込", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_メンバー読込 = False
End Function


'------------------------------------------------------------------
' AS_孤児検出
'   マスタにあるがシフト表に無い氏名を拾う(配置漏れの原因)。
'------------------------------------------------------------------
Private Function AS_孤児検出() As Boolean
    Dim i As Long
    Dim r As Long
    Dim nmCheck As String
    Dim hit As Long
    On Error GoTo ErrHandler

    '=== マスタにあるがシフト表に無い氏名(孤児)を検出 ===
10   r = CFG_FIRST_ROW
20   Do While Len(Trim$(CStr(mCfg.Cells(r, CFG_COL_NAME).Value))) > 0
30       nmCheck = Trim$(CStr(mCfg.Cells(r, CFG_COL_NAME).Value))
40       If Not IsNonName(nmCheck) Then
50           hit = 0
60           For i = 1 To mNP
70               If Not mSkipRow(i) Then
80                   If mName(i) = nmCheck Then hit = 1: Exit For
90               End If
100          Next i
110          If hit = 0 Then mOrphan = mOrphan & "・" & nmCheck & vbCrLf
120      End If
130      r = r + 1
140  Loop

    AS_孤児検出 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_孤児検出", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_孤児検出 = False
End Function


'------------------------------------------------------------------
' AS_事前確認
'   検出した不整合をまとめて提示し、続行の可否を利用者に問う。
'   「いいえ」なら False を返して中断する。
'------------------------------------------------------------------
Private Function AS_事前確認() As Boolean
    Dim msg As String
    On Error GoTo ErrHandler

    '=== 整合性チェックの事前確認 ===
10   If Len(mMissing) > 0 Or Len(mOrphan) > 0 Or Len(mDupName) > 0 Or Len(mBadKind) > 0 Then
20       msg = "設定の整合性に注意点があります。" & vbCrLf & vbCrLf
30       If Len(mDupName) > 0 Then
40           msg = msg & "■ 氏名が重複(先に見つかった設定が適用されます)" & vbCrLf & mDupName & vbCrLf
50       End If
60       If Len(mBadKind) > 0 Then
70           msg = msg & "■ 区分が「" & KIND_PH & "」「" & KIND_CL & "」以外" & vbCrLf & _
                  "　(出勤数に計上されません)" & vbCrLf & mBadKind & vbCrLf
80       End If
90       If Len(mMissing) > 0 Then
100          msg = msg & "■ マスタ未登録(既定値=" & KIND_PH & "・通常・遅番可 で処理)" & vbCrLf & mMissing & vbCrLf
110      End If
120      If Len(mOrphan) > 0 Then
130          msg = msg & "■ マスタにあるがシフト表に無い(行削除・氏名変更?)" & vbCrLf & mOrphan & vbCrLf
140      End If
150      msg = msg & "このまま実行しますか?"
160      If MsgBox(msg, vbYesNo + vbExclamation, "設定チェック") <> vbYes Then
170          AS_事前確認 = False      ' 利用者が中止 → 呼び出し側で処理を止める
180          Exit Function
190      End If
200  End If

    AS_事前確認 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_事前確認", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_事前確認 = False
End Function


'------------------------------------------------------------------
' AS_既存分類
'   既に入力済みのセルを固定出勤/固定休みに分類する(mPlan)。
'   空行・集計行は全日スキップ扱い。
'------------------------------------------------------------------
Private Function AS_既存分類() As Boolean
    Dim i As Long
    Dim j As Long
    Dim v As String
    On Error GoTo ErrHandler

    '=== 既存入力の分類(空行・集計行は全日スキップ) ===
10   ReDim mPlan(1 To mNP, 1 To mND)
20   For i = 1 To mNP
30       For j = 1 To mND
40           If mSkipRow(i) Or Not mDayIn(j) Or mLeave(i) Then
50               mPlan(i, j) = ST_SKIP
60           Else
70               v = Trim$(CStr(mGrid.Cells(i, j).Value))
80               If Len(v) = 0 Then
90                   mPlan(i, j) = 0
100              ElseIf InStr("○◯●▲", v) > 0 Then
110                  mPlan(i, j) = ST_FWORK
120              Else
130                  mPlan(i, j) = ST_FOFF          ' 希休・有休・公休など
140              End If
150          End If
160      Next j
170  Next i

    AS_既存分類 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_既存分類", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_既存分類 = False
End Function


'------------------------------------------------------------------
' AS_ルール適用
'   固定曜日は該当曜日のみ出勤、手動は触らない、
'   それ以外は仮に全日出勤として置く。
'------------------------------------------------------------------
Private Function AS_ルール適用() As Boolean
    Dim i As Long
    Dim j As Long
    On Error GoTo ErrHandler

    '=== ルール適用(固定曜日 / 手動=何もしない / それ以外は仮で全出勤) ===
10   For i = 1 To mNP
20       If Not mSkipRow(i) And Not mLeave(i) Then
30           If mRule(i) = "固定曜日" Then
40               For j = 1 To mND
50                   If mPlan(i, j) = 0 Then
60                       If mWD(i, mDayWD(j)) Then mPlan(i, j) = ST_WORK Else mPlan(i, j) = ST_OFF
70                   End If
80               Next j
90           ElseIf mRule(i) = "手動" Then
                ' 派遣行など: 空白は空白のまま(自動配置しない)
100          Else
110              For j = 1 To mND
120                  If mPlan(i, j) = 0 Then mPlan(i, j) = ST_WORK
130              Next j
140          End If
150      End If
160  Next i

    AS_ルール適用 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_ルール適用", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_ルール適用 = False
End Function


'------------------------------------------------------------------
' AS_予定出勤数
'   日ごとの予定出勤数を薬剤師(mCov)と事務員(mCovG)で別に数える。
'------------------------------------------------------------------
Private Function AS_予定出勤数() As Boolean
    Dim i As Long
    Dim j As Long
    On Error GoTo ErrHandler

    '=== 予定出勤数(薬剤師/事務員 別) ===
10   ReDim mCov(1 To mND): ReDim mCovG(1 To mND)
20   For j = 1 To mND
30       mCov(j) = 0: mCovG(j) = 0
40       For i = 1 To mNP
50           If Not mSkipRow(i) Then
60               If mPlan(i, j) = ST_WORK Or mPlan(i, j) = ST_FWORK Then
70                   If mKind(i) = KIND_PH Then mCov(j) = mCov(j) + 1
80                   If mKind(i) = KIND_CL Then mCovG(j) = mCovG(j) + 1
90               End If
100          End If
110      Next i
120  Next j

    AS_予定出勤数 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_予定出勤数", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_予定出勤数 = False
End Function


'------------------------------------------------------------------
' AS_週N日ルール
'   週N日指定者について、週ごとの勤務日数を指定数まで絞る。
'------------------------------------------------------------------
Private Function AS_週N日ルール() As Boolean
    Dim i As Long
    Dim j As Long
    Dim j0 As Long
    Dim cnt As Long
    Dim wCnt As Long
    Dim tgt As Long
    Dim bj As Long
    Dim bs As Double
    Dim sc As Double
    Dim processed() As Boolean
    Dim autoW As Long
    On Error GoTo ErrHandler

    '=== 週N日ルール: 週ごとに勤務日数を絞る ===
10   For i = 1 To mNP
20       If Not mSkipRow(i) Then
30       If mRule(i) = "週N日" And Not mLeave(i) And mWeekN(i) > 0 Then
40           ReDim processed(1 To mND)
50           For j0 = 1 To mND
60               If mDayIn(j0) And Not processed(j0) Then
70                   cnt = 0: wCnt = 0
80                   For j = 1 To mND
90                       If mDayIn(j) And mWkKey(j) = mWkKey(j0) Then
100                          processed(j) = True
110                          cnt = cnt + 1
120                          If mPlan(i, j) = ST_FWORK Then wCnt = wCnt + 1
130                      End If
140                  Next j
150                  tgt = Int(mWeekN(i) * cnt / 7 + 0.5)
160                  If tgt > cnt Then tgt = cnt
170                  Do
180                      autoW = 0
190                      For j = 1 To mND
200                          If mDayIn(j) And mWkKey(j) = mWkKey(j0) And mPlan(i, j) = ST_WORK Then autoW = autoW + 1
210                      Next j
220                      If wCnt + autoW <= tgt Then Exit Do
230                      bj = 0: bs = -1E+30
240                      For j = 1 To mND
250                          If mDayIn(j) And mWkKey(j) = mWkKey(j0) And mPlan(i, j) = ST_WORK Then
260                              sc = OffScore(i, j)
270                              If sc > bs Then bs = sc: bj = j
280                          End If
290                      Next j
300                      If bj = 0 Then Exit Do
310                      mPlan(i, bj) = ST_OFF
320                      CovAdd i, bj, -1
330                  Loop
340              End If
350          Next j0
360      End If
370      End If
380  Next i

    AS_週N日ルール = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_週N日ルール", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_週N日ルール = False
End Function


'------------------------------------------------------------------
' AS_週リスト
'   出現する週キーを日曜起点で昇順に並べる(以降の週単位処理の軸)。
'------------------------------------------------------------------
Private Function AS_週リスト() As Boolean
    Dim j As Long
    Dim w As Long
    Dim exists As Boolean
    On Error GoTo ErrHandler

    '=== 週リスト(日曜キー昇順) ===
10   ReDim mWkList(1 To mND): mNW = 0
20   For j = 1 To mND
30       If mDayIn(j) Then
40           exists = False
50           For w = 1 To mNW
60               If mWkList(w) = mWkKey(j) Then exists = True: Exit For
70           Next w
80           If Not exists Then mNW = mNW + 1: mWkList(mNW) = mWkKey(j)
90       End If
100  Next j

    AS_週リスト = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_週リスト", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_週リスト = False
End Function
