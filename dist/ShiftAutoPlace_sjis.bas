Attribute VB_Name = "ShiftAutoPlace"
Option Explicit
'==================================================================
'  ShiftAutoPlace v9.9.1
'  公休の配置・均等化アルゴリズムと後半工程。
'  v9.9.1: 玉突きの A 側に CB_対象者か の判定が抜けていたのを修正。
'          固定曜日の人が非固定曜日に引き出され、曜日の約束が
'          崩れる可能性があった(B 側には元から入っていた)。
'  v9.9.0: 1人1日の入替で届かないとき、2人の玉突きを試すようにした。
'          不足日に入れる人が隣接日を手放せば連勤に収まるのに、その
'          隣接日に余裕が無いために抜けない、という詰まり方をしていた。
'          隣接日を別の人が引き受ければ全員の休日数を変えずに解ける。
'          CB_連勤で消えたか が盤面を戻した後に測っており、blkRun と
'          blkOff を取り違えていたのを修正。
'  v9.8.0: 不足を埋める交換のときだけ連勤上限に mRunBonus を足せる
'          ようにした。診断ログが rawPairs=20 / pairs=0 を示し、日単位
'          では成立する交換20組がすべて連勤・連休の上限で消えていた
'          ため。どちらの上限で消えたかを blkRun / blkOff で出す。
'  v9.7.0: AS_連勤緩和 が勤務ルールを見ずに全員へ RepairRuns を
'          かけていたのを修正。固定曜日は対象外にし、週N日は同じ週の
'          中だけで入れ替えるようにした。旧版は週をまたいで日を移す
'          ため、週N日の週あたり勤務日数が崩れていた。
'          CB_診断 に rawPairs(連勤・連休を見る前の組数)を追加。
'  v9.6.0: CoverBalance の連勤・連休の判定を「交換した後の状態」で
'          行うように修正。交換前の状態で見ていたため、抜く日が
'          入れる日の隣にあるとき、まだ出勤のままの抜く日を連勤に
'          数えてしまい、実際には収まる交換まで弾いていた。
'          1手も動かなかったときの切り分け用に CB_診断 を追加し、
'          CoverBalance の LogSuccess に評価値と手数を出す。
'  v9.5.0: 遅番の最低人数を日ごとに決める AP_遅番目標 を追加。
'          混雑日だけ厚くできる(設定が0なら従来どおり全日同じ)。
'          CoverBalance の判定を「不足の二乗和(CB_評価)が下がる交換を
'          選ぶ」方式に置き換え。しきい値 CB_MIN_GAP を廃止したため、
'          過剰+1 の日が残らなくなる。
'          記号の候補選択が同点のとき常に上の行を選び、1人に偏って
'          いたのを、日ごとに走査開始位置をずらして解消。
'  v9.4.0: CoverBalance(日別の過不足を均す工程)を追加。
'          FiveBalance が5診日の出勤数を減らし続ける不具合を修正。
'          事務員の記号割当を「置く記号で候補を並べる」ように修正し、
'          早番人数を設定値(mClerkEarlyN)から読むようにした。
'  v9.3.0: AS_休業行の塗り を追加(休業者の行を灰色にする)。
'  v9.2.0: AS_記号割当(136行)と AS_レポート(106行)を工程ごとに分割。
'          AS_レポート が常に False を返し、呼び出し側の最終 LogSuccess に
'          到達しなかった不具合を修正。
'  ※ ShiftAuto / ShiftAutoPlace / ShiftAutoLog の3本で1組。
'    共有状態は ShiftAuto の Public 変数に置く。
'==================================================================
Private Const MODULE_NAME As String = "ShiftAutoPlace"

' 候補選択の初期値(これより小さいカウントを探す)
Private Const CNT_INF As Long = 32767
' スコア比較の初期値(これより大きいスコアを探す)
Private Const SCORE_INF As Double = -1E+30
' CoverBalance の打ち切り回数。1回の入替で不足の二乗和が必ず減るため
' 通常はこの上限に達する前に収束する(無限ループの保険)
Private Const CB_MAX_PASS As Long = 500
' FiveBalance の打ち切り回数
Private Const FB_MAX_PASS As Long = 100

'--- v9.1.1 追加: 宣言漏れの補完 ---
'  残りの公休ノルマ日数。AS_公休ノルマ で算出し
'  AS_残ノルマ配置 が引き継ぐため、モジュール変数とする。
Private remOff() As Long


'------------------------------------------------------------------
' AS_公休ノルマ
'   通常ルールの公休を配置する。週の基本休(L9)を先に置き、
'   余剰は連休化する。
'------------------------------------------------------------------
Public Function AS_公休ノルマ() As Boolean
    Dim v As String                   ' v9.1.1: 宣言漏れ
    Dim i As Long
    Dim j As Long
    Dim w As Long
    Dim n As Long
    Dim quota As Long
    Dim offN As Long
    Dim tW() As Long
    Dim exW As Long
    Dim dW As Long
    Dim sumB As Long
    Dim needT As Long
    Dim guard As Long
    Dim added As Boolean
    Dim placedB As Boolean
    On Error GoTo ErrHandler

    '=== 通常ルール: 公休ノルマ先行(週(L9)休基本+余剰は連休化) ===
10   ReDim remOff(1 To mNP)
20   For i = 1 To mNP
30       remOff(i) = 0
40       If Not mSkipRow(i) Then
50       If mRule(i) = "通常" And Not mLeave(i) Then
60           quota = mQuota(i): If quota < 0 Then quota = mTargetOff
70           offN = 0
80           For j = 1 To mND
90               If mPlan(i, j) = ST_FOFF Then
100                  v = Trim$(CStr(mGrid.Cells(i, j).Value))
110                  If Not IsPaidOff(v) Then offN = offN + 1
120              End If
130          Next j
140          remOff(i) = quota - offN
150          If remOff(i) < 0 Then remOff(i) = 0
160      End If
170      End If
180  Next i
190  For i = 1 To mNP
200      If remOff(i) > 0 Then
210          needT = remOff(i)
220          ReDim tW(1 To mNW): sumB = 0
230          For w = 1 To mNW
240              exW = 0: dW = 0
250              For j = 1 To mND
260                  If mDayIn(j) And mWkKey(j) = mWkList(w) Then
270                      dW = dW + 1
280                      If mPlan(i, j) = ST_FOFF Or mPlan(i, j) = ST_OFF Then exW = exW + 1
290                  End If
300              Next j
310              tW(w) = mWeekBase - exW
320              If tW(w) > dW - exW Then tW(w) = dW - exW
330              If dW <= 2 And tW(w) > 1 Then tW(w) = 1
340              If tW(w) < 0 Then tW(w) = 0
350              sumB = sumB + tW(w)
360          Next w
370          guard = 0
380          Do While sumB > needT And guard < 100
390              For w = mNW To 1 Step -1
400                  If sumB > needT And tW(w) > 0 Then tW(w) = tW(w) - 1: sumB = sumB - 1
410              Next w
420              guard = guard + 1
430          Loop
440          guard = 0
450          Do While sumB < needT And guard < 100
460              added = False
470              For w = 1 To mNW
480                  If sumB < needT And tW(w) < mMaxOffRun Then tW(w) = tW(w) + 1: sumB = sumB + 1: added = True
490              Next w
500              If Not added Then Exit Do
510              guard = guard + 1
520          Loop
530          For w = 1 To mNW
540              n = tW(w)
550              If n > remOff(i) Then n = remOff(i)
560              Do While n >= 2
570                  placedB = False
580                  If n >= 3 Then
590                      If PlaceOffBlock(i, mWkList(w), 3) Then
600                          remOff(i) = remOff(i) - 3: n = n - 3: placedB = True
610                      End If
620                  End If
630                  If Not placedB Then
640                      If PlaceOffBlock(i, mWkList(w), 2) Then
650                          remOff(i) = remOff(i) - 2: n = n - 2
660                      Else
670                          Exit Do
680                      End If
690                  End If
700              Loop
710              Do While n > 0
720                  If PlaceOffSingle(i, mWkList(w)) Then
730                      remOff(i) = remOff(i) - 1: n = n - 1
740                  Else
750                      Exit Do
760                  End If
770              Loop
780          Next w
790      End If
800  Next i

    AS_公休ノルマ = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_公休ノルマ", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_公休ノルマ = False
End Function


'------------------------------------------------------------------
' AS_残ノルマ配置
'   ノルマ未達分を既存の休みに寄せて1日ずつ配置する(誤差0厳守)。
'------------------------------------------------------------------
Public Function AS_残ノルマ配置() As Boolean
    Dim i As Long
    Dim j As Long
    Dim moved As Boolean
    Dim bj As Long
    Dim bs As Double
    Dim sc As Double
    On Error GoTo ErrHandler

    '=== 残りノルマ: 既存の休みに寄せて1日ずつ必ず配置(誤差0厳守) ===
10   Do
20       moved = False
30       For i = 1 To mNP
40           If remOff(i) > 0 Then
50               bj = 0: bs = -1E+30
60               For j = 1 To mND
70                   If mPlan(i, j) = ST_WORK Then
80                       sc = OffScore(i, j) + AdjBonus(i, j)
90                       If sc > bs Then bs = sc: bj = j
100                  End If
110              Next j
120              If bj > 0 Then
130                  mPlan(i, bj) = ST_OFF
140                  CovAdd i, bj, -1
150                  remOff(i) = remOff(i) - 1
160                  moved = True
170              Else
180                  mUnmet = mUnmet & vbCrLf & "・" & mName(i) & " : あと" & remOff(i) & "日 配置できず"
190                  remOff(i) = 0
200              End If
210          End If
220      Next i
230  Loop While moved

    AS_残ノルマ配置 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_残ノルマ配置", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_残ノルマ配置 = False
End Function


'------------------------------------------------------------------
' AS_連勤緩和
'   連勤上限を超えた箇所を入替で緩和する(最大3巡)。
'------------------------------------------------------------------
Public Function AS_連勤緩和() As Boolean
    Dim i As Long
    Dim guard As Long
    On Error GoTo ErrHandler

    '=== 連勤上限超えの緩和(入替) ===
    '    固定曜日は「その曜日に出る」ことが約束なので触らない。
    '    連勤が上限を超えてもそれが正しい姿で、動かすと約束が崩れる。
10   For guard = 1 To 3
20       For i = 1 To mNP
30           If Not mSkipRow(i) And Not mLeave(i) Then
40               If mRule(i) <> "固定曜日" Then RepairRuns i
50           End If
60       Next i
70   Next guard

    AS_連勤緩和 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_連勤緩和", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_連勤緩和 = False
End Function


'------------------------------------------------------------------
' AS_記号割当
'   出勤日に○●▲を割り当てる。薬剤師は早番(L5)→遅番最低(L6)→
'   残りを均等、事務員は早番1人+2人目以降はL13の記号。
'------------------------------------------------------------------
Public Function AS_記号割当() As Boolean
    Dim j As Long
    On Error GoTo ErrHandler

10  ReDim mSymb(1 To mNP, 1 To mND)
20  ReDim mCntE(1 To mNP): ReDim mCntM(1 To mNP): ReDim mCntL(1 To mNP)
30  AP_既存記号を数える
40  For j = 1 To mND
50      If mDayIn(j) Then
60          AP_薬剤師の記号 j
70          AP_事務員の記号 j
80          AP_残りは早番 j
90      End If
100 Next j

    AS_記号割当 = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_記号割当", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_記号割当 = False
End Function

'------------------------------------------------------------------
' AP_既存記号を数える
'   手入力済み(ST_FWORK)の記号を個人別カウントに取り込む。
'   自動配置はこのカウントを均すように動く。
'------------------------------------------------------------------
Private Sub AP_既存記号を数える()
    Dim i As Long, j As Long, v As String
    On Error GoTo ErrHandler

10  For i = 1 To mNP
20      If Not mSkipRow(i) Then
30          For j = 1 To mND
40              If mPlan(i, j) = ST_FWORK Then
50                  v = Trim$(CStr(mGrid.Cells(i, j).Value))
60                  If IsEarlySym(v) Then mCntE(i) = mCntE(i) + 1
70                  If v = SYM_MID Then mCntM(i) = mCntM(i) + 1
80                  If v = SYM_LATE Then mCntL(i) = mCntL(i) + 1
90              End If
100         Next j
110     End If
120 Next i
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "AP_既存記号を数える", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
End Sub

'------------------------------------------------------------------
' AP_日別既存数
'   その日に手入力済みの記号を区分別に数える。
'   自動で何人足せばよいかの起点になる。
'------------------------------------------------------------------
Private Sub AP_日別既存数(ByVal j As Long, ByVal kind As String, _
                          ByRef nEarly As Long, ByRef nMid As Long, ByRef nLate As Long)
    Dim i As Long, v As String
    On Error GoTo ErrHandler

10  nEarly = 0: nMid = 0: nLate = 0
20  For i = 1 To mNP
30      If Not mSkipRow(i) Then
40          If mKind(i) = kind And mPlan(i, j) = ST_FWORK Then
50              v = Trim$(CStr(mGrid.Cells(i, j).Value))
60              If IsEarlySym(v) Then nEarly = nEarly + 1
70              If v = SYM_MID Then nMid = nMid + 1
80              If v = SYM_LATE Then nLate = nLate + 1
90          End If
100     End If
110 Next i
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "AP_日別既存数", Err.Number, Err.Description, Erl, _
             "j=" & j & "; kind=" & kind & "; i=" & i
End Sub

'------------------------------------------------------------------
' AP_最少候補
'   その日まだ記号が決まっていない人のうち、指定した記号の月合計が
'   最も少ない人を返す(0 = 該当なし)。個人差を詰めるための選び方。
'   lateOnly = True のときは遅番可の人だけを対象にする。
'------------------------------------------------------------------
Private Function AP_最少候補(ByVal j As Long, ByVal kind As String, _
                             ByVal sym As String, ByVal lateOnly As Boolean) As Long
    Dim i As Long, k As Long, best As Long, bestCnt As Long, c As Long
    On Error GoTo ErrHandler

    '--- 走査の開始位置を日ごとにずらす。
    '    同点のとき常に上の行を選ぶと、事務員が2名以上いても
    '    片方に記号が寄り続ける(月合計が並んだ日は必ず先頭が勝つため)。
    '    乱数ではなく日付による巡回にしているのは、同じ入力なら
    '    同じ結果になり、実行を比べられるようにするため。
10  best = 0: bestCnt = CNT_INF
20  For k = 0 To mNP - 1
25      i = ((j + k - 1) Mod mNP) + 1
30      If Not mSkipRow(i) Then
40          If mKind(i) = kind And mPlan(i, j) = ST_WORK And Len(mSymb(i, j)) = 0 Then
50              If (Not lateOnly) Or mCanLate(i) Then
60                  c = SymCnt(i, sym)
70                  If c < bestCnt Then bestCnt = c: best = i
80              End If
90          End If
100     End If
110 Next k
120 AP_最少候補 = best
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_最少候補", Err.Number, Err.Description, Erl, _
             "j=" & j & "; kind=" & kind & "; sym=" & sym & "; i=" & i
    AP_最少候補 = 0
End Function

'------------------------------------------------------------------
' AP_記号を置く
'   記号を確定し、個人別カウントも同時に更新する。
'   (この2つは必ず対で動かす。片方だけ更新すると均等化が壊れる)
'------------------------------------------------------------------
Private Sub AP_記号を置く(ByVal i As Long, ByVal j As Long, ByVal sym As String)
    On Error GoTo ErrHandler

10  mSymb(i, j) = sym
20  AddCnt i, sym, 1
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "AP_記号を置く", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j & "; sym=" & sym
End Sub

'------------------------------------------------------------------
' AP_遅番目標
'   その日の遅番(▲)の最低人数。混雑日(医師 DOC_BUSY_N 名)だけ
'   別の人数にできる。設定が0のときは全日とも通常の最低人数。
'------------------------------------------------------------------
Private Function AP_遅番目標(ByVal j As Long) As Long
    On Error GoTo ErrHandler

10  AP_遅番目標 = mLateMin
20  If mLateBusy > 0 And mDayDoc(j) >= DOC_BUSY_N Then AP_遅番目標 = mLateBusy
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_遅番目標", Err.Number, Err.Description, Erl, _
             "j=" & j & "; lateMin=" & mLateMin & "; lateBusy=" & mLateBusy
    AP_遅番目標 = mLateMin
End Function

'------------------------------------------------------------------
' AP_薬剤師の記号
'   ○(設定人数) → ▲(最低人数) → 残りを ● ▲ で均等、の順に決める。
'------------------------------------------------------------------
Private Sub AP_薬剤師の記号(ByVal j As Long)
    Dim needE As Long, needL As Long
    Dim dayE As Long, dayM As Long, dayL As Long
    Dim bi As Long, symS As String
    On Error GoTo ErrHandler

10  AP_日別既存数 j, KIND_PH, dayE, dayM, dayL

    '--- ○: 設定の早番人数まで ---
20  needE = mEarlyN - dayE
30  Do While needE > 0
40      bi = AP_最少候補(j, KIND_PH, SYM_EARLY, False)
50      If bi = 0 Then Exit Do
60      AP_記号を置く bi, j, SYM_EARLY
70      needE = needE - 1
80  Loop

    '--- ▲: 設定の最低人数まで(遅番可の人のみ・混雑日は別人数) ---
90  needL = AP_遅番目標(j) - dayL
100 Do While needL > 0
110     bi = AP_最少候補(j, KIND_PH, SYM_LATE, True)
120     If bi = 0 Then Exit Do
130     AP_記号を置く bi, j, SYM_LATE
140     dayL = dayL + 1: needL = needL - 1
150 Loop

    '--- 残り: ● と ▲ が同数に近づくよう交互に置く ---
160 Do
170     If dayM <= dayL Then symS = SYM_MID Else symS = SYM_LATE
180     bi = AP_最少候補(j, KIND_PH, symS, True)
190     If bi = 0 Then Exit Do
200     AP_記号を置く bi, j, symS
210     If symS = SYM_MID Then dayM = dayM + 1 Else dayL = dayL + 1
220 Loop
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "AP_薬剤師の記号", Err.Number, Err.Description, Erl, _
             "j=" & j & "; needEarly=" & needE & "; needLate=" & needL
End Sub

'------------------------------------------------------------------
' AP_事務員の記号
'   ○(早番)は設定の人数まで。それ以降は設定の記号(事務員の2人目の記号)。
'   ここでいう早番人数は事務員だけの枠で、薬剤師の早番(mEarlyN)とは別に
'   数える。1日の○の合計は mEarlyN + mClerkEarlyN になる。
'   v9.4.0: 置く記号を先に決め、その記号の月合計が最少の人を選ぶ。
'   旧版は●を置くときも○の回数で候補を並べていた。
'------------------------------------------------------------------
Private Sub AP_事務員の記号(ByVal j As Long)
    Dim dayE As Long, dayM As Long, dayL As Long, bi As Long
    Dim symS As String
    On Error GoTo ErrHandler

10  AP_日別既存数 j, KIND_CL, dayE, dayM, dayL
20  Do
30      If dayE < mClerkEarlyN Then symS = SYM_EARLY Else symS = mGSym
40      bi = AP_最少候補(j, KIND_CL, symS, False)
50      If bi = 0 Then Exit Do
60      AP_記号を置く bi, j, symS
70      If symS = SYM_EARLY Then dayE = dayE + 1
80  Loop
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "AP_事務員の記号", Err.Number, Err.Description, Erl, _
             "j=" & j & "; dayEarly=" & dayE
End Sub

'------------------------------------------------------------------
' AP_残りは早番
'   遅番不可の薬剤師など、ここまでで記号が付かなかった出勤を ○ にする。
'------------------------------------------------------------------
Private Sub AP_残りは早番(ByVal j As Long)
    Dim i As Long
    On Error GoTo ErrHandler

10  For i = 1 To mNP
20      If Not mSkipRow(i) Then
30          If mPlan(i, j) = ST_WORK And Len(mSymb(i, j)) = 0 Then
40              AP_記号を置く i, j, SYM_EARLY
50          End If
60      End If
70  Next i
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "AP_残りは早番", Err.Number, Err.Description, Erl, _
             "j=" & j & "; i=" & i
End Sub


'------------------------------------------------------------------
' AS_書き込み
'   確定した記号をシートに書き、変更前の値を差分ログに残す。
'------------------------------------------------------------------
Public Function AS_書き込み() As Boolean
    Dim i As Long
    Dim j As Long
    Dim lg As Worksheet
    Dim lgR As Long
    Dim tv As String
    Dim cCell As Range
    On Error GoTo ErrHandler

    '=== 書き込み(差分ログを記録しながら) ===
10   Set lg = GetLogSheet()
20   mSess = NextSession(lg)
30   lgR = LogLastRow(lg)
40   mLogged = 0
50   Application.EnableEvents = False
60   Application.ScreenUpdating = False
70   For i = 1 To mNP
80       If Not mSkipRow(i) Then          ' 空行・集計行には一切書き込まない
90           For j = 1 To mND
100              tv = ""
110              If mPlan(i, j) = ST_OFF Then
120                  tv = SYM_OFF
130              ElseIf mPlan(i, j) = ST_WORK Then
140                  tv = mSymb(i, j)
150              End If
160              If Len(tv) > 0 Then
170                  Set cCell = mGrid.Cells(i, j)
180                  If Not cCell.HasFormula Then
190                      If Trim$(CStr(cCell.Value)) <> tv Then
200                          LogChange lg, lgR, mSess, "自動作成", cCell, tv
210                          mLogged = mLogged + 1
220                      End If
230                  End If
240                  StampCell mWs, cCell, tv
250                  mWritten = mWritten + 1
260              End If
270          Next j
280      End If
290  Next i
300  Application.ScreenUpdating = True
310  Application.EnableEvents = True


    AS_書き込み = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_書き込み", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_書き込み = False
End Function



'------------------------------------------------------------------
' AS_休業行の塗り
'   休業(マスタC列の○)のスタッフ行を灰色にし、解除された行は塗りを外す。
'   休業は月ごとに変わるため、毎回すべての行を見て塗り直す。
'
'   塗りを外すのは「マクロが塗った色と同じ場合」だけにする。
'   利用者がパレットの背景色ボタンで付けた色を消さないため。
'   氏名(A列)も含めて塗る。誰が休業かを行の左端で見分けられるようにする。
'------------------------------------------------------------------
Public Function AS_休業行の塗り() As Boolean
    Dim i As Long, r As Long, c As Range, rng As Range
    Dim painted As Long, cleared As Long
    On Error GoTo ErrHandler

10  For i = 1 To mNP
20      If Not mSkipRow(i) Then
30          r = mGrid.Row + i - 1
            '--- A列(氏名)から入力欄の右端まで ---
40          Set rng = mWs.Range(mWs.Cells(r, 1), _
                               mWs.Cells(r, mGrid.Column + mGrid.Columns.Count - 1))
50          For Each c In rng.Cells
60              If mLeave(i) Then
70                  c.Interior.Color = ClrLeaveBg()
80              ElseIf c.Interior.Pattern <> xlNone Then
90                  If c.Interior.Color = ClrLeaveBg() Then c.Interior.Pattern = xlNone
100             End If
110         Next c
120         If mLeave(i) Then painted = painted + 1 Else cleared = cleared + 1
130     End If
140 Next i

    AS_休業行の塗り = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_休業行の塗り", Err.Number, Err.Description, Erl, _
             "i=" & i & "; row=" & r & "; painted=" & painted
    AS_休業行の塗り = False
End Function

'------------------------------------------------------------------
' AS_レポート
'   実行結果(書込数・公休数・過不足・遅番不足など)を集計して表示する。
'------------------------------------------------------------------
Public Function AS_レポート() As Boolean
    Dim msg As String
    On Error GoTo ErrHandler

10  msg = AP_レポート見出し()
20  msg = msg & AP_レポート個人別()
30  msg = msg & AP_レポート警告()
40  msg = msg & vbCrLf & "変更ログ: セッション#" & mSess & " に " & mLogged & "セル記録。" & vbCrLf & _
              "取り消すには「シフト変更を戻す」を実行してください。"
50  MsgBox msg, vbInformation, "シフト自動作成 結果"

    LogSuccess MODULE_NAME, "AS_レポート", _
               "Reported result for " & Format(mMonthDt, "yyyy-mm") & _
               "; written=" & mWritten
    AS_レポート = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AS_レポート", Err.Number, Err.Description, Erl, _
             "np=" & mNP & "; nd=" & mND
    AS_レポート = False
End Function

'--- レポートの見出し(対象月・ノルマ・書込セル数・対象者) ---
Private Function AP_レポート見出し() As String
    On Error GoTo ErrHandler

10  AP_レポート見出し = "対象月: " & Format(mMonthDt, "yyyy年m月") & _
        "　公休ノルマ(土日祝): " & mTargetOff & "日　書込セル: " & mWritten & vbCrLf & _
        "入力範囲: " & mGrid.Address(False, False) & vbCrLf & _
        "対象者: " & mActiveN & "名" & _
        IIf(mSkipN > 0, "　(スキップ行: " & mSkipN & "行)", "") & vbCrLf & vbCrLf
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_レポート見出し", Err.Number, Err.Description, Erl, _
             "np=" & mNP
    AP_レポート見出し = ""
End Function

'--- 個人別の集計行(出勤・休み・連勤/連休max・記号の内訳) ---
Private Function AP_レポート個人別() As String
    Dim i As Long, j As Long, msg As String
    Dim offC As Long, wC As Long, paidC As Long
    On Error GoTo ErrHandler

10  For i = 1 To mNP
20      If mSkipRow(i) Then
            ' 空行・集計行は結果に出さない
30      ElseIf mLeave(i) Then
40          msg = msg & mName(i) & " : 休業(スキップ)" & vbCrLf
50      Else
60          offC = 0: wC = 0: paidC = 0
70          For j = 1 To mND
                Select Case mPlan(i, j)
                    Case ST_WORK, ST_FWORK
                        wC = wC + 1
                    Case ST_OFF, ST_FOFF
                        offC = offC + 1
                        If mPlan(i, j) = ST_FOFF Then
                            If IsPaidOff(Trim$(CStr(mGrid.Cells(i, j).Value))) Then paidC = paidC + 1
                        End If
                End Select
80          Next j
90          msg = msg & mName(i) & " : 出勤" & wC & " 休" & offC & _
                IIf(paidC > 0, "(うちノルマ外" & paidC & ")", "") & _
                " 連勤max" & MaxRun(i) & " 連休max" & MaxOffRun(i) & _
                IIf(mKind(i) = KIND_PH, " 医5日" & FiveCnt(i) & _
                    " ○" & mCntE(i) & " ●" & mCntM(i) & " ▲" & mCntL(i), "") & vbCrLf
100     End If
110 Next i
120 AP_レポート個人別 = msg
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_レポート個人別", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
    AP_レポート個人別 = msg
End Function

'--- 警告(ノルマ未達・事務員不在・遅番不足・必要数不足・設定の不一致) ---
Private Function AP_レポート警告() As String
    Dim msg As String
    Dim gZero As Long, lateShort As Long, lateBad As Long
    Dim shortN As Long, worst As Long
    On Error GoTo ErrHandler

10  If Len(mUnmet) > 0 Then
20      msg = msg & vbCrLf & "■ 公休ノルマ未達(空きセル不足):" & mUnmet & vbCrLf
30  End If

40  gZero = AP_事務員不在日数()
50  If gZero > 0 Then
60      msg = msg & vbCrLf & "■ 事務員が不在の日: " & gZero & "日" & vbCrLf
70  End If

80  lateShort = AP_遅番不足日数(lateBad)
90  If lateShort > 0 Then
100     msg = msg & vbCrLf & "※遅番(" & SYM_LATE & ")が目標" & mLateMin & "名" & _
              IIf(mLateBusy > 0, "(" & DOC_BUSY_N & "診日は" & mLateBusy & "名)", "") & _
              "に届かない日: " & lateShort & "日" & _
              IIf(lateBad > 0, "　■ うち目標-1名未満: " & lateBad & "日", "") & vbCrLf
110 End If

115 msg = msg & AP_連勤上乗せの影響()
120 shortN = AP_必要数不足日数(worst)
130 If shortN > 0 Then
140     msg = msg & vbCrLf & "※必要数(医師数+" & mReqPlus & ")に届かない日: " & shortN & _
              "日(最大不足" & worst & "名)" & vbCrLf & AP_人日収支()
150 End If

160 If Len(mMissing) > 0 Then msg = msg & vbCrLf & "設定未登録(既定値で処理):" & vbCrLf & mMissing
170 If Len(mOrphan) > 0 Then msg = msg & vbCrLf & "マスタにあるがシフト表に無い:" & vbCrLf & mOrphan
180 AP_レポート警告 = msg
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_レポート警告", Err.Number, Err.Description, Erl, _
             "nd=" & mND
    AP_レポート警告 = msg
End Function

'--- 連勤の上乗せを使った結果、通常上限を超えた人を出す ---
'    上乗せは労務上の例外なので、誰が何日になったかを必ず見せる。
Private Function AP_連勤上乗せの影響() As String
    Dim i As Long, n As Long, worst As Long, r As Long, names As String
    On Error GoTo ErrHandler

10  If mRunBonus <= 0 Then Exit Function
20  For i = 1 To mNP
30      If Not mSkipRow(i) And Not mLeave(i) Then
40          r = MaxRun(i)
50          If r > mMaxRun Then
60              n = n + 1
70              If r > worst Then worst = r
80              names = names & "・" & mName(i) & "(" & r & "日)" & vbCrLf
90          End If
100     End If
110 Next i
120 If n = 0 Then Exit Function
130 AP_連勤上乗せの影響 = vbCrLf & "■ 連勤が通常の上限" & mMaxRun & "日を超えた人: " & _
        n & "名(最大" & worst & "日)" & vbCrLf & names & _
        "　→ 不足日を埋めるため上限に" & mRunBonus & "日上乗せしています。" & vbCrLf
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_連勤上乗せの影響", Err.Number, Err.Description, Erl, _
             "i=" & i & "; n=" & n
    AP_連勤上乗せの影響 = ""
End Function

'--- 必要数不足が「人手不足」か「配分の偏り」かを1行で示す ---
'    月の総必要人日と総出勤人日を比べる。総出勤が足りていれば配置を
'    変える余地があり、足りていなければ人を増やす以外に解がない。
Private Function AP_人日収支() As String
    Dim j As Long, totReq As Long, totCov As Long, gap As Long
    On Error GoTo ErrHandler

10  For j = 1 To mND
20      If mDayIn(j) Then
30          totReq = totReq + mDayReq(j)
40          totCov = totCov + mCov(j)
50      End If
60  Next j
70  gap = totReq - totCov
80  AP_人日収支 = "　月の必要人日: " & totReq & "　出勤人日: " & totCov & vbCrLf & _
        IIf(gap > 0, _
            "　→ 人日そのものが" & gap & "日分足りません。" & _
            "公休ノルマを減らすか人を増やさない限り解消しません。", _
            "　→ 人日は足りています。日別の配分の問題です。") & vbCrLf
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_人日収支", Err.Number, Err.Description, Erl, _
             "j=" & j & "; totalReq=" & totReq & "; totalCov=" & totCov
    AP_人日収支 = ""
End Function

'--- 事務員が1人もいない日の数 ---
Private Function AP_事務員不在日数() As Long
    Dim j As Long, n As Long
    On Error GoTo ErrHandler

10  For j = 1 To mND
20      If mDayIn(j) And mCovG(j) < 1 Then n = n + 1
30  Next j
40  AP_事務員不在日数 = n
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_事務員不在日数", Err.Number, Err.Description, Erl, _
             "j=" & j
    AP_事務員不在日数 = 0
End Function

'--- その日の遅番(▲)の人数(自動割当分と手入力分の合計) ---
Private Function AP_日別遅番数(ByVal j As Long) As Long
    Dim i As Long, n As Long
    On Error GoTo ErrHandler

10  For i = 1 To mNP
20      If Not mSkipRow(i) Then
30          If mKind(i) = KIND_PH Then
40              If mPlan(i, j) = ST_WORK Then
50                  If mSymb(i, j) = SYM_LATE Then n = n + 1
60              ElseIf mPlan(i, j) = ST_FWORK Then
70                  If Trim$(CStr(mGrid.Cells(i, j).Value)) = SYM_LATE Then n = n + 1
80              End If
90          End If
100     End If
110 Next i
120 AP_日別遅番数 = n
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_日別遅番数", Err.Number, Err.Description, Erl, _
             "j=" & j & "; i=" & i
    AP_日別遅番数 = 0
End Function

'--- 遅番が目標に届かない日数。bad には「目標-1 名未満」の日数を返す ---
Private Function AP_遅番不足日数(ByRef bad As Long) As Long
    Dim j As Long, n As Long, dayLate As Long, tgt As Long
    On Error GoTo ErrHandler

10  bad = 0
20  For j = 1 To mND
30      If mDayIn(j) Then
40          dayLate = AP_日別遅番数(j): tgt = AP_遅番目標(j)
50          If dayLate < tgt Then n = n + 1
60          If dayLate < tgt - 1 Then bad = bad + 1
70      End If
80  Next j
90  AP_遅番不足日数 = n
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_遅番不足日数", Err.Number, Err.Description, Erl, _
             "j=" & j
    AP_遅番不足日数 = n
End Function

'--- 必要数に届かない日数。worst には最大の不足人数を返す ---
Private Function AP_必要数不足日数(ByRef worst As Long) As Long
    Dim j As Long, n As Long
    On Error GoTo ErrHandler

10  worst = 0
20  For j = 1 To mND
30      If mDayIn(j) And mCov(j) < mDayReq(j) Then
40          n = n + 1
50          If mDayReq(j) - mCov(j) > worst Then worst = mDayReq(j) - mCov(j)
60      End If
70  Next j
80  AP_必要数不足日数 = n
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AP_必要数不足日数", Err.Number, Err.Description, Erl, _
             "j=" & j
    AP_必要数不足日数 = n
End Function


'--- その日を休みにする良さ(大きいほど休み向き) ---
Public Function OffScore(ByVal i As Long, ByVal j As Long) As Double
    On Error GoTo ErrHandler
    Dim s As Double, L As Long, lft As Long, rgt As Long
    If mKind(i) = KIND_PH Then
        s = s + 5# * (mCov(j) - 1 - mDayReq(j))          ' 不足を日別に均す(ソフト)
        If mDayDoc(j) = 5 Then s = s + 3# * (FiveCnt(i) - FiveAvg())
    ElseIf mKind(i) = KIND_CL Then
        If mCovG(j) - 1 < 1 Then
            s = s - 12                                    ' 事務員ゼロの日は強く回避
        Else
            s = s + 4# * (mCovG(j) - 2)                   ' 重なる日を優先して休みに
        End If
    End If
    L = RunLenAt(i, j, lft, rgt)
    If L >= mMaxRun + 1 Then s = s + 8 + IIf(lft < rgt, lft, rgt)
    If mDayWD(j) = 1 Or mDayWD(j) = 7 Or mDayHol(j) Then s = s + 2
    OffScore = s
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "OffScore", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
End Function


'--- 既存休に隣接して連休(上限内)になる位置を優遇 ---
Private Function AdjBonus(ByVal i As Long, ByVal j As Long) As Double
    On Error GoTo ErrHandler
    Dim total As Long
    total = 1 + OffRunBefore(i, j) + OffRunAfter(i, j)
    If total >= 2 And total <= mMaxOffRun Then
        AdjBonus = 4
    ElseIf total > mMaxOffRun Then
        AdjBonus = -3 * (total - mMaxOffRun)
    End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "AdjBonus", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
End Function


'--- 指定週内に size 日連続の公休ブロックを最良位置へ ---
Private Function PlaceOffBlock(ByVal i As Long, ByVal wk As Long, ByVal size As Long) As Boolean
    On Error GoTo ErrHandler
    Dim j As Long, k As Long, ok As Boolean
    Dim s As Double, bs As Double, bj As Long, offRun As Long
    bj = 0: bs = -1E+30
    For j = 1 To mND - size + 1
        ok = True
        For k = j To j + size - 1
            If Not mDayIn(k) Then ok = False
            If ok Then If mWkKey(k) <> wk Then ok = False
            If ok Then If mPlan(i, k) <> ST_WORK Then ok = False
            If Not ok Then Exit For
        Next k
        If ok Then
            s = 0
            For k = j To j + size - 1
                s = s + OffScore(i, k)
            Next k
            offRun = size + OffRunBefore(i, j) + OffRunAfter(i, j + size - 1)
            If offRun > mMaxOffRun Then s = s - 4 * (offRun - mMaxOffRun)
            If s > bs Then bs = s: bj = j
        End If
    Next j
    If bj > 0 Then
        For k = bj To bj + size - 1
            mPlan(i, k) = ST_OFF
            CovAdd i, k, -1
        Next k
        PlaceOffBlock = True
    End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "PlaceOffBlock", Err.Number, Err.Description, Erl, _
             "i=" & i & "; wk=" & wk & "; size=" & size
End Function


'--- 指定週内に公休1日を最良位置へ(既存休に寄せる) ---
Private Function PlaceOffSingle(ByVal i As Long, ByVal wk As Long) As Boolean
    On Error GoTo ErrHandler
    Dim j As Long, bj As Long, bs As Double, sc As Double
    bj = 0: bs = -1E+30
    For j = 1 To mND
        If mDayIn(j) And mWkKey(j) = wk And mPlan(i, j) = ST_WORK Then
            sc = OffScore(i, j) + AdjBonus(i, j)
            If sc > bs Then bs = sc: bj = j
        End If
    Next j
    If bj > 0 Then
        mPlan(i, bj) = ST_OFF
        CovAdd i, bj, -1
        PlaceOffSingle = True
    End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "PlaceOffSingle", Err.Number, Err.Description, Erl, _
             "i=" & i & "; wk=" & wk
End Function


'--- 連勤上限超えの緩和: 連勤の中央を休みにし、他の自動公休と入替 ---
'    v9.7.0: 週N日の人は入替先を同じ週に限る。月内のどこへでも移すと
'    週あたりの勤務日数が変わり、週N日の指定が崩れるため。
'    (週4日は4連勤になり得るので、この処理が実際に発火する)
Private Sub RepairRuns(ByVal i As Long)
    On Error GoTo ErrHandler
    Dim j As Long, L As Long, lft As Long, rgt As Long
    Dim ctr As Long, off As Long, k As Long, cand As Long
    Dim bk As Long, bs As Double, sc As Double
    j = 1
    Do While j <= mND
        If mPlan(i, j) = ST_WORK Or mPlan(i, j) = ST_FWORK Then
            L = RunLenAt(i, j, lft, rgt)
            j = j - lft
            If L >= mMaxRun + 1 Then
                ctr = j + (L - 1) \ 2
                cand = 0
                For off = 0 To L - 1
                    k = ctr + ((off + 1) \ 2) * IIf(off Mod 2 = 0, 1, -1)
                    If k >= j And k <= j + L - 1 Then
                        If mPlan(i, k) = ST_WORK Then cand = k: Exit For
                    End If
                Next off
                If cand > 0 Then
                    bk = 0: bs = -1E+30
                    For k = 1 To mND
                        If mPlan(i, k) = ST_OFF And RR_同じ週か(i, k, cand) Then
                            If WorkRunIf(i, k) <= mMaxRun Then
                                sc = 0
                                If mKind(i) = KIND_PH Then sc = mDayReq(k) - mCov(k)
                                If mKind(i) = KIND_CL Then sc = 1 - mCovG(k)
                                If sc > bs Then bs = sc: bk = k
                            End If
                        End If
                    Next k
                    If bk > 0 Then
                        mPlan(i, cand) = ST_OFF
                        mPlan(i, bk) = ST_WORK
                        CovAdd i, cand, -1
                        CovAdd i, bk, 1
                    End If
                End If
            End If
            j = j + L
        Else
            j = j + 1
        End If
    Loop
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "RepairRuns", Err.Number, Err.Description, Erl, "i=" & i
End Sub


'--- 入替先として使ってよい日か(週N日の人は同じ週の中だけ) ---
Private Function RR_同じ週か(ByVal i As Long, ByVal k As Long, _
                             ByVal cand As Long) As Boolean
    On Error GoTo ErrHandler

10  If mRule(i) <> "週N日" Then
20      RR_同じ週か = True
30  Else
40      RR_同じ週か = (mWkKey(k) = mWkKey(cand))
50  End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "RR_同じ週か", Err.Number, Err.Description, Erl, _
             "i=" & i & "; k=" & k & "; cand=" & cand
    RR_同じ週か = False
End Function


'==================================================================
'  日別の過不足を均す (v9.4.0)
'    個人の休日数は変えずに、余裕のある日の出勤を不足している日へ移す。
'    不足が大きい日から順に埋めるため、必要数(医師数+n)の大きい日
'    すなわち5診日から先に埋まり、次に4診日・3診日となる。
'==================================================================
Public Sub CoverBalance()
    Dim pass As Long, moves As Long, chains As Long
    Dim before As Double
    On Error GoTo ErrHandler

    '--- まず1人1日の入替を試し、届かなければ2人の玉突きに落とす。
    '    玉突きは探索が広いので、単独で解ける限りは使わない ---
10  before = CB_評価()
20  For pass = 1 To CB_MAX_PASS
30      If CB_1名移す() Then
40          moves = moves + 1
50      ElseIf CB_2名移す() Then
60          moves = moves + 1: chains = chains + 1
70      Else
80          Exit For
90      End If
100 Next pass

    '--- 1手も動かないときに、どこで詰まっているかを残す ---
    LogSuccess MODULE_NAME, "CoverBalance", _
               "score " & before & " -> " & CB_評価() & "; moves=" & moves & _
               "; chains=" & chains & "; " & CB_診断()
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "CoverBalance", Err.Number, Err.Description, Erl, _
             "pass=" & pass & "; moves=" & moves & "; chains=" & chains
End Sub

'--- 交換の候補が何組あるかを数える(1手も動かないときの切り分け用) ---
'    shortDays  不足している日数
'    movable    日を動かしてよい人数
'    canWork    公休を出勤に変えられる (人,日) の組
'    canRest    出勤を公休に変えられる (人,日) の組(日単位の条件のみ)
'    rawPairs   日単位の条件だけで成立する交換の組
'    pairs      連勤・連休まで見て実際に成立する交換の組
'    blkRun     連勤の上限で消えた組
'    blkOff     連休の上限で消えた組(連勤は通ったもの)
'    canWork か canRest が 0 ならそちら側で詰まっている。
'    rawPairs=0 なら「その日に出られる人」と「休める日を持つ人」が
'    別人で、1人1日の入替では届かない(2人の玉突きが要る)。
'    rawPairs>0 で pairs=0 なら上限で弾かれている。blkRun が大きければ
'    「不足を埋めるときの連勤上限の上乗せ」を1にすると届く。
Private Function CB_診断() As String
    Dim i As Long, j As Long, jf As Long
    Dim nShort As Long, nMove As Long, nIn As Long, nOut As Long
    Dim nPair As Long, nRaw As Long, nBlkR As Long, nBlkO As Long
    On Error GoTo ErrHandler

10  For j = 1 To mND
20      If mDayIn(j) Then
30          If mDayReq(j) - mCov(j) > 0 Then nShort = nShort + 1
40      End If
50  Next j
60  For i = 1 To mNP
70      If CB_対象者か(i) Then
80          nMove = nMove + 1
90          For j = 1 To mND
100             If CB_入れられるか(i, j) Then
110                 nIn = nIn + 1
120                 For jf = 1 To mND
130                     If CB_抜けるか(i, jf, j) Then
132                         nRaw = nRaw + 1
134                         If CB_交換できるか(i, j, jf) Then
136                             nPair = nPair + 1
138                         ElseIf CB_連勤で消えたか(i, j, jf) Then
140                             nBlkR = nBlkR + 1
142                         Else
144                             nBlkO = nBlkO + 1
146                         End If
150                     End If
160                 Next jf
170             End If
180             If CB_休みにできるか(i, j) Then nOut = nOut + 1
190         Next j
200     End If
210 Next i
220 CB_診断 = "shortDays=" & nShort & "; movable=" & nMove & _
             "; canWork=" & nIn & "; canRest=" & nOut & _
             "; rawPairs=" & nRaw & "; pairs=" & nPair & _
             "; blkRun=" & nBlkR & "; blkOff=" & nBlkO & _
             "; runLimit=" & CB_連勤上限()
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_診断", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
    CB_診断 = "(diagnosis failed)"
End Function

'--- 交換が消えた理由が連勤側か(診断用。連休側と区別する) ---
'    CB_交換できるか は判定後に盤面を戻すため、戻ったあとで測ると
'    抜く日が出勤のままになり連勤を多く数える。ここでも入れ替えてから測る。
Private Function CB_連勤で消えたか(ByVal i As Long, ByVal jTo As Long, _
                                   ByVal jFrom As Long) As Boolean
    Dim over As Boolean
    On Error GoTo ErrHandler

10  mPlan(i, jTo) = ST_WORK
20  mPlan(i, jFrom) = ST_OFF
30  over = (WorkRunIf(i, jTo) > CB_連勤上限())
40  mPlan(i, jTo) = ST_OFF
50  mPlan(i, jFrom) = ST_WORK
60  CB_連勤で消えたか = over
    Exit Function
ErrHandler:
    mPlan(i, jTo) = ST_OFF
    mPlan(i, jFrom) = ST_WORK
    LogError MODULE_NAME, "CB_連勤で消えたか", Err.Number, Err.Description, Erl, _
             "i=" & i & "; jTo=" & jTo & "; jFrom=" & jFrom
    CB_連勤で消えたか = False
End Function

'--- 日別カバレッジの評価値。小さいほど良い ---
'    不足を二乗して合計する。不足2の日1つ(=4)より、不足1の日2つ(=2)の
'    ほうが良いと評価されるため、不足が1日に固まらない。
'    過剰は0扱い(余っていること自体は害ではない)。
Private Function CB_評価() As Double
    Dim j As Long, d As Long, s As Double
    On Error GoTo ErrHandler

10  For j = 1 To mND
20      If mDayIn(j) Then
30          d = mDayReq(j) - mCov(j)
40          If d > 0 Then s = s + CDbl(d) * d
50      End If
60  Next j
70  CB_評価 = s
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_評価", Err.Number, Err.Description, Erl, "j=" & j
    CB_評価 = 0
End Function

'--- 評価が最も改善する1手を打つ。改善しなければ False ---
'    交換は「必要数を保てる日から抜いて、不足している日へ入れる」形に
'    限る(CB_抜けるか)。抜いた日は交換の前後どちらも不足0なので評価に
'    影響せず、交換後の評価は base - (2*d - 1) で確定する
'    (d = 入れる日の不足)。総当たりで CB_評価 を呼び直した場合と同じ
'    答えになり、人数×日数×日数×日数 の総当たりを避けられる。
'    ※ CB_抜けるか の条件を緩めるときは、この式も見直すこと。
Private Function CB_1名移す() As Boolean
    Dim i As Long, jTo As Long, jFrom As Long, d As Long
    Dim base As Double, best As Double, cur As Double
    Dim bi As Long, bf As Long, bt As Long
    On Error GoTo ErrHandler

10  base = CB_評価()
20  best = base: bi = 0: bf = 0: bt = 0
30  For i = 1 To mNP
40      If CB_対象者か(i) Then
50          For jTo = 1 To mND
60              If CB_入れられるか(i, jTo) Then
70                  d = mDayReq(jTo) - mCov(jTo)
80                  cur = base - (2 * CDbl(d) - 1)
90                  If cur < best Then
100                     jFrom = CB_抜ける日(i, jTo)
110                     If jFrom > 0 Then best = cur: bi = i: bf = jFrom: bt = jTo
120                 End If
130             End If
140         Next jTo
150     End If
160 Next i

170 If bi = 0 Then Exit Function
180 mPlan(bi, bt) = ST_WORK
190 mPlan(bi, bf) = ST_OFF
200 CovAdd bi, bt, 1
210 CovAdd bi, bf, -1
220 CB_1名移す = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_1名移す", Err.Number, Err.Description, Erl, _
             "i=" & i & "; from=" & jFrom & "; to=" & jTo
End Function

'--- 自動作成で日を動かしてよい人か ---
'    手動(派遣など)は別軸で決まるので触らない。
'    固定曜日は「その曜日に出る」ことが約束なので、日を移すと約束が
'    崩れる。週N日は週内での移動だけ許す(CB_抜けるか で見る)。
Private Function CB_対象者か(ByVal i As Long) As Boolean
    On Error GoTo ErrHandler

10  If mSkipRow(i) Then Exit Function
20  If mLeave(i) Then Exit Function
30  If mKind(i) <> KIND_PH Then Exit Function
40  If mRule(i) = "手動" Then Exit Function
50  If mRule(i) = "固定曜日" Then Exit Function
60  CB_対象者か = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_対象者か", Err.Number, Err.Description, Erl, "i=" & i
    CB_対象者か = False
End Function

'--- その日を出勤にできるか(日単位の条件だけ) ---
'    自動で置いた公休(ST_OFF)だけを動かす。希望休・有休(ST_FOFF)は
'    絶対に触らない。
'    連勤の上限はここでは見ない。抜く日が隣にあると、まだ出勤のままの
'    抜く日を連勤に数えてしまうため、交換を決めた後に CB_交換できるか
'    で「交換した後の状態」を見る。
Private Function CB_入れられるか(ByVal i As Long, ByVal j As Long) As Boolean
    On Error GoTo ErrHandler

10  If Not mDayIn(j) Then Exit Function
20  If mPlan(i, j) <> ST_OFF Then Exit Function
30  If mDayReq(j) - mCov(j) <= 0 Then Exit Function
40  CB_入れられるか = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_入れられるか", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
    CB_入れられるか = False
End Function

'--- その日を公休にできるか(日単位の条件だけ) ---
'    自動で置いた出勤(ST_WORK)だけを動かす。既存入力(ST_FWORK)は触らない。
'    抜いた後もその日が必要数を保てることを条件にする(不足の付け替えを防ぐ)。
'    連休の上限はここでは見ない(理由は CB_入れられるか と同じ)。
Private Function CB_休みにできるか(ByVal i As Long, ByVal j As Long) As Boolean
    On Error GoTo ErrHandler

10  If Not mDayIn(j) Then Exit Function
20  If mPlan(i, j) <> ST_WORK Then Exit Function
30  If mCov(j) - 1 < mDayReq(j) Then Exit Function
40  CB_休みにできるか = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_休みにできるか", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
    CB_休みにできるか = False
End Function

'--- jTo へ回すために、その日を公休にしてよいか ---
Private Function CB_抜けるか(ByVal i As Long, ByVal j As Long, _
                             ByVal jTo As Long) As Boolean
    On Error GoTo ErrHandler

10  If j = jTo Then Exit Function
20  If Not CB_休みにできるか(i, j) Then Exit Function
    '--- 週N日の人は週をまたいで動かすと週の勤務日数が変わる ---
30  If mRule(i) = "週N日" Then
40      If mWkKey(j) <> mWkKey(jTo) Then Exit Function
50  End If
60  CB_抜けるか = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_抜けるか", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j & "; jTo=" & jTo
    CB_抜けるか = False
End Function

'--- 不足を埋める交換のときに許す連勤の上限 ---
'    通常の配置(公休ノルマ・連勤緩和)は mMaxRun を守る。ここだけは
'    設定した日数まで上乗せを許し、必要数に届かない日を埋める。
'    労務上は例外的な措置なので、既定は上乗せなし(0)。
Private Function CB_連勤上限() As Long
    On Error GoTo ErrHandler

10  CB_連勤上限 = mMaxRun + mRunBonus
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_連勤上限", Err.Number, Err.Description, Erl, _
             "maxRun=" & mMaxRun & "; bonus=" & mRunBonus
    CB_連勤上限 = mMaxRun
End Function

'--- 実際に入れ替えてみて連勤・連休の上限に収まるか(必ず元に戻す) ---
'    交換前の状態で見ると、抜く日が入れる日の隣にあるとき、
'    まだ出勤のままの抜く日を連勤に数えてしまう。
'    例: 8/22公休 8/23出勤 8/24公休 の人を 8/23→8/24 に移すとき、
'    交換前に見ると 8/23 から続く連勤に 8/24 が足され、上限を超えたと
'    判定されて弾かれる。交換後は 8/23 が公休なので実際には収まる。
Private Function CB_交換できるか(ByVal i As Long, ByVal jTo As Long, _
                                 ByVal jFrom As Long) As Boolean
    Dim ok As Boolean
    On Error GoTo ErrHandler

10  mPlan(i, jTo) = ST_WORK
20  mPlan(i, jFrom) = ST_OFF
30  ok = (WorkRunIf(i, jTo) <= CB_連勤上限()) And (OffRunIf(i, jFrom) <= mMaxOffRun)
40  mPlan(i, jTo) = ST_OFF
50  mPlan(i, jFrom) = ST_WORK
60  CB_交換できるか = ok
    Exit Function
ErrHandler:
    '--- 途中で落ちても必ず元の状態に戻す ---
    mPlan(i, jTo) = ST_OFF
    mPlan(i, jFrom) = ST_WORK
    LogError MODULE_NAME, "CB_交換できるか", Err.Number, Err.Description, Erl, _
             "i=" & i & "; jTo=" & jTo & "; jFrom=" & jFrom
    CB_交換できるか = False
End Function

'--- jTo へ回すために公休へ変える出勤日。最も余裕のある日を返す ---
'    どの日を選んでも評価値は同じなので、余裕が大きい日から先に削る。
Private Function CB_抜ける日(ByVal i As Long, ByVal jTo As Long) As Long
    Dim j As Long, best As Long, bs As Long, s As Long
    On Error GoTo ErrHandler

10  best = 0: bs = -1
20  For j = 1 To mND
30      If CB_抜けるか(i, j, jTo) Then
40          If CB_交換できるか(i, jTo, j) Then
50              s = mCov(j) - mDayReq(j)
60              If s > bs Then bs = s: best = j
70          End If
80      End If
90  Next j
100 CB_抜ける日 = best
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_抜ける日", Err.Number, Err.Description, Erl, _
             "i=" & i & "; jTo=" & jTo & "; j=" & j
    CB_抜ける日 = 0
End Function


'==================================================================
'  2人の玉突き (v9.9.0)
'    1人1日の入替で不足日を埋められないときの逃げ道。
'
'      A: 不足日 D を出勤にし、代わりに中継日 X を公休にする
'      B: その X を出勤にし、代わりに余裕のある Y を公休にする
'
'    D は +1、X は A が抜けて B が入るので差し引き 0、Y は -1。
'    Y は余裕のある日に限るので不足には転じない。
'    A も B も出勤日と公休日を1対1で入れ替えるだけなので、
'    誰の月間休日数も変わらない。
'
'    X に余裕が要らないのが肝心なところ。単独の入替では X を手放せず、
'    そのせいで A の連勤が縮まらずに弾かれていた。
'==================================================================
Private Function CB_2名移す() As Boolean
    Dim d As Long, a As Long, x As Long, b As Long, y As Long
    On Error GoTo ErrHandler

10  For d = 1 To mND
20      If CB_不足日か(d) Then
30          For a = 1 To mNP
40              If CB_対象者か(a) And CB_入れられるか(a, d) Then
50                  For x = 1 To mND
60                      If CB_中継日か(a, d, x) Then
70                          For b = 1 To mNP
80                              If b <> a And CB_引き受けられるか(b, x) Then
90                                  y = CB_抜ける日(b, x)
100                                 If y > 0 Then
110                                     CB_玉突きを打つ a, d, x, b, y
120                                     CB_2名移す = True
130                                     Exit Function
140                                 End If
150                             End If
160                         Next b
170                     End If
180                 Next x
190             End If
200         Next a
210     End If
220 Next d
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_2名移す", Err.Number, Err.Description, Erl, _
             "d=" & d & "; a=" & a & "; x=" & x & "; b=" & b
End Function

'--- 必要数に届いていない日か ---
Private Function CB_不足日か(ByVal j As Long) As Boolean
    On Error GoTo ErrHandler

10  If Not mDayIn(j) Then Exit Function
20  CB_不足日か = (mDayReq(j) - mCov(j) > 0)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_不足日か", Err.Number, Err.Description, Erl, "j=" & j
    CB_不足日か = False
End Function

'--- A が不足日 D と引き換えに手放せる中継日か ---
'    ここでは日の余裕を問わない。空いた分は B が埋めるため。
Private Function CB_中継日か(ByVal i As Long, ByVal jTo As Long, _
                             ByVal j As Long) As Boolean
    On Error GoTo ErrHandler

10  If j = jTo Then Exit Function
20  If Not mDayIn(j) Then Exit Function
30  If mPlan(i, j) <> ST_WORK Then Exit Function
    '--- 週N日の人は週をまたいで動かすと週の勤務日数が変わる ---
40  If mRule(i) = "週N日" Then
50      If mWkKey(j) <> mWkKey(jTo) Then Exit Function
60  End If
70  CB_中継日か = CB_交換できるか(i, jTo, j)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_中継日か", Err.Number, Err.Description, Erl, _
             "i=" & i & "; jTo=" & jTo & "; j=" & j
    CB_中継日か = False
End Function

'--- B が中継日 X を引き受けられるか(その日が自動の公休か) ---
Private Function CB_引き受けられるか(ByVal i As Long, ByVal j As Long) As Boolean
    On Error GoTo ErrHandler

10  If Not CB_対象者か(i) Then Exit Function
20  If Not mDayIn(j) Then Exit Function
30  If mPlan(i, j) <> ST_OFF Then Exit Function
40  CB_引き受けられるか = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CB_引き受けられるか", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
    CB_引き受けられるか = False
End Function

'--- 玉突きを確定する(4か所を対で動かす) ---
Private Sub CB_玉突きを打つ(ByVal a As Long, ByVal d As Long, ByVal x As Long, _
                            ByVal b As Long, ByVal y As Long)
    On Error GoTo ErrHandler

10  mPlan(a, d) = ST_WORK: CovAdd a, d, 1
20  mPlan(a, x) = ST_OFF: CovAdd a, x, -1
30  mPlan(b, x) = ST_WORK: CovAdd b, x, 1
40  mPlan(b, y) = ST_OFF: CovAdd b, y, -1
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "CB_玉突きを打つ", Err.Number, Err.Description, Erl, _
             "a=" & a & "; d=" & d & "; x=" & x & "; b=" & b & "; y=" & y
End Sub


'--- 医師5名日の出勤を均等化(通常ルールの薬剤師間で誤差1以内を目標) ---
'    v9.4.0: 日別の人数を減らす向きの入替を制限した。
'    旧版は「最多の人を5診日から降ろす」分岐を先に試し、成功すると
'    その巡回を終えていたため、成功するたびに5診日の出勤が1名ずつ
'    減り続け、5診日が4診日より薄くなっていた。
'    現版は (1)最少の人を混雑日へ乗せる向きを先に試し、
'          (2)降ろす向きは、その日が必要数を保てるときだけ許す。
'    どちらも余裕のある日との入替に限るため、全日が不足している月では
'    何も動かさない(カバレッジを優先し、個人差はそのまま残す)。
Public Sub FiveBalance()
    Dim pass As Long, i As Long, f As Long
    Dim mxI As Long, mnI As Long, mxV As Long, mnV As Long
    Dim swapped As Boolean
    On Error GoTo ErrHandler

10  For pass = 1 To FB_MAX_PASS
20      mxI = 0: mnI = 0: mxV = -1: mnV = CNT_INF
30      For i = 1 To mNP
40          If Not mSkipRow(i) Then
50          If mKind(i) = KIND_PH And Not mLeave(i) And mRule(i) = "通常" Then
60              f = FiveCnt(i)
70              If f > mxV Then mxV = f: mxI = i
80              If f < mnV Then mnV = f: mnI = i
90          End If
100         End If
110     Next i
120     If mxI = 0 Or mnI = 0 Or mxI = mnI Then Exit Sub
130     If mxV - mnV <= 1 Then Exit Sub
140     swapped = FB_混雑日へ乗せる(mnI)
150     If Not swapped Then swapped = FB_混雑日から降ろす(mxI)
160     If Not swapped Then Exit Sub
170 Next pass
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "FiveBalance", Err.Number, Err.Description, Erl, _
             "pass=" & pass & "; maxIdx=" & mxI & "; minIdx=" & mnI
End Sub

'--- 混雑日の公休を出勤に変え、余裕のある非混雑日を休みにする ---
Private Function FB_混雑日へ乗せる(ByVal i As Long) As Boolean
    Dim j As Long, dBusy As Long, dSpare As Long
    Dim bs As Double, d As Double
    On Error GoTo ErrHandler

    '--- 乗せる先: 混雑日のうち最も不足している日 ---
10  dBusy = 0: bs = SCORE_INF
20  For j = 1 To mND
30      If mDayIn(j) And mDayDoc(j) = DOC_BUSY_N And mPlan(i, j) = ST_OFF Then
40          If WorkRunIf(i, j) <= mMaxRun Then
50              d = mDayReq(j) - mCov(j)
60              If d > bs Then bs = d: dBusy = j
70          End If
80      End If
90  Next j
100 If dBusy = 0 Then Exit Function

    '--- 抜く先: 必要数を保てる非混雑日のうち最も余裕のある日 ---
110 dSpare = FB_余裕のある出勤日(i, dBusy)
120 If dSpare = 0 Then Exit Function

130 mPlan(i, dBusy) = ST_WORK
140 mPlan(i, dSpare) = ST_OFF
150 CovAdd i, dBusy, 1
160 CovAdd i, dSpare, -1
170 FB_混雑日へ乗せる = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "FB_混雑日へ乗せる", Err.Number, Err.Description, Erl, _
             "i=" & i & "; busyDay=" & dBusy & "; spareDay=" & dSpare
    FB_混雑日へ乗せる = False
End Function

'--- 混雑日の出勤を休みに変え、代わりに非混雑日へ出勤を移す ---
'    抜いても必要数を割らない混雑日だけを対象にする。
Private Function FB_混雑日から降ろす(ByVal i As Long) As Boolean
    Dim j As Long, dBusy As Long, dPut As Long
    Dim bs As Double, d As Double
    On Error GoTo ErrHandler

    '--- 降ろす先: 抜いても必要数を保てる混雑日のうち最も余裕のある日 ---
10  dBusy = 0: bs = SCORE_INF
20  For j = 1 To mND
30      If mDayIn(j) And mDayDoc(j) = DOC_BUSY_N And mPlan(i, j) = ST_WORK Then
40          If mCov(j) - 1 >= mDayReq(j) Then
50              If OffRunIf(i, j) <= mMaxOffRun Then
60                  d = mCov(j) - mDayReq(j)
70                  If d > bs Then bs = d: dBusy = j
80              End If
90          End If
100     End If
110 Next j
120 If dBusy = 0 Then Exit Function

    '--- 乗せる先: 非混雑日の公休のうち最も不足している日 ---
130 dPut = 0: bs = SCORE_INF
140 For j = 1 To mND
150     If mDayIn(j) And mDayDoc(j) <> DOC_BUSY_N And mPlan(i, j) = ST_OFF Then
160         If WorkRunIf(i, j) <= mMaxRun Then
170             d = mDayReq(j) - mCov(j)
180             If d > bs Then bs = d: dPut = j
190         End If
200     End If
210 Next j
220 If dPut = 0 Then Exit Function

230 mPlan(i, dBusy) = ST_OFF
240 mPlan(i, dPut) = ST_WORK
250 CovAdd i, dBusy, -1
260 CovAdd i, dPut, 1
270 FB_混雑日から降ろす = True
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "FB_混雑日から降ろす", Err.Number, Err.Description, Erl, _
             "i=" & i & "; busyDay=" & dBusy & "; putDay=" & dPut
    FB_混雑日から降ろす = False
End Function

'--- 抜いても必要数を割らない非混雑日のうち、最も余裕のある出勤日 ---
Private Function FB_余裕のある出勤日(ByVal i As Long, ByVal exceptJ As Long) As Long
    Dim j As Long, best As Long
    Dim bs As Double, d As Double
    On Error GoTo ErrHandler

10  best = 0: bs = SCORE_INF
20  For j = 1 To mND
30      If mDayIn(j) And j <> exceptJ And mDayDoc(j) <> DOC_BUSY_N Then
40          If mPlan(i, j) = ST_WORK And mCov(j) - 1 >= mDayReq(j) Then
50              If OffRunIf(i, j) <= mMaxOffRun Then
60                  d = mCov(j) - mDayReq(j)
70                  If d > bs Then bs = d: best = j
80              End If
90          End If
100     End If
110 Next j
120 FB_余裕のある出勤日 = best
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "FB_余裕のある出勤日", Err.Number, Err.Description, Erl, _
             "i=" & i & "; exceptJ=" & exceptJ & "; j=" & j
    FB_余裕のある出勤日 = 0
End Function


'--- ○●▲の個人差を均等化(誤差2以内を目標・同じ日の2人で記号を交換) ---
Public Sub SymbolBalance()
    On Error GoTo ErrHandler
    Dim pass As Long, s As Long, i As Long, j As Long
    Dim mxI As Long, mnI As Long, mxV As Long, mnV As Long, c As Long
    Dim sym As String, other As String, done As Boolean
    For pass = 1 To 300
        done = True
        For s = 1 To 3
            sym = Choose(s, SYM_EARLY, SYM_MID, SYM_LATE)
            mxI = 0: mnI = 0: mxV = -1: mnV = 32767
            For i = 1 To mNP
                If Not mSkipRow(i) Then
                If mKind(i) = KIND_PH And Not mLeave(i) And mRule(i) <> "手動" Then
                    If sym = SYM_EARLY Or mCanLate(i) Then
                        c = SymCnt(i, sym)
                        If c > mxV Then mxV = c: mxI = i
                        If c < mnV Then mnV = c: mnI = i
                    End If
                End If
                End If
            Next i
            If mxI > 0 And mnI > 0 And mxI <> mnI And mxV - mnV > 2 Then
                For j = 1 To mND
                    If mDayIn(j) And mPlan(mxI, j) = ST_WORK And mPlan(mnI, j) = ST_WORK Then
                        If mSymb(mxI, j) = sym And mSymb(mnI, j) <> sym Then
                            other = mSymb(mnI, j)
                            If other = SYM_EARLY Or mCanLate(mxI) Then
                                mSymb(mxI, j) = other: mSymb(mnI, j) = sym
                                AddCnt mxI, sym, -1: AddCnt mxI, other, 1
                                AddCnt mnI, other, -1: AddCnt mnI, sym, 1
                                done = False
                                Exit For
                            End If
                        End If
                    End If
                Next j
            End If
        Next s
        If done Then Exit For
    Next pass
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "SymbolBalance", Err.Number, Err.Description, Erl, ""
End Sub
