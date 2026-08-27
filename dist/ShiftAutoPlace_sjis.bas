Attribute VB_Name = "ShiftAutoPlace"
Option Explicit
'==================================================================
'  ShiftAutoPlace v9.3.0
'  公休の配置・均等化アルゴリズムと後半工程。
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
' 事務員の早番(○)は1日あたり何人までか
Private Const CLERK_EARLY_MAX As Long = 1

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
10   For guard = 1 To 3
20       For i = 1 To mNP
30           If Not mSkipRow(i) And Not mLeave(i) Then RepairRuns i
40       Next i
50   Next guard

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
    Dim i As Long, best As Long, bestCnt As Long, c As Long
    On Error GoTo ErrHandler

10  best = 0: bestCnt = CNT_INF
20  For i = 1 To mNP
30      If Not mSkipRow(i) Then
40          If mKind(i) = kind And mPlan(i, j) = ST_WORK And Len(mSymb(i, j)) = 0 Then
50              If (Not lateOnly) Or mCanLate(i) Then
60                  c = SymCnt(i, sym)
70                  If c < bestCnt Then bestCnt = c: best = i
80              End If
90          End If
100     End If
110 Next i
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

    '--- ▲: 設定の最低人数まで(遅番可の人のみ) ---
90  needL = mLateMin - dayL
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
'   ○(早番)は1日1人まで。2人目以降は設定の記号(L13)にする。
'------------------------------------------------------------------
Private Sub AP_事務員の記号(ByVal j As Long)
    Dim dayE As Long, dayM As Long, dayL As Long, bi As Long
    On Error GoTo ErrHandler

10  AP_日別既存数 j, KIND_CL, dayE, dayM, dayL
20  Do
30      bi = AP_最少候補(j, KIND_CL, SYM_EARLY, False)
40      If bi = 0 Then Exit Do
50      If dayE < CLERK_EARLY_MAX Then
60          AP_記号を置く bi, j, SYM_EARLY
70          dayE = dayE + 1
80      Else
90          AP_記号を置く bi, j, mGSym
100     End If
110 Loop
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
100     msg = msg & vbCrLf & "※遅番(" & SYM_LATE & ")が目標" & mLateMin & "名に届かない日: " & _
              lateShort & "日" & _
              IIf(lateBad > 0, "　■ うち" & mLateMin - 1 & "名未満: " & lateBad & "日", "") & vbCrLf
110 End If

120 shortN = AP_必要数不足日数(worst)
130 If shortN > 0 Then
140     msg = msg & vbCrLf & "※必要数(医師数+" & mReqPlus & ")に届かない日: " & shortN & _
              "日(最大不足" & worst & "名)" & vbCrLf & _
              "　→ 公休ノルマ優先のため。日別は過不足行で確認できます。" & vbCrLf
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
    Dim j As Long, n As Long, dayLate As Long
    On Error GoTo ErrHandler

10  bad = 0
20  For j = 1 To mND
30      If mDayIn(j) Then
40          dayLate = AP_日別遅番数(j)
50          If dayLate < mLateMin Then n = n + 1
60          If dayLate < mLateMin - 1 Then bad = bad + 1
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
                        If mPlan(i, k) = ST_OFF Then
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


'--- 医師5名日の出勤を均等化(通常ルールの薬剤師間で誤差1以内を目標) ---
Public Sub FiveBalance()
    On Error GoTo ErrHandler
    Dim pass As Long, i As Long, j As Long, f As Long
    Dim mxI As Long, mnI As Long, mxV As Long, mnV As Long
    Dim d5 As Long, dx As Long, swapped As Boolean
    For pass = 1 To 100
        mxI = 0: mnI = 0: mxV = -1: mnV = 32767
        For i = 1 To mNP
            If Not mSkipRow(i) Then
            If mKind(i) = KIND_PH And Not mLeave(i) And mRule(i) = "通常" Then
                f = FiveCnt(i)
                If f > mxV Then mxV = f: mxI = i
                If f < mnV Then mnV = f: mnI = i
            End If
            End If
        Next i
        If mxI = 0 Or mnI = 0 Or mxI = mnI Then Exit Sub
        If mxV - mnV <= 1 Then Exit Sub
        swapped = False
        d5 = 0: dx = 0
        For j = 1 To mND
            If mDayIn(j) And mDayDoc(j) = 5 And mPlan(mxI, j) = ST_WORK Then d5 = j: Exit For
        Next j
        For j = 1 To mND
            If mDayIn(j) And mDayDoc(j) <> 5 And mPlan(mxI, j) = ST_OFF Then
                If WorkRunIf(mxI, j) <= mMaxRun Then dx = j: Exit For
            End If
        Next j
        If d5 > 0 And dx > 0 Then
            mPlan(mxI, d5) = ST_OFF: mPlan(mxI, dx) = ST_WORK
            CovAdd mxI, d5, -1: CovAdd mxI, dx, 1
            swapped = True
        End If
        If Not swapped Then
            d5 = 0: dx = 0
            For j = 1 To mND
                If mDayIn(j) And mDayDoc(j) = 5 And mPlan(mnI, j) = ST_OFF Then
                    If WorkRunIf(mnI, j) <= mMaxRun Then d5 = j: Exit For
                End If
            Next j
            For j = 1 To mND
                If mDayIn(j) And mDayDoc(j) <> 5 And mPlan(mnI, j) = ST_WORK Then dx = j: Exit For
            Next j
            If d5 > 0 And dx > 0 Then
                mPlan(mnI, d5) = ST_WORK: mPlan(mnI, dx) = ST_OFF
                CovAdd mnI, d5, 1: CovAdd mnI, dx, -1
                swapped = True
            End If
        End If
        If Not swapped Then Exit Sub
    Next pass
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "FiveBalance", Err.Number, Err.Description, Erl, ""
End Sub


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
