Option Explicit
'==================================================================
'  シフト自動作成  ＜標準モジュール ShiftAuto v6.5.0＞  2026-08-24
'  ルールはすべてシート「自動作成設定」から読込(K列ラベル部分一致):
'   L5=早番(○)人数/日  L6=遅番(▲)最低人数/日(誤差-1)
'   L7=連勤の上限  L8=連休の上限  L9=週の基本休日数
'   L10=必要出勤数(医師数+n)のn  L11=ノルマ外の休み記号(カンマ区切り)
'   L13=事務員の2人目の記号(○は1日1人)
'  メンバー表(5行目?)はA列が空白になるまで読込(人数可変)
'  主要ルール: 公休ノルマ=土日祝と同数(誤差0厳守・ノルマ外記号は上乗せ)/
'   固定曜日・週N日・手動(派遣)ルール/週2休基本+2?3連休化/連勤上限/
'   医5日均等(誤差1)/○●▲均等(誤差2)/▲最低人数(誤差-1)/
'   事務員不在の日をなるべく回避・事務員の○は1日1人
'  v6.5.0 追加:
'   ・設定シートが無い場合は既定値で自動生成(BuildCfgSheet)
'   ・差分ログ「シフト変更ログ」: 変更セル/変更前後/書式を全記録
'   ・「シフト変更を戻す」= 最後のセッションを逆再生して復元(繰返し可)
'   ・「シフト白紙化」= 入力欄を全消去(こちらもログに記録され復元可)
'==================================================================
Private Const CFG_SHEET As String = "自動作成設定"
Private Const HOL_SHEET As String = "祝日マスタ"
Private Const LOG_SHEET As String = "シフト変更ログ"
Private Const DOC_LABEL As String = "医師数(診)"
Private Const SYM_EARLY As String = "○"
Private Const SYM_MID   As String = "●"
Private Const SYM_LATE  As String = "▲"
Private Const SYM_OFF   As String = "公休"
' 予定ステータス
Private Const ST_SKIP  As Long = -1   ' 月外・休業
Private Const ST_WORK  As Long = 1    ' 自動:出勤
Private Const ST_OFF   As Long = 2    ' 自動:公休
Private Const ST_FWORK As Long = 3    ' 既存入力:出勤(○◯●▲)
Private Const ST_FOFF  As Long = 4    ' 既存入力:休み(希休・有休・公休等)
'--- モジュール内共有状態 ---
Private mPlan() As Long, mKind() As String, mRule() As String
Private mLeave() As Boolean, mCanLate() As Boolean
Private mDayIn() As Boolean, mDayWD() As Long, mDayHol() As Boolean
Private mDayDoc() As Long, mDayReq() As Long, mWkKey() As Long
Private mCov() As Long          ' 薬剤師の予定出勤数
Private mCovG() As Long         ' 事務員の予定出勤数
Private mSymb() As String, mCntE() As Long, mCntM() As Long, mCntL() As Long
Private mMaxRun As Long, mMaxOffRun As Long, mPaidSyms As String
Private mNP As Long, mND As Long
Public Sub シフト自動作成()
    Dim ws As Worksheet, cfg As Worksheet, holWs As Worksheet, grid As Range
    Dim i As Long, j As Long, k As Long, r As Long, w As Long, j0 As Long
    Dim dateRow As Long, docRow As Long, monthNum As Long
    Dim mres As Variant, dayDt() As Date, targetOff As Long
    Dim pName() As String, pWD() As Boolean
    Dim pWeekN() As Long, pQuota() As Long, pFound() As Boolean
    Dim missing As String, unmet As String, v As String
    Dim quota As Long, offN As Long, remOff() As Long
    Dim processed() As Boolean, cnt As Long, wCnt As Long, tgt As Long, autoW As Long
    Dim bj As Long, bs As Double, sc As Double
    Dim wkList() As Long, nW As Long, exists As Boolean
    Dim tW() As Long, exW As Long, dW As Long, sumB As Long, needT As Long
    Dim guard As Long, added As Boolean, n As Long
    Dim moved As Boolean
    Dim placedB As Boolean
    Dim earlyN As Long, lateMin As Long, weekBase As Long, reqPlus As Long
    Dim gSym As String, gE As Long, gZero As Long
    Dim needE As Long, needL As Long
    Dim dayL As Long, dayM As Long
    Dim bi As Long, bv As Long, symS As String
    Dim written As Long
    Dim msg As String, offC As Long, wC As Long, paidC As Long
    Dim shortN As Long, worst As Long
    Dim lateShort As Long, lateBad As Long, dayLateN As Long
    Dim lg As Worksheet, sess As Long, lgR As Long, tv As String, cCell As Range
    Dim logged As Long
    '=== 準備 ===
    Set ws = Worksheets("シフト")
    On Error Resume Next
    Set cfg = Worksheets(CFG_SHEET)
    Set holWs = Worksheets(HOL_SHEET)
    On Error GoTo 0
    If cfg Is Nothing Then
        If MsgBox("設定シート「" & CFG_SHEET & "」がありません。" & vbCrLf & _
                  "既定値で自動生成しますか?(氏名はシフト表から取込)" & vbCrLf & vbCrLf & _
                  "※生成後、区分(事務員)・固定曜日・週N日・手動(派遣)・" & vbCrLf & _
                  "　休業・遅番不可 は手動で設定してください。", _
                  vbYesNo + vbQuestion, "設定シート自動生成") = vbYes Then
            Set cfg = BuildCfgSheet(ws)
            MsgBox "設定シート「" & CFG_SHEET & "」を既定値で生成しました。" & vbCrLf & _
                   "内容を確認・修正のうえ、もう一度実行してください。", vbInformation
        End If
        Exit Sub
    End If
    If Not IsDate(ws.Range("A1").Value) Then
        MsgBox "シフト!A1 に対象月の日付がありません。", vbExclamation: Exit Sub
    End If
    '=== 全体設定の読込(K列ラベルの部分一致検索) ===
    earlyN = CLng(CfgNum(cfg, "早番", 1))
    lateMin = CLng(CfgNum(cfg, "遅番", 3))
    mMaxRun = CLng(CfgNum(cfg, "連勤", 3))
    mMaxOffRun = CLng(CfgNum(cfg, "連休", 3))
    weekBase = CLng(CfgNum(cfg, "週の基本", 2))
    reqPlus = CLng(CfgNum(cfg, "必要出勤", 1))
    mPaidSyms = CfgTxt(cfg, "ノルマ外", "有休")
    gSym = CfgTxt(cfg, "2人目", SYM_MID)
    If earlyN < 0 Then earlyN = 0
    If lateMin < 0 Then lateMin = 0
    If mMaxRun < 1 Then mMaxRun = 1
    If mMaxOffRun < 1 Then mMaxOffRun = 1
    If weekBase < 0 Then weekBase = 0
    Set grid = AutoShiftRange(ws)
    mNP = grid.Rows.Count: mND = grid.Columns.Count
    dateRow = grid.Row - 1
    mres = Application.Match(DOC_LABEL, ws.Columns(1), 0)
    If IsError(mres) Then
        MsgBox "「" & DOC_LABEL & "」行が見つかりません。", vbExclamation: Exit Sub
    End If
    docRow = CLng(mres)
    monthNum = Month(ws.Range("A1").Value)
    '=== 日情報 ===
    ReDim dayDt(1 To mND): ReDim mDayIn(1 To mND): ReDim mDayWD(1 To mND)
    ReDim mDayHol(1 To mND): ReDim mDayDoc(1 To mND): ReDim mDayReq(1 To mND)
    ReDim mWkKey(1 To mND)
    targetOff = 0
    For j = 1 To mND
        mDayIn(j) = False
        If IsDate(ws.Cells(dateRow, grid.Column + j - 1).Value) Then
            dayDt(j) = ws.Cells(dateRow, grid.Column + j - 1).Value
            mDayIn(j) = (Month(dayDt(j)) = monthNum)
        End If
        If mDayIn(j) Then
            mDayWD(j) = Weekday(dayDt(j), vbSunday)
            mWkKey(j) = CLng(dayDt(j)) - (mDayWD(j) - 1)
            If Not holWs Is Nothing Then
                mDayHol(j) = (Application.CountIf(holWs.Columns(1), dayDt(j)) > 0)
            End If
            mDayDoc(j) = Val(ws.Cells(docRow, grid.Column + j - 1).Value)
            mDayReq(j) = mDayDoc(j) + reqPlus
            If mDayWD(j) = 1 Or mDayWD(j) = 7 Then
                targetOff = targetOff + 1
            ElseIf mDayHol(j) Then
                targetOff = targetOff + 1
            End If
        End If
    Next j
    '=== メンバー設定読込(A列が空白になるまで・人数可変) ===
    ReDim pName(1 To mNP): ReDim mKind(1 To mNP): ReDim mLeave(1 To mNP)
    ReDim mRule(1 To mNP): ReDim pWD(1 To mNP, 1 To 7): ReDim pWeekN(1 To mNP)
    ReDim pQuota(1 To mNP): ReDim mCanLate(1 To mNP): ReDim pFound(1 To mNP)
    For i = 1 To mNP
        pName(i) = Trim$(CStr(ws.Cells(grid.Row + i - 1, 1).Value))
        mKind(i) = "薬剤師": mRule(i) = "通常": mCanLate(i) = True: pQuota(i) = -1
        r = 5
        Do While Len(Trim$(CStr(cfg.Cells(r, 1).Value))) > 0
            If Trim$(CStr(cfg.Cells(r, 1).Value)) = pName(i) Then
                pFound(i) = True
                If Len(Trim$(CStr(cfg.Cells(r, 2).Value))) > 0 Then mKind(i) = Trim$(CStr(cfg.Cells(r, 2).Value))
                mLeave(i) = (Trim$(CStr(cfg.Cells(r, 3).Value)) <> "")
                If Len(Trim$(CStr(cfg.Cells(r, 4).Value))) > 0 Then mRule(i) = Trim$(CStr(cfg.Cells(r, 4).Value))
                ParseWD CStr(cfg.Cells(r, 5).Value), pWD, i
                pWeekN(i) = Val(cfg.Cells(r, 6).Value)
                If Len(Trim$(CStr(cfg.Cells(r, 7).Value))) > 0 Then pQuota(i) = Val(cfg.Cells(r, 7).Value)
                mCanLate(i) = (Trim$(CStr(cfg.Cells(r, 8).Value)) <> "不可")
                Exit Do
            End If
            r = r + 1
        Loop
        If Not pFound(i) And Len(pName(i)) > 0 Then missing = missing & vbCrLf & "・" & pName(i)
    Next i
    '=== 既存入力の分類 ===
    ReDim mPlan(1 To mNP, 1 To mND)
    For i = 1 To mNP
        For j = 1 To mND
            If Not mDayIn(j) Or mLeave(i) Then
                mPlan(i, j) = ST_SKIP
            Else
                v = Trim$(CStr(grid.Cells(i, j).Value))
                If Len(v) = 0 Then
                    mPlan(i, j) = 0
                ElseIf InStr("○◯●▲", v) > 0 Then
                    mPlan(i, j) = ST_FWORK
                Else
                    mPlan(i, j) = ST_FOFF          ' 希休・有休・公休など
                End If
            End If
        Next j
    Next i
    '=== ルール適用(固定曜日 / 手動=何もしない / それ以外は仮で全出勤) ===
    For i = 1 To mNP
        If Not mLeave(i) Then
            If mRule(i) = "固定曜日" Then
                For j = 1 To mND
                    If mPlan(i, j) = 0 Then
                        If pWD(i, mDayWD(j)) Then mPlan(i, j) = ST_WORK Else mPlan(i, j) = ST_OFF
                    End If
                Next j
            ElseIf mRule(i) = "手動" Then
                ' 派遣行など: 空白は空白のまま(自動配置しない)
            Else
                For j = 1 To mND
                    If mPlan(i, j) = 0 Then mPlan(i, j) = ST_WORK
                Next j
            End If
        End If
    Next i
    '=== 予定出勤数(薬剤師/事務員 別) ===
    ReDim mCov(1 To mND): ReDim mCovG(1 To mND)
    For j = 1 To mND
        mCov(j) = 0: mCovG(j) = 0
        For i = 1 To mNP
            If mPlan(i, j) = ST_WORK Or mPlan(i, j) = ST_FWORK Then
                If mKind(i) = "薬剤師" Then mCov(j) = mCov(j) + 1
                If mKind(i) = "事務員" Then mCovG(j) = mCovG(j) + 1
            End If
        Next i
    Next j
    '=== 週N日ルール: 週ごとに勤務日数を絞る ===
    For i = 1 To mNP
        If mRule(i) = "週N日" And Not mLeave(i) And pWeekN(i) > 0 Then
            ReDim processed(1 To mND)
            For j0 = 1 To mND
                If mDayIn(j0) And Not processed(j0) Then
                    cnt = 0: wCnt = 0
                    For j = 1 To mND
                        If mDayIn(j) And mWkKey(j) = mWkKey(j0) Then
                            processed(j) = True
                            cnt = cnt + 1
                            If mPlan(i, j) = ST_FWORK Then wCnt = wCnt + 1
                        End If
                    Next j
                    tgt = Int(pWeekN(i) * cnt / 7 + 0.5)
                    If tgt > cnt Then tgt = cnt
                    Do
                        autoW = 0
                        For j = 1 To mND
                            If mDayIn(j) And mWkKey(j) = mWkKey(j0) And mPlan(i, j) = ST_WORK Then autoW = autoW + 1
                        Next j
                        If wCnt + autoW <= tgt Then Exit Do
                        bj = 0: bs = -1E+30
                        For j = 1 To mND
                            If mDayIn(j) And mWkKey(j) = mWkKey(j0) And mPlan(i, j) = ST_WORK Then
                                sc = OffScore(i, j)
                                If sc > bs Then bs = sc: bj = j
                            End If
                        Next j
                        If bj = 0 Then Exit Do
                        mPlan(i, bj) = ST_OFF
                        CovAdd i, bj, -1
                    Loop
                End If
            Next j0
        End If
    Next i
    '=== 週リスト(日曜キー昇順) ===
    ReDim wkList(1 To mND): nW = 0
    For j = 1 To mND
        If mDayIn(j) Then
            exists = False
            For w = 1 To nW
                If wkList(w) = mWkKey(j) Then exists = True: Exit For
            Next w
            If Not exists Then nW = nW + 1: wkList(nW) = mWkKey(j)
        End If
    Next j
    '=== 通常ルール: 公休ノルマ先行(週(L9)休基本+余剰は連休化) ===
    ReDim remOff(1 To mNP)
    For i = 1 To mNP
        remOff(i) = 0
        If mRule(i) = "通常" And Not mLeave(i) Then
            quota = pQuota(i): If quota < 0 Then quota = targetOff
            offN = 0
            For j = 1 To mND
                If mPlan(i, j) = ST_FOFF Then
                    v = Trim$(CStr(grid.Cells(i, j).Value))
                    If Not IsPaidOff(v) Then offN = offN + 1
                End If
            Next j
            remOff(i) = quota - offN
            If remOff(i) < 0 Then remOff(i) = 0
        End If
    Next i
    For i = 1 To mNP
        If remOff(i) > 0 Then
            needT = remOff(i)
            ReDim tW(1 To nW): sumB = 0
            For w = 1 To nW
                exW = 0: dW = 0
                For j = 1 To mND
                    If mDayIn(j) And mWkKey(j) = wkList(w) Then
                        dW = dW + 1
                        If mPlan(i, j) = ST_FOFF Or mPlan(i, j) = ST_OFF Then exW = exW + 1
                    End If
                Next j
                tW(w) = weekBase - exW
                If tW(w) > dW - exW Then tW(w) = dW - exW
                If dW <= 2 And tW(w) > 1 Then tW(w) = 1
                If tW(w) < 0 Then tW(w) = 0
                sumB = sumB + tW(w)
            Next w
            guard = 0
            Do While sumB > needT And guard < 100
                For w = nW To 1 Step -1
                    If sumB > needT And tW(w) > 0 Then tW(w) = tW(w) - 1: sumB = sumB - 1
                Next w
                guard = guard + 1
            Loop
            guard = 0
            Do While sumB < needT And guard < 100
                added = False
                For w = 1 To nW
                    If sumB < needT And tW(w) < mMaxOffRun Then tW(w) = tW(w) + 1: sumB = sumB + 1: added = True
                Next w
                If Not added Then Exit Do
                guard = guard + 1
            Loop
            For w = 1 To nW
                n = tW(w)
                If n > remOff(i) Then n = remOff(i)
                Do While n >= 2
                    placedB = False
                    If n >= 3 Then
                        If PlaceOffBlock(i, wkList(w), 3) Then
                            remOff(i) = remOff(i) - 3: n = n - 3: placedB = True
                        End If
                    End If
                    If Not placedB Then
                        If PlaceOffBlock(i, wkList(w), 2) Then
                            remOff(i) = remOff(i) - 2: n = n - 2
                        Else
                            Exit Do
                        End If
                    End If
                Loop
                Do While n > 0
                    If PlaceOffSingle(i, wkList(w)) Then
                        remOff(i) = remOff(i) - 1: n = n - 1
                    Else
                        Exit Do
                    End If
                Loop
            Next w
        End If
    Next i
    '=== 残りノルマ: 既存の休みに寄せて1日ずつ必ず配置(誤差0厳守) ===
    Do
        moved = False
        For i = 1 To mNP
            If remOff(i) > 0 Then
                bj = 0: bs = -1E+30
                For j = 1 To mND
                    If mPlan(i, j) = ST_WORK Then
                        sc = OffScore(i, j) + AdjBonus(i, j)
                        If sc > bs Then bs = sc: bj = j
                    End If
                Next j
                If bj > 0 Then
                    mPlan(i, bj) = ST_OFF
                    CovAdd i, bj, -1
                    remOff(i) = remOff(i) - 1
                    moved = True
                Else
                    unmet = unmet & vbCrLf & "・" & pName(i) & " : あと" & remOff(i) & "日 配置できず"
                    remOff(i) = 0
                End If
            End If
        Next i
    Loop While moved
    '=== 連勤上限超えの緩和(入替) ===
    For guard = 1 To 3
        For i = 1 To mNP
            If Not mLeave(i) Then RepairRuns i
        Next i
    Next guard
    '=== 医師5名日の出勤を均等化(誤差1以内を目標) ===
    FiveBalance
    '=== 記号割当 ===
    ReDim mSymb(1 To mNP, 1 To mND)
    ReDim mCntE(1 To mNP): ReDim mCntM(1 To mNP): ReDim mCntL(1 To mNP)
    For i = 1 To mNP
        For j = 1 To mND
            If mPlan(i, j) = ST_FWORK Then
                v = Trim$(CStr(grid.Cells(i, j).Value))
                If v = SYM_LATE Then mCntL(i) = mCntL(i) + 1
                If v = SYM_MID Then mCntM(i) = mCntM(i) + 1
                If v = SYM_EARLY Or v = "◯" Then mCntE(i) = mCntE(i) + 1
            End If
        Next j
    Next i
    For j = 1 To mND
        If mDayIn(j) Then
            '--- 薬剤師: ○(L5人) → ▲最低人数(L6) → 残り●▲均等 ---
            needE = earlyN: dayL = 0: dayM = 0
            For i = 1 To mNP
                If mKind(i) = "薬剤師" And mPlan(i, j) = ST_FWORK Then
                    v = Trim$(CStr(grid.Cells(i, j).Value))
                    If v = SYM_EARLY Or v = "◯" Then needE = needE - 1
                    If v = SYM_LATE Then dayL = dayL + 1
                    If v = SYM_MID Then dayM = dayM + 1
                End If
            Next i
            Do While needE > 0
                bi = 0: bv = 32767
                For i = 1 To mNP
                    If mKind(i) = "薬剤師" And mPlan(i, j) = ST_WORK And Len(mSymb(i, j)) = 0 Then
                        If mCntE(i) < bv Then bv = mCntE(i): bi = i
                    End If
                Next i
                If bi = 0 Then Exit Do
                mSymb(bi, j) = SYM_EARLY: mCntE(bi) = mCntE(bi) + 1: needE = needE - 1
            Loop
            needL = lateMin - dayL
            Do While needL > 0
                bi = 0: bv = 32767
                For i = 1 To mNP
                    If mKind(i) = "薬剤師" And mPlan(i, j) = ST_WORK _
                       And Len(mSymb(i, j)) = 0 And mCanLate(i) Then
                        If mCntL(i) < bv Then bv = mCntL(i): bi = i
                    End If
                Next i
                If bi = 0 Then Exit Do
                mSymb(bi, j) = SYM_LATE: mCntL(bi) = mCntL(bi) + 1: dayL = dayL + 1
                needL = needL - 1
            Loop
            Do
                If dayM <= dayL Then symS = SYM_MID Else symS = SYM_LATE
                bi = 0: bv = 32767
                For i = 1 To mNP
                    If mKind(i) = "薬剤師" And mPlan(i, j) = ST_WORK _
                       And Len(mSymb(i, j)) = 0 And mCanLate(i) Then
                        If symS = SYM_MID Then
                            If mCntM(i) < bv Then bv = mCntM(i): bi = i
                        Else
                            If mCntL(i) < bv Then bv = mCntL(i): bi = i
                        End If
                    End If
                Next i
                If bi = 0 Then Exit Do
                If symS = SYM_MID Then
                    mSymb(bi, j) = SYM_MID: mCntM(bi) = mCntM(bi) + 1: dayM = dayM + 1
                Else
                    mSymb(bi, j) = SYM_LATE: mCntL(bi) = mCntL(bi) + 1: dayL = dayL + 1
                End If
            Loop
            '--- 事務員: ○(早番)は1日1人、2人目以降はL13の記号 ---
            gE = 0
            For i = 1 To mNP
                If mKind(i) = "事務員" And mPlan(i, j) = ST_FWORK Then
                    v = Trim$(CStr(grid.Cells(i, j).Value))
                    If v = SYM_EARLY Or v = "◯" Then gE = gE + 1
                End If
            Next i
            Do
                bi = 0: bv = 32767
                For i = 1 To mNP
                    If mKind(i) = "事務員" And mPlan(i, j) = ST_WORK And Len(mSymb(i, j)) = 0 Then
                        If mCntE(i) < bv Then bv = mCntE(i): bi = i
                    End If
                Next i
                If bi = 0 Then Exit Do
                If gE < 1 Then
                    mSymb(bi, j) = SYM_EARLY: mCntE(bi) = mCntE(bi) + 1: gE = gE + 1
                Else
                    mSymb(bi, j) = gSym
                    AddCnt bi, gSym, 1
                End If
            Loop
            '--- 残り(遅番不可の薬剤師など)は○ ---
            For i = 1 To mNP
                If mPlan(i, j) = ST_WORK And Len(mSymb(i, j)) = 0 Then
                    mSymb(i, j) = SYM_EARLY: mCntE(i) = mCntE(i) + 1
                End If
            Next i
        End If
    Next j
    '=== ○●▲の個人差を均等化(誤差2以内を目標・同日交換) ===
    SymbolBalance
    '=== 書き込み(差分ログを記録しながら) ===
    Set lg = GetLogSheet()
    sess = NextSession(lg)
    lgR = LogLastRow(lg)
    logged = 0
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    For i = 1 To mNP
        For j = 1 To mND
            tv = ""
            If mPlan(i, j) = ST_OFF Then
                tv = SYM_OFF
            ElseIf mPlan(i, j) = ST_WORK Then
                tv = mSymb(i, j)
            End If
            If Len(tv) > 0 Then
                Set cCell = grid.Cells(i, j)
                If Not cCell.HasFormula Then
                    If Trim$(CStr(cCell.Value)) <> tv Then
                        LogChange lg, lgR, sess, "自動作成", cCell, tv
                        logged = logged + 1
                    End If
                End If
                StampCell ws, cCell, tv
                written = written + 1
            End If
        Next j
    Next i
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    '=== 結果レポート ===
    msg = "対象月: " & Format(ws.Range("A1").Value, "yyyy年m月") & _
          "　公休ノルマ(土日祝): " & targetOff & "日　書込セル: " & written & vbCrLf & vbCrLf
    For i = 1 To mNP
        If mLeave(i) Then
            msg = msg & pName(i) & " : 休業(スキップ)" & vbCrLf
        Else
            offC = 0: wC = 0: paidC = 0
            For j = 1 To mND
                Select Case mPlan(i, j)
                    Case ST_WORK, ST_FWORK: wC = wC + 1
                    Case ST_OFF, ST_FOFF
                        offC = offC + 1
                        If mPlan(i, j) = ST_FOFF Then
                            If IsPaidOff(Trim$(CStr(grid.Cells(i, j).Value))) Then paidC = paidC + 1
                        End If
                End Select
            Next j
            msg = msg & pName(i) & " : 出勤" & wC & " 休" & offC & _
                  IIf(paidC > 0, "(うちノルマ外" & paidC & ")", "") & _
                  " 連勤max" & MaxRun(i) & " 連休max" & MaxOffRun(i) & _
                  IIf(mKind(i) = "薬剤師", " 医5日" & FiveCnt(i) & _
                      " ○" & mCntE(i) & " ●" & mCntM(i) & " ▲" & mCntL(i), "") & vbCrLf
        End If
    Next i
    If Len(unmet) > 0 Then
        msg = msg & vbCrLf & "? 公休ノルマ未達(空きセル不足):" & unmet & vbCrLf
    End If
    gZero = 0
    For j = 1 To mND
        If mDayIn(j) And mCovG(j) < 1 Then gZero = gZero + 1
    Next j
    If gZero > 0 Then
        msg = msg & vbCrLf & "? 事務員が不在の日: " & gZero & "日" & vbCrLf
    End If
    lateShort = 0: lateBad = 0
    For j = 1 To mND
        If mDayIn(j) Then
            dayLateN = 0
            For i = 1 To mNP
                If mKind(i) = "薬剤師" Then
                    If mPlan(i, j) = ST_WORK Then
                        If mSymb(i, j) = SYM_LATE Then dayLateN = dayLateN + 1
                    ElseIf mPlan(i, j) = ST_FWORK Then
                        If Trim$(CStr(grid.Cells(i, j).Value)) = SYM_LATE Then dayLateN = dayLateN + 1
                    End If
                End If
            Next i
            If dayLateN < lateMin Then lateShort = lateShort + 1
            If dayLateN < lateMin - 1 Then lateBad = lateBad + 1
        End If
    Next j
    If lateShort > 0 Then
        msg = msg & vbCrLf & "※遅番(▲)が目標" & lateMin & "名に届かない日: " & lateShort & "日" & _
              IIf(lateBad > 0, "　?うち" & lateMin - 1 & "名未満: " & lateBad & "日", "") & vbCrLf
    End If
    shortN = 0: worst = 0
    For j = 1 To mND
        If mDayIn(j) And mCov(j) < mDayReq(j) Then
            shortN = shortN + 1
            If mDayReq(j) - mCov(j) > worst Then worst = mDayReq(j) - mCov(j)
        End If
    Next j
    If shortN > 0 Then
        msg = msg & vbCrLf & "※必要数(医師数+" & reqPlus & ")に届かない日: " & shortN & "日(最大不足" & worst & "名)" & vbCrLf & _
              "　→ 公休ノルマ優先のため。日別は過不足行で確認できます。" & vbCrLf
    End If
    If Len(missing) > 0 Then msg = msg & vbCrLf & "設定未登録(既定値で処理): " & missing
    msg = msg & vbCrLf & "変更ログ: セッション#" & sess & " に " & logged & "セル記録。" & vbCrLf & _
          "取り消すには「シフト変更を戻す」を実行してください。"
    MsgBox msg, vbInformation, "シフト自動作成 結果"
End Sub
'==================================================================
' 差分ログ・戻す・白紙化(v6.5.0 追加)
'==================================================================
'--- ログシート取得(無ければ作成) ---
Private Function GetLogSheet() As Worksheet
    Dim lg As Worksheet
    On Error Resume Next
    Set lg = Worksheets(LOG_SHEET)
    On Error GoTo 0
    If lg Is Nothing Then
        Set lg = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        lg.Name = LOG_SHEET
        lg.Range("A1:J1").Value = Array("セッション", "日時", "操作", "セル", _
            "変更前", "変更後", "取消済", "前文字色", "前太字", "前塗り色")
        lg.Range("A1:J1").Font.Bold = True
        lg.Range("A1:J1").Interior.Color = RGB(217, 217, 217)
        lg.Columns("E:F").NumberFormat = "@"           ' 記号を文字列として保持
        lg.Columns("B:B").NumberFormat = "yyyy/mm/dd hh:mm:ss"
        lg.Columns("A:J").AutoFit
    End If
    Set GetLogSheet = lg
End Function
Private Function LogLastRow(ByVal lg As Worksheet) As Long
    LogLastRow = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row
    If LogLastRow < 1 Then LogLastRow = 1
End Function
Private Function NextSession(ByVal lg As Worksheet) As Long
    Dim lr As Long
    lr = LogLastRow(lg)
    If lr < 2 Then NextSession = 1 Else NextSession = CLng(Val(lg.Cells(lr, 1).Value)) + 1
End Function
'--- 1セル分の差分を記録(変更前の値と書式を保存してから書き換える前提) ---
Private Sub LogChange(ByVal lg As Worksheet, ByRef lr As Long, ByVal sess As Long, _
                      ByVal op As String, ByVal c As Range, ByVal newV As String)
    lr = lr + 1
    lg.Cells(lr, 1).Value = sess
    lg.Cells(lr, 2).Value = Now
    lg.Cells(lr, 3).Value = op
    lg.Cells(lr, 4).Value = c.Address(False, False)
    lg.Cells(lr, 5).Value = CStr(c.Value)
    lg.Cells(lr, 6).Value = newV
    lg.Cells(lr, 8).Value = c.Font.Color
    lg.Cells(lr, 9).Value = IIf(c.Font.Bold, "TRUE", "FALSE")
    If c.Interior.Pattern <> xlNone Then lg.Cells(lr, 10).Value = c.Interior.Color
End Sub
'--- 最後のセッションを逆再生して復元(繰り返し実行で1回ずつ遡れる) ---
Public Sub シフト変更を戻す()
    Dim ws As Worksheet, lg As Worksheet, lr As Long, r As Long
    Dim sess As Long, cnt As Long, c As Range, oldV As String
    Dim op As String, tm As String
    Set ws = Worksheets("シフト")
    On Error Resume Next
    Set lg = Worksheets(LOG_SHEET)
    On Error GoTo 0
    If lg Is Nothing Then
        MsgBox "変更ログがありません。" & vbCrLf & _
               "(ログは「シフト自動作成」「シフト白紙化」の実行時に記録されます)", vbExclamation
        Exit Sub
    End If
    lr = LogLastRow(lg)
    '--- 最後の未取消セッションを特定 ---
    sess = 0
    For r = lr To 2 Step -1
        If Trim$(CStr(lg.Cells(r, 7).Value)) = "" Then
            sess = CLng(Val(lg.Cells(r, 1).Value))
            op = CStr(lg.Cells(r, 3).Value)
            tm = Format(lg.Cells(r, 2).Value, "yyyy/m/d h:nn")
            Exit For
        End If
    Next r
    If sess = 0 Then
        MsgBox "取り消せる変更がありません(すべて取消済です)。", vbInformation: Exit Sub
    End If
    cnt = 0
    For r = 2 To lr
        If CLng(Val(lg.Cells(r, 1).Value)) = sess And Trim$(CStr(lg.Cells(r, 7).Value)) = "" Then cnt = cnt + 1
    Next r
    If MsgBox("セッション#" & sess & "「" & op & "」(" & tm & "・" & cnt & "セル)を取り消し、" & vbCrLf & _
              "変更前の状態に戻します。よろしいですか?", _
              vbYesNo + vbQuestion, "シフト変更を戻す") <> vbYes Then Exit Sub
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    '--- 逆順(後に書いたセルから)に復元 ---
    For r = lr To 2 Step -1
        If CLng(Val(lg.Cells(r, 1).Value)) = sess And Trim$(CStr(lg.Cells(r, 7).Value)) = "" Then
            Set c = ws.Range(CStr(lg.Cells(r, 4).Value))
            If Not c.HasFormula Then
                oldV = CStr(lg.Cells(r, 5).Value)
                If Len(oldV) = 0 Then c.ClearContents Else c.Value = oldV
                c.Font.Color = CLng(Val(lg.Cells(r, 8).Value))
                c.Font.Bold = (CStr(lg.Cells(r, 9).Value) = "TRUE")
                If Len(Trim$(CStr(lg.Cells(r, 10).Value))) > 0 Then
                    c.Interior.Color = CDbl(lg.Cells(r, 10).Value)
                Else
                    c.Interior.Pattern = xlNone
                End If
            End If
            lg.Cells(r, 7).Value = "○"
        End If
    Next r
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    MsgBox "セッション#" & sess & "「" & op & "」を取り消しました(" & cnt & "セル復元)。" & vbCrLf & _
           "さらに前の状態へ戻すには、もう一度実行してください。", vbInformation
End Sub
'--- 入力欄を白紙化(数式セルは保護・全消去分をログに記録=復元可) ---
Public Sub シフト白紙化()
    Dim ws As Worksheet, grid As Range, c As Range
    Dim lg As Worksheet, sess As Long, lr As Long, cnt As Long
    Dim hasVal As Boolean, hasFmt As Boolean
    Set ws = Worksheets("シフト")
    Set grid = AutoShiftRange(ws)
    If MsgBox("シフト入力欄(" & grid.Address(False, False) & ")の入力をすべて消去します。" & vbCrLf & _
              "消去内容は変更ログに記録され、「シフト変更を戻す」で復元できます。" & vbCrLf & _
              "よろしいですか?", vbYesNo + vbExclamation, "シフト白紙化") <> vbYes Then Exit Sub
    Set lg = GetLogSheet()
    sess = NextSession(lg)
    lr = LogLastRow(lg)
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    For Each c In grid.Cells
        If Not c.HasFormula Then
            hasVal = (Len(Trim$(CStr(c.Value))) > 0)
            hasFmt = (c.Interior.Pattern <> xlNone) Or (c.Font.Bold = True)
            If hasVal Or hasFmt Then
                LogChange lg, lr, sess, "白紙化", c, ""
                cnt = cnt + 1
                c.ClearContents
                c.Font.ColorIndex = xlAutomatic
                c.Font.Bold = False
                c.Interior.Pattern = xlNone
            End If
        End If
    Next c
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    If cnt = 0 Then
        MsgBox "消去対象のセルはありませんでした。", vbInformation
    Else
        MsgBox "白紙にしました(セッション#" & sess & "・" & cnt & "セル記録)。" & vbCrLf & _
               "元に戻すには「シフト変更を戻す」を実行してください。", vbInformation
    End If
End Sub
'==================================================================
' 設定シートの自動生成
'==================================================================
'--- 設定シートを既定値で自動生成(氏名はシフト表の行から取込) ---
Private Function BuildCfgSheet(ByVal ws As Worksheet) As Worksheet
    Dim cfg As Worksheet, grid As Range, i As Long, r As Long, nm As String
    Set grid = AutoShiftRange(ws)
    Set cfg = Worksheets.Add(After:=Worksheets(Worksheets.Count))
    cfg.Name = CFG_SHEET

    cfg.Range("A1").Value = "シフト自動作成 設定"
    cfg.Range("A1").Font.Bold = True: cfg.Range("A1").Font.Size = 14
    cfg.Range("A2").Value = "希望休はシフト表に「希休」スタンプで入力してください。実行時、入力済みのセルはすべて保持され、空白セルのみ自動で埋まります。"
    cfg.Range("A3").Value = "派遣スタッフの行をシフト表に追加したら、この表にも氏名を追加し「勤務ルール=手動」を設定してください。"
    cfg.Range("A2:A3").Font.Italic = True
    cfg.Range("A2:A3").Font.Size = 9
    cfg.Range("A2:A3").Font.Color = RGB(128, 128, 128)

    '--- メンバー表ヘッダー(4行目) ---
    cfg.Range("A4:I4").Value = Array("氏名", "区分", "休業(○=休業中)", "勤務ルール", _
        "固定曜日(例:月火金土)", "週勤務日数", "月間休日数(空欄=土日祝と同数)", _
        "遅番・遅半 可否", "備考")

    '--- メンバー行: シフト表A列から氏名を取込、既定値を設定 ---
    r = 5
    For i = 1 To grid.Rows.Count
        nm = Trim$(CStr(ws.Cells(grid.Row + i - 1, 1).Value))
        If Len(nm) > 0 Then
            cfg.Cells(r, 1).Value = nm
            cfg.Cells(r, 2).Value = "薬剤師"   ' 事務員は手動で変更
            cfg.Cells(r, 4).Value = "通常"
            cfg.Cells(r, 8).Value = "可"
            r = r + 1
        End If
    Next i

    '--- 全体設定(K列ラベル / L列値) ---
    cfg.Range("K4").Value = "全体設定(自動作成ルール)": cfg.Range("L4").Value = "値"
    cfg.Range("K5:L13").Value = Application.Transpose(Application.Transpose(Array( _
        Array("早番(○) 人数/日(基本1・最大2)", 1), _
        Array("遅番(▲) 最低人数/日(誤差-1)", 3), _
        Array("連勤の上限(日)", 3), _
        Array("連休の上限(日)", 3), _
        Array("週の基本休日数", 2), _
        Array("必要出勤数(医師数+n)の n", 1), _
        Array("ノルマ外の休み記号(カンマ区切り)", "有休"), _
        Array("週の定義(固定)", "日曜始まり・土曜終わり"), _
        Array("事務員の2人目の記号(○は1日1人)", "●"))))

    cfg.Range("A4:I4,K4:L4").Font.Bold = True
    cfg.Range("A4:I4,K4:L4").Interior.Color = RGB(217, 217, 217)
    cfg.Columns("A:L").AutoFit

    Set BuildCfgSheet = cfg
End Function
'==================================================================
' 設定読込ヘルパー(K列ラベルの部分一致 → L列の値)
'==================================================================
Private Function CfgNum(ByVal cfg As Worksheet, ByVal key As String, ByVal dft As Double) As Double
    Dim r As Long
    For r = 4 To 30
        If InStr(CStr(cfg.Cells(r, 11).Value), key) > 0 Then
            If Len(Trim$(CStr(cfg.Cells(r, 12).Value))) > 0 Then
                If IsNumeric(cfg.Cells(r, 12).Value) Then
                    CfgNum = CDbl(cfg.Cells(r, 12).Value)
                    Exit Function
                End If
            End If
        End If
    Next r
    CfgNum = dft
End Function
Private Function CfgTxt(ByVal cfg As Worksheet, ByVal key As String, ByVal dft As String) As String
    Dim r As Long
    For r = 4 To 30
        If InStr(CStr(cfg.Cells(r, 11).Value), key) > 0 Then
            If Len(Trim$(CStr(cfg.Cells(r, 12).Value))) > 0 Then
                CfgTxt = Trim$(CStr(cfg.Cells(r, 12).Value))
                Exit Function
            End If
        End If
    Next r
    CfgTxt = dft
End Function
'==================================================================
' スコアリング・配置ヘルパー
'==================================================================
'--- 予定出勤数の更新(薬剤師/事務員 別) ---
Private Sub CovAdd(ByVal i As Long, ByVal j As Long, ByVal d As Long)
    If mKind(i) = "薬剤師" Then
        mCov(j) = mCov(j) + d
    ElseIf mKind(i) = "事務員" Then
        mCovG(j) = mCovG(j) + d
    End If
End Sub
'--- ノルマ外の休み記号(L11・カンマ区切り・部分一致) ---
Private Function IsPaidOff(ByVal v As String) As Boolean
    Dim parts() As String, p As Long
    parts = Split(Replace(mPaidSyms, "、", ","), ",")
    For p = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(p))) > 0 Then
            If InStr(v, Trim$(parts(p))) > 0 Then IsPaidOff = True: Exit Function
        End If
    Next p
End Function
'--- その日を休みにする良さ(大きいほど休み向き) ---
Private Function OffScore(ByVal i As Long, ByVal j As Long) As Double
    Dim s As Double, L As Long, lft As Long, rgt As Long
    If mKind(i) = "薬剤師" Then
        s = s + 5# * (mCov(j) - 1 - mDayReq(j))          ' 不足を日別に均す(ソフト)
        If mDayDoc(j) = 5 Then s = s + 3# * (FiveCnt(i) - FiveAvg())
    ElseIf mKind(i) = "事務員" Then
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
End Function
'--- 既存休に隣接して連休(上限内)になる位置を優遇 ---
Private Function AdjBonus(ByVal i As Long, ByVal j As Long) As Double
    Dim total As Long
    total = 1 + OffRunBefore(i, j) + OffRunAfter(i, j)
    If total >= 2 And total <= mMaxOffRun Then
        AdjBonus = 4
    ElseIf total > mMaxOffRun Then
        AdjBonus = -3 * (total - mMaxOffRun)
    End If
End Function
'--- 指定週内に size 日連続の公休ブロックを最良位置へ ---
Private Function PlaceOffBlock(ByVal i As Long, ByVal wk As Long, ByVal size As Long) As Boolean
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
End Function
'--- 指定週内に公休1日を最良位置へ(既存休に寄せる) ---
Private Function PlaceOffSingle(ByVal i As Long, ByVal wk As Long) As Boolean
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
End Function
'--- 連勤上限超えの緩和: 連勤の中央を休みにし、他の自動公休と入替 ---
Private Sub RepairRuns(ByVal i As Long)
    Dim j As Long, L As Long, lft As Long, rgt As Long
    Dim ctr As Long, off As Long, k As Long, cand As Long, bk As Long, bs As Double, sc As Double
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
                                If mKind(i) = "薬剤師" Then sc = mDayReq(k) - mCov(k)
                                If mKind(i) = "事務員" Then sc = 1 - mCovG(k)
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
End Sub
'--- 医師5名日の出勤を均等化(通常ルールの薬剤師間で誤差1以内を目標) ---
Private Sub FiveBalance()
    Dim pass As Long, i As Long, j As Long, f As Long
    Dim mxI As Long, mnI As Long, mxV As Long, mnV As Long
    Dim d5 As Long, dx As Long, swapped As Boolean
    For pass = 1 To 100
        mxI = 0: mnI = 0: mxV = -1: mnV = 32767
        For i = 1 To mNP
            If mKind(i) = "薬剤師" And Not mLeave(i) And mRule(i) = "通常" Then
                f = FiveCnt(i)
                If f > mxV Then mxV = f: mxI = i
                If f < mnV Then mnV = f: mnI = i
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
End Sub
'--- ○●▲の個人差を均等化(誤差2以内を目標・同じ日の2人で記号を交換) ---
Private Sub SymbolBalance()
    Dim pass As Long, s As Long, i As Long, j As Long
    Dim mxI As Long, mnI As Long, mxV As Long, mnV As Long, c As Long
    Dim sym As String, other As String, done As Boolean
    For pass = 1 To 300
        done = True
        For s = 1 To 3
            sym = Choose(s, SYM_EARLY, SYM_MID, SYM_LATE)
            mxI = 0: mnI = 0: mxV = -1: mnV = 32767
            For i = 1 To mNP
                If mKind(i) = "薬剤師" And Not mLeave(i) And mRule(i) <> "手動" Then
                    If sym = SYM_EARLY Or mCanLate(i) Then
                        c = SymCnt(i, sym)
                        If c > mxV Then mxV = c: mxI = i
                        If c < mnV Then mnV = c: mnI = i
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
End Sub
Private Function SymCnt(ByVal i As Long, ByVal sym As String) As Long
    If sym = SYM_EARLY Then SymCnt = mCntE(i)
    If sym = SYM_MID Then SymCnt = mCntM(i)
    If sym = SYM_LATE Then SymCnt = mCntL(i)
End Function
Private Sub AddCnt(ByVal i As Long, ByVal sym As String, ByVal d As Long)
    If sym = SYM_EARLY Then mCntE(i) = mCntE(i) + d
    If sym = SYM_MID Then mCntM(i) = mCntM(i) + d
    If sym = SYM_LATE Then mCntL(i) = mCntL(i) + d
End Sub
'==================================================================
' 計測ヘルパー
'==================================================================
Private Function RunLenAt(ByVal i As Long, ByVal j As Long, ByRef lft As Long, ByRef rgt As Long) As Long
    Dim a As Long, b As Long
    a = j: b = j
    Do While a > 1
        If mPlan(i, a - 1) = ST_WORK Or mPlan(i, a - 1) = ST_FWORK Then a = a - 1 Else Exit Do
    Loop
    Do While b < mND
        If mPlan(i, b + 1) = ST_WORK Or mPlan(i, b + 1) = ST_FWORK Then b = b + 1 Else Exit Do
    Loop
    lft = j - a: rgt = b - j
    RunLenAt = b - a + 1
End Function
Private Function WorkRunIf(ByVal i As Long, ByVal k As Long) As Long
    Dim a As Long, b As Long
    a = k: b = k
    Do While a > 1
        If mPlan(i, a - 1) = ST_WORK Or mPlan(i, a - 1) = ST_FWORK Then a = a - 1 Else Exit Do
    Loop
    Do While b < mND
        If mPlan(i, b + 1) = ST_WORK Or mPlan(i, b + 1) = ST_FWORK Then b = b + 1 Else Exit Do
    Loop
    WorkRunIf = b - a + 1
End Function
Private Function OffRunBefore(ByVal i As Long, ByVal j As Long) As Long
    Dim a As Long
    a = j - 1
    Do While a >= 1
        If mPlan(i, a) = ST_OFF Or mPlan(i, a) = ST_FOFF Then
            OffRunBefore = OffRunBefore + 1: a = a - 1
        Else
            Exit Do
        End If
    Loop
End Function
Private Function OffRunAfter(ByVal i As Long, ByVal j As Long) As Long
    Dim b As Long
    b = j + 1
    Do While b <= mND
        If mPlan(i, b) = ST_OFF Or mPlan(i, b) = ST_FOFF Then
            OffRunAfter = OffRunAfter + 1: b = b + 1
        Else
            Exit Do
        End If
    Loop
End Function
Private Function FiveCnt(ByVal i As Long) As Long
    Dim j As Long
    For j = 1 To mND
        If mDayIn(j) And mDayDoc(j) = 5 Then
            If mPlan(i, j) = ST_WORK Or mPlan(i, j) = ST_FWORK Then FiveCnt = FiveCnt + 1
        End If
    Next j
End Function
Private Function FiveAvg() As Double
    Dim i As Long, t As Long, c As Long
    For i = 1 To mNP
        If mKind(i) = "薬剤師" And Not mLeave(i) Then
            t = t + FiveCnt(i): c = c + 1
        End If
    Next i
    If c > 0 Then FiveAvg = t / c
End Function
Private Function MaxRun(ByVal i As Long) As Long
    Dim j As Long, cur As Long
    For j = 1 To mND
        If mPlan(i, j) = ST_WORK Or mPlan(i, j) = ST_FWORK Then
            cur = cur + 1
            If cur > MaxRun Then MaxRun = cur
        Else
            cur = 0
        End If
    Next j
End Function
Private Function MaxOffRun(ByVal i As Long) As Long
    Dim j As Long, cur As Long
    For j = 1 To mND
        If mPlan(i, j) = ST_OFF Or mPlan(i, j) = ST_FOFF Then
            cur = cur + 1
            If cur > MaxOffRun Then MaxOffRun = cur
        Else
            cur = 0
        End If
    Next j
End Function
'==================================================================
' 共通ヘルパー(動的範囲・スタンプ)
'==================================================================
Private Function AutoShiftRange(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set AutoShiftRange = ws.Parent.Names("シフトパレット範囲").RefersToRange
    On Error GoTo 0
    If AutoShiftRange Is Nothing Then Set AutoShiftRange = ws.Range("B13:AF22")
End Function
Private Function AutoPaletteRange(ByVal ws As Worksheet) As Range
    On Error Resume Next
    Set AutoPaletteRange = ws.Parent.Names("シフトパレット").RefersToRange
    On Error GoTo 0
    If AutoPaletteRange Is Nothing Then Set AutoPaletteRange = ws.Range("B28:M28")
End Function
Private Sub StampCell(ByVal ws As Worksheet, ByVal c As Range, ByVal v As String)
    Dim pal As Range, i As Long, src As Range
    If c.HasFormula Then Exit Sub
    If Len(v) = 0 Then Exit Sub
    Set pal = AutoPaletteRange(ws)
    For i = 1 To pal.Cells.Count
        If Trim$(CStr(pal.Cells(1, i).Value)) = v Then Set src = pal.Cells(1, i): Exit For
    Next i
    c.Value = v
    If Not src Is Nothing Then
        c.Font.Color = src.Font.Color
        c.Font.Bold = src.Font.Bold
        If src.Interior.Pattern <> xlNone Then c.Interior.Color = src.Interior.Color
    End If
End Sub
Private Sub ParseWD(ByVal s As String, ByRef pWD() As Boolean, ByVal i As Long)
    Dim k As Long, w As Long
    Const WDS As String = "日月火水木金土"
    For k = 1 To Len(Trim$(s))
        w = InStr(WDS, Mid$(Trim$(s), k, 1))
        If w >= 1 And w <= 7 Then pWD(i, w) = True
    Next k
End Sub

'==================================================================
' 手動変更ログ(シートモジュール「シフト」から呼び出し)
'==================================================================
Public Function ShiftGrid() As Range
    Set ShiftGrid = AutoShiftRange(Worksheets("シフト"))
End Function
'--- 手動変更を1セッションとして記録(複数セル同時変更=同一セッション) ---
Public Sub LogManualSession(ByVal addrs As Variant, ByVal oldVals As Variant, _
                            ByVal oldFonts As Variant, ByVal oldBolds As Variant, _
                            ByVal oldFills As Variant, ByVal n As Long)
    Dim lg As Worksheet, lr As Long, sess As Long, k As Long
    Dim ws As Worksheet, c As Range
    Set ws = Worksheets("シフト")
    Set lg = GetLogSheet()
    lr = LogLastRow(lg)
    sess = NextSession(lg)
    For k = 1 To n
        Set c = ws.Range(CStr(addrs(k)))
        If CStr(c.Value) <> CStr(oldVals(k)) Then   ' 実際に値が変わったセルのみ
            lr = lr + 1
            lg.Cells(lr, 1).Value = sess
            lg.Cells(lr, 2).Value = Now
            lg.Cells(lr, 3).Value = "手動"
            lg.Cells(lr, 4).Value = CStr(addrs(k))
            lg.Cells(lr, 5).Value = CStr(oldVals(k))
            lg.Cells(lr, 6).Value = CStr(c.Value)
            lg.Cells(lr, 8).Value = oldFonts(k)
            lg.Cells(lr, 9).Value = IIf(oldBolds(k), "TRUE", "FALSE")
            If Not IsEmpty(oldFills(k)) Then lg.Cells(lr, 10).Value = oldFills(k)
        End If
    Next k
End Sub