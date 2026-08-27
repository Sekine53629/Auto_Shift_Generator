Attribute VB_Name = "ShiftAutoLog"
Option Explicit
'==================================================================
'  ShiftAutoLog v9.2.0
'  設定チェック・変更ログ・白紙化と共通の小物ヘルパ。
'  ※ ShiftAuto / ShiftAutoPlace / ShiftAutoLog の3本で1組。
'    共有状態は ShiftAuto の Public 変数に置く。
'==================================================================
Private Const MODULE_NAME As String = "ShiftAutoLog"



'==================================================================
' レイアウト整合性チェック(単体実行用)
'   実行せずに、シフト表とマスタの対応だけを確認できます
'==================================================================
Public Sub シフト設定チェック()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, cfg As Worksheet, grid As Range
    Dim i As Long, r As Long, nm As String, kd As String
    Dim nmList() As String, nN As Long, ii As Long, hit As Long
    Dim missing As String, orphan As String, dupName As String, badKind As String
    Dim blankRows As String, labelRows As String, msg As String
    Dim docRow As Long, drift As Long, warnRange As String

    Set ws = ShiftSheet()
    Set cfg = SheetOrNothing(SHT_CFG)
    If cfg Is Nothing Then
        MsgBox "設定シート「" & SHT_CFG & "」がありません。" & vbCrLf & _
               "ShiftSchema_不足シート生成 で作成できます。", vbExclamation: Exit Sub
    End If
    Set grid = ShiftInputRange(ws)
    If grid Is Nothing Then
        MsgBox "シフト入力欄を特定できません。" & vbCrLf & _
               "B列の開始日の数式、またはA列の「" & LBL_NOTE & "」「" & LBL_DOC & _
               "」を確認してください。", vbExclamation: Exit Sub
    End If

    '--- 範囲が集計行に近すぎないか(下端は医師数の DOC_GAP 行上まで) ---
    docRow = ShiftDocRow(ws)
    drift = ShiftRangeDrift(ws)
    If docRow > 0 And drift > 0 Then
        warnRange = "■ 入力範囲が「" & LBL_DOC & "」行(" & docRow & "行)の" & _
                    DOC_GAP & "行上を " & drift & " 行超えています" & vbCrLf & _
                    "　終端は " & (docRow - DOC_GAP) & "行 が正しい位置です" & vbCrLf & vbCrLf
    End If

    ReDim nmList(1 To grid.Rows.Count): nN = 0
    For i = 1 To grid.Rows.Count
        nm = Trim$(CStr(ws.Cells(grid.Row + i - 1, 1).Value))
        If Len(nm) = 0 Then
            blankRows = blankRows & IIf(Len(blankRows) > 0, ", ", "") & (grid.Row + i - 1) & "行"
        ElseIf IsNonName(nm) Then
            labelRows = labelRows & IIf(Len(labelRows) > 0, ", ", "") & _
                        (grid.Row + i - 1) & "行(" & nm & ")"
        Else
            nN = nN + 1: nmList(nN) = nm
            For ii = 1 To nN - 1
                If nmList(ii) = nm Then
                    If InStr(dupName, "・" & nm & vbCrLf) = 0 Then dupName = dupName & "・" & nm & vbCrLf
                End If
            Next ii
            hit = 0: kd = ""
            r = CFG_FIRST_ROW
            Do While Len(Trim$(CStr(cfg.Cells(r, CFG_COL_NAME).Value))) > 0
                If Trim$(CStr(cfg.Cells(r, CFG_COL_NAME).Value)) = nm Then
                    hit = 1: kd = Trim$(CStr(cfg.Cells(r, CFG_COL_KIND).Value)): Exit Do
                End If
                r = r + 1
            Loop
            If hit = 0 Then
                missing = missing & "・" & nm & vbCrLf
            ElseIf Len(kd) > 0 And kd <> KIND_PH And kd <> KIND_CL Then
                badKind = badKind & "・" & nm & " : 区分「" & kd & "」" & vbCrLf
            End If
        End If
    Next i

    r = CFG_FIRST_ROW
    Do While Len(Trim$(CStr(cfg.Cells(r, CFG_COL_NAME).Value))) > 0
        nm = Trim$(CStr(cfg.Cells(r, CFG_COL_NAME).Value))
        If Not IsNonName(nm) Then
            hit = 0
            For ii = 1 To nN
                If nmList(ii) = nm Then hit = 1: Exit For
            Next ii
            If hit = 0 Then orphan = orphan & "・" & nm & vbCrLf
        End If
        r = r + 1
    Loop

    msg = "シフト入力範囲 : " & grid.Address(False, False) & vbCrLf & _
          "　(上端=再掲日付行の1行下 / 下端=" & LBL_DOC & "の" & DOC_GAP & "行上)" & vbCrLf & _
          "対象者(氏名有) : " & nN & "名" & vbCrLf & _
          "空行(スキップ) : " & IIf(Len(blankRows) > 0, blankRows, "なし") & vbCrLf & _
          "集計行(除外)   : " & IIf(Len(labelRows) > 0, labelRows, "なし") & vbCrLf & vbCrLf
    msg = msg & warnRange
    If Len(dupName) > 0 Then msg = msg & "■ 氏名の重複" & vbCrLf & dupName & vbCrLf
    If Len(badKind) > 0 Then msg = msg & "■ 区分が「" & KIND_PH & "」「" & KIND_CL & "」以外" & vbCrLf & badKind & vbCrLf
    If Len(missing) > 0 Then msg = msg & "■ マスタ未登録(シフト表にあるが設定が無い)" & vbCrLf & missing & vbCrLf
    If Len(orphan) > 0 Then msg = msg & "■ 孤児(マスタにあるがシフト表に無い)" & vbCrLf & orphan & vbCrLf
    If Len(warnRange) = 0 And Len(dupName) = 0 And Len(badKind) = 0 _
       And Len(missing) = 0 And Len(orphan) = 0 Then
        msg = msg & "整合性の問題は見つかりませんでした。"
    End If
    MsgBox msg, vbInformation, AUTOCHECK_MACRO

    LogSuccess MODULE_NAME, "シフト設定チェック", _
               "Checked settings: names=" & nN & "; range=" & grid.Address(False, False)
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "シフト設定チェック", Err.Number, Err.Description, Erl, ""
End Sub


'==================================================================
' 差分ログ・戻す・白紙化
'==================================================================
'--- ログシート取得(無ければ ShiftSchema に作らせる) ---
Public Function GetLogSheet() As Worksheet
    On Error GoTo ErrHandler
    Dim lg As Worksheet

10  Set lg = SheetOrNothing(SHT_LOG)
20  If lg Is Nothing Then
        '--- 生成は ShiftSchema に委譲(見出し・書式の二重定義を避ける) ---
30      ShiftSchema_変更ログ生成
40      Set lg = SheetOrNothing(SHT_LOG)
50  End If
    '--- 拡張列(取消済/前書式)は本モジュール固有なので必要なら補う ---
60  If Not lg Is Nothing Then
70      If Len(Trim$(CStr(lg.Cells(LOG_HDR_ROW, 7).Value))) = 0 Then
80          lg.Range(lg.Cells(LOG_HDR_ROW, 7), lg.Cells(LOG_HDR_ROW, 10)).Value = _
                Array("取消済", "前文字色", "前太字", "前塗り色")
90          lg.Range(lg.Cells(LOG_HDR_ROW, 7), lg.Cells(LOG_HDR_ROW, 10)).Font.Bold = True
100         lg.Range(lg.Cells(LOG_HDR_ROW, 7), lg.Cells(LOG_HDR_ROW, 10)) _
                .Interior.Color = RGB(217, 217, 217)
110     End If
120     lg.Columns(LOG_COL_BEFORE).NumberFormat = "@"
130     lg.Columns(LOG_COL_AFTER).NumberFormat = "@"
140 End If
150 Set GetLogSheet = lg
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "GetLogSheet", Err.Number, Err.Description, Erl, ""
End Function


Public Function LogLastRow(ByVal lg As Worksheet) As Long
    On Error GoTo ErrHandler
10  LogLastRow = lg.Cells(lg.Rows.Count, LOG_COL_TIME).End(xlUp).Row
20  If LogLastRow < LOG_HDR_ROW Then LogLastRow = LOG_HDR_ROW
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "LogLastRow", Err.Number, Err.Description, Erl, ""
    LogLastRow = LOG_HDR_ROW
End Function


Public Function NextSession(ByVal lg As Worksheet) As Long
    On Error GoTo ErrHandler
    Dim lr As Long
10  lr = LogLastRow(lg)
20  If lr < LOG_FIRST_ROW Then
30      NextSession = 1
40  Else
50      NextSession = CLng(Val(lg.Cells(lr, LOG_COL_TIME).Value)) + 1
60  End If
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "NextSession", Err.Number, Err.Description, Erl, ""
    NextSession = 1
End Function


'--- 1セル分の差分を記録(変更前の値と書式を保存してから書き換える前提) ---
'    列: 1=セッション 2=日時 3=操作 4=セル 5=変更前 6=変更後
'        7=取消済 8=前文字色 9=前太字 10=前塗り色
Public Sub LogChange(ByVal lg As Worksheet, ByRef lr As Long, ByVal sess As Long, _
                      ByVal op As String, ByVal c As Range, ByVal newV As String)
    On Error GoTo ErrHandler
10  lr = lr + 1
20  lg.Cells(lr, 1).Value = sess
30  lg.Cells(lr, 2).Value = Now
40  lg.Cells(lr, 3).Value = op
50  lg.Cells(lr, 4).Value = c.Address(False, False)
60  lg.Cells(lr, 5).Value = CStr(c.Value)
70  lg.Cells(lr, 6).Value = newV
80  lg.Cells(lr, 8).Value = c.Font.Color
90  lg.Cells(lr, 9).Value = IIf(c.Font.Bold, "TRUE", "FALSE")
100 If c.Interior.Pattern <> xlNone Then lg.Cells(lr, 10).Value = c.Interior.Color
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "LogChange", Err.Number, Err.Description, Erl, _
             "lr=" & lr & "; sess=" & sess & "; op=" & op & "; newV=" & newV
End Sub


'--- 最後のセッションを逆再生して復元(繰り返し実行で1回ずつ遡れる) ---
Public Sub シフト変更を戻す()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, lg As Worksheet, lr As Long, r As Long
    Dim sess As Long, cnt As Long, c As Range, oldV As String
    Dim op As String, tm As String

    Set ws = ShiftSheet()
    Set lg = SheetOrNothing(SHT_LOG)
    If lg Is Nothing Then
        MsgBox "変更ログがありません。" & vbCrLf & _
               "(ログは「シフト自動作成」「シフト白紙化」の実行時に記録されます)", vbExclamation
        Exit Sub
    End If
    lr = LogLastRow(lg)

    '--- 最後の未取消セッションを特定 ---
    sess = 0
    For r = lr To LOG_FIRST_ROW Step -1
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
    For r = LOG_FIRST_ROW To lr
        If CLng(Val(lg.Cells(r, 1).Value)) = sess And Trim$(CStr(lg.Cells(r, 7).Value)) = "" Then
            cnt = cnt + 1
        End If
    Next r
    If MsgBox("セッション#" & sess & "「" & op & "」(" & tm & "・" & cnt & "セル)を取り消し、" & vbCrLf & _
              "変更前の状態に戻します。よろしいですか?", _
              vbYesNo + vbQuestion, "シフト変更を戻す") <> vbYes Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    '--- 逆順(後に書いたセルから)に復元 ---
    For r = lr To LOG_FIRST_ROW Step -1
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
    LogSuccess MODULE_NAME, "シフト変更を戻す", "Reverted session #" & sess & " (" & cnt & " cells)"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "シフト変更を戻す", Err.Number, Err.Description, Erl, ""
End Sub


'--- 入力欄を白紙化(数式セルは保護・全消去分をログに記録=復元可) ---
Public Sub シフト白紙化()
    On Error GoTo ErrHandler
    Dim ws As Worksheet, grid As Range, c As Range
    Dim lg As Worksheet, sess As Long, lr As Long, cnt As Long
    Dim hasVal As Boolean, hasFmt As Boolean

    Set ws = ShiftSheet()
    Set grid = ShiftInputRange(ws)
    If grid Is Nothing Then
        MsgBox "シフト入力欄を特定できません。" & vbCrLf & _
               "B列の開始日の数式、またはA列の「" & LBL_NOTE & "」「" & LBL_DOC & _
               "」を確認してください。", vbExclamation: Exit Sub
    End If
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
    LogSuccess MODULE_NAME, "シフト白紙化", "Cleared " & cnt & " cells in " & grid.Address(False, False)
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "シフト白紙化", Err.Number, Err.Description, Erl, ""
End Sub


'==================================================================
' 設定シートの生成(ShiftSchema に委譲)
'==================================================================
'--- 旧 BuildCfgSheet。シート生成の実体は ShiftSchema に一本化した。
'    氏名の取込だけは本モジュールが行う(シフト表の並びに依存するため)。
Public Function BuildCfgSheet(ByVal ws As Worksheet) As Worksheet
    On Error GoTo ErrHandler
    Dim cfg As Worksheet, grid As Range, i As Long, r As Long, nm As String

10  Set grid = ShiftInputRange(ws)
20  If grid Is Nothing Then Exit Function

    '--- 見出し・全体設定・入力規則は ShiftSchema が作る ---
30  ShiftSchema_自動作成設定生成
40  Set cfg = SheetOrNothing(SHT_CFG)
50  If cfg Is Nothing Then Exit Function

    '--- 氏名をシフト表A列から取込(空行・集計行は飛ばす) ---
60  r = CFG_FIRST_ROW
70  For i = 1 To grid.Rows.Count
80      nm = Trim$(CStr(ws.Cells(grid.Row + i - 1, 1).Value))
90      If Len(nm) > 0 And Not IsNonName(nm) Then
100         If Len(Trim$(CStr(cfg.Cells(r, CFG_COL_NAME).Value))) = 0 Then
110             cfg.Cells(r, CFG_COL_NAME).NumberFormat = "@"
120             cfg.Cells(r, CFG_COL_NAME).Value = nm
130             cfg.Cells(r, CFG_COL_KIND).Value = KIND_PH   ' 事務員は手動で変更
140             cfg.Cells(r, CFG_COL_RULE).Value = "通常"
150             cfg.Cells(r, CFG_COL_LATE).Value = "可"
160         End If
170         r = r + 1
180     End If
190 Next i
200 cfg.Columns(CFG_COL_NAME).AutoFit
210 Set BuildCfgSheet = cfg

    LogSuccess MODULE_NAME, "BuildCfgSheet", _
               "Delegated to ShiftSchema; imported " & (r - CFG_FIRST_ROW) & " member rows"
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "BuildCfgSheet", Err.Number, Err.Description, Erl, _
             "r=" & r & "; i=" & i
End Function


'==================================================================
' 設定読込ヘルパー(K列ラベルの部分一致 → L列の値)
'   ※CFG_SCAN_ROWS の宣言は H5 冒頭の定数部にあります
'     (VBA はモジュールレベル宣言をプロシージャより前に置く必要がある)
'==================================================================
Public Function CfgNum(ByVal cfg As Worksheet, ByVal key As String, _
                        ByVal dft As Double) As Double
    On Error GoTo ErrHandler
    Dim r As Long
10  For r = CFG_SET_ROW To CFG_SET_ROW + CFG_SCAN_ROWS
20      If InStr(CStr(cfg.Cells(r, CFG_COL_SETK).Value), key) > 0 Then
30          If Len(Trim$(CStr(cfg.Cells(r, CFG_COL_SETV).Value))) > 0 Then
40              If IsNumeric(cfg.Cells(r, CFG_COL_SETV).Value) Then
50                  CfgNum = CDbl(cfg.Cells(r, CFG_COL_SETV).Value)
60                  Exit Function
70              End If
80          End If
90      End If
100 Next r
110 CfgNum = dft
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CfgNum", Err.Number, Err.Description, Erl, _
             "key=" & key & "; dft=" & dft
    CfgNum = dft
End Function


Public Function CfgTxt(ByVal cfg As Worksheet, ByVal key As String, _
                        ByVal dft As String) As String
    On Error GoTo ErrHandler
    Dim r As Long
10  For r = CFG_SET_ROW To CFG_SET_ROW + CFG_SCAN_ROWS
20      If InStr(CStr(cfg.Cells(r, CFG_COL_SETK).Value), key) > 0 Then
30          If Len(Trim$(CStr(cfg.Cells(r, CFG_COL_SETV).Value))) > 0 Then
40              CfgTxt = Trim$(CStr(cfg.Cells(r, CFG_COL_SETV).Value))
50              Exit Function
60          End If
70      End If
80  Next r
90  CfgTxt = dft
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "CfgTxt", Err.Number, Err.Description, Erl, _
             "key=" & key & "; dft=" & dft
    CfgTxt = dft
End Function


'==================================================================
' 手動変更ログ(シートモジュール「シフト」から呼び出し)
'==================================================================
'--- 手動変更を1セッションとして記録(複数セル同時変更=同一セッション) ---
Public Sub LogManualSession(ByVal addrs As Variant, ByVal oldVals As Variant, _
                            ByVal oldFonts As Variant, ByVal oldBolds As Variant, _
                            ByVal oldFills As Variant, ByVal n As Long)
    On Error GoTo ErrHandler
    Dim lg As Worksheet, lr As Long, sess As Long, k As Long
    Dim ws As Worksheet, c As Range
    Set ws = ShiftSheet()
    Set lg = GetLogSheet()
    If lg Is Nothing Then Exit Sub
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
    LogSuccess MODULE_NAME, "LogManualSession", "Logged manual session for " & n & " cells"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "LogManualSession", Err.Number, Err.Description, Erl, "n=" & n
End Sub


'--- ログ全消去(期替わり用リセット) ---
Public Sub シフトログリセット(Optional ByVal ask As Boolean = True)
    On Error GoTo ErrHandler
    Dim lg As Worksheet, lr As Long
    Set lg = SheetOrNothing(SHT_LOG)
    If lg Is Nothing Then Exit Sub
    lr = lg.Cells(lg.Rows.Count, LOG_COL_TIME).End(xlUp).Row
    If lr < LOG_FIRST_ROW Then Exit Sub
    If ask Then
        If MsgBox("変更ログ(" & lr - LOG_HDR_ROW & "行)をすべて消去します。" & vbCrLf & _
                  "消去後は「シフト変更を戻す」で過去の状態に戻せなくなります。" & vbCrLf & _
                  "よろしいですか?", _
                  vbYesNo + vbExclamation, "シフトログ リセット") <> vbYes Then Exit Sub
    End If
    lg.Rows(LOG_FIRST_ROW & ":" & lr).Delete
    LogSuccess MODULE_NAME, "シフトログリセット", _
               "Cleared change log (" & lr - LOG_HDR_ROW & " rows)"
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "シフトログリセット", Err.Number, Err.Description, Erl, _
             "ask=" & ask
End Sub


Public Function SymCnt(ByVal i As Long, ByVal sym As String) As Long
    On Error GoTo ErrHandler
    If sym = SYM_EARLY Then SymCnt = mCntE(i)
    If sym = SYM_MID Then SymCnt = mCntM(i)
    If sym = SYM_LATE Then SymCnt = mCntL(i)
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "SymCnt", Err.Number, Err.Description, Erl, _
             "i=" & i & "; sym=" & sym
End Function


Public Sub AddCnt(ByVal i As Long, ByVal sym As String, ByVal d As Long)
    On Error GoTo ErrHandler
    If sym = SYM_EARLY Then mCntE(i) = mCntE(i) + d
    If sym = SYM_MID Then mCntM(i) = mCntM(i) + d
    If sym = SYM_LATE Then mCntL(i) = mCntL(i) + d
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "AddCnt", Err.Number, Err.Description, Erl, _
             "i=" & i & "; sym=" & sym & "; d=" & d
End Sub


'==================================================================
' 計測ヘルパー
'==================================================================
Public Function RunLenAt(ByVal i As Long, ByVal j As Long, _
                          ByRef lft As Long, ByRef rgt As Long) As Long
    On Error GoTo ErrHandler
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
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "RunLenAt", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
End Function


Public Function WorkRunIf(ByVal i As Long, ByVal k As Long) As Long
    On Error GoTo ErrHandler
    Dim a As Long, b As Long
    a = k: b = k
    Do While a > 1
        If mPlan(i, a - 1) = ST_WORK Or mPlan(i, a - 1) = ST_FWORK Then a = a - 1 Else Exit Do
    Loop
    Do While b < mND
        If mPlan(i, b + 1) = ST_WORK Or mPlan(i, b + 1) = ST_FWORK Then b = b + 1 Else Exit Do
    Loop
    WorkRunIf = b - a + 1
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "WorkRunIf", Err.Number, Err.Description, Erl, _
             "i=" & i & "; k=" & k
End Function


Public Function OffRunBefore(ByVal i As Long, ByVal j As Long) As Long
    On Error GoTo ErrHandler
    Dim a As Long
    a = j - 1
    Do While a >= 1
        If mPlan(i, a) = ST_OFF Or mPlan(i, a) = ST_FOFF Then
            OffRunBefore = OffRunBefore + 1: a = a - 1
        Else
            Exit Do
        End If
    Loop
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "OffRunBefore", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
End Function


Public Function OffRunAfter(ByVal i As Long, ByVal j As Long) As Long
    On Error GoTo ErrHandler
    Dim b As Long
    b = j + 1
    Do While b <= mND
        If mPlan(i, b) = ST_OFF Or mPlan(i, b) = ST_FOFF Then
            OffRunAfter = OffRunAfter + 1: b = b + 1
        Else
            Exit Do
        End If
    Loop
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "OffRunAfter", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j
End Function


Public Function FiveCnt(ByVal i As Long) As Long
    On Error GoTo ErrHandler
    Dim j As Long
    For j = 1 To mND
        If mDayIn(j) And mDayDoc(j) = 5 Then
            If mPlan(i, j) = ST_WORK Or mPlan(i, j) = ST_FWORK Then FiveCnt = FiveCnt + 1
        End If
    Next j
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "FiveCnt", Err.Number, Err.Description, Erl, "i=" & i
End Function


Public Function FiveAvg() As Double
    On Error GoTo ErrHandler
    Dim i As Long, t As Long, c As Long
    For i = 1 To mNP
        If Not mSkipRow(i) Then
        If mKind(i) = KIND_PH And Not mLeave(i) Then
            t = t + FiveCnt(i): c = c + 1
        End If
        End If
    Next i
    If c > 0 Then FiveAvg = t / c
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "FiveAvg", Err.Number, Err.Description, Erl, ""
End Function


Public Function MaxRun(ByVal i As Long) As Long
    On Error GoTo ErrHandler
    Dim j As Long, cur As Long
    For j = 1 To mND
        If mPlan(i, j) = ST_WORK Or mPlan(i, j) = ST_FWORK Then
            cur = cur + 1
            If cur > MaxRun Then MaxRun = cur
        Else
            cur = 0
        End If
    Next j
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "MaxRun", Err.Number, Err.Description, Erl, "i=" & i
End Function


Public Function MaxOffRun(ByVal i As Long) As Long
    On Error GoTo ErrHandler
    Dim j As Long, cur As Long
    For j = 1 To mND
        If mPlan(i, j) = ST_OFF Or mPlan(i, j) = ST_FOFF Then
            cur = cur + 1
            If cur > MaxOffRun Then MaxOffRun = cur
        Else
            cur = 0
        End If
    Next j
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "MaxOffRun", Err.Number, Err.Description, Erl, "i=" & i
End Function


'==================================================================
' 共通ヘルパー(スタンプ・曜日解析)
'==================================================================
'--- パレットの書式でセルに記号を書き込む ---
Public Sub StampCell(ByVal ws As Worksheet, ByVal c As Range, ByVal v As String)
    On Error GoTo ErrHandler
    Dim pal As Range, i As Long, src As Range
    If c.HasFormula Then Exit Sub
    If Len(v) = 0 Then Exit Sub
    Set pal = PaletteRange(ws)
    If Not pal Is Nothing Then
        For i = 1 To pal.Cells.Count
            If Trim$(CStr(pal.Cells(1, i).Value)) = v Then Set src = pal.Cells(1, i): Exit For
        Next i
    End If
    c.Value = v
    If Not src Is Nothing Then
        c.Font.Color = src.Font.Color
        c.Font.Bold = src.Font.Bold
        If src.Interior.Pattern <> xlNone Then c.Interior.Color = src.Interior.Color
    End If
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "StampCell", Err.Number, Err.Description, Erl, "v=" & v
End Sub


'--- 「月火金土」等の文字列を曜日フラグに変換 ---
Public Sub ParseWD(ByVal s As String, ByRef pWD() As Boolean, ByVal i As Long)
    On Error GoTo ErrHandler
    Dim k As Long, w As Long
    Const WDS As String = "日月火水木金土"
    For k = 1 To Len(Trim$(s))
        w = InStr(WDS, Mid$(Trim$(s), k, 1))
        If w >= 1 And w <= 7 Then pWD(i, w) = True
    Next k
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "ParseWD", Err.Number, Err.Description, Erl, _
             "s=" & s & "; i=" & i
End Sub


'==================================================================
' スコアリング・配置ヘルパー
'==================================================================
'--- 予定出勤数の更新(薬剤師/事務員 別) ---
Public Sub CovAdd(ByVal i As Long, ByVal j As Long, ByVal d As Long)
    On Error GoTo ErrHandler
    If mSkipRow(i) Then Exit Sub
    If mKind(i) = KIND_PH Then
        mCov(j) = mCov(j) + d
    ElseIf mKind(i) = KIND_CL Then
        mCovG(j) = mCovG(j) + d
    End If
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "CovAdd", Err.Number, Err.Description, Erl, _
             "i=" & i & "; j=" & j & "; d=" & d
End Sub


'--- ノルマ外の休み記号(L11・カンマ区切り・部分一致) ---
Public Function IsPaidOff(ByVal v As String) As Boolean
    On Error GoTo ErrHandler
    Dim parts() As String, p As Long
    parts = Split(Replace(mPaidSyms, "、", ","), ",")
    For p = LBound(parts) To UBound(parts)
        If Len(Trim$(parts(p))) > 0 Then
            If InStr(v, Trim$(parts(p))) > 0 Then IsPaidOff = True: Exit Function
        End If
    Next p
    Exit Function
ErrHandler:
    LogError MODULE_NAME, "IsPaidOff", Err.Number, Err.Description, Erl, "v=" & v
End Function
