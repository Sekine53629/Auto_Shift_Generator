Option Explicit
'==================================================================
'  シフト表 クリック入力 + 手動変更ログ + 期替わりリセット
'  ＜シートモジュール v2.0＞  2026-08-27
'  ※シフト表のシート見出しを右クリック →「コードの表示」で開き、
'    ここに貼り付けること(標準モジュールではない)
'
'  入力欄の範囲は ShiftCommon.ShiftInputRange が一元管理する。
'  イベントは高頻度で発生するため LogSuccess は呼ばない
'  (エラー時のみ LogError を記録する)。
'
'  v2.0 変更:
'   ・期替わり判定の対象セルを Me.Range("A1") 固定から
'     ShiftCommon.MonthCell(Me) 経由に変更
'     実シートの年月は「年月・タイトル行」のA列(A4)にあり、A1は空。
'     旧実装では A4 を書き換えてもログリセットの確認が出なかった。
'==================================================================
Private Const MODULE_NAME As String = "Sheet1"

' 一度に記録する変更セルの上限(全選択などの大量操作は記録対象外)
Private Const MAX_CACHE_CELLS As Long = 200

'--- 手動変更の記録用キャッシュ(選択時に変更前の状態を退避) ---
Private mN As Long
Private mAddrs() As String, mVals() As String
Private mFonts() As Long, mBolds() As Boolean, mFills() As Variant

'--- セルを選んだだけ(シングルクリック) ---
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Handled As Boolean
    On Error GoTo ErrHandler

10  CacheGridState Target            ' ログ用: 先に変更前の状態を退避
20  ShiftClick_Handle Target, "select", Handled
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "Worksheet_SelectionChange", Err.Number, Err.Description, Erl, _
             "target=" & Target.Address(False, False)
End Sub

'--- ダブルクリック → 次の記号へ／スタンプを押す ---
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim Handled As Boolean
    On Error GoTo ErrHandler

10  ShiftClick_Handle Target, "double", Handled
20  If Handled Then Cancel = True
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "Worksheet_BeforeDoubleClick", Err.Number, Err.Description, Erl, _
             "target=" & Target.Address(False, False)
End Sub

'--- 右クリック → 1つ前に戻す／選択範囲にまとめてスタンプ ---
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Dim Handled As Boolean
    On Error GoTo ErrHandler

10  ShiftClick_Handle Target, "right", Handled
20  If Handled Then Cancel = True
    Exit Sub
ErrHandler:
    LogError MODULE_NAME, "Worksheet_BeforeRightClick", Err.Number, Err.Description, Erl, _
             "target=" & Target.Address(False, False)
End Sub

'--- 値が変わった → 年月セルなら期替わりリセット確認 / 入力欄なら手動ログ記録 ---
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim g As Range, grid As Range, k As Long, hit As Long
    Dim aAddrs() As String, aVals() As String, aFonts() As Long
    Dim aBolds() As Boolean, aFills() As Variant
    Dim mc As Range
    On Error GoTo ErrHandler

    '--- 年月セル(実シートはA4)の変更 = 期替わり → ログリセット確認 ---
    '    位置は ShiftCommon.MonthCell が解決する(A1固定にしない)
5   Set mc = MonthCell(Me)
10  If Not mc Is Nothing Then
15  If Not Application.Intersect(Target, mc) Is Nothing Then
20      If IsDate(mc.Value) Then
30          If MsgBox("対象月が変わりました。変更ログをリセットしますか?" & vbCrLf & _
                      "(前の期のログはすべて消去されます)", _
                      vbYesNo + vbQuestion, "期替わり") = vbYes Then
40              Application.EnableEvents = False
50              シフトログリセット False
60              Application.EnableEvents = True
70          End If
80      End If
85  End If
90  End If

    '--- 入力欄の手動変更を記録 ---
100 Set grid = ShiftInputRange(Me)
110 If grid Is Nothing Then GoTo CleanUp
120 Set g = Application.Intersect(Target, grid)
130 If g Is Nothing Or mN = 0 Then GoTo CleanUp

    '--- キャッシュ済みセルのうち今回変更された範囲に含まれるものを抽出 ---
140 ReDim aAddrs(1 To mN): ReDim aVals(1 To mN): ReDim aFonts(1 To mN)
150 ReDim aBolds(1 To mN): ReDim aFills(1 To mN)
160 For k = 1 To mN
170     If Not Application.Intersect(g, Me.Range(mAddrs(k))) Is Nothing Then
180         hit = hit + 1
190         aAddrs(hit) = mAddrs(k): aVals(hit) = mVals(k)
200         aFonts(hit) = mFonts(k): aBolds(hit) = mBolds(k): aFills(hit) = mFills(k)
210     End If
220 Next k
230 If hit > 0 Then
240     Application.EnableEvents = False
250     LogManualSession aAddrs, aVals, aFonts, aBolds, aFills, hit
260     Application.EnableEvents = True
270 End If

CleanUp:
    On Error Resume Next
    Application.EnableEvents = True
    On Error GoTo 0
    CacheGridState Target            ' 連続編集に備えて再キャッシュ
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "Worksheet_Change", Err.Number, Err.Description, Erl, _
             "target=" & Target.Address(False, False) & "; cached=" & mN & "; hit=" & hit
    Resume CleanUp
End Sub

'--- 変更前の値・書式を退避(入力欄の内側のみ) ---
Private Sub CacheGridState(ByVal rng As Range)
    Dim g As Range, grid As Range, c As Range, k As Long
    On Error GoTo ErrHandler

10  mN = 0
20  Set grid = ShiftInputRange(Me)
30  If grid Is Nothing Then Exit Sub
40  Set g = Application.Intersect(rng, grid)
50  If g Is Nothing Then Exit Sub
60  If g.Cells.Count > MAX_CACHE_CELLS Then Exit Sub

70  ReDim mAddrs(1 To g.Cells.Count): ReDim mVals(1 To g.Cells.Count)
80  ReDim mFonts(1 To g.Cells.Count): ReDim mBolds(1 To g.Cells.Count)
90  ReDim mFills(1 To g.Cells.Count)
100 For Each c In g.Cells
110     k = k + 1
120     mAddrs(k) = c.Address(False, False)
130     mVals(k) = CStr(c.Value)
140     mFonts(k) = c.Font.Color
150     mBolds(k) = (c.Font.Bold = True)
160     If c.Interior.Pattern <> xlNone Then mFills(k) = c.Interior.Color Else mFills(k) = Empty
170 Next c
180 mN = k
    Exit Sub

ErrHandler:
    LogError MODULE_NAME, "CacheGridState", Err.Number, Err.Description, Erl, _
             "range=" & rng.Address(False, False) & "; k=" & k
    mN = 0
End Sub

