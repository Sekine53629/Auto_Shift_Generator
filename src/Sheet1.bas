Option Explicit
'==================================================================
'  シフト表 クリック入力マクロ + 手動変更ログ  ＜シートモジュール＞
'  ※ シフト表のシート見出しを右クリック →「コードの表示」で開き、
'    ここに貼り付けてください（標準モジュールではありません）
'==================================================================
'--- 手動変更の記録用キャッシュ(選択時に変更前の状態を退避) ---
Private mN As Long
Private mAddrs() As String, mVals() As String
Private mFonts() As Long, mBolds() As Boolean, mFills() As Variant

' セルを選んだだけ（シングルクリック）
'   → 既定では何もしません。CYCLE_TRIGGER を "single" にしたときだけ動きます。
Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim h As Boolean
    CacheGridState Target            ' ログ用: 先に変更前の状態を退避
    ShiftClick_Handle Target, "select", h
End Sub

' ダブルクリック → 次の記号へ／スタンプを押す
Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
    Dim h As Boolean
    ShiftClick_Handle Target, "double", h
    If h Then Cancel = True
End Sub

' 右クリック → 1つ前に戻す／選択範囲にまとめてスタンプ
Private Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Dim h As Boolean
    ShiftClick_Handle Target, "right", h
    If h Then Cancel = True
End Sub

' 値が変わった → キャッシュと比較して「手動」としてログに記録
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim g As Range, k As Long, hit As Long
    Dim aAddrs() As String, aVals() As String, aFonts() As Long
    Dim aBolds() As Boolean, aFills() As Variant
    On Error GoTo done
    Set g = Application.Intersect(Target, ShiftGrid())
    If g Is Nothing Or mN = 0 Then GoTo done
    '--- キャッシュ済みセルのうち今回変更された範囲に含まれるものを抽出 ---
    ReDim aAddrs(1 To mN): ReDim aVals(1 To mN): ReDim aFonts(1 To mN)
    ReDim aBolds(1 To mN): ReDim aFills(1 To mN)
    For k = 1 To mN
        If Not Application.Intersect(g, Me.Range(mAddrs(k))) Is Nothing Then
            hit = hit + 1
            aAddrs(hit) = mAddrs(k): aVals(hit) = mVals(k)
            aFonts(hit) = mFonts(k): aBolds(hit) = mBolds(k): aFills(hit) = mFills(k)
        End If
    Next k
    If hit > 0 Then
        Application.EnableEvents = False
        LogManualSession aAddrs, aVals, aFonts, aBolds, aFills, hit
        Application.EnableEvents = True
    End If
done:
    Application.EnableEvents = True
    CacheGridState Target            ' 連続編集(同じセルの連打など)に備えて再キャッシュ
End Sub

'--- 変更前の値・書式を退避(入力欄=シフトパレット範囲 内のみ) ---
Private Sub CacheGridState(ByVal rng As Range)
    Dim g As Range, c As Range, k As Long
    mN = 0
    On Error Resume Next
    Set g = Application.Intersect(rng, ShiftGrid())
    On Error GoTo 0
    If g Is Nothing Then Exit Sub
    If g.Cells.Count > 200 Then Exit Sub   ' 大量選択(全選択など)は記録対象外
    ReDim mAddrs(1 To g.Cells.Count): ReDim mVals(1 To g.Cells.Count)
    ReDim mFonts(1 To g.Cells.Count): ReDim mBolds(1 To g.Cells.Count)
    ReDim mFills(1 To g.Cells.Count)
    For Each c In g.Cells
        k = k + 1
        mAddrs(k) = c.Address(False, False)
        mVals(k) = CStr(c.Value)
        mFonts(k) = c.Font.Color
        mBolds(k) = (c.Font.Bold = True)
        If c.Interior.Pattern <> xlNone Then mFills(k) = c.Interior.Color Else mFills(k) = Empty
    Next c
    mN = k
End Sub