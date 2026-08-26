Attribute VB_Name = "DataUpdater"
Option Explicit

' =====================================================================
' 債券RVモデル - データ更新マクロ
' 米財務省 / 日本財務省 / NY連銀 / Yahoo からヒストリカル込みで再取得し、
' 「履歴データ」シートを書き換える。全シートは数式参照なので自動再計算。
' BDP関数は不使用。
'
' 設計要点:
'  - HTTPは MSXML2.ServerXMLHTTP を使用(WinHttpはCOM自動化環境でError 429)。
'  - 日本財務省CSVはShift-JISのため ADODB.Stream でバイト→テキスト変換。
'  - 履歴データは「行削除」せず、固定範囲をクリア→一括書き込み。
'    (行削除は全シートの数式参照を再配線してハングするため)
'  - 分析データエンジンは容量ベース(IFガード)なので行数増減に耐える。
' =====================================================================

Private Const START_ISO As String = "2023-01-01"
Private Const SHEET_HIST As String = "履歴データ"
Private Const CLEAR_LAST As Long = 2000   ' クリアする最大行(容量に余裕を持たせる)

' ---------------------------------------------------------------------
' メインエントリ(ユーザー用: MsgBoxあり)
' ---------------------------------------------------------------------
Public Sub UpdateData()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Dim ust As Object, jgb As Object, fund As Object
    Dim usdjpy As Object, eurusd As Object, zn As Object
    Dim pens As Object, jgbF As Object

    Application.StatusBar = "データ更新: 米国財務省 取得中..."
    Set ust = FetchUST()

    Application.StatusBar = "データ更新: 日本財務省(JGB) 取得中..."
    Set jgb = FetchJGB()

    Application.StatusBar = "データ更新: NY連銀(ファンディング) 取得中..."
    Set fund = FetchFunding()

    Application.StatusBar = "データ更新: Yahoo(FX・先物) 取得中..."
    Set usdjpy = FetchYahoo("USDJPY=X")
    Set eurusd = FetchYahoo("EURUSD=X")
    Set zn = FetchYahoo("ZN=F")
    Application.StatusBar = "データ更新: 2510.T(JGB先物) 取得中..."
    Set jgbF = FetchYahoo("2510.T")

    Application.StatusBar = "データ更新: Pensford(ASWスワップ) 取得中..."
    Set pens = FetchPensford()

    Application.StatusBar = "データ更新: 履歴構築・書き込み中..."
    Dim nRows As Long
    nRows = WriteMaster(ust, jgb, fund, usdjpy, eurusd, zn, jgbF)

    Application.StatusBar = "データ更新: ASWデータ書き込み中..."
    Call WriteAsw(pens)

    Application.StatusBar = "データ更新: 再計算中..."
    Application.CalculateFull

    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "データ更新が完了しました。" & vbCrLf & _
           "履歴データ行数: " & nRows & " 行" & vbCrLf & _
           "全シートを再計算しました。", vbInformation, "データ更新"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "データ更新中にエラーが発生しました。" & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "データ更新"
End Sub

' ---------------------------------------------------------------------
' テスト用: MsgBoxなし版(自動化検証用)。進捗をテキストファイルにログ出力。
' ---------------------------------------------------------------------
Private Sub LogMsg(ByVal s As String)
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open ThisWorkbook.Path & "\vba_progress.log" For Append As #f
    Print #f, Format(Now, "hh:mm:ss") & " " & s
    Close #f
    On Error GoTo 0
End Sub

Public Sub UpdateDataSilent()
    On Error GoTo ErrHandler
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    On Error Resume Next
    Kill ThisWorkbook.Path & "\vba_progress.log"
    On Error GoTo 0

    Dim ust As Object, jgb As Object, fund As Object
    Dim usdjpy As Object, eurusd As Object, zn As Object
    Dim pens As Object, jgbF As Object

    LogMsg "start FetchUST"
    Set ust = FetchUST()
    LogMsg "FetchUST done count=" & ust.Count

    LogMsg "start FetchJGB"
    Set jgb = FetchJGB()
    LogMsg "FetchJGB done count=" & jgb.Count

    LogMsg "start FetchFunding"
    Set fund = FetchFunding()
    LogMsg "FetchFunding done count=" & fund.Count

    LogMsg "start FetchYahoo USDJPY"
    Set usdjpy = FetchYahoo("USDJPY=X")
    LogMsg "USDJPY done count=" & usdjpy.Count

    LogMsg "start FetchYahoo EURUSD"
    Set eurusd = FetchYahoo("EURUSD=X")
    LogMsg "EURUSD done count=" & eurusd.Count

    LogMsg "start FetchYahoo ZN"
    Set zn = FetchYahoo("ZN=F")
    LogMsg "start FetchYahoo 2510.T"
    Set jgbF = FetchYahoo("2510.T")
    LogMsg "2510.T done count=" & jgbF.Count

    LogMsg "start FetchPensford"
    Set pens = FetchPensford()
    LogMsg "FetchPensford done count=" & pens.Count
    LogMsg "ZN done count=" & zn.Count

    LogMsg "start WriteMaster"
    Dim nRows As Long
    nRows = WriteMaster(ust, jgb, fund, usdjpy, eurusd, zn, jgbF)
    LogMsg "WriteMaster done rows=" & nRows

    LogMsg "start WriteAsw"
    Call WriteAsw(pens)
    LogMsg "WriteAsw done"
    LogMsg "WriteMaster done rows=" & nRows

    LogMsg "start CalculateFull"
    Application.CalculateFull
    LogMsg "CalculateFull done"

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_HIST)
    ws.Range("T1").Value = "OK rows=" & nRows & " ust=" & ust.Count & " jgb=" & jgb.Count & _
        " fund=" & fund.Count & " fx=" & usdjpy.Count & " zn=" & zn.Count
    LogMsg "ALL DONE"
    Exit Sub

ErrHandler:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    ThisWorkbook.Sheets(SHEET_HIST).Range("T1").Value = "ERR " & Err.Number & ": " & Err.Description
    LogMsg "ERROR " & Err.Number & ": " & Err.Description
End Sub

' ---------------------------------------------------------------------
' HTTP GET -> バイト配列 (MSXML2.ServerXMLHTTP)
' ---------------------------------------------------------------------
Private Function HttpGetBytes(ByVal url As String) As Variant
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("MSXML2.ServerXMLHTTP")
    If Err.Number <> 0 Then
        Err.Clear
        Set http = CreateObject("MSXML2.XMLHTTP")
    End If
    If Err.Number <> 0 Then HttpGetBytes = Null: Exit Function
    On Error Resume Next
    http.setTimeouts 10000, 15000, 20000, 30000
    On Error GoTo 0
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.setRequestHeader "Accept", "*/*"
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        HttpGetBytes = http.ResponseBody
    Else
        HttpGetBytes = Null
    End If
    On Error GoTo 0
End Function

' バイト配列 -> テキスト (ADODB.Stream で文字コード指定)
Private Function BytesToText(ByVal byt As Variant, ByVal charset As String) As String
    On Error Resume Next
    If IsNull(byt) Then BytesToText = "": Exit Function
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1  ' adTypeBinary
    st.Open
    st.Write byt
    st.Position = 0
    st.Type = 2  ' adTypeText
    st.charset = charset
    BytesToText = st.ReadText(-1)
    st.Close
    On Error GoTo 0
End Function

' ---------------------------------------------------------------------
' 米国財務省 イールドカーブ (年ごとにCSV取得, UTF-8)
' 戻り値: Dictionary(ISO日付 -> Array(2Y,5Y,10Y,20Y,30Y))
' ---------------------------------------------------------------------
Private Function FetchPensford() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim txt As String
    txt = BytesToText(HttpGetBytes("https://pensford.com/api/live-rates"), "utf-8")
    If Len(txt) = 0 Then Set FetchPensford = dict: Exit Function
    ' "name":{"...","quote":数値} のペアを抽出(JSONはフラットな1段)
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = """([A-Za-z0-9]+)""\s*:\s*\{[^}]*?""quote""\s*:\s*([0-9.\-]+)"
    Dim m As Object
    For Each m In re.Execute(txt)
        dict(m.SubMatches(0)) = CDbl(m.SubMatches(1))
    Next m
    ' quoteDate (MM/DD/YYYY) を ISO に変換して格納
    re.Pattern = """quoteDate""\s*:\s*""([^""]+)"""
    If re.Test(txt) Then
        Dim qd As String
        qd = re.Execute(txt)(0).SubMatches(0)
        Dim qp() As String
        qp = Split(qd, "/")
        If UBound(qp) = 2 Then dict("quoteDate") = qp(2) & "-" & qp(0) & "-" & qp(1)
    End If
    Set FetchPensford = dict
End Function

' ---------------------------------------------------------------------
' ASWデータシートへPensfordスナップショット書き込み
' (B列4..19行がbuild_v3.pyのasw_fieldsと対応。quoteDate=4行目から)
' ---------------------------------------------------------------------
Private Sub WriteAsw(ByVal pens As Object)
    On Error GoTo ErrHandler
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("ASWデータ")
    Dim fields As Variant
    fields = Array("quoteDate", "dailySofr", "termSofr1M", "termSofr3M", "fedFunds", _
                   "treasury2Y", "treasury3Y", "treasury5Y", "treasury7Y", "treasury10Y", _
                   "oisSwap2Y", "oisSwap3Y", "oisSwap5Y", "oisSwap7Y", "oisSwap10Y", "bankSwap10Y")
    Dim r As Long
    For r = 0 To UBound(fields)
        If pens.Exists(fields(r)) Then ws.Cells(4 + r, 2).Value = pens(fields(r))
    Next r
    ws.Cells(20, 2).Value = "lastFetched " & Format(Now, "yyyy-mm-dd hh:nn")
    Exit Sub
ErrHandler:
    LogMsg "WriteAsw ERROR " & Err.Number & ": " & Err.Description
End Sub


Private Function FetchUST() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim y As Long
    For y = 2023 To Year(Date)
        Dim url As String
        url = "https://home.treasury.gov/resource-center/data-chart-center/interest-rates/" & _
              "daily-treasury-rates.csv/" & y & "/all?type=daily_treasury_yield_curve" & _
              "&field_tdr_date_value=" & y & "&page&_format=csv"
        Dim txt As String
        txt = BytesToText(HttpGetBytes(url), "utf-8")
        If Len(txt) > 0 Then Call ParseUSTCsv(txt, dict)
    Next y
    Set FetchUST = dict
End Function

Private Sub ParseUSTCsv(ByVal txt As String, ByRef dict As Object)
    Dim lines() As String
    lines = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    If UBound(lines) < 1 Then Exit Sub
    Dim hdr() As String
    hdr = Split(lines(0), ",")
    Dim c2 As Long, c5 As Long, c10 As Long, c20 As Long, c30 As Long
    c2 = -1: c5 = -1: c10 = -1: c20 = -1: c30 = -1
    Dim i As Long
    For i = 0 To UBound(hdr)
        Dim h As String
        h = UCase(Trim(Replace(hdr(i), """", "")))
        Select Case h
            Case "2 YR": c2 = i
            Case "5 YR": c5 = i
            Case "10 YR": c10 = i
            Case "20 YR": c20 = i
            Case "30 YR": c30 = i
        End Select
    Next i
    If c2 < 0 Or c10 < 0 Or c30 < 0 Then Exit Sub
    Dim r As Long
    For r = 1 To UBound(lines)
        If Len(Trim(lines(r))) = 0 Then GoTo NextLine
        Dim f() As String
        f = Split(lines(r), ",")
        If UBound(f) < c30 Then GoTo NextLine
        Dim iso As String
        iso = USTDateToISO(f(0))
        If iso = "" Or iso < START_ISO Then GoTo NextLine
        Dim arr(0 To 4) As Variant
        arr(0) = ToNum(f(c2))
        arr(1) = ToNum(f(c5))
        arr(2) = ToNum(f(c10))
        arr(3) = ToNum(f(c20))
        arr(4) = ToNum(f(c30))
        dict(iso) = arr
NextLine:
    Next r
End Sub

' MM/DD/YYYY -> YYYY-MM-DD
Private Function USTDateToISO(ByVal s As String) As String
    s = Trim(Replace(s, """", ""))
    Dim p() As String
    p = Split(s, "/")
    If UBound(p) <> 2 Then USTDateToISO = "": Exit Function
    USTDateToISO = p(2) & "-" & Format(CLng(p(0)), "00") & "-" & Format(CLng(p(1)), "00")
End Function

' ---------------------------------------------------------------------
' 日本財務省 JGB (全履歴CSV, Shift-JIS, 和暦日付)
' 戻り値: Dictionary(ISO日付 -> Array(2Y,10Y,20Y,30Y))
' ---------------------------------------------------------------------
Private Function FetchJGB() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim byt As Variant
    byt = HttpGetBytes("https://www.mof.go.jp/jgbs/reference/interest_rate/data/jgbcm_all.csv")
    If IsNull(byt) Then
        LogMsg "JGB: http FAILED (null bytes)"
        Set FetchJGB = dict: Exit Function
    End If
    LogMsg "JGB: bytes=" & (UBound(byt) + 1)
    Dim txt As String
    txt = BytesToText(byt, "shift_jis")
    LogMsg "JGB: text len=" & Len(txt) & " head=[" & Left(txt, 30) & "]"
    If Len(txt) = 0 Then Set FetchJGB = dict: Exit Function
    Dim lines() As String
    lines = Split(Replace(txt, vbCrLf, vbLf), vbLf)
    Dim c2 As Long, c10 As Long, c20 As Long, c30 As Long
    c2 = -1: c10 = -1: c20 = -1: c30 = -1
    Dim i As Long
    For i = 0 To UBound(lines)
        If InStr(lines(i), "基準日") > 0 Then
            Dim hdr() As String
            hdr = Split(lines(i), ",")
            Dim j As Long
            For j = 0 To UBound(hdr)
                Dim h As String
                h = Trim(hdr(j))
                Select Case h
                    Case "2年": c2 = j
                    Case "10年": c10 = j
                    Case "20年": c20 = j
                    Case "30年": c30 = j
                End Select
            Next j
            Dim r As Long
            For r = i + 1 To UBound(lines)
                If Len(Trim(lines(r))) = 0 Then GoTo NextJ
                Dim f() As String
                f = Split(lines(r), ",")
                If UBound(f) < c30 Then GoTo NextJ
                Dim iso As String
                iso = EraToISO(Trim(f(0)))
                If iso = "" Or iso < START_ISO Then GoTo NextJ
                Dim arr(0 To 3) As Variant
                arr(0) = ToNum(f(c2))
                arr(1) = ToNum(f(c10))
                arr(2) = ToNum(f(c20))
                arr(3) = ToNum(f(c30))
                dict(iso) = arr
NextJ:
            Next r
            Exit For
        End If
    Next i
    Set FetchJGB = dict
End Function

' 和暦 "R8.8.3" / "H31.4.30" -> "2026-08-03"
Private Function EraToISO(ByVal s As String) As String
    EraToISO = ""
    If Len(s) < 3 Then Exit Function
    Dim era As String
    era = Left(s, 1)
    Dim base As Long
    Select Case era
        Case "R": base = 2018
        Case "H": base = 1988
        Case "S": base = 1925
        Case "T": base = 1911
        Case Else: Exit Function
    End Select
    Dim rest As String
    rest = Mid(s, 2)
    Dim p() As String
    p = Split(rest, ".")
    If UBound(p) <> 2 Then Exit Function
    Dim yy As Long, mm As Long, dd As Long
    On Error Resume Next
    yy = CLng(p(0)): mm = CLng(p(1)): dd = CLng(p(2))
    On Error GoTo 0
    If yy = 0 Or mm = 0 Or dd = 0 Then Exit Function
    Dim yr As Long
    yr = base + yy
    On Error Resume Next
    Dim chk As Date
    chk = DateSerial(yr, mm, dd)
    If Err.Number <> 0 Then EraToISO = "": Exit Function
    On Error GoTo 0
    EraToISO = Format(chk, "yyyy-mm-dd")
End Function

' ---------------------------------------------------------------------
' NY連銀 ファンディング (SOFR/EFFR/TGCR/BGCR, UTF-8 JSON)
' 戻り値: Dictionary(ISO日付 -> Array(SOFR,EFFR,TGCR,BGCR))
' ---------------------------------------------------------------------
Private Function FetchFunding() As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim prods As Variant
    prods = Array(Array("sofr", "secured", 0), Array("effr", "unsecured", 1), _
                  Array("tgcr", "secured", 2), Array("bgcr", "secured", 3))
    Dim endDate As String
    endDate = Format(Date, "yyyy-mm-dd")
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.Pattern = """effectiveDate""\s*:\s*""([^""]+)""[^}]*""percentRate""\s*:\s*([0-9.\-]+|null)"
    Dim k As Long
    For k = 0 To UBound(prods)
        Dim prod As String, kind As String, idx As Long
        prod = prods(k)(0): kind = prods(k)(1): idx = prods(k)(2)
        Dim url As String
        url = "https://markets.newyorkfed.org/api/rates/" & kind & "/" & prod & _
              "/search.json?startDate=" & START_ISO & "&endDate=" & endDate & "&limit=1200"
        Dim txt As String
        txt = BytesToText(HttpGetBytes(url), "utf-8")
        If Len(txt) = 0 Then GoTo NextProd
        Dim matches As Object
        Set matches = re.Execute(txt)
        Dim m As Object
        For Each m In matches
            Dim dIso As String, rate As Variant
            dIso = m.SubMatches(0)
            If m.SubMatches(1) = "null" Then
                rate = Empty
            Else
                rate = CDbl(m.SubMatches(1))
            End If
            If dIso >= START_ISO Then
                Dim arr As Variant
                If dict.Exists(dIso) Then
                    arr = dict(dIso)
                Else
                    Dim blank(0 To 3) As Variant
                    arr = blank
                End If
                arr(idx) = rate
                dict(dIso) = arr
            End If
        Next m
NextProd:
    Next k
    Set FetchFunding = dict
End Function

' ---------------------------------------------------------------------
' Yahoo (FX・先物, UTF-8 JSON)
' 戻り値: Dictionary(ISO日付 -> close)
' ---------------------------------------------------------------------
Private Function FetchYahoo(ByVal sym As String) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    Dim t1 As Double, t2 As Double
    t1 = UnixTs(DateSerial(2023, 1, 1))
    t2 = UnixTs(Now)
    Dim url As String
    url = "https://query1.finance.yahoo.com/v8/finance/chart/" & sym & _
          "?period1=" & CStr(t1) & "&period2=" & CStr(t2) & "&interval=1d&events=history"
    Dim txt As String
    txt = BytesToText(HttpGetBytes(url), "utf-8")
    If Len(txt) = 0 Then Set FetchYahoo = dict: Exit Function
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.Pattern = """timestamp""\s*:\s*\[([^\]]*)\]"
    Dim tsStr As String
    If re.Test(txt) Then
        tsStr = re.Execute(txt)(0).SubMatches(0)
    Else
        Set FetchYahoo = dict: Exit Function
    End If
    re.Pattern = """close""\s*:\s*\[([^\]]*)\]"
    Dim closeStr As String
    If re.Test(txt) Then
        closeStr = re.Execute(txt)(0).SubMatches(0)
    Else
        Set FetchYahoo = dict: Exit Function
    End If
    Dim tsArr() As String, clArr() As String
    tsArr = Split(tsStr, ",")
    clArr = Split(closeStr, ",")
    Dim n As Long
    n = UBound(tsArr)
    If UBound(clArr) < n Then n = UBound(clArr)
    Dim i As Long
    For i = 0 To n
        Dim ts As Double, cv As String
        ts = Val(Trim(tsArr(i)))
        cv = Trim(clArr(i))
        If ts > 0 And cv <> "null" And Len(cv) > 0 Then
            Dim iso As String
            iso = UnixToISO(ts)
            If iso >= START_ISO Then dict(iso) = CDbl(cv)
        End If
    Next i
    Set FetchYahoo = dict
End Function

' Date -> Unix timestamp
Private Function UnixTs(ByVal d As Date) As Double
    UnixTs = (CDbl(d) - 25569#) * 86400#
End Function

' Unix timestamp -> ISO日付
Private Function UnixToISO(ByVal ts As Double) As String
    Dim serial As Double
    serial = ts / 86400# + 25569#
    UnixToISO = Format(CDate(serial), "yyyy-mm-dd")
End Function

' ISO "YYYY-MM-DD" -> Date (ロケール非依存)
Private Function ISOToDate(ByVal iso As String) As Date
    Dim p() As String
    p = Split(iso, "-")
    ISOToDate = DateSerial(CInt(p(0)), CInt(p(1)), CInt(p(2)))
End Function

' 数値変換(空/"-"はEmpty)
Private Function ToNum(ByVal s As String) As Variant
    s = Trim(Replace(s, """", ""))
    If Len(s) = 0 Or s = "-" Or s = "ND" Then
        ToNum = Empty
    Else
        On Error Resume Next
        ToNum = CDbl(s)
        If Err.Number <> 0 Then ToNum = Empty
        On Error GoTo 0
    End If
End Function

' ---------------------------------------------------------------------
' マスター構築 + 履歴データ書き込み
' UST営業日を共通インデックスにし、JGB・FXは前方補完。
' 【重要】行削除はせず、A2:S(CLEAR_LAST)をクリア→配列一括書き込み。
' ---------------------------------------------------------------------
Private Function WriteMaster(ByVal ust As Object, ByVal jgb As Object, _
        ByVal fund As Object, ByVal usdjpy As Object, ByVal eurusd As Object, _
        ByVal zn As Object, ByVal jgbF As Object) As Long
    LogMsg "WM: enter"
    ' UST日付 ∩ fund日付 を収集
    Dim keys As Variant
    keys = ust.keys
    LogMsg "WM: got keys ubound=" & UBound(keys)
    Dim dates() As String
    Dim cnt As Long
    cnt = 0
    ReDim dates(0 To UBound(keys))
    Dim i As Long
    For i = 0 To UBound(keys)
        If fund.Exists(keys(i)) Then
            dates(cnt) = keys(i)
            cnt = cnt + 1
        End If
    Next i
    LogMsg "WM: collected cnt=" & cnt
    If cnt = 0 Then WriteMaster = 0: Exit Function
    ReDim Preserve dates(0 To cnt - 1)
    LogMsg "WM: sorting..."
    Call SortStrArr(dates)
    LogMsg "WM: sorted"

    Dim ws As Worksheet
    LogMsg "WM: getting sheet"
    Set ws = ThisWorkbook.Sheets(SHEET_HIST)
    LogMsg "WM: got sheet"

    ' クリア前に既存のR/S列(BUND_FUT/JGB_FUT)を退避(前方補完用)
    Dim oldR() As Variant, oldS() As Variant
    ReDim oldR(1 To CLEAR_LAST)
    ReDim oldS(1 To CLEAR_LAST)
    Dim oc As Long
    For oc = 2 To CLEAR_LAST
        oldR(oc) = ws.Cells(oc, 18).Value
        oldS(oc) = ws.Cells(oc, 19).Value
    Next oc

    ' 出力バッファ(19列)を配列に溜めて一括書き込み
    Dim nDays As Long
    nDays = UBound(dates) + 1
    Dim out() As Variant
    ReDim out(1 To nDays, 1 To 19)
    LogMsg "WM: redim out " & nDays & "x19"

    Dim prevJ(0 To 3) As Variant
    Dim prevFx(0 To 2) As Variant
    Dim prevBund As Variant, prevJgb As Variant
    Dim u As Variant, j As Variant, f As Variant
    Dim d As Long
    For d = 0 To UBound(dates)
        If d Mod 200 = 0 Then LogMsg "WM: loop d=" & d
        Dim iso As String
        iso = dates(d)
        j = Empty
        f = Empty
        u = ust(iso)
        If jgb.Exists(iso) Then j = jgb(iso)
        If fund.Exists(iso) Then f = fund(iso)
        ' JGB前方補完
        Dim t As Long
        If Not IsEmpty(j) Then
            For t = 0 To 3
                If Not IsEmpty(j(t)) Then prevJ(t) = j(t)
            Next t
        End If
        ' FX前方補完
        If usdjpy.Exists(iso) Then prevFx(0) = usdjpy(iso)
        If eurusd.Exists(iso) Then prevFx(1) = eurusd(iso)
        If zn.Exists(iso) Then prevFx(2) = zn(iso)

        Dim rr As Long
        rr = d + 1
        out(rr, 1) = ISOToDate(iso)
        ' UST 2Y,5Y,10Y,20Y,30Y -> 列2..6
        For t = 0 To 4
            If Not IsEmpty(u(t)) Then out(rr, 2 + t) = u(t)
        Next t
        ' JGB 2Y,10Y,20Y,30Y -> 列7..10
        For t = 0 To 3
            Dim jv As Variant
            jv = Empty
            If Not IsEmpty(j) Then
                If Not IsEmpty(j(t)) Then jv = j(t)
            End If
            If IsEmpty(jv) Then jv = prevJ(t)
            If Not IsEmpty(jv) Then out(rr, 7 + t) = jv
        Next t
        ' SOFR,EFFR,TGCR,BGCR -> 列11..14
        If Not IsEmpty(f) Then
            For t = 0 To 3
                If Not IsEmpty(f(t)) Then out(rr, 11 + t) = f(t)
            Next t
        End If
        ' USDJPY,EURUSD,ZN -> 列15..17
        For t = 0 To 2
            If Not IsEmpty(prevFx(t)) Then out(rr, 15 + t) = prevFx(t)
        Next t
        ' BUND_FUT(18列): 前回値の前方補完(Eurexデータはオフライン構築のため)
        If Not IsEmpty(oldR(rr)) Then prevBund = oldR(rr)
        If Not IsEmpty(prevBund) Then out(rr, 18) = prevBund
        ' JGB_FUT(19列): 2510.T終値をインプライド利回りに換算
        ' (parse_data.pyと同じ2点キャリブレーション: y = 16.874 - 0.01754*P)
        If Not jgbF Is Nothing Then
            If jgbF.Exists(iso) Then prevJgb = 16.874 - 0.01754 * jgbF(iso)
        End If
        If IsEmpty(prevJgb) And Not IsEmpty(oldS(rr)) Then prevJgb = oldS(rr)
        If Not IsEmpty(prevJgb) Then out(rr, 19) = prevJgb
    Next d

    ' 【重要】行削除せず、固定範囲をクリア→一括書き込み
    LogMsg "WM: clearing A2:S" & CLEAR_LAST
    Application.EnableEvents = False
    ws.Range("A2:S" & CLEAR_LAST).ClearContents
    LogMsg "WM: cleared, writing " & nDays & " rows"
    ws.Range("A2").Resize(nDays, 19).Value = out
    LogMsg "WM: wrote array, setting date format"
    ws.Range("A2").Resize(nDays, 1).NumberFormat = "yyyy-mm-dd"
    Application.EnableEvents = True
    LogMsg "WM: done writing"

    ' 注記を再追加(クリア範囲の外)
    Dim noteRow As Long
    noteRow = CLEAR_LAST + 2
    ws.Cells(noteRow, 1).Value = "★このシートが全モデルの唯一のデータ源。ダッシュボードの【データ更新】ボタン(VBA)が米財務省/日本財務省/NY連銀/Yahoo/Pensfordからヒストリカル込みで再取得し、ここを書き換える。書き換え後は全シートが自動再計算。"
    ws.Cells(noteRow + 1, 1).Value = "出典: 米国財務省(daily treasury yield curve)/日本財務省(国債金利情報,日本の休日は前方補完)/NY連銀(SOFR・EFFR・TGCR・BGCR)/Yahoo(USDJPY・EURUSD・ZN先物・2510.T)/Pensford(ASWスワップ,ASWデータシート)。BUND_FUTはEurexからオフライン構築した値を前方補完、JGB_FUTは2510.T終値からインプライド利回りに換算。"

    WriteMaster = nDays
End Function

' 文字列配列の昇順ソート(ISO日付は辞書順=時系列)
Private Sub SortStrArr(ByRef arr() As String)
    Dim i As Long, j As Long
    Dim tmp As String
    For i = LBound(arr) + 1 To UBound(arr)
        tmp = arr(i)
        j = i - 1
        Do While j >= LBound(arr)
            If arr(j) > tmp Then
                arr(j + 1) = arr(j)
                j = j - 1
            Else
                Exit Do
            End If
        Loop
        arr(j + 1) = tmp
    Next i
End Sub
