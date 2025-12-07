' =====================================================
' 주식 현재가 추적 VBA 스크립트 (Windows용 - 네이버 금융)
'
' 사용법:
' 1. 엑셀에서 Alt + F11로 VBA 편집기 열기
' 2. 삽입 > 모듈
' 3. 이 코드 붙여넣기
' 4. UpdateStockPrices 실행
'
' 데이터 시트 구조:
' 1행: 헤더 (종목명, 종목코드)
' 2행부터: 데이터 (A열: 종목명, B열: 종목코드)
' =====================================================

Option Explicit

' =====================================================
' 메인 함수
' =====================================================
Sub UpdateStockPrices()
    Dim wsData As Worksheet
    Dim wsToday As Worksheet
    Dim todayName As String
    Dim lastRow As Long
    Dim i As Long
    Dim rowNum As Long
    Dim stockName As String
    Dim stockCode As String
    Dim currentPrice As String
    Dim priceChange As String
    Dim changePercent As String
    Dim processedCount As Long

    On Error GoTo ErrorHandler

    todayName = Format(Date, "yyyy-mm-dd")

    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("데이터")
    On Error GoTo ErrorHandler

    If wsData Is Nothing Then
        MsgBox "'데이터' 시트를 찾을 수 없습니다.", vbExclamation, "오류"
        Exit Sub
    End If

    Set wsToday = GetOrCreateSheet(todayName)
    SetupHeader wsToday

    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "데이터 시트에 종목 데이터가 없습니다.", vbExclamation, "오류"
        Exit Sub
    End If

    Application.ScreenUpdating = False
    Application.StatusBar = "주식 현재가 업데이트 중..."

    processedCount = 0
    rowNum = 2

    For i = 2 To lastRow
        stockName = Trim(CStr(wsData.Cells(i, 1).Value))
        stockCode = Trim(CStr(wsData.Cells(i, 2).Value))

        If stockName = "" And stockCode = "" Then
            GoTo NextRow
        End If

        If stockCode <> "" Then
            Application.StatusBar = "처리 중: " & stockName & " (" & processedCount + 1 & "/" & (lastRow - 1) & ")"
            DoEvents

            stockCode = CleanStockCode(stockCode)

            ' 네이버 금융에서 현재가 가져오기
            Call GetNaverStockPrice(stockCode, currentPrice, priceChange, changePercent)

            wsToday.Cells(rowNum, 1).Value = stockName
            wsToday.Cells(rowNum, 2).Value = "'" & stockCode
            wsToday.Cells(rowNum, 3).NumberFormat = "@"
            wsToday.Cells(rowNum, 3).Value = currentPrice
            wsToday.Cells(rowNum, 4).NumberFormat = "@"
            wsToday.Cells(rowNum, 4).Value = priceChange
            wsToday.Cells(rowNum, 5).NumberFormat = "@"
            wsToday.Cells(rowNum, 5).Value = changePercent
            wsToday.Cells(rowNum, 6).Value = Format(Now, "hh:mm:ss")

            ApplyPriceColor wsToday, rowNum, priceChange

            rowNum = rowNum + 1
            processedCount = processedCount + 1

            Delay 0.5
        End If

NextRow:
    Next i

    wsToday.Columns("A:F").AutoFit

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "완료! " & processedCount & "개 종목 처리됨", vbInformation, "완료"

    wsToday.Activate
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "오류: " & Err.Description, vbCritical, "오류"
End Sub

' =====================================================
' 네이버 금융 API (JSON)
' =====================================================
Private Sub GetNaverStockPrice(stockCode As String, ByRef price As String, ByRef change As String, ByRef changePercent As String)
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim curPrice As String
    Dim diffPrice As String
    Dim pctVal As String
    Dim isUp As Boolean

    On Error GoTo NaverError

    price = "-"
    change = "-"
    changePercent = "-"

    ' 네이버 모바일 주식 API (JSON)
    url = "https://m.stock.naver.com/api/stock/" & stockCode & "/basic"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.send

    If http.Status = 200 Then
        response = http.responseText

        ' JSON에서 문자열 값 추출 (콤마 포함)
        curPrice = ExtractJsonString(response, "closePrice")
        diffPrice = ExtractJsonString(response, "compareToPreviousClosePrice")
        pctVal = ExtractJsonString(response, "fluctuationsRatio")

        ' 상승/하락 확인
        isUp = InStr(1, response, """text"":""상승""", vbTextCompare) > 0

        If Len(curPrice) > 0 Then
            price = curPrice  ' 이미 콤마 포함된 형식

            If Len(diffPrice) > 0 Then
                If isUp Then
                    change = "+" & diffPrice
                Else
                    change = "-" & diffPrice
                End If
            End If

            If Len(pctVal) > 0 Then
                If isUp Then
                    changePercent = "+" & pctVal & "%"
                Else
                    changePercent = "-" & pctVal & "%"
                End If
            End If
        End If
    End If

    Set http = Nothing
    Exit Sub

NaverError:
    price = "오류"
    change = "-"
    changePercent = "-"
    If Not http Is Nothing Then Set http = Nothing
End Sub

' JSON에서 문자열 값 추출 (따옴표 안의 값)
Private Function ExtractJsonString(json As String, key As String) As String
    Dim searchKey As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String

    On Error GoTo ExtractErr

    searchKey = """" & key & """:"""

    startPos = InStr(1, json, searchKey, vbTextCompare)
    If startPos = 0 Then
        ExtractJsonString = ""
        Exit Function
    End If

    startPos = startPos + Len(searchKey)

    ' 닫는 따옴표 찾기
    endPos = InStr(startPos, json, """", vbTextCompare)
    If endPos = 0 Then
        ExtractJsonString = ""
        Exit Function
    End If

    value = Mid(json, startPos, endPos - startPos)
    ExtractJsonString = value

    Exit Function

ExtractErr:
    ExtractJsonString = ""
End Function

' =====================================================
' 시트 관리
' =====================================================
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    Else
        lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
        If lastRow > 1 Then
            ws.Range("A2:F" & lastRow).ClearContents
            ws.Range("A2:F" & lastRow).Font.Color = RGB(0, 0, 0)
        End If
    End If

    Set GetOrCreateSheet = ws
End Function

Private Sub SetupHeader(ws As Worksheet)
    ws.Cells(1, 1).Value = "종목명"
    ws.Cells(1, 2).Value = "종목코드"
    ws.Cells(1, 3).Value = "현재가"
    ws.Cells(1, 4).Value = "전일대비"
    ws.Cells(1, 5).Value = "등락률"
    ws.Cells(1, 6).Value = "업데이트시간"

    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
End Sub

' =====================================================
' 유틸리티
' =====================================================
Private Function CleanStockCode(ByVal code As String) As String
    Dim result As String
    Dim i As Long
    Dim c As String

    code = CStr(code)
    result = ""

    For i = 1 To Len(code)
        c = Mid(code, i, 1)
        If c >= "0" And c <= "9" Then
            result = result & c
        End If
    Next i

    Do While Len(result) < 6
        result = "0" & result
    Loop

    CleanStockCode = result
End Function

Private Sub Delay(seconds As Double)
    Dim endTime As Double
    endTime = Timer + seconds
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

Private Sub ApplyPriceColor(ws As Worksheet, rowNum As Long, priceChange As String)
    Dim changeVal As Double

    On Error Resume Next
    changeVal = Val(Replace(Replace(priceChange, "+", ""), ",", ""))
    On Error GoTo 0

    If changeVal > 0 Then
        ws.Cells(rowNum, 4).Font.Color = RGB(255, 0, 0)
        ws.Cells(rowNum, 5).Font.Color = RGB(255, 0, 0)
    ElseIf changeVal < 0 Then
        ws.Cells(rowNum, 4).Font.Color = RGB(0, 0, 255)
        ws.Cells(rowNum, 5).Font.Color = RGB(0, 0, 255)
    End If
End Sub

' =====================================================
' 테스트 함수
' =====================================================
Sub TestSingleStock()
    Dim code As String
    Dim price As String
    Dim change As String
    Dim pct As String

    code = InputBox("종목코드 (예: 005930):", "테스트", "005930")
    If code = "" Then Exit Sub

    code = CleanStockCode(code)
    Call GetNaverStockPrice(code, price, change, pct)

    MsgBox "종목코드: " & code & vbCrLf & _
           "현재가: " & price & vbCrLf & _
           "전일대비: " & change & vbCrLf & _
           "등락률: " & pct, vbInformation, "결과"
End Sub

' 인터넷 연결 테스트
Sub TestConnection()
    Dim http As Object

    On Error GoTo ConnError

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", "https://www.naver.com", False
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.send

    MsgBox "연결 성공! HTTP 상태: " & http.Status, vbInformation
    Set http = Nothing
    Exit Sub

ConnError:
    MsgBox "연결 실패: " & Err.Description, vbCritical
End Sub

' =====================================================
' 유틸리티 매크로
' =====================================================
Sub DeleteDateSheet()
    Dim sheetName As String
    sheetName = InputBox("삭제할 시트:", "삭제", Format(Date, "yyyy-mm-dd"))
    If sheetName = "" Then Exit Sub

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Sub ClearAllDateSheets()
    Dim ws As Worksheet
    Dim count As Long

    If MsgBox("모든 날짜 시트 삭제?", vbYesNo) = vbNo Then Exit Sub

    Application.DisplayAlerts = False
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "데이터" And Len(ws.Name) = 10 Then
            If Mid(ws.Name, 5, 1) = "-" Then ws.Delete: count = count + 1
        End If
    Next ws
    Application.DisplayAlerts = True
    MsgBox count & "개 삭제됨"
End Sub
