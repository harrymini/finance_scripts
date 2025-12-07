' =====================================================
' 주식 현재가 추적 VBA 스크립트 (Windows용 - 네이버 금융)
' 버전: 2.1
' 최종 수정: 2025-12-07
'
' 사용법:
' 1. 엑셀에서 Alt + F11로 VBA 편집기 열기
' 2. 삽입 > 모듈
' 3. 이 코드 붙여넣기
' 4. UpdateStockPrices 실행 (현재가 조회)
'    또는 UpdateStockPricesByDate 실행 (과거 날짜별 조회)
'
' 데이터 시트 구조:
' 1행: 헤더 (종목명, 종목코드, 조회날짜)
' 2행부터: 데이터 (A열: 종목명, B열: 종목코드, C열: 조회날짜)
' C열의 고유 날짜별로 탭이 생성됨 (UpdateStockPricesByDate용)
'
' 변경 이력:
' v2.1 - C열 날짜를 세로로 읽어 고유 날짜별 탭 생성, 전일대비/등락률 계산 수정
' v2.0 - 과거 날짜별 시세 조회 기능 추가
' v1.0 - 현재가 조회 기능
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
                ' 이미 -가 포함되어 있으면 그대로 사용
                If Left(diffPrice, 1) = "-" Then
                    change = diffPrice
                ElseIf isUp Then
                    change = "+" & diffPrice
                Else
                    change = "-" & diffPrice
                End If
            End If

            If Len(pctVal) > 0 Then
                ' 이미 -가 포함되어 있으면 그대로 사용
                If Left(pctVal, 1) = "-" Then
                    changePercent = pctVal & "%"
                ElseIf isUp Then
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
' 과거 날짜별 시세 조회 (C열 날짜 → 고유 날짜별 탭 생성)
' 데이터 시트: A열=종목명, B열=종목코드, C열=조회날짜
' C열의 고유 날짜마다 탭을 생성하고 해당 날짜의 모든 종목 시세 조회
' =====================================================
Sub UpdateStockPricesByDate()
    Dim wsData As Worksheet
    Dim wsResult As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim j As Long
    Dim rowNum As Long
    Dim stockName As String
    Dim stockCode As String
    Dim targetDate As String
    Dim targetDateFormatted As String
    Dim currentPrice As String
    Dim priceChange As String
    Dim changePercent As String
    Dim totalProcessed As Long
    Dim uniqueDates As Object
    Dim dateKey As Variant
    Dim rawDate As Variant

    On Error GoTo ErrorHandler

    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("데이터")
    On Error GoTo ErrorHandler

    If wsData Is Nothing Then
        MsgBox "'데이터' 시트를 찾을 수 없습니다.", vbExclamation, "오류"
        Exit Sub
    End If

    ' 종목 수 확인 (A열)
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "데이터 시트에 종목 데이터가 없습니다.", vbExclamation, "오류"
        Exit Sub
    End If

    ' C열에서 고유 날짜 수집 (Dictionary 사용)
    Set uniqueDates = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRow
        rawDate = wsData.Cells(i, 3).Value
        If rawDate <> "" Then
            targetDate = ConvertToYYYYMMDD(rawDate)
            If Len(targetDate) = 8 And Not uniqueDates.exists(targetDate) Then
                uniqueDates.Add targetDate, targetDate
            End If
        End If
    Next i

    If uniqueDates.count = 0 Then
        MsgBox "C열에 조회할 날짜가 없습니다.", vbExclamation, "오류"
        Exit Sub
    End If

    If MsgBox("총 " & uniqueDates.count & "개 날짜를 조회합니다." & vbCrLf & _
              "각 날짜별로 탭이 생성됩니다." & vbCrLf & vbCrLf & _
              "계속하시겠습니까?", vbYesNo + vbQuestion, "확인") = vbNo Then
        Exit Sub
    End If

    Application.ScreenUpdating = False
    totalProcessed = 0

    ' 각 고유 날짜별로 탭 생성 및 데이터 조회
    For Each dateKey In uniqueDates.Keys
        targetDate = CStr(dateKey)

        ' 탭 이름은 yyyy-mm-dd 형식
        targetDateFormatted = Left(targetDate, 4) & "-" & Mid(targetDate, 5, 2) & "-" & Right(targetDate, 2)

        Application.StatusBar = "탭 생성 중: " & targetDateFormatted
        DoEvents

        ' 해당 날짜 탭 생성/가져오기
        Set wsResult = GetOrCreateSheet(targetDateFormatted)
        SetupHeaderForHistorical wsResult

        rowNum = 2

        ' 모든 종목에 대해 해당 날짜 시세 조회
        For j = 2 To lastRow
            stockName = Trim(CStr(wsData.Cells(j, 1).Value))
            stockCode = Trim(CStr(wsData.Cells(j, 2).Value))

            If stockName = "" And stockCode = "" Then GoTo NextStock
            If stockCode = "" Then GoTo NextStock

            Application.StatusBar = targetDateFormatted & " - " & stockName
            DoEvents

            stockCode = CleanStockCode(stockCode)

            ' 과거 시세 가져오기
            Call GetNaverHistoricalPrice(stockCode, targetDate, currentPrice, priceChange, changePercent)

            wsResult.Cells(rowNum, 1).Value = stockName
            wsResult.Cells(rowNum, 2).Value = "'" & stockCode
            wsResult.Cells(rowNum, 3).NumberFormat = "@"
            wsResult.Cells(rowNum, 3).Value = currentPrice
            wsResult.Cells(rowNum, 4).NumberFormat = "@"
            wsResult.Cells(rowNum, 4).Value = priceChange
            wsResult.Cells(rowNum, 5).NumberFormat = "@"
            wsResult.Cells(rowNum, 5).Value = changePercent
            wsResult.Cells(rowNum, 6).Value = Format(Now, "yyyy-mm-dd hh:mm:ss")

            ApplyPriceColorByDate wsResult, rowNum, priceChange

            rowNum = rowNum + 1
            totalProcessed = totalProcessed + 1

            Delay 0.3

NextStock:
        Next j

        wsResult.Columns("A:F").AutoFit
    Next dateKey

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "완료!" & vbCrLf & _
           "생성된 탭: " & uniqueDates.count & "개" & vbCrLf & _
           "처리된 데이터: " & totalProcessed & "건", vbInformation, "완료"

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "오류: " & Err.Description, vbCritical, "오류"
End Sub

' 날짜를 yyyymmdd 형식으로 변환
Private Function ConvertToYYYYMMDD(rawDate As Variant) As String
    Dim result As String
    Dim dateParts() As String
    Dim y As String, m As String, d As String

    On Error GoTo ConvertErr

    If IsDate(rawDate) Then
        result = Format(CDate(rawDate), "yyyymmdd")
    Else
        result = Trim(CStr(rawDate))
        result = Replace(result, " ", "")

        ' yyyy-mm-dd 또는 yyyy/mm/dd 형식
        If InStr(result, "-") > 0 Then
            dateParts = Split(result, "-")
        ElseIf InStr(result, "/") > 0 Then
            dateParts = Split(result, "/")
        Else
            ' 이미 yyyymmdd 형식
            ConvertToYYYYMMDD = result
            Exit Function
        End If

        If UBound(dateParts) >= 2 Then
            y = dateParts(0)
            m = dateParts(1)
            d = dateParts(2)

            ' 월, 일이 한 자리면 0 추가
            If Len(m) = 1 Then m = "0" & m
            If Len(d) = 1 Then d = "0" & d

            result = y & m & d
        End If
    End If

    ConvertToYYYYMMDD = result
    Exit Function

ConvertErr:
    ConvertToYYYYMMDD = ""
End Function

' 과거 시세 가져오기 (네이버 일별 시세 API) - 한 번의 API 호출로 당일+전일 시세 모두 조회
Private Sub GetNaverHistoricalPrice(stockCode As String, targetDate As String, ByRef price As String, ByRef change As String, ByRef changePercent As String)
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim lines() As String
    Dim values() As String
    Dim i As Long
    Dim curPrice As Double
    Dim prevPrice As Double
    Dim diff As Double
    Dim pct As Double
    Dim cleanLine As String
    Dim prevDate As String
    Dim dataRows() As String
    Dim dataCount As Long
    Dim targetDateStr As String
    Dim foundIdx As Long

    On Error GoTo HistError

    price = "-"
    change = "-"
    changePercent = "-"

    ' 10일 전부터 조회 (주말/공휴일 고려)
    prevDate = Format(DateAdd("d", -10, CDate(Left(targetDate, 4) & "-" & Mid(targetDate, 5, 2) & "-" & Right(targetDate, 2))), "yyyymmdd")

    url = "https://fchart.stock.naver.com/siseJson.nhn?symbol=" & stockCode & "&requestType=1&startTime=" & prevDate & "&endTime=" & targetDate & "&timeframe=day"

    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.send

    If http.Status = 200 Then
        response = http.responseText

        ' 응답 정리
        response = Replace(response, "[", "")
        response = Replace(response, "]", "")
        response = Replace(response, "'", "")
        response = Replace(response, """", "")
        response = Replace(response, vbLf, "|")
        response = Replace(response, vbCr, "")

        lines = Split(response, "|")

        ' 유효한 데이터 행만 추출
        ReDim dataRows(1 To UBound(lines))
        dataCount = 0

        For i = 1 To UBound(lines)
            cleanLine = Trim(lines(i))
            If Len(cleanLine) > 10 Then
                values = Split(cleanLine, ",")
                If UBound(values) >= 4 Then
                    dataCount = dataCount + 1
                    dataRows(dataCount) = cleanLine
                End If
            End If
        Next i

        ' 마지막 데이터가 조회 대상 날짜
        If dataCount >= 1 Then
            values = Split(dataRows(dataCount), ",")
            curPrice = Val(Trim(values(4)))

            If curPrice > 0 Then
                price = Format(curPrice, "#,##0")

                ' 전일 데이터 (마지막에서 두 번째)
                If dataCount >= 2 Then
                    values = Split(dataRows(dataCount - 1), ",")
                    prevPrice = Val(Trim(values(4)))

                    If prevPrice > 0 Then
                        diff = curPrice - prevPrice
                        pct = (diff / prevPrice) * 100

                        If diff > 0 Then
                            change = "+" & Format(diff, "#,##0")
                            changePercent = "+" & Format(pct, "0.00") & "%"
                        ElseIf diff < 0 Then
                            change = Format(diff, "#,##0")
                            changePercent = Format(pct, "0.00") & "%"
                        Else
                            change = "0"
                            changePercent = "0.00%"
                        End If
                    End If
                End If
            End If
        End If
    End If

    Set http = Nothing
    Exit Sub

HistError:
    price = "오류"
    change = "-"
    changePercent = "-"
    If Not http Is Nothing Then Set http = Nothing
End Sub

' 과거 시세 조회용 헤더 설정 (탭 이름이 날짜이므로 조회날짜 컬럼 제외)
Private Sub SetupHeaderForHistorical(ws As Worksheet)
    ws.Cells(1, 1).Value = "종목명"
    ws.Cells(1, 2).Value = "종목코드"
    ws.Cells(1, 3).Value = "종가"
    ws.Cells(1, 4).Value = "전일대비"
    ws.Cells(1, 5).Value = "등락률"
    ws.Cells(1, 6).Value = "업데이트시간"

    With ws.Range("A1:F1")
        .Font.Bold = True
        .Interior.Color = RGB(100, 149, 237)  ' Cornflower Blue
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter
    End With
End Sub

' 날짜 조회용 색상 적용 (전일대비=4열, 등락률=5열)
Private Sub ApplyPriceColorByDate(ws As Worksheet, rowNum As Long, priceChange As String)
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
