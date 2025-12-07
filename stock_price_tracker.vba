' =====================================================
' 주식 현재가 추적 VBA 스크립트
'
' 사용법:
' 1. 엑셀에서 Alt + F11로 VBA 편집기 열기
' 2. 삽입 > 모듈
' 3. 이 코드 붙여넣기
' 4. UpdateStockPrices 실행 (F5 또는 매크로 실행)
'
' 전제조건:
' - '데이터' 시트에 A열: 종목명, B열: 종목코드
' - 인터넷 연결 필요
' =====================================================

Option Explicit

' 메인 함수: 주식 현재가 업데이트
Sub UpdateStockPrices()
    Dim wsData As Worksheet
    Dim wsToday As Worksheet
    Dim todayName As String
    Dim lastRow As Long
    Dim i As Long
    Dim stockName As String
    Dim stockCode As String
    Dim currentPrice As String
    Dim priceChange As String
    Dim changePercent As String

    On Error GoTo ErrorHandler

    ' 오늘 날짜로 시트 이름 생성 (예: 2025-12-07)
    todayName = Format(Date, "YYYY-MM-DD")

    ' 데이터 시트 확인
    Set wsData = Nothing
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("데이터")
    On Error GoTo ErrorHandler

    If wsData Is Nothing Then
        MsgBox "'데이터' 시트를 찾을 수 없습니다." & vbCrLf & _
               "A열: 종목명, B열: 종목코드가 있는 '데이터' 시트를 만들어주세요.", _
               vbExclamation, "오류"
        Exit Sub
    End If

    ' 오늘 날짜 시트 확인 또는 생성
    Set wsToday = GetOrCreateSheet(todayName)

    ' 헤더 설정
    SetupHeader wsToday

    ' 데이터 시트의 마지막 행 찾기
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "데이터 시트에 종목 데이터가 없습니다.", vbExclamation, "오류"
        Exit Sub
    End If

    ' 상태 표시
    Application.ScreenUpdating = False
    Application.StatusBar = "주식 현재가 업데이트 중..."

    ' 각 종목에 대해 현재가 가져오기
    For i = 2 To lastRow
        stockName = Trim(wsData.Cells(i, 1).Value)
        stockCode = Trim(wsData.Cells(i, 2).Value)

        If stockCode <> "" Then
            ' 상태 업데이트
            Application.StatusBar = "처리 중: " & stockName & " (" & i - 1 & "/" & lastRow - 1 & ")"

            ' 종목코드 정리 (숫자만 추출)
            stockCode = CleanStockCode(stockCode)

            ' 현재가 가져오기
            Call GetStockPrice(stockCode, currentPrice, priceChange, changePercent)

            ' 결과 기록
            wsToday.Cells(i, 1).Value = stockName
            wsToday.Cells(i, 2).Value = stockCode
            wsToday.Cells(i, 3).Value = currentPrice
            wsToday.Cells(i, 4).Value = priceChange
            wsToday.Cells(i, 5).Value = changePercent
            wsToday.Cells(i, 6).Value = Format(Now, "HH:MM:SS")

            ' 색상 적용 (상승/하락)
            If InStr(priceChange, "+") > 0 Or Val(priceChange) > 0 Then
                wsToday.Cells(i, 4).Font.Color = RGB(255, 0, 0)  ' 빨간색 (상승)
                wsToday.Cells(i, 5).Font.Color = RGB(255, 0, 0)
            ElseIf InStr(priceChange, "-") > 0 Or Val(priceChange) < 0 Then
                wsToday.Cells(i, 4).Font.Color = RGB(0, 0, 255)  ' 파란색 (하락)
                wsToday.Cells(i, 5).Font.Color = RGB(0, 0, 255)
            End If

            ' 서버 부하 방지를 위한 딜레이
            Application.Wait Now + TimeValue("00:00:00.3")
        End If
    Next i

    ' 열 너비 자동 조정
    wsToday.Columns("A:F").AutoFit

    ' 숫자 형식 적용
    wsToday.Range("C2:C" & lastRow).NumberFormat = "#,##0"
    wsToday.Range("D2:D" & lastRow).NumberFormat = "+#,##0;-#,##0;0"
    wsToday.Range("E2:E" & lastRow).NumberFormat = "+0.00%;-0.00%;0.00%"

    ' 정리
    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' 완료 메시지
    MsgBox "주식 현재가 업데이트 완료!" & vbCrLf & _
           "총 " & lastRow - 1 & "개 종목 처리됨" & vbCrLf & _
           "시트: " & todayName, vbInformation, "완료"

    ' 오늘 시트로 이동
    wsToday.Activate

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "오류 발생: " & Err.Description, vbCritical, "오류"
End Sub

' 시트 가져오기 또는 생성
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' 새 시트 생성
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    Else
        ' 기존 시트 데이터 영역만 초기화 (헤더 제외)
        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If lastRow > 1 Then
            ws.Range("A2:F" & lastRow).ClearContents
        End If
    End If

    Set GetOrCreateSheet = ws
End Function

' 헤더 설정
Private Sub SetupHeader(ws As Worksheet)
    With ws.Range("A1:F1")
        .Value = Array("종목명", "종목코드", "현재가", "전일대비", "등락률", "업데이트시간")
        .Font.Bold = True
        .Interior.Color = RGB(70, 130, 180)  ' Steel Blue
        .Font.Color = RGB(255, 255, 255)     ' White
        .HorizontalAlignment = xlCenter
    End With
End Sub

' 종목코드 정리 (숫자 6자리로)
Private Function CleanStockCode(code As String) As String
    Dim result As String
    Dim i As Integer

    result = ""
    For i = 1 To Len(code)
        If IsNumeric(Mid(code, i, 1)) Then
            result = result & Mid(code, i, 1)
        End If
    Next i

    ' 6자리로 패딩
    If Len(result) < 6 Then
        result = String(6 - Len(result), "0") & result
    End If

    CleanStockCode = result
End Function

' 네이버 금융에서 주식 현재가 가져오기
Private Sub GetStockPrice(stockCode As String, ByRef price As String, ByRef change As String, ByRef changePercent As String)
    Dim http As Object
    Dim url As String
    Dim response As String

    On Error GoTo PriceError

    ' 기본값 설정
    price = "N/A"
    change = "N/A"
    changePercent = "N/A"

    ' 네이버 금융 API URL
    url = "https://finance.naver.com/item/main.naver?code=" & stockCode

    ' HTTP 요청
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send

    If http.Status = 200 Then
        response = http.responseText

        ' 현재가 파싱
        price = ParsePrice(response)
        change = ParseChange(response)
        changePercent = ParseChangePercent(response)
    End If

    Set http = Nothing
    Exit Sub

PriceError:
    price = "오류"
    change = "-"
    changePercent = "-"
    Set http = Nothing
End Sub

' HTML에서 현재가 추출
Private Function ParsePrice(html As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempStr As String

    On Error GoTo ParseError

    ' 현재가 찾기 (네이버 금융 HTML 구조)
    startPos = InStr(html, "no_today")
    If startPos > 0 Then
        startPos = InStr(startPos, html, "<span class=""blind"">")
        If startPos > 0 Then
            startPos = startPos + Len("<span class=""blind"">")
            endPos = InStr(startPos, html, "</span>")
            If endPos > startPos Then
                tempStr = Mid(html, startPos, endPos - startPos)
                tempStr = Replace(tempStr, ",", "")
                tempStr = Trim(tempStr)
                If IsNumeric(tempStr) Then
                    ParsePrice = tempStr
                    Exit Function
                End If
            End If
        End If
    End If

    ' 대체 방법: sise_new API 사용
    ParsePrice = "N/A"
    Exit Function

ParseError:
    ParsePrice = "N/A"
End Function

' HTML에서 전일대비 추출
Private Function ParseChange(html As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempStr As String
    Dim isUp As Boolean

    On Error GoTo ParseError

    ' 상승/하락 여부 확인
    isUp = InStr(html, "no_exday up") > 0

    ' 전일대비 찾기
    startPos = InStr(html, "no_exday")
    If startPos > 0 Then
        startPos = InStr(startPos, html, "<span class=""blind"">")
        If startPos > 0 Then
            startPos = startPos + Len("<span class=""blind"">")
            endPos = InStr(startPos, html, "</span>")
            If endPos > startPos Then
                tempStr = Mid(html, startPos, endPos - startPos)
                tempStr = Replace(tempStr, ",", "")
                tempStr = Trim(tempStr)
                If IsNumeric(tempStr) Then
                    If isUp Then
                        ParseChange = "+" & tempStr
                    Else
                        ParseChange = "-" & tempStr
                    End If
                    Exit Function
                End If
            End If
        End If
    End If

    ParseChange = "0"
    Exit Function

ParseError:
    ParseChange = "N/A"
End Function

' HTML에서 등락률 추출
Private Function ParseChangePercent(html As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempStr As String
    Dim isUp As Boolean

    On Error GoTo ParseError

    ' 상승/하락 여부 확인
    isUp = InStr(html, "no_exday up") > 0

    ' 등락률 찾기
    startPos = InStr(html, "no_exday")
    If startPos > 0 Then
        ' 두 번째 blind span 찾기 (등락률)
        startPos = InStr(startPos, html, "<span class=""blind"">")
        If startPos > 0 Then
            startPos = InStr(startPos + 1, html, "<span class=""blind"">")
            If startPos > 0 Then
                startPos = startPos + Len("<span class=""blind"">")
                endPos = InStr(startPos, html, "</span>")
                If endPos > startPos Then
                    tempStr = Mid(html, startPos, endPos - startPos)
                    tempStr = Replace(tempStr, "%", "")
                    tempStr = Trim(tempStr)
                    If IsNumeric(tempStr) Then
                        If isUp Then
                            ParseChangePercent = Val(tempStr) / 100
                        Else
                            ParseChangePercent = -Val(tempStr) / 100
                        End If
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    ParseChangePercent = 0
    Exit Function

ParseError:
    ParseChangePercent = "N/A"
End Function

' =====================================================
' 대체 방법: 네이버 시세 API 사용 (더 안정적)
' =====================================================

Sub UpdateStockPricesAPI()
    Dim wsData As Worksheet
    Dim wsToday As Worksheet
    Dim todayName As String
    Dim lastRow As Long
    Dim i As Long
    Dim stockName As String
    Dim stockCode As String
    Dim priceData As Variant

    On Error GoTo ErrorHandler

    todayName = Format(Date, "YYYY-MM-DD")

    Set wsData = Nothing
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("데이터")
    On Error GoTo ErrorHandler

    If wsData Is Nothing Then
        MsgBox "'데이터' 시트를 찾을 수 없습니다.", vbExclamation, "오류"
        Exit Sub
    End If

    Set wsToday = GetOrCreateSheet(todayName)
    SetupHeader wsToday

    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row

    Application.ScreenUpdating = False
    Application.StatusBar = "주식 현재가 업데이트 중..."

    For i = 2 To lastRow
        stockName = Trim(wsData.Cells(i, 1).Value)
        stockCode = Trim(wsData.Cells(i, 2).Value)

        If stockCode <> "" Then
            Application.StatusBar = "처리 중: " & stockName & " (" & i - 1 & "/" & lastRow - 1 & ")"

            stockCode = CleanStockCode(stockCode)
            priceData = GetStockPriceAPI(stockCode)

            wsToday.Cells(i, 1).Value = stockName
            wsToday.Cells(i, 2).Value = stockCode
            wsToday.Cells(i, 3).Value = priceData(0)  ' 현재가
            wsToday.Cells(i, 4).Value = priceData(1)  ' 전일대비
            wsToday.Cells(i, 5).Value = priceData(2)  ' 등락률
            wsToday.Cells(i, 6).Value = Format(Now, "HH:MM:SS")

            ' 색상 적용
            If Val(priceData(1)) > 0 Then
                wsToday.Cells(i, 4).Font.Color = RGB(255, 0, 0)
                wsToday.Cells(i, 5).Font.Color = RGB(255, 0, 0)
            ElseIf Val(priceData(1)) < 0 Then
                wsToday.Cells(i, 4).Font.Color = RGB(0, 0, 255)
                wsToday.Cells(i, 5).Font.Color = RGB(0, 0, 255)
            End If

            Application.Wait Now + TimeValue("00:00:00.3")
        End If
    Next i

    wsToday.Columns("A:F").AutoFit
    wsToday.Range("C2:C" & lastRow).NumberFormat = "#,##0"
    wsToday.Range("D2:D" & lastRow).NumberFormat = "+#,##0;-#,##0;0"
    wsToday.Range("E2:E" & lastRow).NumberFormat = "+0.00%;-0.00%;0.00%"

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "주식 현재가 업데이트 완료!" & vbCrLf & _
           "총 " & lastRow - 1 & "개 종목 처리됨", vbInformation, "완료"

    wsToday.Activate
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "오류 발생: " & Err.Description, vbCritical, "오류"
End Sub

' 네이버 시세 API로 가격 가져오기
Private Function GetStockPriceAPI(stockCode As String) As Variant
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim result(2) As Variant
    Dim lines() As String
    Dim fields() As String

    On Error GoTo APIError

    result(0) = "N/A"
    result(1) = 0
    result(2) = 0

    ' 네이버 시세 CSV API
    url = "https://fchart.stock.naver.com/siseJson.nhn?symbol=" & stockCode & "&requestType=1&count=1"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send

    If http.Status = 200 Then
        response = http.responseText

        ' JSON 파싱 (간단한 방식)
        ' 응답 형식: [["날짜","시가","고가","저가","종가","거래량"]]
        response = Replace(response, "[", "")
        response = Replace(response, "]", "")
        response = Replace(response, """", "")
        response = Replace(response, "'", "")
        response = Replace(response, vbLf, "")
        response = Replace(response, vbCr, "")

        lines = Split(response, ",")

        If UBound(lines) >= 4 Then
            result(0) = Val(Trim(lines(4)))  ' 종가(현재가)
        End If
    End If

    ' 전일대비와 등락률은 별도 API 호출 필요
    Call GetPriceChange(stockCode, result)

    Set http = Nothing
    GetStockPriceAPI = result
    Exit Function

APIError:
    result(0) = "오류"
    result(1) = 0
    result(2) = 0
    Set http = Nothing
    GetStockPriceAPI = result
End Function

' 전일대비 가져오기
Private Sub GetPriceChange(stockCode As String, ByRef result As Variant)
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim startPos As Long, endPos As Long
    Dim prevClose As Double, curPrice As Double

    On Error GoTo ChangeError

    ' 간이 시세 API
    url = "https://finance.naver.com/item/sise.naver?code=" & stockCode

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send

    If http.Status = 200 Then
        response = http.responseText

        ' 전일종가 찾기
        startPos = InStr(response, "전일가</th>")
        If startPos > 0 Then
            startPos = InStr(startPos, response, "<td class=""num"">")
            If startPos > 0 Then
                startPos = startPos + Len("<td class=""num"">")
                endPos = InStr(startPos, response, "</td>")
                If endPos > startPos Then
                    prevClose = Val(Replace(Mid(response, startPos, endPos - startPos), ",", ""))
                End If
            End If
        End If

        ' 현재가로 등락 계산
        If result(0) <> "N/A" And result(0) <> "오류" And prevClose > 0 Then
            curPrice = Val(result(0))
            result(1) = curPrice - prevClose
            result(2) = (curPrice - prevClose) / prevClose
        End If
    End If

    Set http = Nothing
    Exit Sub

ChangeError:
    Set http = Nothing
End Sub

' =====================================================
' 버튼/단축키 등록용 래퍼 함수들
' =====================================================

' 빠른 업데이트 (HTML 파싱)
Sub QuickUpdate()
    UpdateStockPrices
End Sub

' API 업데이트 (더 안정적)
Sub APIUpdate()
    UpdateStockPricesAPI
End Sub

' 특정 시트 삭제
Sub DeleteDateSheet()
    Dim sheetName As String
    sheetName = InputBox("삭제할 시트 이름 (날짜):", "시트 삭제", Format(Date, "YYYY-MM-DD"))

    If sheetName = "" Then Exit Sub

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    MsgBox "시트 '" & sheetName & "' 삭제됨", vbInformation
End Sub

' 모든 날짜 시트 삭제 (데이터 시트 제외)
Sub ClearAllDateSheets()
    Dim ws As Worksheet
    Dim count As Integer

    If MsgBox("모든 날짜 시트를 삭제하시겠습니까?" & vbCrLf & _
              "(데이터 시트는 유지됩니다)", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    count = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "데이터" And IsDateSheet(ws.Name) Then
            ws.Delete
            count = count + 1
        End If
    Next ws

    Application.DisplayAlerts = True
    MsgBox count & "개 시트 삭제됨", vbInformation
End Sub

' 날짜 형식 시트인지 확인
Private Function IsDateSheet(sheetName As String) As Boolean
    On Error GoTo NotDate

    If Len(sheetName) = 10 And Mid(sheetName, 5, 1) = "-" And Mid(sheetName, 8, 1) = "-" Then
        IsDateSheet = True
        Exit Function
    End If

NotDate:
    IsDateSheet = False
End Function
