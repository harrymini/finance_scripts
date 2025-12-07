' =====================================================
' 주식 현재가 추적 VBA 스크립트 v5 (KRX API)
'
' 사용법:
' 1. 엑셀에서 Alt + F11로 VBA 편집기 열기
' 2. 삽입 > 모듈
' 3. 이 코드 붙여넣기
' 4. UpdateStockPrices 실행 (F5 또는 매크로 실행)
'
' 전제조건:
' - '데이터' 시트: 1행 헤더, 2행부터 데이터
'   A열: 종목명, B열: 종목코드
' - 인터넷 연결 필요
' =====================================================

Option Explicit

' =====================================================
' 메인 함수: 주식 현재가 업데이트
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

            ' KRX API로 현재가 가져오기
            Call GetKRXStockPrice(stockCode, currentPrice, priceChange, changePercent)

            wsToday.Cells(rowNum, 1).Value = stockName
            wsToday.Cells(rowNum, 2).Value = "'" & stockCode
            wsToday.Cells(rowNum, 3).Value = currentPrice
            wsToday.Cells(rowNum, 4).Value = priceChange
            wsToday.Cells(rowNum, 5).Value = changePercent
            wsToday.Cells(rowNum, 6).Value = Format(Now, "hh:mm:ss")

            ApplyPriceColor wsToday, rowNum, priceChange

            rowNum = rowNum + 1
            processedCount = processedCount + 1

            Delay 0.3
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
' KRX API로 주가 데이터 가져오기
' =====================================================
Private Sub GetKRXStockPrice(stockCode As String, ByRef price As String, ByRef change As String, ByRef changePercent As String)
    Dim http As Object
    Dim url As String
    Dim postData As String
    Dim response As String
    Dim curPrice As Double
    Dim prevPrice As Double
    Dim diffPrice As Double
    Dim diffPct As Double

    On Error GoTo KRXError

    price = "-"
    change = "-"
    changePercent = "-"

    ' KRX 정보데이터시스템 API
    url = "http://data.krx.co.kr/comm/bldAttendant/getJsonData.cmd"

    ' POST 데이터
    postData = "bld=dbms/MDC/STAT/standard/MDCSTAT01501" & _
               "&locale=ko_KR" & _
               "&isuCd=KR7" & stockCode & "003" & _
               "&isuCd2=KR7" & stockCode & "003" & _
               "&strtDd=" & Format(Date, "yyyymmdd") & _
               "&endDd=" & Format(Date, "yyyymmdd") & _
               "&share=1" & _
               "&money=1"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/x-www-form-urlencoded; charset=UTF-8"
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.setRequestHeader "Referer", "http://data.krx.co.kr/contents/MDC/MDI/mdiLoader/index.cmd?menuId=MDC0201010101"
    http.send postData

    If http.Status = 200 Then
        response = http.responseText

        ' JSON에서 값 추출
        curPrice = ExtractKRXValue(response, "TDD_CLSPRC")
        prevPrice = ExtractKRXValue(response, "TDD_OPNPRC")

        If curPrice = 0 Then
            ' 대체 필드 시도
            curPrice = ExtractKRXValue(response, "CLSPRC")
        End If

        If curPrice > 0 Then
            price = Format(curPrice, "#,##0")

            ' 전일대비 계산
            diffPrice = ExtractKRXValue(response, "CMPPREVDD_PRC")
            diffPct = ExtractKRXValue(response, "FLUC_RT")

            If diffPrice <> 0 Or diffPct <> 0 Then
                If diffPrice > 0 Then
                    change = "+" & Format(diffPrice, "#,##0")
                ElseIf diffPrice < 0 Then
                    change = Format(diffPrice, "#,##0")
                Else
                    change = "0"
                End If

                If diffPct > 0 Then
                    changePercent = "+" & Format(diffPct, "0.00") & "%"
                ElseIf diffPct < 0 Then
                    changePercent = Format(diffPct, "0.00") & "%"
                Else
                    changePercent = "0.00%"
                End If
            End If
        Else
            ' KRX 실패시 네이버로 폴백
            Call GetNaverStockPrice(stockCode, price, change, changePercent)
        End If
    Else
        ' HTTP 실패시 네이버로 폴백
        Call GetNaverStockPrice(stockCode, price, change, changePercent)
    End If

    Set http = Nothing
    Exit Sub

KRXError:
    ' 에러시 네이버로 폴백
    Call GetNaverStockPrice(stockCode, price, change, changePercent)
    If Not http Is Nothing Then Set http = Nothing
End Sub

' KRX JSON에서 값 추출
Private Function ExtractKRXValue(json As String, key As String) As Double
    Dim searchKey As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String
    Dim i As Long
    Dim c As String

    On Error GoTo ExtractErr

    searchKey = """" & key & """:"

    startPos = InStr(1, json, searchKey, vbTextCompare)
    If startPos = 0 Then
        ExtractKRXValue = 0
        Exit Function
    End If

    startPos = startPos + Len(searchKey)

    ' 공백 건너뛰기
    Do While Mid(json, startPos, 1) = " " Or Mid(json, startPos, 1) = """"
        startPos = startPos + 1
    Loop

    ' 값 추출 (숫자와 마이너스, 소수점만)
    value = ""
    For i = startPos To Len(json)
        c = Mid(json, i, 1)
        If (c >= "0" And c <= "9") Or c = "-" Or c = "." Then
            value = value & c
        ElseIf c = "," And Len(value) > 0 Then
            ' 천단위 구분자는 무시
        ElseIf Len(value) > 0 Then
            Exit For
        End If
    Next i

    If Len(value) > 0 Then
        ExtractKRXValue = Val(value)
    Else
        ExtractKRXValue = 0
    End If

    Exit Function

ExtractErr:
    ExtractKRXValue = 0
End Function

' =====================================================
' 네이버 금융 (폴백용)
' =====================================================
Private Sub GetNaverStockPrice(stockCode As String, ByRef price As String, ByRef change As String, ByRef changePercent As String)
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim curPrice As Double
    Dim prevPrice As Double

    On Error GoTo NaverError

    price = "-"
    change = "-"
    changePercent = "-"

    ' 네이버 금융 시세 조회
    url = "https://fchart.stock.naver.com/siseJson.nhn?symbol=" & stockCode & "&requestType=1&count=2"

    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send

    If http.Status = 200 Then
        response = http.responseText

        ' 간단한 파싱 (응답: [["날짜","시가","고가","저가","종가","거래량"], ...])
        curPrice = ExtractNaverPrice(response, 1)
        prevPrice = ExtractNaverPrice(response, 2)

        If curPrice > 0 Then
            price = Format(curPrice, "#,##0")

            If prevPrice > 0 Then
                Dim diff As Double
                Dim pct As Double
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

    Set http = Nothing
    Exit Sub

NaverError:
    price = "오류"
    change = "-"
    changePercent = "-"
    If Not http Is Nothing Then Set http = Nothing
End Sub

' 네이버 시세 JSON에서 종가 추출
Private Function ExtractNaverPrice(response As String, rowNum As Long) As Double
    Dim lines() As String
    Dim dataLine As String
    Dim values() As String
    Dim i As Long
    Dim lineCount As Long
    Dim cleanResponse As String

    On Error GoTo ExtractNaverErr

    ' 응답 정리
    cleanResponse = response
    cleanResponse = Replace(cleanResponse, "[", "")
    cleanResponse = Replace(cleanResponse, "]", "")
    cleanResponse = Replace(cleanResponse, """", "")
    cleanResponse = Replace(cleanResponse, "'", "")
    cleanResponse = Replace(cleanResponse, vbLf, "")
    cleanResponse = Replace(cleanResponse, vbCr, "")
    cleanResponse = Replace(cleanResponse, Chr(10), "")
    cleanResponse = Replace(cleanResponse, Chr(13), "")

    ' 줄 단위로 분리 (각 행은 콤마로 구분된 6개 값)
    ' 헤더 행 건너뛰고 데이터 행 찾기
    Dim allValues() As String
    allValues = Split(cleanResponse, ",")

    ' 각 행은 6개 값 (날짜,시가,고가,저가,종가,거래량)
    ' rowNum=1이면 첫번째 데이터 (인덱스 6~11), rowNum=2면 두번째 데이터 (인덱스 12~17)
    Dim idx As Long
    idx = 6 + (rowNum - 1) * 6 + 4  ' 종가 위치

    If idx <= UBound(allValues) Then
        ExtractNaverPrice = Val(Trim(allValues(idx)))
    Else
        ExtractNaverPrice = 0
    End If

    Exit Function

ExtractNaverErr:
    ExtractNaverPrice = 0
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

    code = InputBox("종목코드 입력 (예: 005930):", "테스트", "005930")
    If code = "" Then Exit Sub

    code = CleanStockCode(code)
    Call GetKRXStockPrice(code, price, change, pct)

    MsgBox "종목코드: " & code & vbCrLf & _
           "현재가: " & price & vbCrLf & _
           "전일대비: " & change & vbCrLf & _
           "등락률: " & pct, vbInformation, "테스트 결과"
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
