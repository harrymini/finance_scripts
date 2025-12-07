' =====================================================
' 주식 현재가 추적 VBA 스크립트 v2
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

    ' 오늘 날짜로 시트 이름 생성 (예: 2025-12-07)
    todayName = Format(Date, "yyyy-mm-dd")

    ' 데이터 시트 확인
    On Error Resume Next
    Set wsData = ThisWorkbook.Worksheets("데이터")
    On Error GoTo ErrorHandler

    If wsData Is Nothing Then
        MsgBox "'데이터' 시트를 찾을 수 없습니다." & vbCrLf & vbCrLf & _
               "1행: 헤더 (종목명, 종목코드)" & vbCrLf & _
               "2행부터: 데이터", vbExclamation, "오류"
        Exit Sub
    End If

    ' 오늘 날짜 시트 확인 또는 생성
    Set wsToday = GetOrCreateSheet(todayName)

    ' 헤더 설정
    SetupHeader wsToday

    ' 데이터 시트의 마지막 행 찾기 (A열 기준)
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "데이터 시트에 종목 데이터가 없습니다." & vbCrLf & _
               "2행부터 데이터를 입력해주세요.", vbExclamation, "오류"
        Exit Sub
    End If

    ' 상태 표시
    Application.ScreenUpdating = False
    Application.StatusBar = "주식 현재가 업데이트 중..."

    processedCount = 0
    rowNum = 2  ' 결과 시트의 시작 행

    ' 2행부터 데이터 읽기 (1행은 헤더)
    For i = 2 To lastRow
        stockName = Trim(CStr(wsData.Cells(i, 1).Value))
        stockCode = Trim(CStr(wsData.Cells(i, 2).Value))

        ' 빈 행 건너뛰기
        If stockName = "" And stockCode = "" Then
            GoTo NextRow
        End If

        If stockCode <> "" Then
            ' 상태 업데이트
            Application.StatusBar = "처리 중: " & stockName & " (" & processedCount + 1 & "/" & (lastRow - 1) & ")"
            DoEvents

            ' 종목코드 정리 (숫자 6자리)
            stockCode = CleanStockCode(stockCode)

            ' 현재가 가져오기
            Call GetStockPrice(stockCode, currentPrice, priceChange, changePercent)

            ' 결과 기록
            wsToday.Cells(rowNum, 1).Value = stockName
            wsToday.Cells(rowNum, 2).Value = "'" & stockCode  ' 텍스트로 저장
            wsToday.Cells(rowNum, 3).Value = currentPrice
            wsToday.Cells(rowNum, 4).Value = priceChange
            wsToday.Cells(rowNum, 5).Value = changePercent
            wsToday.Cells(rowNum, 6).Value = Format(Now, "hh:mm:ss")

            ' 색상 적용 (상승/하락)
            ApplyPriceColor wsToday, rowNum, priceChange

            rowNum = rowNum + 1
            processedCount = processedCount + 1

            ' 서버 부하 방지 딜레이 (0.5초)
            Delay 0.5
        End If

NextRow:
    Next i

    ' 열 너비 자동 조정
    wsToday.Columns("A:F").AutoFit

    ' 숫자 형식 적용
    If rowNum > 2 Then
        wsToday.Range("C2:C" & rowNum - 1).NumberFormat = "#,##0"
    End If

    ' 정리
    Application.StatusBar = False
    Application.ScreenUpdating = True

    ' 완료 메시지
    MsgBox "주식 현재가 업데이트 완료!" & vbCrLf & vbCrLf & _
           "처리된 종목: " & processedCount & "개" & vbCrLf & _
           "시트: " & todayName, vbInformation, "완료"

    ' 오늘 시트로 이동
    wsToday.Activate

    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "오류 발생: " & Err.Description & vbCrLf & _
           "오류 번호: " & Err.Number, vbCritical, "오류"
End Sub

' =====================================================
' 시트 관리 함수들
' =====================================================

' 시트 가져오기 또는 생성
Private Function GetOrCreateSheet(sheetName As String) As Worksheet
    Dim ws As Worksheet
    Dim lastRow As Long

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        ' 새 시트 생성
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    Else
        ' 기존 시트 데이터 영역만 초기화 (헤더 제외)
        lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row
        If lastRow > 1 Then
            ws.Range("A2:F" & lastRow).ClearContents
            ws.Range("A2:F" & lastRow).Font.Color = RGB(0, 0, 0)  ' 색상 초기화
        End If
    End If

    Set GetOrCreateSheet = ws
End Function

' 헤더 설정
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
' 유틸리티 함수들
' =====================================================

' 종목코드 정리 (숫자 6자리로)
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

    ' 6자리로 패딩
    Do While Len(result) < 6
        result = "0" & result
    Loop

    CleanStockCode = result
End Function

' 딜레이 함수 (초 단위)
Private Sub Delay(seconds As Double)
    Dim endTime As Double
    endTime = Timer + seconds
    Do While Timer < endTime
        DoEvents
    Loop
End Sub

' 가격 색상 적용
Private Sub ApplyPriceColor(ws As Worksheet, rowNum As Long, priceChange As String)
    Dim changeVal As Double

    On Error Resume Next
    changeVal = Val(Replace(Replace(priceChange, "+", ""), ",", ""))
    On Error GoTo 0

    If changeVal > 0 Then
        ws.Cells(rowNum, 4).Font.Color = RGB(255, 0, 0)  ' 빨간색 (상승)
        ws.Cells(rowNum, 5).Font.Color = RGB(255, 0, 0)
    ElseIf changeVal < 0 Then
        ws.Cells(rowNum, 4).Font.Color = RGB(0, 0, 255)  ' 파란색 (하락)
        ws.Cells(rowNum, 5).Font.Color = RGB(0, 0, 255)
    End If
End Sub

' =====================================================
' 주가 데이터 가져오기
' =====================================================

' 네이버 금융에서 주식 현재가 가져오기
Private Sub GetStockPrice(stockCode As String, ByRef price As String, ByRef change As String, ByRef changePercent As String)
    Dim http As Object
    Dim url As String
    Dim response As String

    On Error GoTo PriceError

    ' 기본값 설정
    price = "-"
    change = "-"
    changePercent = "-"

    ' 네이버 금융 URL
    url = "https://finance.naver.com/item/main.naver?code=" & stockCode

    ' HTTP 요청
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    http.send

    If http.Status = 200 Then
        response = http.responseText

        ' 현재가 파싱
        price = ParseCurrentPrice(response)
        change = ParsePriceChange(response)
        changePercent = ParseChangePercent(response)
    End If

    Set http = Nothing
    Exit Sub

PriceError:
    price = "오류"
    change = "-"
    changePercent = "-"
    If Not http Is Nothing Then Set http = Nothing
End Sub

' HTML에서 현재가 추출
Private Function ParseCurrentPrice(html As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempStr As String

    On Error GoTo ParseError

    ' 현재가 찾기 - "no_today" 클래스 내의 blind 스팬
    startPos = InStr(1, html, "no_today", vbTextCompare)
    If startPos > 0 Then
        startPos = InStr(startPos, html, "blind", vbTextCompare)
        If startPos > 0 Then
            startPos = InStr(startPos, html, ">", vbTextCompare)
            If startPos > 0 Then
                startPos = startPos + 1
                endPos = InStr(startPos, html, "<", vbTextCompare)
                If endPos > startPos Then
                    tempStr = Mid(html, startPos, endPos - startPos)
                    tempStr = Trim(Replace(tempStr, ",", ""))
                    If IsNumeric(tempStr) Then
                        ParseCurrentPrice = Format(Val(tempStr), "#,##0")
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    ParseCurrentPrice = "-"
    Exit Function

ParseError:
    ParseCurrentPrice = "-"
End Function

' HTML에서 전일대비 추출
Private Function ParsePriceChange(html As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempStr As String
    Dim isUp As Boolean
    Dim isDown As Boolean

    On Error GoTo ParseError

    ' 상승/하락 여부 확인
    isUp = InStr(1, html, "no_exday up", vbTextCompare) > 0
    isDown = InStr(1, html, "no_exday down", vbTextCompare) > 0

    ' 전일대비 찾기
    startPos = InStr(1, html, "no_exday", vbTextCompare)
    If startPos > 0 Then
        startPos = InStr(startPos, html, "blind", vbTextCompare)
        If startPos > 0 Then
            startPos = InStr(startPos, html, ">", vbTextCompare)
            If startPos > 0 Then
                startPos = startPos + 1
                endPos = InStr(startPos, html, "<", vbTextCompare)
                If endPos > startPos Then
                    tempStr = Mid(html, startPos, endPos - startPos)
                    tempStr = Trim(Replace(tempStr, ",", ""))
                    If IsNumeric(tempStr) Then
                        If isUp Then
                            ParsePriceChange = "+" & Format(Val(tempStr), "#,##0")
                        ElseIf isDown Then
                            ParsePriceChange = "-" & Format(Val(tempStr), "#,##0")
                        Else
                            ParsePriceChange = Format(Val(tempStr), "#,##0")
                        End If
                        Exit Function
                    End If
                End If
            End If
        End If
    End If

    ParsePriceChange = "0"
    Exit Function

ParseError:
    ParsePriceChange = "-"
End Function

' HTML에서 등락률 추출
Private Function ParseChangePercent(html As String) As String
    Dim startPos As Long
    Dim endPos As Long
    Dim tempStr As String
    Dim isUp As Boolean
    Dim isDown As Boolean
    Dim searchStart As Long

    On Error GoTo ParseError

    ' 상승/하락 여부 확인
    isUp = InStr(1, html, "no_exday up", vbTextCompare) > 0
    isDown = InStr(1, html, "no_exday down", vbTextCompare) > 0

    ' 등락률 찾기 (두 번째 blind 스팬)
    startPos = InStr(1, html, "no_exday", vbTextCompare)
    If startPos > 0 Then
        ' 첫 번째 blind 건너뛰기
        searchStart = InStr(startPos, html, "blind", vbTextCompare)
        If searchStart > 0 Then
            searchStart = InStr(searchStart + 5, html, "blind", vbTextCompare)
            If searchStart > 0 Then
                startPos = InStr(searchStart, html, ">", vbTextCompare)
                If startPos > 0 Then
                    startPos = startPos + 1
                    endPos = InStr(startPos, html, "<", vbTextCompare)
                    If endPos > startPos Then
                        tempStr = Mid(html, startPos, endPos - startPos)
                        tempStr = Trim(Replace(tempStr, "%", ""))
                        If IsNumeric(tempStr) Then
                            If isUp Then
                                ParseChangePercent = "+" & tempStr & "%"
                            ElseIf isDown Then
                                ParseChangePercent = "-" & tempStr & "%"
                            Else
                                ParseChangePercent = tempStr & "%"
                            End If
                            Exit Function
                        End If
                    End If
                End If
            End If
        End If
    End If

    ParseChangePercent = "0%"
    Exit Function

ParseError:
    ParseChangePercent = "-"
End Function

' =====================================================
' 유틸리티 매크로
' =====================================================

' 특정 날짜 시트 삭제
Sub DeleteDateSheet()
    Dim sheetName As String

    sheetName = InputBox("삭제할 시트 이름 (날짜):", "시트 삭제", Format(Date, "yyyy-mm-dd"))

    If sheetName = "" Then Exit Sub

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True

    If Err.Number = 0 Then
        MsgBox "시트 '" & sheetName & "' 삭제됨", vbInformation
    Else
        MsgBox "시트를 찾을 수 없습니다: " & sheetName, vbExclamation
    End If
    On Error GoTo 0
End Sub

' 모든 날짜 시트 삭제 (데이터 시트 제외)
Sub ClearAllDateSheets()
    Dim ws As Worksheet
    Dim count As Long

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
    ' yyyy-mm-dd 형식 확인
    If Len(sheetName) = 10 Then
        If Mid(sheetName, 5, 1) = "-" And Mid(sheetName, 8, 1) = "-" Then
            IsDateSheet = True
            Exit Function
        End If
    End If
    IsDateSheet = False
End Function
