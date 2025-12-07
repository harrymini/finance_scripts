' =====================================================
' 주식 현재가 추적 VBA 스크립트 v3
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

    ' 오늘 날짜로 시트 이름 생성
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

    ' 데이터 시트의 마지막 행 찾기
    lastRow = wsData.Cells(wsData.Rows.count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "데이터 시트에 종목 데이터가 없습니다.", vbExclamation, "오류"
        Exit Sub
    End If

    ' 상태 표시
    Application.ScreenUpdating = False
    Application.StatusBar = "주식 현재가 업데이트 중..."

    processedCount = 0
    rowNum = 2

    ' 2행부터 데이터 읽기
    For i = 2 To lastRow
        stockName = Trim(CStr(wsData.Cells(i, 1).Value))
        stockCode = Trim(CStr(wsData.Cells(i, 2).Value))

        If stockName = "" And stockCode = "" Then
            GoTo NextRow
        End If

        If stockCode <> "" Then
            Application.StatusBar = "처리 중: " & stockName & " (" & processedCount + 1 & "/" & (lastRow - 1) & ")"
            DoEvents

            ' 종목코드 정리
            stockCode = CleanStockCode(stockCode)

            ' 현재가 가져오기 (새 API 사용)
            Call GetStockPriceNew(stockCode, currentPrice, priceChange, changePercent)

            ' 결과 기록
            wsToday.Cells(rowNum, 1).Value = stockName
            wsToday.Cells(rowNum, 2).Value = "'" & stockCode
            wsToday.Cells(rowNum, 3).Value = currentPrice
            wsToday.Cells(rowNum, 4).Value = priceChange
            wsToday.Cells(rowNum, 5).Value = changePercent
            wsToday.Cells(rowNum, 6).Value = Format(Now, "hh:mm:ss")

            ' 색상 적용
            ApplyPriceColor wsToday, rowNum, priceChange

            rowNum = rowNum + 1
            processedCount = processedCount + 1

            ' 딜레이
            Delay 0.3
        End If

NextRow:
    Next i

    wsToday.Columns("A:F").AutoFit

    Application.StatusBar = False
    Application.ScreenUpdating = True

    MsgBox "주식 현재가 업데이트 완료!" & vbCrLf & vbCrLf & _
           "처리된 종목: " & processedCount & "개" & vbCrLf & _
           "시트: " & todayName, vbInformation, "완료"

    wsToday.Activate
    Exit Sub

ErrorHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "오류 발생: " & Err.Description & vbCrLf & _
           "오류 번호: " & Err.Number, vbCritical, "오류"
End Sub

' =====================================================
' 주가 데이터 가져오기 (네이버 모바일 API - JSON)
' =====================================================
Private Sub GetStockPriceNew(stockCode As String, ByRef price As String, ByRef change As String, ByRef changePercent As String)
    Dim http As Object
    Dim url As String
    Dim response As String
    Dim json As String

    On Error GoTo PriceError

    price = "-"
    change = "-"
    changePercent = "-"

    ' 네이버 모바일 주식 API (JSON 반환)
    url = "https://m.stock.naver.com/api/stock/" & stockCode & "/basic"

    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.send

    If http.Status = 200 Then
        json = http.responseText

        ' JSON에서 값 추출
        price = ExtractJsonValue(json, "closePrice")
        If price = "" Then price = ExtractJsonValue(json, "currentPrice")

        change = ExtractJsonValue(json, "compareToPreviousClosePrice")
        changePercent = ExtractJsonValue(json, "fluctuationsRatio")

        ' 포맷팅
        If price <> "" And price <> "-" Then
            price = Format(Val(Replace(price, ",", "")), "#,##0")
        End If

        If change <> "" And change <> "-" Then
            Dim changeVal As Double
            changeVal = Val(Replace(change, ",", ""))
            If changeVal > 0 Then
                change = "+" & Format(changeVal, "#,##0")
            ElseIf changeVal < 0 Then
                change = Format(changeVal, "#,##0")
            Else
                change = "0"
            End If
        End If

        If changePercent <> "" And changePercent <> "-" Then
            Dim pctVal As Double
            pctVal = Val(changePercent)
            If pctVal > 0 Then
                changePercent = "+" & Format(pctVal, "0.00") & "%"
            ElseIf pctVal < 0 Then
                changePercent = Format(pctVal, "0.00") & "%"
            Else
                changePercent = "0.00%"
            End If
        End If
    End If

    Set http = Nothing
    Exit Sub

PriceError:
    price = "오류"
    change = "-"
    changePercent = "-"
    If Not http Is Nothing Then Set http = Nothing
End Sub

' JSON에서 특정 키의 값 추출 (간단한 파서)
Private Function ExtractJsonValue(json As String, key As String) As String
    Dim searchKey As String
    Dim startPos As Long
    Dim endPos As Long
    Dim value As String

    On Error GoTo ExtractError

    ' "key": 또는 "key":로 검색
    searchKey = """" & key & """:"

    startPos = InStr(1, json, searchKey, vbTextCompare)
    If startPos = 0 Then
        ExtractJsonValue = ""
        Exit Function
    End If

    startPos = startPos + Len(searchKey)

    ' 공백 건너뛰기
    Do While Mid(json, startPos, 1) = " "
        startPos = startPos + 1
    Loop

    ' 값 타입 확인
    If Mid(json, startPos, 1) = """" Then
        ' 문자열 값
        startPos = startPos + 1
        endPos = InStr(startPos, json, """", vbTextCompare)
    ElseIf Mid(json, startPos, 1) = "n" Then
        ' null
        ExtractJsonValue = ""
        Exit Function
    Else
        ' 숫자 값
        endPos = startPos
        Do While endPos <= Len(json)
            Dim c As String
            c = Mid(json, endPos, 1)
            If c = "," Or c = "}" Or c = "]" Or c = " " Or c = vbCr Or c = vbLf Then
                Exit Do
            End If
            endPos = endPos + 1
        Loop
    End If

    If endPos > startPos Then
        value = Mid(json, startPos, endPos - startPos)
        ExtractJsonValue = Trim(value)
    Else
        ExtractJsonValue = ""
    End If

    Exit Function

ExtractError:
    ExtractJsonValue = ""
End Function

' =====================================================
' 시트 관리 함수들
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
' 유틸리티 함수들
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
' 유틸리티 매크로
' =====================================================

Sub DeleteDateSheet()
    Dim sheetName As String
    sheetName = InputBox("삭제할 시트 이름:", "시트 삭제", Format(Date, "yyyy-mm-dd"))
    If sheetName = "" Then Exit Sub

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Worksheets(sheetName).Delete
    Application.DisplayAlerts = True

    If Err.Number = 0 Then
        MsgBox "삭제됨: " & sheetName, vbInformation
    Else
        MsgBox "시트를 찾을 수 없습니다", vbExclamation
    End If
    On Error GoTo 0
End Sub

Sub ClearAllDateSheets()
    Dim ws As Worksheet
    Dim count As Long

    If MsgBox("모든 날짜 시트를 삭제하시겠습니까?", vbYesNo + vbQuestion) = vbNo Then
        Exit Sub
    End If

    Application.DisplayAlerts = False
    count = 0

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> "데이터" And Len(ws.Name) = 10 Then
            If Mid(ws.Name, 5, 1) = "-" And Mid(ws.Name, 8, 1) = "-" Then
                ws.Delete
                count = count + 1
            End If
        End If
    Next ws

    Application.DisplayAlerts = True
    MsgBox count & "개 시트 삭제됨", vbInformation
End Sub
