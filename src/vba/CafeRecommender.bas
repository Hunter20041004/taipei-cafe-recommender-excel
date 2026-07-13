Attribute VB_Name = "CafeRecommender"
'======================================================
' 新店、文山、大安咖啡廳推薦系統 - VBA 巨集模組
' 功能：篩選、排序、超連結、統計摘要、條件式格式化
'======================================================
Option Explicit

Sub 搜尋咖啡廳()
    Dim wsDash As Worksheet
    Dim wsData As Worksheet
    Dim lastRow As Long
    Dim i As Long, j As Long
    
    Set wsDash = ThisWorkbook.Sheets("儀表板")
    Set wsData = ThisWorkbook.Sheets("全部資料")
    
    ' ── 讀取使用者篩選條件 ──
    Dim priceFilter As String
    Dim commuteFilter As String
    Dim timeFilter As String
    Dim areaFilter As String
    
    priceFilter = Trim(CStr(wsDash.Range("D5").Value))
    commuteFilter = Trim(CStr(wsDash.Range("D6").Value))
    timeFilter = Trim(CStr(wsDash.Range("D7").Value))
    areaFilter = Trim(CStr(wsDash.Range("D8").Value))
    
    ' ── 清除舊結果 ──
    wsDash.Range("B11:C25").ClearContents
    ' 清除舊超連結
    Dim hl As Hyperlink
    For Each hl In wsDash.Hyperlinks
        If Not Intersect(hl.Range, wsDash.Range("C11:C25")) Is Nothing Then
            hl.Delete
        End If
    Next hl
    ' 清除舊統計摘要
    wsDash.Range("B27:I28").ClearContents
    ' 清除舊條件式格式化 (I欄)
    wsDash.Range("I11:I25").FormatConditions.Delete
    
    ' ── 取得資料列數 ──
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "全部資料表中沒有資料！", vbExclamation, "錯誤"
        Exit Sub
    End If
    
    ' ── 篩選符合條件的咖啡廳 ──
    Dim matchNames() As String
    Dim matchRatings() As Double
    Dim matchUrls() As String
    Dim matchCount As Long
    matchCount = 0
    ReDim matchNames(1 To lastRow)
    ReDim matchRatings(1 To lastRow)
    ReDim matchUrls(1 To lastRow)
    
    For i = 2 To lastRow
        Dim cafeName As String
        Dim cafeArea As String
        Dim commute As Double
        Dim price As Double
        Dim timeLimit As String
        Dim rating As Double
        Dim cafeUrl As String
        
        cafeName = CStr(wsData.Cells(i, 1).Value)
        If Trim(cafeName) = "" Then GoTo NextRow
        If Not IsNumeric(wsData.Cells(i, 4).Value) Or _
           Not IsNumeric(wsData.Cells(i, 5).Value) Or _
           Not IsNumeric(wsData.Cells(i, 9).Value) Then GoTo NextRow
        
        cafeArea = Trim(CStr(wsData.Cells(i, 2).Value))
        commute = Val(wsData.Cells(i, 4).Value)
        price = Val(wsData.Cells(i, 5).Value)
        timeLimit = Trim(CStr(wsData.Cells(i, 8).Value))
        rating = Val(wsData.Cells(i, 9).Value)
        cafeUrl = Trim(CStr(wsData.Cells(i, 10).Value))
        
        ' 檢查飲料均價
        Dim priceOK As Boolean
        Select Case priceFilter
            Case "不限":       priceOK = True
            Case "150元以下":  priceOK = (price <= 150)
            Case "150-200元":  priceOK = (price > 150 And price <= 200)
            Case "200元以上":  priceOK = (price > 200)
            Case Else:         priceOK = True
        End Select
        
        ' 檢查通勤時間
        Dim commuteOK As Boolean
        Select Case commuteFilter
            Case "不限":         commuteOK = True
            Case "15分鐘以內":   commuteOK = (commute <= 15)
            Case "15-30分鐘":    commuteOK = (commute > 15 And commute <= 30)
            Case "30-45分鐘":    commuteOK = (commute > 30 And commute <= 45)
            Case "45分鐘以上":   commuteOK = (commute > 45)
            Case Else:           commuteOK = True
        End Select
        
        ' 檢查限時偏好
        Dim timeOK As Boolean
        Select Case timeFilter
            Case "不拘":   timeOK = True
            Case "不限時": timeOK = (timeLimit = "不限時")
            Case "有限時": timeOK = (Left(timeLimit, 3) = "有限時")
            Case Else:     timeOK = True
        End Select
        
        ' 檢查區域偏好
        Dim areaOK As Boolean
        Select Case areaFilter
            Case "不限":   areaOK = True
            Case Else:     areaOK = (cafeArea = areaFilter)
        End Select
        
        ' 四個條件都符合才收錄
        If priceOK And commuteOK And timeOK And areaOK Then
            matchCount = matchCount + 1
            matchNames(matchCount) = cafeName
            matchRatings(matchCount) = rating
            matchUrls(matchCount) = cafeUrl
        End If
NextRow:
    Next i
    
    ' ── 無結果處理 ──
    If matchCount = 0 Then
        wsDash.Range("C11").Value = "沒有符合條件的咖啡廳，請調整篩選條件"
        更新圖表 0
        MsgBox "沒有找到符合條件的咖啡廳！" & vbCrLf & _
               "請嘗試放寬篩選條件。", vbInformation, "搜尋結果"
        Exit Sub
    End If
    
    ' ── 依 Google 星級排序（泡沫排序，由高到低）──
    Dim tempName As String
    Dim tempRating As Double
    Dim tempUrl As String
    For i = 1 To matchCount - 1
        For j = 1 To matchCount - i
            If matchRatings(j) < matchRatings(j + 1) Then
                tempName = matchNames(j)
                tempRating = matchRatings(j)
                tempUrl = matchUrls(j)
                matchNames(j) = matchNames(j + 1)
                matchRatings(j) = matchRatings(j + 1)
                matchUrls(j) = matchUrls(j + 1)
                matchNames(j + 1) = tempName
                matchRatings(j + 1) = tempRating
                matchUrls(j + 1) = tempUrl
            End If
        Next j
    Next i
    
    ' ── 輸出推薦結果（最多 15 筆）──
    Dim maxOutput As Long
    If matchCount > 15 Then
        maxOutput = 15
    Else
        maxOutput = matchCount
    End If
    
    For i = 1 To maxOutput
        wsDash.Cells(10 + i, 2).Value = i
        wsDash.Cells(10 + i, 3).Value = matchNames(i)
        ' Google Maps 超連結
        If matchUrls(i) <> "" Then
            wsDash.Hyperlinks.Add _
                Anchor:=wsDash.Cells(10 + i, 3), _
                Address:=matchUrls(i), _
                TextToDisplay:=matchNames(i), _
                ScreenTip:="點擊開啟 Google Maps"
        End If
    Next i
    
    ' ── 星級條件式格式化 ──
    設定星級格式 wsDash, maxOutput
    
    ' ── 統計摘要 ──
    輸出統計摘要 wsDash, matchCount, matchRatings, maxOutput
    
    ' ── 更新圖表 ──
    更新圖表 maxOutput
    
    MsgBox "搜尋完成！" & vbCrLf & _
           "共找到 " & matchCount & " 間符合條件的咖啡廳，" & vbCrLf & _
           "已顯示評分最高的前 " & maxOutput & " 間。", _
           vbInformation, "搜尋結果"
End Sub

Sub 設定星級格式(wsDash As Worksheet, dataCount As Long)
    Dim rng As Range
    Set rng = wsDash.Range("I11:I" & 10 + dataCount)
    rng.FormatConditions.Delete
    
    ' 4.7 以上：綠色
    Dim fc1 As FormatCondition
    Set fc1 = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreaterEqual, Formula1:="4.7")
    fc1.Interior.Color = RGB(198, 239, 206)
    fc1.Font.Color = RGB(0, 97, 0)
    fc1.Font.Bold = True
    
    ' 4.4 ~ 4.6：黃色
    Dim fc2 As FormatCondition
    Set fc2 = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlBetween, Formula1:="4.4", Formula2:="4.6")
    fc2.Interior.Color = RGB(255, 235, 156)
    fc2.Font.Color = RGB(156, 101, 0)
    fc2.Font.Bold = True
    
    ' 4.3 以下：橘色
    Dim fc3 As FormatCondition
    Set fc3 = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLessEqual, Formula1:="4.3")
    fc3.Interior.Color = RGB(255, 199, 142)
    fc3.Font.Color = RGB(156, 56, 0)
    fc3.Font.Bold = True
End Sub

Sub 輸出統計摘要(wsDash As Worksheet, totalCount As Long, matchRatings() As Double, maxOutput As Long)
    Dim sumRating As Double
    Dim maxRating As Double
    Dim minRating As Double
    Dim i As Long
    
    sumRating = 0
    maxRating = 0
    minRating = 5
    
    For i = 1 To totalCount
        sumRating = sumRating + matchRatings(i)
        If matchRatings(i) > maxRating Then maxRating = matchRatings(i)
        If matchRatings(i) < minRating Then minRating = matchRatings(i)
    Next i
    
    Dim avgRating As Double
    avgRating = Round(sumRating / totalCount, 2)
    
    ' 輸出到結果下方 (第 27 列)
    wsDash.Range("B27").Value = "搜尋統計摘要"
    wsDash.Range("B27").Font.Bold = True
    wsDash.Range("B27").Font.Size = 12
    
    wsDash.Range("B28").Value = "符合條件："
    wsDash.Range("C28").Value = totalCount & " 間"
    wsDash.Range("D28").Value = "平均評分："
    wsDash.Range("E28").Value = avgRating
    wsDash.Range("F28").Value = "最高評分："
    wsDash.Range("G28").Value = maxRating
    wsDash.Range("H28").Value = "最低評分："
    wsDash.Range("I28").Value = minRating
    
    ' 設定格式
    Dim c As Range
    For Each c In wsDash.Range("B28:I28")
        c.Font.Name = "微軟正黑體"
        c.Font.Size = 11
    Next c
    wsDash.Range("B28").Font.Bold = True
    wsDash.Range("D28").Font.Bold = True
    wsDash.Range("F28").Font.Bold = True
    wsDash.Range("H28").Font.Bold = True
End Sub

Sub 更新圖表(dataCount As Long)
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("儀表板")
    
    If wsDash.ChartObjects.Count = 0 Then Exit Sub
    
    Dim cht As ChartObject
    Set cht = wsDash.ChartObjects(1)

    If dataCount = 0 Then
        With cht.Chart
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            .HasTitle = True
            .ChartTitle.Text = "推薦咖啡廳 Google 星級比較"
        End With
        Exit Sub
    End If
    
    Dim catRange As Range
    Dim valRange As Range
    Set catRange = wsDash.Range("C11:C" & 10 + dataCount)
    Set valRange = wsDash.Range("I11:I" & 10 + dataCount)
    
    With cht.Chart
        .SetSourceData Source:=valRange
        If .SeriesCollection.Count > 0 Then
            .SeriesCollection(1).XValues = catRange
            .SeriesCollection(1).Name = "Google 星級"
        End If
        .ChartTitle.Text = "推薦咖啡廳 Google 星級比較（前 " & dataCount & " 名）"
    End With
End Sub

Sub 清除結果()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("儀表板")
    wsDash.Range("B11:C25").ClearContents
    wsDash.Range("B27:I28").ClearContents
    wsDash.Range("I11:I25").FormatConditions.Delete
    ' 清除超連結
    Dim hl As Hyperlink
    For Each hl In wsDash.Hyperlinks
        If Not Intersect(hl.Range, wsDash.Range("C11:C25")) Is Nothing Then
            hl.Delete
        End If
    Next hl
    更新圖表 0
    MsgBox "已清除所有搜尋結果！", vbInformation, "清除完成"
End Sub

Sub 顯示全部咖啡廳()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets("儀表板")
    
    wsDash.Range("D5").Value = "不限"
    wsDash.Range("D6").Value = "不限"
    wsDash.Range("D7").Value = "不拘"
    wsDash.Range("D8").Value = "不限"
    
    搜尋咖啡廳
End Sub
