Sub 組合せ評価処理()
    Dim wsDataMaster As Worksheet
    Dim ws検索 As Worksheet
    Dim ws組合せ評価 As Worksheet
    
    Set wsDataMaster = ThisWorkbook.Worksheets("データマスター")
    Set ws検索 = ThisWorkbook.Worksheets("検索")
    Set ws組合せ評価 = ThisWorkbook.Worksheets("組合せ評価")
    
    ' データマスターを配列に格納
    Dim dataRange As Range
    Set dataRange = wsDataMaster.UsedRange
    Dim dataMasterArray As Variant
    dataMasterArray = dataRange.Value
    
    ' 評価対象の資料番号を取得
    Dim targetDocNo As String
    targetDocNo = ws検索.Range("F2").Value
    
    ' ヘッダー行を取得（1行目）
    Dim headers As Variant
    ReDim headers(1 To UBound(dataMasterArray, 2))
    For i = 1 To UBound(dataMasterArray, 2)
        headers(i) = dataMasterArray(1, i)
    Next i
    
    ' 組合せ評価の処理
    Dim lastRow As Long
    lastRow = ws組合せ評価.Cells(ws組合せ評価.Rows.Count, 1).End(xlUp).Row
    
    For evalRow = 2 To lastRow
        Dim spec1 As String, spec2 As String
        Dim threshold As Long
        
        spec1 = ws組合せ評価.Cells(evalRow, 1).Value ' 仕様項目_1
        spec2 = ws組合せ評価.Cells(evalRow, 2).Value ' 仕様項目_2
        threshold = ws組合せ評価.Cells(evalRow, 3).Value ' 閾値
        
        ' カラムインデックスを取得
        Dim spec1ColIndex As Long, spec2ColIndex As Long
        spec1ColIndex = GetColumnIndex(headers, spec1)
        spec2ColIndex = GetColumnIndex(headers, spec2)
        
        If spec1ColIndex = 0 Or spec2ColIndex = 0 Then
            ws組合せ評価.Cells(evalRow, 4).Value = "エラー"
            ws組合せ評価.Cells(evalRow, 5).Value = "カラムが見つかりません"
            GoTo NextEvalRow
        End If
        
        ' 組合せ数カウント配列を作成（評価対象を除外）
        Dim combinationCount As Object
        Set combinationCount = CreateObject("Scripting.Dictionary")
        
        For dataRow = 2 To UBound(dataMasterArray, 1)
            ' 評価対象の資料番号を除外
            If dataMasterArray(dataRow, 1) <> targetDocNo Then
                Dim key As String
                key = dataMasterArray(dataRow, spec1ColIndex) & "|" & dataMasterArray(dataRow, spec2ColIndex)
                
                If combinationCount.Exists(key) Then
                    combinationCount(key) = combinationCount(key) + 1
                Else
                    combinationCount(key) = 1
                End If
            End If
        Next dataRow
        
        ' 評価対象の組合せを取得
        Dim targetSpec1 As String, targetSpec2 As String
        For dataRow = 2 To UBound(dataMasterArray, 1)
            If dataMasterArray(dataRow, 1) = targetDocNo Then
                targetSpec1 = dataMasterArray(dataRow, spec1ColIndex)
                targetSpec2 = dataMasterArray(dataRow, spec2ColIndex)
                Exit For
            End If
        Next dataRow
        
        ' 評価対象の組合せのカウントを取得
        Dim targetKey As String
        targetKey = targetSpec1 & "|" & targetSpec2
        
        Dim targetCount As Long
        If combinationCount.Exists(targetKey) Then
            targetCount = combinationCount(targetKey)
        Else
            targetCount = 0
        End If
        
        ' アラート判定
        If targetCount <= threshold Then
            ws組合せ評価.Cells(evalRow, 4).Value = "NG"
            ws組合せ評価.Cells(evalRow, 5).Value = "資料番号の仕様項目の組合せは過去実績で閾値以下の件数です。"
        Else
            ws組合せ評価.Cells(evalRow, 4).Value = "OK"
            ws組合せ評価.Cells(evalRow, 5).Value = ""
        End If
        
NextEvalRow:
    Next evalRow
    
    MsgBox "組合せ評価処理が完了しました。"
End Sub

' ヘッダー配列からカラムインデックスを取得する関数
Function GetColumnIndex(headers As Variant, columnName As String) As Long
    Dim i As Long
    For i = 1 To UBound(headers)
        If headers(i) = columnName Then
            GetColumnIndex = i
            Exit Function
        End If
    Next i
    GetColumnIndex = 0 ' 見つからない場合
End Function
