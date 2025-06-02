Sub 評価アラートチェック()
    Dim wsData As Worksheet, wsSearch As Worksheet, wsEval As Worksheet
    Dim dataArr As Variant, resultArr() As Variant
    Dim targetID As String
    Dim lastRow As Long, evalLastRow As Long
    Dim dicts() As Object
    Dim i As Long, j As Long
    Dim colIdx1 As Long, colIdx2 As Long
    Dim header As Object
    Dim key As String
    Dim countDict As Object
    Dim item1 As String, item2 As String
    Dim alertMsg As String
    
    Set wsData = ThisWorkbook.Sheets("データマスター")
    Set wsSearch = ThisWorkbook.Sheets("検索")
    Set wsEval = ThisWorkbook.Sheets("組合せ評価")
    
    ' 評価対象の資料番号
    targetID = wsSearch.Range("F2").Value
    
    ' データマスター配列に格納（ヘッダ含む）
    With wsData
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        dataArr = .Range("A1", .Cells(lastRow, .Cells(1, .Columns.Count).End(xlToLeft).Column)).Value
    End With
    
    ' ヘッダの列名と列番号をマッピング
    Set header = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(dataArr, 2)
        header(dataArr(1, j)) = j
    Next j
    
    ' 組合せ評価 最終行取得
    evalLastRow = wsEval.Cells(wsEval.Rows.Count, 1).End(xlUp).Row
    
    ' 組合せごとのループ
    For i = 2 To evalLastRow
        item1 = wsEval.Cells(i, 1).Value ' 仕様項目_1
        item2 = wsEval.Cells(i, 2).Value ' 仕様項目_2
        threshold = wsEval.Cells(i, 3).Value
        
        ' 対象列のインデックス取得
        If Not header.exists(item1) Or Not header.exists(item2) Then
            wsEval.Cells(i, 4).Value = "NG"
            wsEval.Cells(i, 5).Value = "仕様項目が見つかりません。"
            GoTo NextCombination
        End If
        
        colIdx1 = header(item1)
        colIdx2 = header(item2)
        
        ' 組合せカウント用辞書作成
        Set countDict = CreateObject("Scripting.Dictionary")
        
        For j = 2 To UBound(dataArr, 1)
            ' 資料番号が対象資料の場合スキップ
            If dataArr(j, 1) = targetID Then GoTo SkipRow
            
            key = dataArr(j, colIdx1) & "||" & dataArr(j, colIdx2)
            If Not countDict.exists(key) Then
                countDict(key) = 0
            End If
            countDict(key) = countDict(key) + 1
            
SkipRow:
        Next j
        
        ' 対象資料番号の仕様項目値取得
        For j = 2 To UBound(dataArr, 1)
            If dataArr(j, 1) = targetID Then
                key = dataArr(j, colIdx1) & "||" & dataArr(j, colIdx2)
                Exit For
            End If
        Next j
        
        ' 閾値比較
        If countDict.exists(key) Then
            If countDict(key) <= threshold Then
                wsEval.Cells(i, 4).Value = "NG"
                wsEval.Cells(i, 5).Value = "資料番号の仕様項目の組合せは過去実績で閾値以下の件数です。"
            Else
                wsEval.Cells(i, 4).Value = "OK"
                wsEval.Cells(i, 5).ClearContents
            End If
        Else
            ' 組合せが過去に存在しない場合もNGとする
            wsEval.Cells(i, 4).Value = "NG"
            wsEval.Cells(i, 5).Value = "資料番号の仕様項目の組合せは過去実績で閾値以下の件数です。"
        End If
        
NextCombination:
    Next i
    
    MsgBox "評価完了しました。", vbInformation
End Sub
