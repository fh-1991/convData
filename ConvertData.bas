Sub カテゴリ変換_動的配列対応()

    Dim wsData As Worksheet
    Dim wsConv As Worksheet
    Dim wsOut As Worksheet
    Dim dataArr As Variant
    Dim outputArr() As Variant
    Dim lastRow As Long, lastCol As Long
    Dim i As Long, j As Long
    Dim value As Variant
    Dim headerArr As Variant
    Dim categoryMap As Object
    Dim convRow As Long, convCol As Long
    Dim colIndex As Long
    Dim colName As String
    Dim thresholdDict As Object
    Set thresholdDict = CreateObject("Scripting.Dictionary")

    ' ワークシート設定
    Set wsData = ThisWorkbook.Sheets("データマスター")
    Set wsConv = ThisWorkbook.Sheets("カテゴリ変換")

    ' 出力シートの初期化
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("データマスター_変換後").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set wsOut = ThisWorkbook.Sheets.Add(After:=wsConv)
    wsOut.Name = "データマスター_変換後"

    ' データマスター読み込み（配列）
    lastRow = wsData.Cells(wsData.Rows.Count, 1).End(xlUp).Row
    lastCol = wsData.Cells(1, wsData.Columns.Count).End(xlToLeft).Column
    dataArr = wsData.Range(wsData.Cells(1, 1), wsData.Cells(lastRow, lastCol)).Value

    ' 出力用配列初期化
    ReDim outputArr(1 To UBound(dataArr, 1), 1 To UBound(dataArr, 2))

    ' ヘッダー抽出
    For j = 1 To UBound(dataArr, 2)
        outputArr(1, j) = dataArr(1, j)
    Next j

    ' カテゴリ変換テーブルの読み込み
    convRow = 2
    Do While wsConv.Cells(convRow, 1).Value <> ""
        colName = wsConv.Cells(convRow, 1).Value
        convCol = wsConv.Cells(convRow, wsConv.Columns.Count).End(xlToLeft).Column
        Dim thresholds() As Double
        ReDim thresholds(1 To convCol - 1)
        For j = 2 To convCol
            If IsNumeric(wsConv.Cells(convRow, j).Value) Then
                thresholds(j - 1) = CDbl(wsConv.Cells(convRow, j).Value)
            End If
        Next j
        thresholdDict(colName) = thresholds
        convRow = convRow + 1
    Loop

    ' カラム名から列番号を取得
    Dim colNameToIndex As Object
    Set colNameToIndex = CreateObject("Scripting.Dictionary")
    For j = 1 To UBound(dataArr, 2)
        colNameToIndex(dataArr(1, j)) = j
    Next j

    ' データ変換処理
    For i = 2 To UBound(dataArr, 1)
        For j = 1 To UBound(dataArr, 2)
            value = dataArr(i, j)
            colName = dataArr(1, j)

            If thresholdDict.Exists(colName) Then
                Dim cats() As Double
                cats = thresholdDict(colName)
                Dim k As Long
                Dim category As Long
                category = 1
                For k = 1 To UBound(cats)
                    If value <= cats(k) Then
                        Exit For
                    End If
                    category = category + 1
                Next k
                outputArr(i, j) = category
            Else
                outputArr(i, j) = value ' 変換対象でなければそのまま
            End If
        Next j
    Next i

    ' 出力配列を書き込み
    wsOut.Range("A1").Resize(UBound(outputArr, 1), UBound(outputArr, 2)).Value = outputArr

    MsgBox "カテゴリ変換が完了しました。", vbInformation

End Sub
