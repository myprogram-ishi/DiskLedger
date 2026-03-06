Attribute VB_Name = "Module_common"
Public Sub testRunPython()



End Sub

'*******************************
'
'*******************************
Public Sub getTopAndEndRow(shName As String, topRow As Integer, topCol As Integer)

    Set targetSHeet = Sheets(shName)
    
    With targetSHeet
        
        startRow = .Cells(1, topCol).End(xlDown)
        endRow = .Cells(Rows.Count, topCol).End(xlUp)
    
    End With

End Sub

'*******************************
'   結果シートの初期化
'*******************************
Public Sub initializeResultSheet()

    resultSheetName = UserForm_Top.ComboBox_resultCompareSht.Value
    
    Set resultSheet = Sheets(resultSheetName)
        
    resultSheet.Activate
    ActiveSheet.Cells.Select
    Selection.ClearContents
    resultSheet.Range("A1").Select
    
    result_top_row = 1
    With resultSheet
        
        .Cells(result_top_row, 1) = "比較シート"
        .Cells(result_top_row, 2) = UserForm_Top.ComboBox_compareSht.Value      '比較シート名は、初期化の時点で書き込む
        
        result_top_row = result_top_row + 1
        .Cells(result_top_row, 1) = "フォルダ数 hardDisk"
        result_top_row = result_top_row + 1
        .Cells(result_top_row, 1) = "フォルダ数 DVD"
        result_top_row = result_top_row + 1
        .Cells(result_top_row, 1) = "〇"
        result_top_row = result_top_row + 1
        .Cells(result_top_row, 1) = "×"
    End With

End Sub

'**************************************************************
'          シート一覧
'　シート（メイン）に、このブックにあるシート名を一覧表示する
'**************************************************************
Public Sub geberateWorkSheetList()

Set shtListSht = Sheets("メイン")


    topRow = 10
    col = 1

    With shtListSht
    
        .Columns(1).Clear
        
        .Cells(topRow - 1, col) = "シート一覧"
        .Cells(topRow - 1, col).Font.Bold = True
        .Cells(topRow - 1, col).HorizontalAlignment = xlCenter
    
    
        For index = 0 To (Worksheets.Count - 1)
        
            Set targetSHeet = ActiveWorkbook.Sheets(index + 1)
            Set sheetNaneCell = .Cells(topRow + index, col)
    
            sheetNaneCell.Value = targetSHeet.Name
            
            sheetNaneCell.HorizontalAlignment = xlRight
            
            'シート名が書かれたセルの背景色を設定する。
            If targetSHeet.Visible = True Then
            '表示シートのシート名は、白
                Range(sheetNaneCell, sheetNaneCell).Interior.Color = rgbWhite
            Else
            '非表示シートのシート名は、薄い灰色
                Range(sheetNaneCell, sheetNaneCell).Interior.Color = rgbLightGray   '薄い灰色
            End If
    
        Next index
    End With

End Sub

'************************************************
'   エラーメッセージ生成
'主にメッセージボックスに表示する目的の文字列
'************************************************
Public Function generateErrorMessage(RrrObj As Object) As String

    generateErrorMessage _
        = "エラー番号 : " & RrrObj.Number & vbCrLf & RrrObj.Description

End Function





