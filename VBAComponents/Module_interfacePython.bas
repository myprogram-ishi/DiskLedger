Attribute VB_Name = "Module_interfacePython"
'*****************************************************************************
'   ユーザーフォームのキャプションを書き換える
'
'   userformIndex : ユーザーフォームを指定する必要がある場合使用する。
'*****************************************************************************
Public Sub updateUserformCaption(userformIndex As Integer, message As String)

    If Module_Data.doShow_UserForm_Top = True Then     '現状未使用
        UserForm_Top.Label_Message.Caption = message
        DoEvents
    End If

End Sub

'**************************************************
'   引数で指定されたシートの列の最後の行を返却
'**************************************************
Public Function getEndOfRowCount(shtName As String, col As Integer) As Integer

    getEndOfRowCount = Sheets(shtName).Cells(Rows.Count, col).End(xlUp).row

End Function

'**********************************************************
'   引数で指定されたシートの列の最後の行のデータを追加
'**********************************************************
Public Sub getAppendDataToLasRow(shtName As String, col As Integer, setVal As Variant)

    row = Sheets(shtName).Cells(Rows.Count, col).End(xlUp).row
    row = row + 1
    
    Sheets(shtName).Cells(row, col).Value = setVal


End Sub
'***********************************************************************
'   ファイル数表示のための、フォルダフルパス表事行のオフセット設定
'   UserForm_getFileCntのときのみ値を返す。それ以外はゼロ。
'***********************************************************************
Public Function getFullpathWriteStartRowForFileCount() As Integer

    If Module_Data.currentUserForm = StateUserForm.getFileCnt Then
        getFullpathWriteStartRowForFileCount = UserForm_getFileCnt.topRow
    Else
        getFullpathWriteStartRowForFileCount = 0
    End If

End Function
'**********************************************************
'   引数で指定されたセルの背景色、フォント色を変更する
'**********************************************************
Public Sub setCellInterior(shtName As String, r As Integer, c As Integer, colorPtrn As Integer)

    Dim selClrPtrn As Integer

    Sheets(shtName).Cells(r, c).Select

    Select Case colorPtrn
    
        Case 0
        
            selClrPtrn = Module_ColorControl.bkClr_red
        
        Case 1
        
            selClrPtrn = Module_ColorControl.bkClr_blue
        
        End Select

    Module_ColorControl.setBackColorAndFontColor (selClrPtrn)

End Sub

'**********************************************************
'   フォルダごとのファイル数一覧のデータフレームを作る
'**********************************************************
Public Function generateFileCntDataFrame(shtName As String, topRow As Integer)

Dim retFRomPython As Variant

    Set shtFileCnt = Sheets(shtName)
    
    endRow = shtFileCnt.Cells(Rows.Count, 1).End(xlUp).row
    rowCnt = endRow - topRow + 1
    
    retFRomPython = excelIO_UDF_generateFileCntByFolderList(shtName, topRow, rowCnt, Module_Data.upperLimit_fileCnt)
   
    
    If retFRomPython <> 0 Then
    
        Select Case (VarType(retFRomPython))
        
            Case vbString
            
                MsgBox retFRomPython
                Stop
                returnVal = 0
                
            Case vbInteger
                
                returnVal = dfLength
                
            Case Else
            
                'MsgBox "typeVal : " + CStr(VarType(retFRomPython))
                
                ret = MsgBox("typeVal : " & CStr(VarType(retFRomPython)) & vbCrLf _
                                & "価：" & CStr(retFRomPython), _
                                vbOKCancel, "generateFileCntDataFrame")
                
                If ret = vbCnancel Then Stop
        
        End Select
    
    End If
    
    generateFileCntDataFrame = dfLength

End Function

'**********************************************************
'   検索用のキーワードを取得する
'**********************************************************
Public Function getSearchKeyword() As String

    getSearchKeyword = UserForm_seatchFolder.TextBox_keyWord.Value

End Function

'**********************************************************
'
'**********************************************************
Public Sub writeSearchResults_to_workSheet(resultData As Variant, targetSheetName As String)

Dim resultSheetName As String

    resultSheetName = UserForm_seatchFolder.ComboBox_shtSearchResults.Value
    Set resultSheet = Sheets(resultSheetName)
    
    Set targetSHeet = Sheets(targetSheetName)

    startRow = 10
    
    If resultSheet.Cells(startRow, 1) = "" Then
        StartCol = 1    '最初の１回目（初期値）
    Else
        StartCol = resultSheet.Cells(startRow, Columns.Count).End(xlToLeft).column + 1
    End If

    n = ActiveSheet.Name

    For i = 0 To UBound(resultData)
    
        On Error Resume Next
        
        currRow = startRow + i
        
        Set resultSHeet_cells = resultSheet.Cells(startRow + i, StartCol)
        Set targetSHeet_cells = targetSHeet.Cells(CInt(resultData(i)), StartCol)
        
        resultSHeet_cells_adrs = resultSHeet_cells.Address()
        targetSHeet_cells_adrs = targetSHeet_cells.Address()
        
        resultSHeet_cells.Value = targetSHeet_cells.Value
            
        If targetSHeet_cells_adrs <> Empty Then
        
            destHyperLink = "'" & targetSheetName & "'" & "!" & targetSHeet_cells_adrs
        
            ActiveSheet.Hyperlinks.Add Anchor:=ActiveSheet.Cells(currRow, StartCol), _
                Address:="", SubAddress:=destHyperLink
        End If
        
        If Err.Number > 0 Then
            resultSheet.Cells(currRow, StartCol).Value = resultData(i)
        End If
        
    Next i

End Sub

'**********************************************************
'
'**********************************************************
Public Function getSearchRootFolder()

    getSearchStartFolder = UserForm_Top.TextBox_searchRootFolder.Value

End Function
