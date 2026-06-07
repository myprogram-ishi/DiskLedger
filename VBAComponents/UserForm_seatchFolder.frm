VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_seatchFolder 
   Caption         =   "フォルダ検索"
   ClientHeight    =   5340
   ClientLeft      =   144
   ClientTop       =   568
   ClientWidth     =   7144
   OleObjectBlob   =   "UserForm_seatchFolder.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm_seatchFolder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_close_Click()

    Unload Me

End Sub

Private Sub CommandButton_end_Click()

    End

End Sub

Private Sub OptionButton_HardDisk_Click()

    setAllItemListBoxStatus (False)
    
    enableChechBox_argumentSheetName (Module_Data.shatName_base)

    'Me.ListBox_searhTargetSheet.Enabled = False
    Me.Frame_searhTargetSheet.Enabled = False

End Sub
Private Sub OptionButton_manual_Click()

     Me.ListBox_searhTargetSheet.Enabled = True
     Me.Frame_searhTargetSheet.Enabled = True
     

End Sub
'*******************************************************************
'　リストボックスについているチェックボックスを全て同じ状態にする
'*******************************************************************
Private Sub setAllItemListBoxStatus(status As Boolean)

    shtCnt = Worksheets.Count
    
    With Me.ListBox_searhTargetSheet
    
        listCnt = .ListCount
    
        For i = 0 To listCnt - 1
            .Selected(i) = status
        Next i
    
    End With

End Sub

'************************************************************
'　　表示中のシートのみ、リストボックスのチェックをつける
'************************************************************
Private Sub OptionButton_shoeSheets_Click()

    shtCnt = Worksheets.Count
    
    With Me.ListBox_searhTargetSheet
    
        For i = 0 To shtCnt - 1
            .Selected(i) = Sheets(.List(i)).Visible
        Next i
    
    End With

End Sub

'************************************************************
' 引数に指定したシートのみ、リストボックスのチェックをつける
'************************************************************
Private Sub enableChechBox_argumentSheetName(shtName As String)

    shtCnt = Worksheets.Count
    
    With Me.ListBox_searhTargetSheet
    
        For i = 0 To shtCnt - 1
            
            If Sheets(i + 1).Name = shtName Then
                .Selected(i) = True
            End If
        
        Next i
    
    End With

End Sub

'***************************************
'      検索開始
'***************************************
Private Sub CmdBtn_Start_Click()
    
Dim targetSHeet As String
Dim targetSHeets() As Variant
Dim sheetCnt As Integer
Dim oneDimArray As Variant


    Sheets(UserForm_seatchFolder.ComboBox_shtSearchResults.Value).Activate
    With ActiveSheet
        .Cells.Select
        Selection.ClearContents
        .Range("A1").Select
    End With


'-----------------------------------------
'       検索対象のシート一覧を生成
'-----------------------------------------
    With Me.ListBox_searhTargetSheet
    
        sheetCnt = 0
    
        For i = 0 To (.ListCount - 1)
        
            If .Selected(i) = True Then
                
                ReDim Preserve targetSHeets(sheetCnt)
                targetSHeets(sheetCnt) = .List(i)
                sheetCnt = sheetCnt + 1
            
            End If
        
        Next i
    
    End With
    
    '選択数が１つのときは、python側では、listとして扱われず、
    '文字データとして扱われるようなので、配列の先頭だけを引数に設定する。
    If UBound(targetSHeets) > 0 Then
        Call excelIO_UDF_search(ActiveWorkbook.Name, targetSHeets)
    Else
        targetSHeet = targetSHeets(0)
        oneDimArray = excelIO_UDF_search(ActiveWorkbook.Name, targetSHeets(0))
        'targetSHeet = targetSHeets(0)
        'Call writeSearchResultsfoWorksheet(targetSHeet, oneDimArray)
    End If
    
    MsgBox "検査終了"

End Sub

'************************************************************
'   検索結果をワークシートに記録する
'************************************************************
Private Sub writeSearchResultsfoWorksheet__(targetSheetName As String, resultData As Variant)

Dim resultSheetName As String

    resultSheetName = Me.ComboBox_shtSearchResults.Value

    startRow = 10
    StartCol = 1

    For i = 0 To UBound(resultData)
    
        On Error Resume Next
    
        Sheets(resultSheetName).Cells(startRow + i, StartCol).Value _
            = Sheets(targetSheetName).Cells(CInt(resultData(i, 1)), 1).Value

        If Err.Number > 0 Then
            Sheets(resultSheetName).Cells(startRow + i, StartCol).Value = resultData(i, 1)
        End If
        
    Next i
    
End Sub


'************************************************************
'       フォーム初期化
'************************************************************
Private Sub UserForm_Initialize()

    shtCnt = Worksheets.Count
    
    With Me.ListBox_searhTargetSheet
    
        Me.TextBox_keyWord.Value = "桜"
    
        For i = 0 To shtCnt - 1
            .AddItem Sheets(i + 1).Name
            .Selected(i) = False
        Next i
    
    End With
    
    For i = 1 To Sheets.Count
        Me.ComboBox_shtSearchResults.AddItem Sheets(i).Name
        
        If Sheets(i).Name = "検索結果" Then default_Index = i - 1
        
    Next i
    
    Me.ComboBox_shtSearchResults.ListIndex = default_Index
    
    Me.OptionButton_HardDisk.Value = True

End Sub
