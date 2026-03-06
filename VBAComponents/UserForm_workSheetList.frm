VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_workSheetList 
   Caption         =   "ワークシート操作"
   ClientHeight    =   4330
   ClientLeft      =   17110
   ClientTop       =   660
   ClientWidth     =   5610
   OleObjectBlob   =   "UserForm_workSheetList.frx":0000
End
Attribute VB_Name = "UserForm_workSheetList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_Hide_Click()

    Unload Me
    
End Sub
'**************************************************
'       シートの表示切替
'**************************************************
Private Sub CommandButton_SheetFileCntActivate_Click()

    Sheets("ファイル数").Activate

End Sub
'**************************************************
'       シートの表示切替
'**************************************************
Private Sub CommandButton_SheetMainActivate_Click()

    Sheets("メイン").Activate

End Sub
'**************************************************
'
'**************************************************
Private Sub ListBox_workSheetList_AfterUpdate()

    ListItemValue = Me.ListBox_workSheetList.Value
    
End Sub
'********************************************************************
'   チェックボックスの状態により、シートの表示／非表示を切り替える
'********************************************************************
Private Sub ListBox_workSheetList_Change()

    With Me.ListBox_workSheetList
        
        ListItemIndex = .ListIndex
        ListItemValue = .List(ListItemIndex, 0)
        
        Sheets(ListItemValue).Visible = .Selected(ListItemIndex)
    
    End With
    
End Sub
'**************************************************
'
'**************************************************
Private Sub ListBox_workSheetList_Click()

 ListItemValue = Me.ListBox_workSheetList.Value

End Sub

Private Sub ListBox_workSheetList_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

 ListItemValue = Me.ListBox_workSheetList.Value
 
End Sub

Private Sub UserForm_Initialize()

    shtCnt = Worksheets.Count
    
    With Me.ListBox_workSheetList
    
        For i = 0 To shtCnt - 1
            .AddItem Sheets(i + 1).Name
            '.List(i, 0) = 1
            '.List(i, 0) = Sheets(i + 1).Name
            .Selected(i) = Sheets(i + 1).Visible
            
        Next i
    
    End With

End Sub
