VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_addFolderTree 
   Caption         =   "フォルダツリー追加"
   ClientHeight    =   4176
   ClientLeft      =   80
   ClientTop       =   500
   ClientWidth     =   6640
   OleObjectBlob   =   "UserForm_addFolderTree.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm_addFolderTree"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************
'　フォルダツリーを追加するシート選択
'***************************************
Private Sub selectSheetAddFolderTree()

    'フォルダを追加するシートを選択する
    shtName = Me.ComboBox_addFolderTreeSht.Value
    ActiveWorkbook.Sheets(shtName).Activate
    
    '追加する列番号を取得（最終列の次）
    addCol = Sheets(shtName).Cells(Module_Data.Row_FolderTreeTop, Columns.Count).End(xlToLeft).column + 1
    ActiveSheet.Cells(1, addCol).Select

End Sub

'***************************************
'       シート選択コンボボックス
'***************************************
Private Sub ComboBox_addFolderTreeSht_Change()

    'フォルダを追加するシートを選択する
    selectSheetAddFolderTree

End Sub

'***************************************
'       フォルダツリーを追加する
'***************************************
Private Sub CommandButton_addFolderTree_Click()
    
    If Me.TextBox_rootFolder <> "" Then
    
        'フォルダツリーを追加する列番号を検索
        Set addSht = Sheets(Me.ComboBox_addFolderTreeSht.Value)
        
        '追加する列番号を取得（最終列の次）
        addCol = addSht.Cells(Module_Data.Row_addFolderTreeTop, Columns.Count).End(xlToLeft).column + 1
        
        'pythonにフォルダパスを文字列として渡すために、エスケープシーケンスを２つ重ねて書く
        path_forPython = Me.TextBox_rootFolder
        'path_forPython = Replace(Me.TextBox_rootFolder, "\", "\\")
        
        Call excelIO_UDF_addFolderTree(path_forPython, ActiveWorkbook.Name, _
            Me.ComboBox_addFolderTreeSht.Value, Module_Data.Row_addFolderTreeTop, addCol)
        
    
        addSht.Activate
        
        MsgBox "終了しました"
    
    Else
    
        MsgBox "追加するフォルダが選択されていません"
    End If

End Sub
'*****************************************
'   取得するフォルダツリーの先頭を取得
'*****************************************
Private Sub CommandButton_getRoot_Click()

    With Application.FileDialog(msoFileDialogFolderPicker)

        If .Show = True Then
            UserForm_addFolderTree.TextBox_rootFolder.Text = .SelectedItems(1)
        End If
    
    End With

End Sub
'*******************
'       閉じる
'*******************
Private Sub CommandButton_close_Click()
    
    Unload Me
    
End Sub

'*************************************
'   ユーザーフォームの状態設定
'*************************************
Private Sub UserForm_Click()
    
    'ユーザーフォームの状態設定
    Module_Data.currentUserForm = StateUserForm.addFolderTree
    
End Sub

'*************************************
'       ユーザーフォーム初期化
'*************************************
Private Sub UserForm_Initialize()

    UserForm_Click
    
    addShtName = UserForm_Top.ComboBox_compareSht.Value

With Me

    .TextBox_rootFolder.Value = "E:\"

    For i = 1 To ActiveWorkbook.Sheets.Count
        
        compareShtName = ActiveWorkbook.Sheets(i).Name
        
        .ComboBox_addFolderTreeSht.AddItem compareShtName
        
        'If InStr(1, compareShtName, UserForm_Top.TextBox_keyWord.Text) Then defaultIndex = i
        If compareShtName = addShtName Then defaultIndex = i
        
    Next i

    .ComboBox_addFolderTreeSht.ListIndex = defaultIndex - 1
    
    selectSheetAddFolderTree
    
End With

End Sub

