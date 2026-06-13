VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_getFileCnt 
   Caption         =   "ファイル数取得"
   ClientHeight    =   4128
   ClientLeft      =   160
   ClientTop       =   590
   ClientWidth     =   6750
   OleObjectBlob   =   "UserForm_getFileCnt.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm_getFileCnt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public topRow As Integer

'フォルダ一つ当たりのファイル数限度値（最大値）
Dim MaxLimitFileCnt As Integer '= 100

Const col_fileCnt  As Integer = 2
Const col_fileFullPath As Integer = 3
'*********************************************
'   ファイル数取得
'*********************************************
Private Sub CommandButton_getFileCnt_Click()

'Dim rootPath As String
Dim shtNameFileCnt As String
Dim fileCnt As Integer

    Application.ScreenUpdating = False
    
    initializeFileCountSheet
        
    shtNameFileCnt = "ファイル数"
    Set shtFileCnt = Sheets(shtNameFileCnt)
    rootPath = Me.TextBox_rootFolder.Value
    
    FolderTreeCol = 1
    folderFullPathCol = FolderTreeCol + 3
    
    '展開用のフォルダツリー（参照用）を取得
    Call excelIO_UDF_addFolderTree(rootPath, ActiveWorkbook.Name, shtNameFileCnt, topRow, FolderTreeCol)
    
    '展開用のフォルダツリー（実際の展開作業用）に、folderFullPathCol列へコピーする
    With ActiveSheet
        .Columns(FolderTreeCol).Select
        Selection.Copy
        .Columns(folderFullPathCol).Select
        .Paste
        'コピーして上書きされた項目名を元に戻す
        .Cells(topRow - 1, folderFullPathCol - 1) = "新フォルダ名"
        .Cells(topRow - 1, folderFullPathCol) = "フルパス表示"
        .Range("A1").Select
    End With
    Application.CutCopyMode = False
    'フォルダツリーを展開する。（展開結果をfolderFullPathCol列に上書きする）
    Call excelIO_UDF_expandFolderTree(ActiveWorkbook.Name, shtNameFileCnt, shtNameFileCnt, topRow, folderFullPathCol)
    
Dim fso As Object
Dim r As Integer
Dim FullPath As String
Dim ColorIndex As Integer
Dim LimitOver As Boolean

    Set fso = CreateObject("Scripting.FileSystemObject")
    
    With ActiveSheet
    r = topRow
    
        On Error Resume Next
        
        Do
            FullPath = .Cells(r, folderFullPathCol)
            If FullPath = "" Then Exit Do
            
            'フォルダ内のファイル数を取得
            fileCnt = fso.GetFolder(FullPath).Files.Count
            
            If Err.Number > o Then
                'MsgBox "エラー番号 : " & Err.Number
                ErrorMsg = Module_common.generateErrorMessage(Err)
                
                ret = MsgBox(ErrorMsg & vbcelf & "終了しますか？", vbYesNoCancel)
                
                Select Case ret
                    Case vbOK
                        End
                    Case vbCancel
                        Stop
                    End Select
            End If
            
            .Cells(r, Module_FolderControl.col_fileCntCol) = fileCnt
            
            '末尾が、アンダーバーのときはフォルダ名未定義として扱う
            If Right(FullPath, 1) = "_" Then
                .Cells(r, Module_FolderControl.col_newBranchName) = Module_FolderControl.UndefileFolderName
            End If
            
            'フォルダの権限を取得
            '.Cells(r, Module_FolderControl.col_fileCntCol + 1) = excelIO_UDF_getFolderPermissions(FullPath)
            
            'ファイル数チェック
            MaxLimitFileCnt = Module_Data.upperLimit_fileCnt
            ret = checkNumberOfFilesExceedsTheLowerLimit(fileCnt, ColorIndex, LimitOver)
            
            'リストボックスに最大値を超えたフォルダ名を追加する
            If LimitOver = True Then
                UserForm_getFileCnt.ListBox_LimitOverFolder.AddItem FullPath
            End If
                        
            'ファイル数に応じて、セルの背景、フォントの色を変更する
            If ret = True Then
                Application.ScreenUpdating = False
                '設定関数を実行
                .Cells(r, Module_FolderControl.col_fileCntCol).Select
                Call Module_ColorControl.setBackColorAndFontColor(ColorIndex)
                Application.ScreenUpdating = True
            End If

            'ハイパーリンクを設定する
            linkPath = .Cells(r, 3)
            Call Module_HyperLink.SetHyperLink(r, Module_FolderControl.col_fileCntCol, FullPath)
            
            r = r + 1
        Loop
    End With
    
    ' 後始末
    'Set fso = Nothing
    
    '上限値越えフォルダを表示するリストボックスを更新
    Module_Data.upperLimit_fileCnt = Val(Me.TextBox_FileCntUpperLimit.Value)
    addListBoxItem_FplderPath
    
    Application.ScreenUpdating = True
    
    DoEvents
    
    MsgBox "完了しました"
    
End Sub

'*********************************************************
'   ファイル数チェック
'*********************************************************
Private Function checkNumberOfFilesExceedsTheLowerLimit(fileCnt As Integer, ByRef ColorIndex As Integer, ByRef isLimitOver As Boolean) As Boolean

    overLowerLimit = True
    isLimitOver = False
    
    If fileCnt > MaxLimitFileCnt Then
    
        ColorIndex = Module_ColorControl.bkClr_red
        isLimitOver = True
    
    ElseIf fileCnt > 90 Then
    
        ColorIndex = Module_ColorControl.bkClr_redPurpl
    
    ElseIf fileCnt > 80 Then
    
        ColorIndex = Module_ColorControl.bkClr_org
    
    ElseIf fileCnt > 70 Then
    
        ColorIndex = Module_ColorControl.bkClr_yew
        
    Else
    
        overLowerLimit = False
    
    End If
    
    checkNumberOfFilesExceedsTheLowerLimit = overLowerLimit

End Function

'*********************************************
'   フォルダフルパス一覧取得    実験用
'*********************************************
Private Sub getAllFolerPath(root As Variant, shtName As String, writeRow As Integer, writetCol As Integer)

Dim fso As Object
Dim folder As Object
Dim row As Integer


Set fso = CreateObject("Scripting.FileSystemObject")
 
    For Each folder In fso.GetFolder(root).SubFolders
    
        Sheets(shtName).Cells(writeRow, writetCol) = folder
        row = writeRow + 1
        Call getAllFolerPath(folder.Path, shtName, row, writetCol)
    Next

End Sub

'*********************************************
'   フォルダ取得
'*********************************************
Private Sub CommandButton_SelFolder_Click()

    With Application.FileDialog(msoFileDialogFolderPicker)
        
        .InitialFileName = "D:\旅日記\2024\"
         
        If .Show = True Then
            Me.TextBox_rootFolder.Text = .SelectedItems(1)
        End If
    
        
    End With
    
    Me.ListBox_LimitOverFolder.Clear

End Sub
'*********************************************
'          閉じる
'*********************************************
Private Sub CommandButton1_Click()
    
    Unload Me

End Sub

'**********************************************
'   リストボックスのアイテムをダブルクリック
'**********************************************
Private Sub ListBox_LimitOverFolder_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    ListItemValue = Me.ListBox_LimitOverFolder.Value
    
    
    'リストボックスで選択されたアイテム（フォルダ）の表示セルにフォーカスを移動する
    row = 5
    isMatch = False
    '選択されたフォルダ名が表示されたシートの位置（セル）を検索する
    With Sheets("ファイル数")
        Do
            If .Cells(row, col_fileFullPath) = ListItemValue Then
                isMatch = True
                Exit Do
            End If
            
            row = row + 1
            
            If .Cells(row, col_fileFullPath) = "" Then Exit Do
        Loop
    End With
    
    If isMatch = True Then
        Cells(row, 1).Select
    Else
        Cells(1, 1).Select
    End If

    'リストボックスで選択されたアイテム（フォルダ）をエクスプローラーで開く
    FullPath = Me.TextBox_rootFolder & ListItemValue
    Shell "C:\Windows\Explorer.exe " & FullPath, vbNormalFocus

End Sub
'*************************************
'　先頭フォルダ選択オプションボタン
'*************************************
Private Sub OptionButton_LatestFolder_Click()

    seaarTop = "D:\旅日記\2024"
    retFolder = excelIO_UDF_getLatestFolder(seaarTop)

    Me.TextBox_rootFolder = seaarTop & "\" & retFolder

End Sub
'*************************************
'　先頭フォルダ選択オプションボタン
'*************************************
Private Sub OptionButton_worksheet_Click()

    Me.TextBox_rootFolder.Text = Sheets("ファイル数").Range("A5")

End Sub

'*************************************
'   ユーザーフォームの状態設定
'*************************************
Private Sub UserForm_Click()

    Module_Data.currentUserForm = StateUserForm.getFileCnt

End Sub
'*************************************
'       結果保存シートの初期化
'*************************************
Sub initializeFileCountSheet()

    Sheets("ファイル数").Activate
    
    With ActiveSheet
    
        .Cells.Select
        Selection.ClearContents
        
        'セルの背景とフォント色を初期化
        Module_ColorControl.initializeBackColorAndFontColor
        
        .Range("A1").Select
        
        itemNameRow = topRow - 1
        .Cells(itemNameRow, 1) = "フォルダツリー"
        .Cells(itemNameRow, 2) = "ファイル数"
        .Cells(itemNameRow, 3) = "フルパス表示"
    
    End With
    
    Rows("4:4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092390
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A4").Select
    
End Sub

'***************************************************
'　リストボックスにアイテム（フォルダパス）を追加
'***************************************************
Private Sub addListBoxItem_FplderPath()

Dim row As Integer

    Me.ListBox_LimitOverFolder.Clear
    
    commonFolder = Me.TextBox_rootFolder
    cmnFlderLen = Len(commonFolder)
    
    row = 5
    
    'MethodUsingVBA (row)
    
    MethodUsingPython (row)

End Sub
'********************************************
'　pythonによるリストボックスへのアイテム追加
'********************************************
Private Sub MethodUsingPython(topRow As Integer)

Dim row As Integer

    row = topRow

    dfLength = Module_interfacePython.generateFileCntDataFrame("ファイル数", row)
    
    If dfLength = 0 Then Exit Sub

    rowListBox = 0
    Do
        With Me.ListBox_LimitOverFolder
                    
            '.AddItem cntOverFolderName
            .AddItem ""
            .List(rowListBox, 0) = excelIO_UDF_df_fileCntByFolderItem(rowListBox, 0)
            .List(rowListBox, 1) = excelIO_UDF_df_fileCntByFolderItem(rowListBox, 1)
            rowListBox = rowListBox + 1
            
            If rowListBox >= dfLength Then Exit Do
            If rowListBox > 1000 Then Exit Do  '暴走防止
        End With
    Loop
        
End Sub

Private Sub MethodUsingPython_(topRow As Integer)

    row = topRow

    rowListBox = 0
    With Sheets("ファイル数")
        Do
            fileCnt = .Cells(row, col_fileCnt)
            If fileCnt >= MaxLimitFileCnt Then
                With Me.ListBox_LimitOverFolder
                    
                    cntOverFolderName = Sheets("ファイル数").Cells(row, col_fileFullPath)
                    
                    If InStr(1, cntOverFolderName, commonFolder) > 0 Then
                        cntOverFolderName = Right(cntOverFolderName, Len(cntOverFolderName) - cmnFlderLen)
                    End If
                    
                    '.AddItem cntOverFolderName
                    .AddItem ""
                    .List(rowListBox, 0) = excelIO_UDF_df_fileCntByFolderItem(row - 5, 0)
                    .List(rowListBox, 1) = excelIO_UDF_df_fileCntByFolderItem(row - 5, 1)
                    rowListBox = rowListBox + 1
                    
                End With
            End If
            
            row = row + 1
            
            If .Cells(row, col_fileCnt) = "" Then Exit Do
            If row > 1000 Then Exit Do  '暴走防止
        Loop
                
        'Me.ListBox_LimitOverFolder.ColumnWidths = Me.ListBox_LimitOverFolder.Width - 5
        
    End With
    
End Sub

'********************************************
'　VBAによるリストボックスへのアイテム追加
'********************************************
Private Sub MethodUsingVBA(topRow As Integer)

    row = topRow

    rowListBox = 0
    With Sheets("ファイル数")
        Do
            fileCnt = .Cells(row, col_fileCnt)
            If fileCnt >= MaxLimitFileCnt Then
                With Me.ListBox_LimitOverFolder
                    
                    cntOverFolderName = Sheets("ファイル数").Cells(row, col_fileFullPath)
                    
                    If InStr(1, cntOverFolderName, commonFolder) > 0 Then
                        cntOverFolderName = Right(cntOverFolderName, Len(cntOverFolderName) - cmnFlderLen)
                    End If
                    
                    '.AddItem cntOverFolderName
                    .AddItem ""
                    .List(rowListBox, 0) = Sheets("ファイル数").Cells(row, col_fileFullPath)
                    .List(rowListBox, 1) = Sheets("ファイル数").Cells(row, col_fileCnt)
                    rowListBox = rowListBox + 1
                    
                End With
            End If
            
            row = row + 1
            
            If .Cells(row, col_fileCnt) = "" Then Exit Do
            If row > 1000 Then Exit Do  '暴走防止
        Loop
                
        'Me.ListBox_LimitOverFolder.ColumnWidths = Me.ListBox_LimitOverFolder.Width - 5
        
    End With
    
End Sub

'*************************************
'       初期化
'*************************************
Private Sub UserForm_Initialize()

    Dim retFolder As Variant

    'ユーザーフォームの状態設定
    UserForm_Click
    
    Me.Height = 240
    Me.Width = 350
    
    topRow = 5
    
    Do While (Me.ListBox_LimitOverFolder.ListCount > 0)
        Me.ListBox_LimitOverFolder.RemoveItem 0
    Loop
    
    '先頭フォルダの選択
    Me.OptionButton_worksheet.Value = True
    'Me.OptionButton_LatestFolder.Value = True
    
    '価未設定のときは、100にする
    If Module_Data.upperLimit_fileCnt < 1 Then Module_Data.upperLimit_fileCnt = 100
    
    Me.TextBox_FileCntUpperLimit.Value = Module_Data.upperLimit_fileCnt

    'Me.TextBox_RootFolder = "D:\旅日記\2023\0429～0507_木次線、芸備線"
    
    row = 5
    
    Set shtFileCnt = Sheets("ファイル数")
        
    With Me.ListBox_LimitOverFolder
        .ColumnWidths = .Width - 10
    End With
    
    addListBoxItem_FplderPath
    
End Sub

