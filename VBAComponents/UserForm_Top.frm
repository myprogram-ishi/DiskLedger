VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_Top 
   Caption         =   "メイン"
   ClientHeight    =   6924
   ClientLeft      =   264
   ClientTop       =   1240
   ClientWidth     =   11776
   OleObjectBlob   =   "UserForm_Top.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "UserForm_Top"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const currYear As String = "2024"
Const resultMeans As String = "result"

'********************************************
'   フォルダツリーへハイパーリンクをつける
'********************************************
Private Sub CmdBtn_addHYperLink_Click()

Dim FolderTreeCol As Integer

    '展開したフォルダdツリーをワークシートに出さない
    excelIO_UDF_set_expandFolderTree_status (False)
     
    FolderTreeCol = Me.ComboBox_folderTreeTop_for_addHypLink.ListIndex + 1
    
    'フォルダツリーを展開して、ハイパーリンクをつける
    Module_FolderControl.expandFolderTree_on_base (FolderTreeCol)
    
    MsgBox "終了しました"
    
End Sub

'******************************************
'   サクラエディタでフォルダツリーを開く
'******************************************
Private Sub CmdBtn_showFolderTree_with_sakura_Click()

    Dim TopFolderName As String
    
    TopFolderName = Me.TextBox_TopFolder_showSakura.Value

    UserForm_Top.Label_Message.Caption = TopFolderName & "以下、フォルダツリーを展開を開始します"
    DoEvents

    Call excelIO_UDF_filrOpen_with_sakuraEditor(Me.TextBox_TopFolder_showSakura.Value, "d:\folderTree.txt")
     
    'Call excelIO_UDF_filrOpen_with_sakuraEditor("D:\git\testData\", "folderTree.txt")

    UserForm_Top.Label_Message.Caption = TopFolderName & "フォルダツリーの展開が完了しました"
    DoEvents
    
End Sub

'***********************************************
'
'***********************************************
Private Sub ComboBox_BaseFolderTreeTop_Change()

    If isRunChangeCombobox() = False Then Exit Sub

End Sub

'**********************************************************
'   基準シートのフォルダツリー一覧
'ComboBox_baseShtの値が決まった状態で実行する必要がある
'**********************************************************
Private Sub addItem_ComboBox_BaseFolderTreeTop(topRow As Integer)

Dim split_itemName As Variant

    With UserForm_Top
        
        baseShtName = .ComboBox_baseSht.Value
        
        If .ComboBox_BaseFolderTreeTop.ListCount > 0 Then .ComboBox_BaseFolderTreeTop.Clear
        
        col = 1
        Do
            'フォルダ名は、短縮形にして表示する
            itemName = Sheets(baseShtName).Cells(topRow, col)
            split_itemName = Split(itemName, "\")
            'itemName = split_itemName(0) & "\..\" & split_itemName(2) & "\" & split_itemName(3)
            itemName = split_itemName(2) & "\" & split_itemName(3)
            .ComboBox_BaseFolderTreeTop.AddItem itemName
            col = col + 1
            
            If Sheets(baseShtName).Cells(topRow, col) = "" Then Exit Do
        Loop
        
        '最期のアイテムに、「全ての列」を意味する項目を追加
        .ComboBox_BaseFolderTreeTop.AddItem "All Colu,ms"
        
        .ComboBox_BaseFolderTreeTop.ListIndex = 0
        
    End With

End Sub

Private Sub addItem_ComboBox_folderTreeTop_for_addHypLink(topRow As Integer)

    With UserForm_Top
        
        baseShtName = .ComboBox_baseSht.Value
        
        
        If .ComboBox_folderTreeTop_for_addHypLink.ListCount > 0 Then .ComboBox_folderTreeTop_for_addHypLink.Clear
        
        col = 1
        Do
            .ComboBox_folderTreeTop_for_addHypLink.AddItem Sheets(baseShtName).Cells(topRow, col)
            col = col + 1
            
            If Sheets(baseShtName).Cells(topRow, col) = "" Then Exit Do
        Loop
        
        '最期のアイテムに、「全ての列」を意味する項目を追加
        .ComboBox_folderTreeTop_for_addHypLink.AddItem "All Colu,ms"
        
        .ComboBox_folderTreeTop_for_addHypLink.ListIndex = 0
        
    End With

End Sub
'*************************************
'   コンボボックスチェンジイベント
'   比較されるシート選択
'*************************************
Private Sub ComboBox_baseSht_Change()

    If isRunChangeCombobox() = False Then Exit Sub
    
    addItem_ComboBox_BaseFolderTreeTop (Module_Data.baseSht_TreeTopRow)
    
End Sub
'*************************************
'   コンボボックスチェンジイベント
'   比較するシートの選択
'*************************************
Private Sub ComboBox_compareSht_Change()

    If isRunChangeCombobox() = False Then Exit Sub
    
End Sub
'*************************************
'   コンボボックスチェンジイベント
'   比結果のシートの選択
'*************************************
Private Sub ComboBox_resultCompareSht_Change()

    If isRunChangeCombobox() = False Then Exit Sub

End Sub
'***********************************************************************
'  コンボボックスのチェンジイベントを実行するかどうかを返却値で返す
'  シート選択の組み合わせコンボボックスの表示
'***********************************************************************
Private Function isRunChangeCombobox() As Boolean

    If Module_Data.is_UserForm_Top_Initialize = False Then
        isRunChangeCombobox = False
        Exit Function
    End If
    
    'シートの選択が変更された場合は、組み合わせのコンボボックスの組み合わせ名は非表示
    If Module_Data.is_CombinationValid = True Then
        Me.ComboBox_sheetCombination.ListIndex = 0
    End If

    isRunChangeCombobox = True

End Function

Private Sub ComboBox_romDisk_Change()

    getCombobocItem

End Sub

'*************************************
'　　使用するシート組み合わせ選択
'*************************************
Private Sub ComboBox_sheetCombination_Change()

    If Me.ComboBox_sheetCombination.ListIndex = 0 Then Exit Sub

    Call Module_Data.selectSheetNameInCombobox(Me.ComboBox_sheetCombination.ListIndex)

End Sub

'*************************************
'
'*************************************
Private Sub CommandButton_addFolderTree_Click()
    
    UserForm_addFolderTree.Show (0)
    
End Sub

'*************************************
'
'*************************************
Private Sub CommandButton_end_Click()

    excelIO_UDF__loggerFinalize

    End

End Sub

'*********************************************
'
'*********************************************
Private Sub CommandButton_getFileCnt_Click()

    UserForm_getFileCnt.Show (0)

End Sub
'*********************************************
'   比較結果ファイル出力フォルダを取得する
'*********************************************
Private Sub CommandButton_GetFolder_Click()

    With Application.FileDialog(msoFileDialogFolderPicker)
        If .Show = True Then
            UserForm_Top.TextBox_saveFolderName.Text = .SelectedItems(1)
        End If
    End With

End Sub
'*********************************************
'          ハイパーリンクをつける
'*********************************************
Private Sub CommandButton_HyperLink_Click()

Dim get_listTopFolderList As Variant

    sheetName = UserForm_Top.ComboBox_HyperLinkSht.Value

    get_listTopFolderList = _
        excelIO_UDF_getDestHyperLinkRow(ActiveWorkbook.Name, sheetName _
        , UserForm_Top.TextBox_HyperLinkDestFolder.Value)
    
    
    UserForm_Top.Label_Message.Caption = get_listTopFolderList
    
    Sheets(sheetName).Activate
    
    MsgBox "ハイパーリンク付与完了r"

End Sub
'*********************************************
'  結果シートをすべて新規ファイルに出力する
' シート名に"result"を含むシートを保存する
'*********************************************
Private Sub CommandButton_outputResult_Click()

    'Private Sub outputCompareResultToFile()

    '保存するファイルパスを生成
    saveFileName = UserForm_Top.TextBox_saveFileName.Text
    
    splitFileName = Split(saveFileName, ".")
    
    'ファイル名にを"result"追加する
    saveFileName = splitFileName(0) & "_result." & splitFileName(1)
    saveFileFullPath = UserForm_Top.TextBox_saveFolderName.Text & "\" & saveFileName

    sheetCount = ThisWorkbook.Sheets.Count
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    
    '保存用ワークブックを生成
    Workbooks.Add
    
    'シート名に、"result"を含むシートをコピーする
    For i = 1 To sheetCount
        If InStr(1, ThisWorkbook.Sheets(i).Name, "result") > 0 Then
            ThisWorkbook.Sheets(i).Copy After:=ActiveWorkbook.Sheets(Sheets.Count)
        End If
    Next i
    
    '名前を付けて保存する
    ActiveWorkbook.SaveAs saveFileFullPath
    
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
'*********************************************
'フルパスフォルダ名の最後のフォルダ名のみ変更
'*********************************************
Private Sub CommandButton_renameLastBranchName_Click()

    Module_FolderControl.renameEndOfFolderName (5)

End Sub

'*********************************************
'           検索
'*********************************************
Private Sub CommandButton_ssarchStart_Click()

    Application.ScreenUpdating = False
    
    'Worksheets("Expand_base").Cells.Clear
    'Worksheets("Expand_target").Cells.Clear
    Worksheets(ComboBox_resultCompareSht.Value).Cells.Clear
    
    With Sheets("debug_log")
    
        If .Visible = True Then
            .Activate
            .Cells.Select
            Selection.ClearContents
            .Cells(1, 1).Select
        End If
    
    End With

    'フォルダツリーの展開
    excelIO_UDF_set_expandFolderTree_status (True)  '展開したフォルダdツリーをワークシートに出す
    
    UserForm_Top.Label_Message.Caption = "フォルダツリーの展開を開始します"
    Module_FolderControl.expandFolderTreeExpand
    
    'フォルダツリーの比較
    UserForm_Top.Label_Message.Caption = "フォルダツリーを比較を開始します"
    Module_FolderControl.searchFullpath
    
    
    '結果をファイル出力する
    UserForm_Top.Label_Message.Caption = "比較結果をファイル出力します"
    outputCompareResultToFile
    
    '結果シートの保存
    If Me.CheckBox_copyResultSheet.Value = True Then
        
        copuResultSHeet
        
    End If
    
    Application.ScreenUpdating = True

    MsgBox "フォルダツリーの比較が完了しました"

End Sub
'*********************************************
'   比較結果シートをコピーして名前を変える
'*********************************************
Private Sub copuResultSHeet()

    resultSHtName = Me.ComboBox_resultCompareSht.Value
    diskSheetName = Me.ComboBox_compareSht.Value
    copySHeetNAme = resultSHtName & "_" & diskSheetName

    On Error Resume Next

    Sheets("result").Select
    Sheets("result").Copy After:=Sheets(resultSHtName)
    
    ActiveSheet.Name = copySHeetNAme
    If Err.Number = 1004 Then
    '同名のシートがあった場合は、既存のシートを消して、再度、リネーム
        Application.DisplayAlerts = False
        Sheets(copySHeetNAme).Delete
        Application.DisplayAlerts = True
        
        ActiveSheet.Name = copySHeetNAme
    End If
    
    Application.ScreenUpdating = True

End Sub
'*********************************************
'   比較結果をファイル出力
'*********************************************
Private Sub outputCompareResultToFile()

    '保存するファイルパスを生成
    saveFileName = UserForm_Top.TextBox_saveFileName.Text
    'saveFileFullPath = "D:\git\diff_FolderTree_pythonProject\" & saveFileName
    saveFileFullPath = UserForm_Top.TextBox_saveFolderName.Text & "\" & saveFileName

    resultSHtName = UserForm_Top.ComboBox_resultCompareSht.Value
    Set resultSheet = Sheets(resultSHtName)

    If Dir(saveFileFullPath) <> "" Then
    
        srcWorkBook = ActiveWorkbook.Name
        
    
        Workbooks.Open saveFileFullPath
        
        Workbooks(srcWorkBook).Activate
        resultSheet.Select
        resultSheet.Copy After:=Workbooks(saveFileName).Sheets(1)
        
    Else
        resultSheet.Select
       resultSheet.Copy
    End If

    '結果を保存するシートの名前
    outputResultSheetName = "Result_" & UserForm_Top.ComboBox_compareSht.Value
    
    On Error Resume Next
    ActiveSheet.Name = outputResultSheetName
    'シートの有無確認
    If Err.Number = 1004 Then
    'シートがすでに存在していた場合、既にあるものを削除して、新しくシートを作る
        Application.DisplayAlerts = False
        Sheets(outputResultSheetName).Delete
        Application.DisplayAlerts = True
        ActiveSheet.Name = outputResultSheetName
    End If
    
    '結果をファイルに保存する
    Application.DisplayAlerts = False
    ChDir "D:\git\diff_FolderTree_pythonProject"
    ActiveWorkbook.SaveAs Filename:=saveFileFullPath, FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
    ActiveWindow.Close
    Application.DisplayAlerts = True


End Sub

'**************************************************************
'   コンボボックスに指定されているフォルダツリーを最新化する
'**************************************************************
Private Sub CommandButton_UpdataBaseFolderTree_Click()

    'シート状の列番号は、コンボボックスのインデックス＋１
    treeColumn = Me.ComboBox_BaseFolderTreeTop.ListIndex + 1
    
    baseSht = Me.ComboBox_baseSht.Value
    
    UserForm_Top.Label_Message.Caption = "シート（" & baseSht & "）のフォルダツリー最新化"
    
    Set treeTopCell = Sheets(baseSht).Cells(Module_Data.baseSht_TreeTopRow, treeColumn)
    treeTopPath = treeTopCell.Value
    driveName = Left(treeTopPath, 3)
    
    BaseFolderTreeTop = Me.ComboBox_BaseFolderTreeTop
    
    'トップフォルダチェックのとき、フォルダ名にドライブ名を含めるかどうか。
'    If Me.CheckBox_DriveNameIgnore = False Then
'        '２文字目と３文字目が、":\"　である場合は先頭がドライブ名（C:\, D:\など）とみなし、文字列から省く
'        If Mid(treeTopPath, 2, 2) = ":\" Then
'            temp_treeTopPath = Mid(treeTopPath, 4)
'        Else
'            temp_treeTopPath = treeTopPath
'        End If
'
'        If Mid(BaseFolderTreeTop, 2, 2) = ":\" Then BaseFolderTreeTop = Mid(BaseFolderTreeTop, 3)
'    End If
    
    If treeTopPath <> BaseFolderTreeTop Then
        
        '基本的に、不一致になることはない。なったら、考える。
        ret = MsgBox("ドライブ名：" & driveName & " を付加してやり直しますか？" & vbCrLf _
        & "はい：終了, いいえ：終了, キャンセル：中断", vbYesNoCancel, "確認")
        
        'treeTopPath
    
        If ret = vbYes Then
            Debug.Print "treeTopPath : " & treeTopPath
            Debug.Print "ComboBox_BaseFolderTreeTop : " & BaseFolderTreeTop
           ' Stop
        ElseIf ret = vbNo Then
            End
        ElseIf ret = vbancel Then
            Stop
        Else
            Stop
        End If
        
        'MsgBox "フォルダを検索する仕組みを実装してください"
        'Stop
        'End
    End If
    
    Sheets(baseSht).Activate
    
    ActiveSheet.Columns(treeColumn).Select
    Selection.ClearContents
    treeTopCell.Select
    
    Call excelIO_UDF_addFolderTree(treeTopPath, ActiveWorkbook.Name, baseSht, Module_Data.baseSht_TreeTopRow, treeColumn)

    MsgBox "終了しました"
    UserForm_Top.Label_Message.Caption = ""

End Sub

'*************************************
'       キーワード変更イベント
'     シート名検索キーワード変更
'*************************************
Private Sub TextBox_keyWord_Change()

    getCombobocItem

    'Call getComboboxIndex(UserForm_Top.ComboBox_compareSht, "Disk")

    'Call getComboboxIndex(UserForm_Top.ComboBox_resultCompareSht, "result")

End Sub

'****************************************************************
'       コンボボックスのアイテム検索
'   条件に合ったものを検索して,インデックスを設定する
'****************************************************************
Private Sub getComboboxIndex(currCombo As Object, kinds As String)

Dim diskType As String
Dim keyWord As String

    test1 = UserForm_Top.ComboBox_compareSht.List(1)
    test2 = currCombo.List(1)
    
    With UserForm_Top
    
        diskType = .ComboBox_romDisk.Value
        keyWord = .TextBox_keyWord.Text
        
    End With

    For i = 0 To (currCombo.ListCount - 1)
    
        itemName = currCombo.List(i)
        
        ret_diskType = InStr(1, itemName, diskType)
        ret_kinds = InStr(1, itemName, kinds)
        ret_keyWord = InStr(1, itemName, keyWord)
        
        '掛け算結果が正の数だったら、すべての条件が満たされたことになる。
        If (ret_diskType * ret_kinds * ret_keyWord) > 0 Then
            currCombo.ListIndex = i
            Exit For
        End If

    Next i

End Sub

Private Sub old_()  'TextBox_keyWord_Change()

Dim diskName As String
Dim keyWord As String

    With UserForm_Top
    
        diskName = .ComboBox_romDisk.Value
        keyWord = .TextBox_keyWord.Text

        For i = 1 To ActiveWorkbook.Sheets.Count - 1
            
            itemName = .ComboBox_compareSht.List(i)
            
            If (InStr(1, itemName, diskName) > 0) Then
               
                If (InStr(1, itemName, keyWord) > 0) And (InStr(1, itemName, "result") <= 0) Then
                    .ComboBox_compareSht.ListIndex = i
                End If
                
                itemName = .ComboBox_resultCompareSht.List(i)
                
                If (InStr(1, itemName, "result") > 0) And (InStr(1, itemName, keyWord) > 0) Then
                    .ComboBox_resultCompareSht.ListIndex = i
                End If
            
            End If
            
        Next i
        
        Call Module_Data.selectcComboBox_sheetCombination(keyWord)

    End With


End Sub



'*************************************
'   ユーザーフォームの状態設定
'*************************************
Private Sub UserForm_Click()

    Module_Data.currentUserForm = StateUserForm.Top
    
End Sub

'*************************************
'       ユーザーフォーム初期化
'*************************************
Private Sub UserForm_Initialize()

    'python側の変数初期化
    excelIO_UDF_initialize (ActiveWorkbook.Name)
       
    Module_Data.upperLimit_fileCnt = 100
    
    Module_Data.is_UserForm_Top_Initialize = False  'ユーザーフォーム初期化処理中
    
    'ユーザーフォームの状態設定
    UserForm_Click


    Call Module_Data.initializeParametor
    
    Call Module_Data.initialize_ComboBox_sheetCombination

With Me

    '.TextBox_searchRootFolder = "SPB_17.2"
    '.TextBox_searchRootFolder = "AUTOSAR"
    '.TextBox_searchRootFolder = "Program Files"
    .TextBox_searchRootFolder = "旅日記"
    .TextBox_HyperLinkDestFolder = "旅日記"
    
    .TextBox_saveFolderName.Text = ThisWorkbook.Path
    .TextBox_saveFileName.Text = "比較結果.xlsx"
    
    .TextBox_TopFolder_showSakura.Value = "E:\"
    
    Module_Data.TopFolderName = .TextBox_searchRootFolder
    
    .CheckBox_ExpandBaseFolderTree.Value = False
    .CheckBox_ExpandCompFolderTree.Value = True
    
    addComboboxItm
    
    'フォルダツリーの一覧（ComboBox_baseShtの決定後に実行）
    addItem_ComboBox_BaseFolderTreeTop (Module_Data.baseSht_TreeTopRow)
    
    addItem_ComboBox_folderTreeTop_for_addHypLink (Module_Data.baseSht_TreeTopRow)
    
    '.ComboBox_compareSht.ListIndex = defaultIndex - 1
    
    '.ComboBox_resultCompareSht.ListIndex = defaultIndex - 1

    'ComboBox_resultCompareSht.ListIndex = resultSHt_index

    'テキストボックスのチェンジイベントを期待するため、すべての設定が終わった最後に実行する
    .TextBox_keyWord = currYear
End With
    
    
    Module_Data.is_UserForm_Top_Initialize = True   'ユーザーフォーム初期化処理完了

    'ブックに含まれるシート一覧を作成する
    Module_common.geberateWorkSheetList
    
    
    Module_Data.doShow_UserForm_Top = True
End Sub

'**************************************************************
'    コンボボックスのアイテム追加
'簡易的な並べ替えのためにシート名をキーワード検索をしている。
'**************************************************************
Private Sub addComboboxItm()

With Me

    'DVD Blue rayディスク一覧
    For i = 0 To UBound(Module_Data.romDiskList)
        .ComboBox_romDisk.AddItem Module_Data.romDiskList(i)
    Next i
    
    .ComboBox_romDisk.ListIndex = 0

      For i = 1 To ActiveWorkbook.Sheets.Count
        
        compareShtName = ActiveWorkbook.Sheets(i).Name
        
        ' "基準"シートのシート番号から、コンボボックスのインデックスを算出する（シートインデックスをデクリメント）
        If compareShtName = Module_Data.shatName_base Then baseShtIndex = i - 1
        
        .ComboBox_baseSht.AddItem compareShtName
        
        If InStr(1, compareShtName, "Disk") > 0 Then
            .ComboBox_compareSht.AddItem compareShtName, 0
        Else
             .ComboBox_compareSht.AddItem compareShtName
        End If
        
        
        '比較結果シート
        If InStr(1, compareShtName, resultMeans) > 0 Then
            .ComboBox_resultCompareSht.AddItem compareShtName, 0
        Else
            .ComboBox_resultCompareSht.AddItem compareShtName
        End If
        
        .ComboBox_HyperLinkSht.AddItem compareShtName
        
        If compareShtName = "result" Then
            resultSHt_index = i - 1
        End If
        
        '今年の西暦年が含まれているものを使う
        If InStr(1, compareShtName, currYear) > 0 Then
        
            If InStr(1, compareShtName, "Disk") > 0 Then
                defaultIndex_ComboBox_HyperLinkSht = i - 1
            End If
            
        End If
        
        
        If InStr(1, compareShtName, .TextBox_keyWord.Text) Then
            defaultIndex = i
        End If
        
    Next i
    
    .ComboBox_baseSht.ListIndex = baseShtIndex
    .ComboBox_HyperLinkSht.ListIndex = defaultIndex_ComboBox_HyperLinkSht
    
End With

End Sub

'*******************************************
'   コンボボックスのアイテム選択
'*******************************************
Private Sub getCombobocItem()

Dim diskKind As String
Dim keyWord As String

    diskKind = Me.ComboBox_romDisk.Value
    keyWord = Me.TextBox_keyWord.Value

'基準側（ハードディスク側）フォルダ選択コンボボックス
With Me.ComboBox_BaseFolderTreeTop
    
    For i = 0 To (.ListCount - 1)
    
        itemName = .List(i)
    
        If (InStr(1, itemName, keyWord) > 0) Then
            .ListIndex = i
            Exit For
        End If
        
    Next i
    
End With

'比較対象シート選択コンボボックス
With Me.ComboBox_compareSht
    
    For i = 0 To (.ListCount - 1)
    
        itemName = .List(i)
    
        If (InStr(1, itemName, "Disk") > 0) And (InStr(1, itemName, keyWord) > 0) And (InStr(1, itemName, diskKind) > 0) _
            And (InStr(1, itemName, resultMeans) = 0) Then
            .ListIndex = i
            Exit For
        End If
        
    Next i
    
End With

'結果シート選択コンボボックス
With Me.ComboBox_resultCompareSht
    
    For i = 0 To (.ListCount - 1)
    
        itemName = .List(i)
    
        If (InStr(1, itemName, resultMeans) > 0) And (InStr(1, itemName, keyWord) > 0) And (InStr(1, itemName, diskKind) > 0) Then
            .ListIndex = i
            Exit For
        End If
        
    Next i
    
End With

'ハイパーリンクコンボボックス
With Me.ComboBox_HyperLinkSht
    
    For i = 0 To (.ListCount - 1)
    
        itemName = .List(i)
    
        If (InStr(1, itemName, keyWord) > 0) And (InStr(1, itemName, diskKind) > 0) _
             And (InStr(1, itemName, resultMeans) = 0) Then
            .ListIndex = i
            Exit For
        End If
        
    Next i
    
End With

End Sub
'*******************************************
'       ユーザーフォームを閉じる
'*******************************************
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)

    Module_Data.doShow_UserForm_Top = False

End Sub
