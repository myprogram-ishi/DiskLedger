Attribute VB_Name = "Module_FolderControl"


'Sheets("ファイル数") で使う列番号
Public Const col_fileCntCol As Integer = 2
Public Const col_newBranchName As Integer = col_fileCntCol + 1
Public Const UndefileFolderName = "未定義"

'***********************************************************
'　比較対象の２つのシートのフォルダツリーを展開する
'***********************************************************
Public Sub expandFolderTreeExpand()

Dim col As Integer
    
    currentSheet = ActiveSheet.Name
    
    'Application.ScreenUpdating = False

    With UserForm_Top
    
        'python側のバッファ（配列）クリア
        excelIO_UDF_clear_lst_expand_base
        excelIO_UDF_clear_lst_expand_target
    
        'ハードディスク側　フォルダツリー展開
        If .CheckBox_ExpandBaseFolderTree.Value = True Then
            Worksheets("Expand_base").Cells.Clear
            col = .ComboBox_BaseFolderTreeTop.ListIndex + 1
            expandFolderTree_on_base (col)
        End If
        
        'ROMディスク側　フォルダツリー展開
        If .CheckBox_ExpandCompFolderTree.Value = True Then
            Worksheets("Expand_target").Cells.Clear
            expandFolderTree_on_OneTimeWriteDisc
        End If
    
    End With

    Application.ScreenUpdating = True

End Sub

'*************************************************************************
'   base側（ハードディスク側）のフォルダツリーを展開して、
'   ハイパーリンクをつける
'   引数：folderTreeIndex:フォルダツリーの記載されたシートの列名に相当
'*************************************************************************
Public Sub expandFolderTree_on_base(folderTreeIndex As Integer)

Dim shtTargetSheet As String
Dim shtNameExpand As String

Dim hyplink  As Hyperlink

    With UserForm_Top

        shtTargetSheet = CStr(.ComboBox_baseSht.Value)
        shtNameExpand = "Expand_base"
            
        If excelIO_UDF_get_expandFolderTree_status = True Then
        'シートへデータを書くとき、事前にクリアする
            Worksheets(shtNameExpand).Activate
            ActiveSheet.Cells.Select
            Selection.ClearContents
            ActiveSheet.Range("A1").Select
        End If
        
        If InStr(1, "All", .ComboBox_BaseFolderTreeTop.Value) > 0 Then
            Call expandFolderTree_OneSheet(shtTargetSheet, shtNameExpand, 0, 0)
        Else
            Call expandFolderTree_OneSheet(shtTargetSheet, shtNameExpand, folderTreeIndex, folderTreeIndex)
        End If
        
    End With

    'ハイパーリンクをつける
    If excelIO_UDF_get_expandFolderTree_status = True Then
        Call addHyperLink_to_foldeerTree_fullPath_on_sheets(shtTargetSheet, shtNameExpand)
    Else
        
        addHyperLink_to_foldeerTree_fullPath_on_array (shtTargetSheet)
    End If

End Sub
'*************************************************************
'   フォルダツリーにﾊｲﾊﾟｰﾘﾝｸをつける
'   リンクは、シートに展開されたフルバスを使う
'*************************************************************
Private Sub addHyperLink_to_foldeerTree_fullPath_on_array(shtTargetSheet As String)

Dim getFolderList As Variant

'セルへのハイパーリンク設定
    hpLink_r = Module_Data.baseSht_TreeTopRow + 1
    hpLink_c = UserForm_Top.ComboBox_folderTreeTop_for_addHypLink.ListIndex + 1

    getFolderList = excelIO_UDF_getExpandFolderTree

    Sheets(shtTargetSheet).Activate
    indwxUbound = UBound(getFolderList)
    For i = o To indwxUbound
    
        With Sheets(shtTargetSheet)
        On Error Resume Next
        Set hyplink = .Hyperlinks.Add(Anchor:=.Cells(hpLink_r, hpLink_c), Address:=getFolderList(indwxUbound - i))
        End With
    
        hpLink_r = hpLink_r + 1
    Next i
    
End Sub


Private Sub addHyperLink_to_foldeerTree_fullPath_on_sheets(shtTargetSheet As String, shtNameExpand As String)
'セルへのハイパーリンク設定
    hpLink_r = Module_Data.baseSht_TreeTopRow + 1
    hpLink_c = 1
    srcRow = 1
    srcCol = 1

    'Sheets(shtTargetSheet).Activate
    Do
        With Sheets(shtTargetSheet)
        Set hyplink = .Hyperlinks.Add(Anchor:=.Cells(hpLink_r, hpLink_c), _
                                              Address:=Sheets(shtNameExpand).Cells(srcRow, srcCol))
                                              ') 'TextToDisplay:="侍エンジニア"
        End With
    
        hpLink_r = hpLink_r + 1
        srcRow = srcRow + 1
        If Sheets(shtNameExpand).Cells(srcRow, srcCol) = "" Then Exit Do
    Loop
    
End Sub


'*************************************************************
'   DVD, blue ray側（一回書き込みディスク側）のﾌｫﾙﾀﾞﾂﾘｰ展開
'*************************************************************
Private Sub expandFolderTree_on_OneTimeWriteDisc()

Dim shtTargetSheet As String
Dim shtNameExpand As String

    With UserForm_Top
    
        shtTargetSheet = CStr(.ComboBox_compareSht.Value)
        shtNameExpand = "Expand_target"
        Worksheets(shtNameExpand).Activate
        ActiveSheet.Cells.Select
        Selection.ClearContents
        ActiveSheet.Range("A1").Select
        
        '展開したフォルダフルパスを保存するリストをクリアする
        excelIO_UDF_clear_lst_expand_target
        
        Call expandFolderTree_OneSheet(shtTargetSheet, shtNameExpand, 0, 0)
        
    End With

End Sub

'***************************************************************
'       フォルダツリー展開
' 引数で指定されたシート名のstartColからendColまでを展開する
'ただし、startCol = 0 かつ、endCol = 0のときはシート全て
'***************************************************************
Private Sub expandFolderTree_OneSheet(shtTargetSheet As String, shtNameExpand As String, startColumn As Integer, endColumn As Integer)
    
    Const LoopCntLimit As Integer = 5000
    
    Application.ScreenUpdating = False
    
    '展開するフォルダツリーがあるシート
    Set targetSHeet = Sheets(shtTargetSheet)
    
    'フォルダツリー展開結果を書き込むシート
    Set currentSheet = Sheets(shtNameExpand)
    
    If excelIO_UDF_get_expandFolderTree_status = True Then
        currentSheet.Activate
        Cells.Select
        Selection.ClearContents
        Range("A1").Select
    End If

    'シートの最終列
    targetSHeet.Activate
    LimitCol = targetSHeet.Cells(Module_Data.Row_FolderTreeTop, Columns.Count).End(xlToLeft).column
    'フォルダツリーを展開する範囲指定（開始列、終了列）
    If startColumn = 0 And endColumn = 0 Then
        col = 2
        endCol = LimitCol
    Else
        col = startColumn
        endCol = endColumn
    End If
    
    
    '=======================================================
    '   １枚のシートに記録された全てのフォルダツリー展開
    '=======================================================
    loopCnt = 1
    Do
    
loopTop:
        
        'ColumnName = Split(Columns(loopCnt).Address, "$")(2)
        'showColInfo = ColumnName & "[列] ==> " & loopCnt & "[列目]"
        
        '------------------
        '   先頭行の検索
        '------------------
        topRow = 1
        Do
            If InStr(1, targetSHeet.Cells(topRow, col), Module_Data.TopFolderName) > 0 Then Exit Do
            topRow = topRow + 1
            
            '暴走防止
            If topRow > LoopCntLimit Then Exit Do
        Loop
        
        test = targetSHeet.Cells(topRow, col)
        
        If topRow > LoopCntLimit Then
            col = col + 1
            loopCnt = loopCnt + 1
            
            If col < 100 Then
                GoTo loopTop
            Else
                MsgBox "先頭行が見つかりません。"
                Exit Sub
            End If
        End If
        
        startRow = topRow
        
        '------------------------------------------------
        '   １列のフォルダツリー展開（ディスク１枚分）
        '------------------------------------------------
        ret = excelIO_UDF_test(ActiveWorkbook.Name, row + 1, col)
        
        ColumnName = Split(Columns(col).Address, "$")(2)
        showColInfo = ColumnName & "[列] ==> " & col & "[列目]"
        UserForm_Top.Label_Message.Caption = "シート(" & shtTargetSheet & ") のフォルダツリーを展開中です。" & showColInfo
        DoEvents
        
        nextRow = excelIO_UDF_expandFolderTree(ActiveWorkbook.Name, shtTargetSheet, shtNameExpand, startRow, col)
        'startRow = nextRow
        
        col = col + 1
        If col > endCol Then Exit Do
        
        '暴走防止
        If col > LimitCol Then Exit Do
        
        loopCnt = loopCnt + 1
    Loop

    'RunPython "import inerfaceExcel; inerfaceExcel.excelIO_main()"
   
    currentSheet.Activate

    Application.ScreenUpdating = True
    
End Sub

'***********************************************************
'　フォルダツリーを展開したフルパスの比較
'***********************************************************
Public Sub searchFullpath()

    Application.ScreenUpdating = False

    With UserForm_Top
        baseShtName = .ComboBox_baseSht.Value
        compShtName = .ComboBox_compareSht.Value
        
        topFolder = .TextBox_searchRootFolder.Value
        resultSheet = .ComboBox_resultCompareSht.Value
    End With
    
        'シートのクリア
        Module_common.initializeResultSheet
        
    UserForm_Top.Label_Message.Caption = "基準側フォルダパスを取得中"
    With Sheets("Expand_base")
    
        col = UserForm_Top.ComboBox_BaseFolderTreeTop.ListIndex + 1
        
        Set topCell = .Cells(1, col)
        
        If topCell <> "" Then
        '先頭行にデータがある場合は、先頭から
            startRow = 1
        Else
        '先頭行にデータがない場合は、先頭行を検索する
            startRow = topCell.End(xlDown).row
        End If
        
        endRow = .Cells(Rows.Count, col).End(xlUp).row
        
        'Base側のフルパスをpythonのデータ（リスト）へ保存する
        For r = startRow To endRow
            excelIO_UDF_setBaseFullPath (CStr(.Cells(r, col)))
        Next r
        
    End With

'デバグ
'debug_getFullPath

    UserForm_Top.Label_Message.Caption = "フォルダツリーを比較中"
    
    '返却値は、比較結果を記録した先頭行
    dataTopRow = excelIO_UDF_searchBranch(ActiveWorkbook.Name, topFolder, resultSheet)
    
    
    '×の行を探す
    Set shtResult = Sheets(resultSheet)
    row_notFound = 5
    If shtResult.Cells(row_notFound, 1) = "×" Then
    
        isFound = True
    Else
        isFound = False
        row_notFound = 1
        Do
            If shtResult.Cells(row_notFound, 1) = "×" Then
                isFound = True
                Exit Do
            End If
            
            row_notFound = row_notFound + 1
            If row_notFound > 100 Then Exit Do  '暴走防止
        Loop
    End If
    
    If isFound = True And shtResult.Cells(row_notFound, 2) > 0 Then
        '×がゼロではないとき、先頭行の一つ上にフィルタを追加する
        Rows(dataTopRow - 1).Select
        Selection.AutoFilter
        Range("A9").Select
        ActiveSheet.Range("$A$9:$E$2059").AutoFilter Field:=1, Criteria1:="×"
    End If

    Application.ScreenUpdating = True

End Sub

'*********************************************************
'   フォルダ名変更
'   フォルダ名フルパスの末尾のフォルダ名前のみを変更する
'*********************************************************
Public Sub renameEndOfFolderName(startRow As Integer)

    endRow = Sheets("ファイル数").Cells(Rows.Count, col_newBranchName).End(xlUp).row

    For r = startRow To endRow
    
        With Sheets("ファイル数")
            oldNamedBranch = .Cells(r, 1)
            oldNamedFullPath = .Cells(r, 4)
            newName = .Cells(r, col_newBranchName)
        End With
    
        If newName <> "" And newName <> UndefileFolderName Then
            newFullPathName = excelIO_UDF_generateNewFileFullPath(oldNamedFullPath, newName)
            Debug.Print (oldNamedFullPath)
            Debug.Print (newFullPathName)
            
            On Error Resume Next
            
            'フォルダ名変更
            'Name "D:\旅日記\2024\1101～1104_第七十六回正倉院展\1103_高取、飛鳥\03_明日香村\02_" As "D:\旅日記\2024\1101～1104_第七十六回正倉院展\1103_高取、飛鳥\03_明日香村\02_万葉文化館"
            Name CStr(oldNamedFullPath) As CStr(newFullPathName)
            
            If Err.Number = 0 Then
                branch = Left(oldNamedBranch, InStr(1, oldNamedBranch, "―"))
                Sheets("ファイル数").Cells(r, 1) = branch & newName
                Stop
            Else
                Debug.Print Err.Description
                Stop
            End If
            
        End If
        
    Next r
Stop
End Sub


'***************************************************************************************
'   デバッグ用関数
'   python側で、リストデータとして保存されているデータを取得し、エクセルシートに出す。
'***************************************************************************************
Private Sub debug_getFullPath()

UserForm_Top.Label_Message.Caption = "[debug]フルパス取得"

With Sheets("debug_log")
    
    dbg_col = 10
    
    For r = 1 To (endRow - startRow)
        .Cells(r, dbg_col) = excelIO_UDF_outputBaseFullPath(r - 1)
        .Cells(r, dbg_col + 1) = excelIO_UDF_outputTargetFullPath(r - 1)
    Next r
        

End With

End Sub

