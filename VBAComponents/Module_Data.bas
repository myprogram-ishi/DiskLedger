Attribute VB_Name = "Module_Data"
Type sheetCombination
    year As String
    ptrnName As String
    basSht As String
    targetSht As String
    expandbaseSht As String
    expandTargetSht As String
    resultSht As String
End Type

Type networkDrive
    driveName As String
    drivePath As String
End Type


Dim comparePtrn1 As sheetCombination
Dim comparePtrn2 As sheetCombination

Dim networkDrivePath As networkDrive

Dim sheetCombinationArray() As sheetCombination


Public currentUserForm As Integer

Enum StateUserForm
    Top = 0
    addFolderTree = 1
    getFileCnt = 2
End Enum

Public doShow_UserForm_Top As Boolean

'DVD, Blue rayディスク種類一覧
Public romDiskList As Variant
'Public Const romDiskList As Variant = Array("25GB", "50GB")

Public Const shatName_base As String = "base"

Public upperLimit_fileCnt As Integer

'基準シートのフォルダツリー先頭行
Public Const baseSht_TreeTopRow  As Integer = 5

Public TopFolderName As String

'UserForm_Topの初期化関数が完了しているかどうか。
Public is_UserForm_Top_Initialize As Boolean

Public is_CombinationValid As Boolean


'各シートにフォルダツリーを書き込むの先頭行の共通定義
Public Const Row_FolderTreeTop  As Integer = 10

Public Const Row_addFolderTreeTop  As Integer = 5


Sub setNetworkDriveName()

    ReDim networkDrivePath(0)
    networkDrivePath(0).driveName = "Z:\"
    networkDrivePath(0).drivePath = "\\LANDISK-8DEF17\disk1\旅日記:\"

    ReDim networkDrivePath(1)
    networkDrivePath(1).driveName = "Y:\"
    networkDrivePath(1).drivePath = "\\LANDISK-8DEF17\disk1\旅日記:\"

End Sub

'**************************************************************
'   シートの組み合わせ選択変更によるコンボボックスへの反映
'**************************************************************
Sub selectSheetNameInCombobox(CombinationNo As Integer)

    If is_UserForm_Top_Initialize = False Then Exit Sub
    
Dim comboItemList As Variant
    
    is_CombinationValid = False
    
    ret = UserForm_Top.ComboBox_sheetCombination.ListIndex

    With UserForm_Top
    
        '基準シート
        targetSHeet = sheetCombinationArray(CombinationNo).basSht
        For i = 0 To (.ComboBox_baseSht.ListCount - 1)
            .ComboBox_baseSht.ListIndex = i
            If .ComboBox_baseSht.Value = targetSHeet Then Exit For
        Next i
        
        'ハードディスク側のフォルダを検索する（エクセル上の列番号検索）
        listCnt = .ComboBox_BaseFolderTreeTop.ListCount
        comboItemList = .ComboBox_BaseFolderTreeTop.List(0)
        
        For i = 0 To (listCnt - 1)
            comboItem = .ComboBox_BaseFolderTreeTop.List(i)
            y = sheetCombinationArray(.ComboBox_sheetCombination.ListIndex).year
            If InStr(1, comboItem, "2023") > 0 Then
                .ComboBox_BaseFolderTreeTop.ListIndex = i
                Exit For
            End If
        Next i
        
        ret = .ComboBox_sheetCombination.ListIndex
        
        '比較シート
        targetSHeet = sheetCombinationArray(CombinationNo).targetSht
        For i = 0 To (.ComboBox_compareSht.ListCount - 1)
            .ComboBox_compareSht.ListIndex = i
            If .ComboBox_compareSht.Value = targetSHeet Then Exit For
        Next i
        
        ret = .ComboBox_sheetCombination.ListIndex
        
        '結果シート
        'targetSheet = sheetCombinationArray(CombinationNo).resultSht
        'For i = 0 To (.ComboBox_resultCompareSht.ListCount - 1)
        '    .ComboBox_resultCompareSht.ListIndex = i
        '    If .ComboBox_resultCompareSht.Value = targetSheet Then Exit For
        'Next i
        
        'ret = .ComboBox_sheetCombination.ListIndex
        
    End With
    
    is_CombinationValid = True

End Sub

'*******************************************************
'    キーワードに応じたシートの組み合わせを選択する
'*******************************************************
Sub selectcComboBox_sheetCombination(keyWord As String)

    With UserForm_Top
    
      For i = 0 To (UBound(sheetCombinationArray) - 1)
      
        conbiName = sheetCombinationArray(i).ptrnName
        
        '結果シート
        If (InStr(1, keyWord, "result") > 0) Then
            If (InStr(1, keyWord, conbiName) > 0) Then
                .ComboBox_sheetCombination.ListIndex = i
            End If
          
        Else
            If (InStr(1, keyWord, conbiName) > 0) Then
                .ComboBox_sheetCombination.ListIndex = i
            End If
        End If
      Next i
    
    End With

End Sub
'******************************************
'  シート組み合わせコンボボックスの初期化
'******************************************
Sub initialize_ComboBox_sheetCombination()

    ptrnCnt = UBound(sheetCombinationArray)
    
    With UserForm_Top
    
        For i = 0 To (ptrnCnt - 1)
            .ComboBox_sheetCombination.AddItem (sheetCombinationArray(i).ptrnName)
            
        Next i
        
        .ComboBox_sheetCombination.ListIndex = 0
    
    End With
    

End Sub

'*************************************
'       変数類の初期化
'*************************************
Sub initializeParametor()

    romDiskList = Array("25GB", "50GB")

    is_CombinationValid = True

    'comparePtrn21.ptrnName = "ptrn11"
    'comparePtrn21.basSht = shatName_base
    'comparePtrn21.targetSht = "2023"
    'comparePtrn21.expandbaseSht = "Expand_base"
    'comparePtrn21.expandTargetSht = "Expand_target"
    'comparePtrn21.resultSht = "result_2023"

    'comparePtrn22.ptrnName = "ptrn12"
    'comparePtrn22.basSht = shatName_base
    'comparePtrn22.targetSht = "2022"
    'comparePtrn22.expandbaseSht = "Expand_base"
    'comparePtrn22.expandTargetSht = "Expand_target"
    'comparePtrn22.resultSht = "result_2022"
    
    index = 0
    ReDim sheetCombinationArray(1)
    sheetCombinationArray(index).year = ""
    sheetCombinationArray(index).ptrnName = "-----"
    sheetCombinationArray(index).basSht = ""
    sheetCombinationArray(index).targetSht = ""
    sheetCombinationArray(index).expandbaseSht = ""
    sheetCombinationArray(index).expandTargetSht = ""
    sheetCombinationArray(index).resultSht = ""

    index = index + 1
    sheetCombinationArray(index).year = "2023"
    ReDim Preserve sheetCombinationArray(index + 1)
    sheetCombinationArray(index).ptrnName = "2023_25GB"
    sheetCombinationArray(index).basSht = shatName_base
    sheetCombinationArray(index).targetSht = "2023_Disk25GB"
    sheetCombinationArray(index).expandbaseSht = "Expand_base"
    sheetCombinationArray(index).expandTargetSht = "Expand_target"
    sheetCombinationArray(index).resultSht = "result_2023"
    
    index = index + 1
    sheetCombinationArray(index).year = "2023"
    ReDim Preserve sheetCombinationArray(index + 1)
    sheetCombinationArray(index).ptrnName = "2023_50GB"
    sheetCombinationArray(index).basSht = shatName_base
    sheetCombinationArray(index).targetSht = "2023_Disk50GB"
    sheetCombinationArray(index).expandbaseSht = "Expand_base"
    sheetCombinationArray(index).expandTargetSht = "Expand_target"
    sheetCombinationArray(index).resultSht = "result_2023"

    index = index + 1
    sheetCombinationArray(index).year = "2022"
    ReDim Preserve sheetCombinationArray(index + 1)
    sheetCombinationArray(index).ptrnName = "2022_25GB"
    sheetCombinationArray(index).basSht = shatName_base
    sheetCombinationArray(index).targetSht = "2022_Disk25GB" '"2022"
    sheetCombinationArray(index).expandbaseSht = "Expand_base"
    sheetCombinationArray(index).expandTargetSht = "Expand_target"
    sheetCombinationArray(index).resultSht = "result_2022"
    
    index = index + 1
    sheetCombinationArray(index).year = "2022"
    ReDim Preserve sheetCombinationArray(index + 1)
    sheetCombinationArray(index).ptrnName = "2022_50GB"
    sheetCombinationArray(index).basSht = shatName_base
    sheetCombinationArray(index).targetSht = "2022_Disk50GB"
    sheetCombinationArray(index).expandbaseSht = "Expand_base"
    sheetCombinationArray(index).expandTargetSht = "Expand_target"
    sheetCombinationArray(index).resultSht = "result_2022"
    
End Sub
