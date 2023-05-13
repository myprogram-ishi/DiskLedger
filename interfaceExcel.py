
import xlwings as xlw
import pandas as pdFldr
import data
import folderTreeControl

########################################################
#      初期化　変数クリアなど
########################################################
@xlw.func
def excelIO_UDF_initialize():
    data.dataClear()
@xlw.func
def excelIO_UDF_clear_lst_expand_base():
    data.lst_expandFolderTreeBase.clear()
@xlw.func
def excelIO_UDF_clear_lst_expand_target():
    data.lst_expandFolderTreeTarget.clear()

########################################################
#      エクセルとの入出力メイン関数
########################################################
@xlw.func
def excelIO_UDF_main(srcExcel):

    data.currentExcel = srcExcel

    wb = xlw.Book(data.currentExcel)
    sheet = wb.sheets[data.shtName_main]
    sheet.range('A1').value = excelIO_UDF_main.__name__

########################################################
#       フォルダツリーを追加する
########################################################
@xlw.func
def excelIO_UDF_addFolderTree(TopFolder, srcExcel, dstSheet, startRow, StartCol):

    if srcExcel == None:
        srcExcel = data.currentExcel

    #フォルダツリーの生成（結果は、data.lst_generateToAddFolderTree　に保存される）
    folderTreeControl.generateFolderTree(path=TopFolder, layer=0, is_last=False,indent_current=data.indent_tree)

    r = startRow
    for branch in data.lst_generateToAddFolderTree:
        excelIO_UDF_outputdebugLog(srcExcel=srcExcel, shtName=dstSheet, row=r, col=StartCol, outval=str(branch))
        r = r + 1

########################################################
#      フォルダツリーを展開
########################################################
@xlw.func
def excelIO_UDF_expandFolderTree(srcExcel= None, srcSheet= None, dstSheet= None, startRow= None, StartCol= None):

    row = startRow
    col = StartCol

    if srcExcel != '' and srcExcel != None:
        data.currentExcel = srcExcel

    if srcSheet != '' and srcSheet != None:
        data.shtName_base = srcSheet

    if dstSheet != '' and dstSheet != None:
        data.shtName_Expand = dstSheet

    excelIO_UDF_appendDataToLasRow(data.shtName_dbgLog, 1, excelIO_UDF_expandFolderTree.__name__)

    #エクセルシートに定義されたフォルダツリーを１行ずつ、リストへ追加する。
    wb = xlw.Book(data.currentExcel)
    sheet_base = wb.sheets[data.shtName_base]
    sheet_pyLog = wb.sheets[data.shtName_dbgLog]

    #フォルダツリーの最後の行を取得する
    macro = wb.macro(data.xlInterface + '.' + data.getEndRowCout)
    EndRow = macro(data.shtName_base, col)

    branchName = sheet_base.cells(row,col).value
    data.lst_branch.clear()
    while row <= EndRow:
        data.lst_branch.append(branchName)
        row = row + 1
        branchName = sheet_base.cells(row, col).value
        if branchName == '':
            break
        else:
            #展開したブランチの数を数える
            dataCnt = row - startRow

    for index in range(len(data.lst_branch) - 1):
        sheet_pyLog.cells(index + 1, 5).value = data.lst_branch[index]

    if data.searchRootFolder == '':
        data.searchRootFolder = excelIO_UDF_getSearchRootFolder()

    top = folderTreeControl.expandFolderTree(data.searchRootFolder, col)
    #sheet_pyLog.cells(row, col).value = 'top ='
    #sheet_pyLog.cells(row, col + 1).value = top

    #この返却値は、次に書き込みをする行番号の算出に使う
    return dataCnt

########################################################
#
########################################################
offsetResultRow = 5
@xlw.func
def excelIO_UDF_searchBranch(srcExcel=None, searchTopFolder=None, resultSheet=None):

    lst_work_expandFolderTreeBase = []
    lst_work_expandFolderTreeTarget = []

    if srcExcel == None:
        srcExcel = data.currentExcel

    wb = xlw.Book(srcExcel)
    sheet_result = wb.sheets[resultSheet]

    data.lst_searchResults.clear()

    if searchTopFolder != None:

        sheet_result.cells(1, 5).value = 'A'

    #検索開始フォルダが指定されている場合は、そのフォルダ名でリストにフィルタをかけて抽出する
        lst_work_expandFolderTreeBase \
            = [item for item in data.lst_expandFolderTreeBase if searchTopFolder in item]

        lst_work_expandFolderTreeTarget \
            = [item for item in data.lst_expandFolderTreeTarget if searchTopFolder in item]
    else:

        sheet_result.cells(1, 5).value = 'B'

    #フォルダ指定の無い場合は、そのまま使用する
        lst_work_expandFolderTreeBase = data.lst_expandFolderTreeBase
        lst_work_expandFolderTreeTarget = data.lst_expandFolderTreeTarget

    col = 2
    foundCount = 0
    sheet_result.cells(2, col).value = len(lst_work_expandFolderTreeBase)
    sheet_result.cells(3, col).value = len(lst_work_expandFolderTreeTarget)

    #展開したフルパスをシートに記録
    for row, item in enumerate(lst_work_expandFolderTreeBase):

        sheet_result.cells(row + offsetResultRow, col).value = item

        # ドライブ名は無視して、比較する
        drive = item.find(':\\')
        if drive > 0:
            item = item[drive+2:]
        # 検索するフォルダが含まれているかどうかをここで判定する。
        if item in lst_work_expandFolderTreeTarget:
            sheet_result.cells(row + offsetResultRow, 1).value = '〇'
            foundCount = foundCount + 1
        else:
            sheet_result.cells(row + offsetResultRow, 1).value = '×'

    sheet_result.cells(1, 2).value = foundCount

    col = 3
    sheet_result.cells(2, col).value = len(lst_work_expandFolderTreeTarget)
    for row, item in enumerate(lst_work_expandFolderTreeTarget):
        sheet_result.cells(row + offsetResultRow, col).value = item

##########################################################################################
def debug():
#基準側のループ
    sheet_result_row = 1
    for itemBase in (data.lst_expandFolderTreeBase):
        posTopFolder = itemBase.find(searchTopFolder)
        posColon = itemBase.find(r':')
        if posTopFolder > 0:
            #検索開始フォルダが指定されている場合は、それ以降を対象とする
            work_itemBase = itemBase[posTopFolder:]
        elif posColon > 0:
            work_itemBase = itemBase[posColon:]
        else:
            work_itemBase = itemBase

        sheet_result.cells(sheet_result_row, 2).value = work_itemBase
        sheet_result_row = sheet_result_row + 1

        for itemTarget in (data.lst_expandFolderTreeTarget):
            posTopFolder = itemTarget.find(searchTopFolder)
            posColon = itemBase.find(r':')
            if posTopFolder > 0:
                # 検索開始フォルダが指定されている場合は、それ以降を対象とする
                work_itemTarget = itemTarget[posTopFolder:]
            elif posColon > 0:
                work_itemTarget = itemTarget[posColon:]
            else:
                work_itemTarget = itemTarget

            if work_itemTarget in work_itemBase:
                data.lst_searchResults.append(work_itemTarget)
                break

    sheet_result.cells(1, 1).value = len(data.lst_expandFolderTreeBase)
    sheet_result.cells(2, 1).value = len(data.lst_expandFolderTreeTarget)

    row = 3
    for item in data.lst_searchResults:
        sheet_result.cells(row, 1).value = item
        row = row + 1

########################################################
#
########################################################
@xlw.func
def generateFullPath():
    print(generateFullPath.__name__)

########################################################
#
########################################################
@xlw.func
def excelIO_UDF_appendDataToLasRow(shtName,col, setVal):

    #エクセルシートに定義されたフォルダツリーを１行ずつ、リストへ追加する。

    wb = xlw.Book(data.currentExcel)
    #sheet_pyLog = wb.sheets[shtName]
    # sheet_pyLog.cells(row, col).value = setVal

    #ふぉるdツリーの最後の行を取得する
    macro = wb.macro(data.xlInterface + '.' + 'getAppendDataToLasRow')
    row = macro(shtName, col, setVal)

########################################################
#   検索開始フォルダの取得
########################################################
@xlw.func
def excelIO_UDF_getSearchRootFolder():
    wb = xlw.Book(data.currentExcel)
    macro = wb.macro(data.xlInterface + '.' + 'getSearchRootFolder')
    data.searchRootFolder = macro()

########################################################
# 比較基準になるフォルダフルパス文字列をエクセルシートから受け寄る
########################################################
@xlw.func
def excelIO_UDF_setBaseFullPath( item ):
    data.lst_expandFolderTreeBase.append(item)

########################################################
#      デバグ用　エクセルシートにリストの値を返す
########################################################
@xlw.func
def excelIO_UDF_outputBaseFullPath( index ):
    return data.lst_expandFolderTreeBase[index]

@xlw.func
def excelIO_UDF_outputTargetFullPath( index ):

    if index < len(data.lst_expandFolderTreeTarget):
        return data.lst_expandFolderTreeTarget[index]
    else:
        return 'out of range'

########################################################
#      指定したエクセルの場所（シート、行、列）に値を書き込む
########################################################
def excelIO_UDF_outputdebugLog(srcExcel=None, shtName=None, row=1, col=1, outval='No data'):

    if shtName == '' or shtName == None:
        shtName = data.shtName_pyDbgLog

    if srcExcel != '' and srcExcel != None:
        data.currentExcel = srcExcel

    outputExcel = data.currentExcel

    #エクセルシートに定義されたフォルダツリーを１行ずつ、リストへ追加する。
    wb = xlw.Book(outputExcel)
    sheet_pyLog = wb.sheets[shtName]
    sheet_pyLog.cells(row, col).value = outval


@xlw.func
def excelIO_UDF_test(srcExcel,row, col):

    data.currentExcel = srcExcel

    #srcExcelr = 'c:\Users\localuser\PycharmProjects\pythonProject\検索_python.xlsm'
    wb = xlw.Book(data.currentExcel)
    sheet = wb.sheets['debug_log']

    sheet.cells(row, col).value = excelIO_UDF_test.__name__
