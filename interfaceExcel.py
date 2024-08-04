import logging
import os
import openpyxl
import xlwings as xlw
import inspect
#import pandas as pdFldr
import data

import pandas as pd
import pandas as pdExcel
import numpy as np
import csv
import folderTreeControl
import searchItem

import testFunction

########################################################
#      初期化　変数クリアなど
########################################################
@xlw.func
def excelIO_UDF_initialize(srcExcel):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    data.currentExcel = srcExcel

    testFunction.initialize_loggerObject()

    data.dataClear()

@xlw.func
def excelIO_UDF_clear_lst_expand_base():
    data.lst_expandFolderTreeBase.clear()
@xlw.func
def excelIO_UDF_clear_lst_expand_target():
    data.lst_expandFolderTreeTarget.clear()

@xlw.func
def excelIO_UDF__loggerFinalize():
    testFunction.finalize_loggerObject()

########################################################
#      エクセルとの入出力メイン関数
########################################################
@xlw.func
def excelIO_UDF_main(srcExcel):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=r'excelIO_UDF_main')
    data.currentExcel = srcExcel

    wb = xlw.Book(data.currentExcel)
    sheet = wb.sheets[data.shtName_main]
    sheet.range('A1').value = excelIO_UDF_main.__name__

########################################################
#   VBA側のユーザーフォームのキャプションを書き換える
########################################################
@xlw.func
def excelIO_UDF_updataUserformCaption(userformIndex=None, msssage=None):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)

def _excelIO_UDF_updataUserformCaption(userformIndex=None, msssage=None):
    wb = xlw.Book(data.currentExcel)
    #wb = xlw.Book("検索_python.xlsm")

    macro = wb.macro(data.xlInterface + '.' + 'updateUserformCaption')
    if userformIndex == None:
        UFindex = 0
    else:
        UFindex = userformIndex

    if msssage == None:
        msgforCaption = ""
    else:
        msgforCaption = msssage

    macro(UFindex, msgforCaption)

########################################################
#   一番日付の新しいフォルダを取得する
########################################################
@xlw.func
def excelIO_UDF_getLatestFolder(targetFolder):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    #フォルダ一覧
    lst_Folder = [f for f in os.listdir(targetFolder) if os.path.isdir(os.path.join(targetFolder, f))]
    #並べ替え
    lst_Folder.sort(reverse=True)

    #フォルダ名先頭文字が、数字のフォルダを取得する
    for folder in lst_Folder:
        if str.isdecimal(folder[0]) == True:
            break

    return folder

########################################################
#       フォルダツリーを追加する
########################################################
@xlw.func
def excelIO_UDF_addFolderTree(TopFolder, srcExcel, dstSheet, startRow, StartCol):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    if srcExcel == None:
        srcExcel = data.currentExcel

    #フォルダツリーの生成（結果は、data.lst_generateToAddFolderTree　に保存される）
    folderTreeControl.generateFolderTree(path=TopFolder, layer=0, is_last=False,
                                         indent_current=data.indent_tree)

    #testFunction.test_output_list_to_textFile(None, r'lst_generateToAddFolderTree.txt',
    #                                          data.lst_expandFolderTreeTarget)
    r = startRow
    isLastTrip = False  #一番最後のお出かけブランチ（日帰り、宿泊問わず最後に出かけた日のフォルダ）
    for branch in data.lst_generateToAddFolderTree:

        firstCharWithoutSpace = (branch.strip())[0]

        if isLastTrip == False:
        #フォルダツリー最後の、旅行フォルダではない場合
            if firstCharWithoutSpace == '├':
                isCaptionUpdata = True
            elif firstCharWithoutSpace == '└':
                isCaptionUpdata = True
                #フォルダツリーの末尾（最後に出かけた日のフォルダに到達）
                isLastTrip = True
            else:
                isCaptionUpdata = False

            #ユーザーフォーム上のキャプション更新（書き換え）
            if isCaptionUpdata == True:
                posWakibou = branch.find(r'―')
                if posWakibou > 0:
                    messageForCaption = branch[posWakibou + 1:] + r'　以下を処理しています'
                    excelIO_UDF_updataUserformCaption(msssage=messageForCaption)
                    
        excelIO_UDF_writeDataOnWorksheet(srcExcel=srcExcel, shtName=dstSheet, row=r, col=StartCol,
                                         outval=str(branch) )
        r = r + 1

########################################################
#      フォルダツリーを展開先選択
# statusの設定で、フォルダツリー展開結果の出力先を変える
#   True : ワークシートに出す
#   False : ワークシートに出さない。データフレームに出す
########################################################
@xlw.func
def excelIO_UDF_set_expandFolderTree_status(status=None):

    data.outExpFolderTreeToWrkSht = status

    if status == False:
        data.lst_expanfFolderTree.clear()
@xlw.func
def excelIO_UDF_get_expandFolderTree_status():
    return data.outExpFolderTreeToWrkSht

@xlw.func
def excelIO_UDF_getExpandFolderTree():
    return data.lst_expandFolderTreeTarget

########################################################
#      フォルダツリーを展開
########################################################
@xlw.func
def excelIO_UDF_expandFolderTree(srcExcel= None, srcSheet= None, dstSheet= None, startRow= None, StartCol= None):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
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
offsetResultRow = 10
@xlw.func
def excelIO_UDF_searchBranch(srcExcel=None, searchTopFolder=None, resultSheet=None):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    lst_work_expandFolderTreeBase = []
    lst_work_expandFolderTreeTarget = []

    if srcExcel == None:
        srcExcel = data.currentExcel

    wb = xlw.Book(srcExcel)
    sheet_result = wb.sheets[resultSheet]

    data.lst_searchResults.clear()

    if searchTopFolder != None:

        for item in data.lst_expandFolderTreeBase:
            topPos = item.find(searchTopFolder)
            if topPos > 0:
                lst_work_expandFolderTreeBase.append(item[topPos:])
            else:
                lst_work_expandFolderTreeBase.append(item)

        for item in data.lst_expandFolderTreeTarget:
            topPos = item.find(searchTopFolder)
            if topPos > 0:
                lst_work_expandFolderTreeTarget.append(item[topPos:])
            else:
                lst_work_expandFolderTreeTarget.append(item)

        # 検索開始フォルダが指定されている場合は、そのフォルダ名でリストにフィルタをかけて抽出する
        # lst_work_expandFolderTreeBase
        # = [item for item in data.lst_expandFolderTreeBase if searchTopFolder in item]

        # lst_work_expandFolderTreeTarget \
        #    = [item for item in data.lst_expandFolderTreeTarget if searchTopFolder in item]

        fileNameListBase = r'lst_work_expandFolderTreeBase_filter.txt'
        fileNameListTarget = r'lst_work_expandFolderTreeTarget_filter.txt'

    else:
    #フォルダ指定の無い場合は、そのまま使用する
        lst_work_expandFolderTreeBase = data.lst_expandFolderTreeBase
        lst_work_expandFolderTreeTarget = data.lst_expandFolderTreeTarget

        fileNameListBase = r'lst_work_expandFolderTreeBase.txt'
        fileNameListTarget = r'lst_work_expandFolderTreeTarget.txt'

    testFunction.test_output_list_to_textFile(None, fileNameListBase, lst_work_expandFolderTreeBase)
    testFunction.test_output_list_to_textFile(None, fileNameListTarget, lst_work_expandFolderTreeTarget)
    col = 2
    result_item_row = 2

    #比較フォルダ数、比較結果表示
    sheet_result.cells(result_item_row, col).value = len(lst_work_expandFolderTreeBase)
    result_item_row = result_item_row + 1
    sheet_result.cells(result_item_row, col).value = len(lst_work_expandFolderTreeTarget)
    result_item_row = result_item_row + 1

    foundCount = 0
    notFoundCount = 0

    lst_NGitem = []

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
            colorPtrn = 1   #背景[青]　文字[白]
        else:
            lst_NGitem.append(item)
            sheet_result.cells(row + offsetResultRow, 1).value = '×'
            colorPtrn = 0   #背景[赤]　文字[白]
            notFoundCount = notFoundCount + 1

        macro = wb.macro(data.xlInterface + '.' + 'setCellInterior')
        macro(resultSheet, row + offsetResultRow, 1, colorPtrn)

    testFunction.test_output_list_to_textFile(None, r'lst_NGitem.txt', lst_NGitem)

    #比較結果の表示
    #見つかったフォルダ数の数（バックアップ済みフォルダ数）
    sheet_result.cells(result_item_row, 2).value = foundCount
    result_item_row = result_item_row + 1
    #見つからなかったフォルダ数の数（バックアップ未済みフォルダ数）
    sheet_result.cells(result_item_row, 2).value = notFoundCount

    col = 3
    #sheet_result.cells(2, col).value = len(lst_work_expandFolderTreeTarget)
    for row, item in enumerate(lst_work_expandFolderTreeTarget):
        sheet_result.cells(row + offsetResultRow, col).value = item

    return offsetResultRow

########################################################
#
########################################################
@xlw.func
def generateFullPath():

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)

########################################################
#
########################################################
@xlw.func
def excelIO_UDF_appendDataToLasRow(shtName,col, setVal):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)

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

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    wb = xlw.Book(data.currentExcel)
    macro = wb.macro(data.xlInterface + '.' + 'getSearchRootFolder')
    data.searchRootFolder = macro()

########################################################
# 比較基準になるフォルダフルパス文字列をエクセルシートから受け寄る
########################################################
@xlw.func
def excelIO_UDF_setBaseFullPath( item ):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
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
#    指定したエクセルの場所（シート、行、列）に値を書き込む
########################################################
def excelIO_UDF_writeDataOnWorksheet(srcExcel=None, shtName=None, row=1, col=1, outval='No data'):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    if shtName == '' or shtName == None:
        shtName = data.shtName_pyDbgLog

    if srcExcel != '' and srcExcel != None:
        data.currentExcel = srcExcel

    outputExcel = data.currentExcel

    #エクセルシートに定義されたフォルダツリーを１行ずつ、リストへ追加する。
    wb = xlw.Book(outputExcel)
    sheet_pyLog = wb.sheets[shtName]
    sheet_pyLog.cells(row, col).value = outval

########################################################
#
########################################################
@xlw.func
def excelIO_UDF_getDestHyperLinkRow(srcExcel= None, srcSheet= None, desstFolderNane= None):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    lst_diskID = []
    lst_posDesstFolder = []
    lst_error = []

    if srcExcel != '' and srcExcel != None:
        data.currentExcel = srcExcel

    if srcSheet != '' and srcSheet != None:
        data.shtName_base = srcSheet

    if desstFolderNane == '' and desstFolderNane == None:
        desstFolderNane = data.searchRootFolder

    excelIO_UDF_appendDataToLasRow(data.shtName_dbgLog, 1, excelIO_UDF_expandFolderTree.__name__)

    wb = xlw.Book.caller()
    ws = wb.sheets[srcSheet]

    "エクセルシートのデータを取得する"
    df_sheet = ws.range(ws.cells(1, 1), ws.cells(5000, 255)).options(pd.DataFrame, index=False).value

    df_sheet = df_sheet.drop(df_sheet.index[[0, 1, 2]])

    #ディスクIDおまとめた、リストを取得
    lst_diskID = [x for x in list(df_sheet.columns) if x != None]

    df_sheet.to_csv(os.path.join(data.outoutFolder_df_tocsv, r'df_sheet.csv'))
    df_sheet_T = df_sheet.T

    a = 0
    df_True = pd.DataFrame
    dict_posDesstFolder = {}
    lst_posDesstFolder.clear()
    for index, diskID in enumerate(lst_diskID):

        if index > 0:

            outFileName = r'df_sheet_org_' + str(index) + '_' + str(diskID) + '.csv'
            df_sheet[str(diskID)].to_csv(os.path.join(data.outoutFolder_df_tocsv, outFileName))

            try:
                outFileName = r'series_sheet1_' + str(diskID) + '.csv'

                df_ret_dropna = df_sheet.dropna(how='all')

                df_ret = df_ret_dropna[df_ret_dropna[diskID].str.contains(desstFolderNane, na=False)]
                df_ret= df_ret.dropna(how='all')
                df_ret.to_csv(os.path.join(data.outoutFolder_df_tocsv, '___' + outFileName))

                lst_posDesstFolder.append(df_ret.index.values)

                # ハイパーリンク
                adjust_df_to_row = 2    #データフレームでの行番号と、エクセルシートでの行番号の調整値（オフセット）
                try:
                    destRow = int(lst_posDesstFolder[len(lst_posDesstFolder) - 1]) + adjust_df_to_row
                except TypeError as error:
                    #ここに来るときは、設定したキーワードを含むセルが複数あるということなので、いったんリストにする
                    list_KeywordCell = list(lst_posDesstFolder[len(lst_posDesstFolder) - 1])
                    try:
                        #リストの最後をハイパーリンクのリンク先に指定する
                        destRow = int(list_KeywordCell[len(list_KeywordCell) - 1]) + adjust_df_to_row
                        #debugFunction_list_toCsv(list_KeywordCell,r'D:\pythonDebugOut\listTemp.csv')
                    except IndexError as error:
                        # ここが実行されるときは、キーワードが見つからなかったとき。適当な値を入れておく
                        destRow = 5

                except IndexError as error:
                    # ここが実行されるときは、キーワードが見つからなかったとき。適当な値を入れておく
                    destRow = 5
                #else:
                #    #暴走防止
                #    destRow = 1

                if index <= 25:     #列名が一文字（"Z"まで）
                    destCol = chr(index + 65)
                else:
                # upper, lowerともに、内側カッコ内の結果（割り算の商または余り）が、１のとき、"A"になるようにする。
                    upperColName = chr((index // 25) + 64)
                    lowerColName = chr((index % 25) + 64)
                    destCol = upperColName + lowerColName

                #ハイパーリンク先を範囲指定することで、画面上部に本来、目的のセルが表示されるようにする。
                destHyperLink = srcExcel + '#' + srcSheet + '!' + destCol + str(
                    destRow) + ":" + destCol + str(destRow + 20)

                #destHyperLink = srcExcel + '#' + srcSheet + '!' + chr(index + 65) + str(destRow + adjust_df_to_row) #str(destRow)

                wbHyp = openpyxl.Workbook()
                wsHyp = wbHyp.active

                ws.cells(1, index + 1).add_hyperlink(destHyperLink, diskID)

                ws.cells(1, index + 1).api.Font.ColorIndex = 5
                ws.cells(1, index + 1).api.Font.Size = 20
                ws.cells(1, index + 1).api.Font.Bold = True

                dict_posDesstFolder[diskID] = df_ret.index.values

                    #break
            except ValueError as error:
                outFileName = r'ValueError_' + str(index) + '_' + str(diskID) + '.csv'
                df_sheet[str(diskID)].to_csv(os.path.join(data.outoutFolder_df_tocsv, outFileName))
                a = a + 1
            except AttributeError as error:
                outFileName = r'AttributeError_' + str(index) + '_' + str(diskID) + '.csv'
                df_sheet.to_csv(os.path.join(data.outoutFolder_df_tocsv, outFileName))
                a = a + 100
                break
            else:
                continue

    with open(os.path.join(data.outoutFolder_df_tocsv, 'lst_posDesstFolder.csv'), 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerows(lst_posDesstFolder)

    return 0

@xlw.func
def excelIO_UDF_generateFileCntByFolderList(srcSheet= None, topRow = None, rowCount = None, fileCntUpperLimit = 0):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
#エクセルシート状の列番号
    col_fileCnt = 2     #ファイル数の列
    col_folder = 3   #フォルダパスの列絵

    if srcSheet != '' and srcSheet != None:
        data.shtName_base = srcSheet

    if topRow == None:
        topRow = 1

    if rowCount == None:
        rowCount = 1000

    if fileCntUpperLimit == 0:
        fileCntUpperLimit = 100

    wb = xlw.Book.caller()
    ws = wb.sheets[srcSheet]

   # データフレームを初期化
    data.df_fileCntByFolder = pd.DataFrame(columns=data.cols_dfFileCnt)

    df_temp = pd.DataFrame([[ws.cells(topRow, col_fileCnt).value, ws.cells(topRow, col_folder).value]],
                       columns=data.cols_dfFileCnt)

    data.df_fileCntByFolder.append(df_temp)
    #data.df_fileCntByFolder.columns = ['path', 'fileCnt']

    index = 1
    for index in range(1, rowCount):
        #ファイル数取得
        fileCount = ws.cells(topRow + index, col_fileCnt).value
        if fileCount == None or fileCount == "":
            fileCount = 0

        df_temp = pd.DataFrame(
            [[fileCount, ws.cells(topRow + index, col_folder).value]],
            columns=data.cols_dfFileCnt)

        data.df_fileCntByFolder = data.df_fileCntByFolder.append(df_temp)

    try:
        errorNum = 0
        data.df_fileCntByFolder.to_csv(os.path.join(data.outoutFolder_df_tocsv, r'df_fileCntByFolder.csv'))

        #'ファイル数'の列を整数型へ返還
        errorNum = 1
        data.df_fileCntByFolder['fileCnt'] = data.df_fileCntByFolder['fileCnt'].astype('int')
        #しきい価を超えた行を抽出
        errorNum = 2
        data.df_fileCntByFolder = data.df_fileCntByFolder.loc[data.df_fileCntByFolder['fileCnt'] > fileCntUpperLimit]
        #大きい順に並べ替え
        errorNum = 3
        data.df_fileCntByFolder = data.df_fileCntByFolder.sort_values('fileCnt', ascending=False)

    except Exception as e:
        return e.__str__()
        #return data.df_fileCntByFolder['fileCnt']

    return (len(data.df_fileCntByFolder.index))
    #return ("index len : " + str(len(data.df_fileCntByFolder.index)))


@xlw.func
def excelIO_UDF_getWorkSheetToDataFrame(srcExcel=None, srcSheet=None, row_colName=None):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    if srcExcel == None:
        srcExcel = xlw.Book.caller()
        ws = wb.sheets[srcSheet]

    if row_colName == None:
        row_colName = 10

    if srcSheet == None:
        srcSheet = '1996～2018'

    dfExcel = pd.read_excel(srcExcel,srcSheet, header=row_colName, index_col=None)

    print(dfExcel.head(20))


@xlw.func
def excelIO_UDF_df_fileCntByFolderItem(row, column):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    try:
        if column == 0:
            retValue = data.df_fileCntByFolder['path'].iloc[row]
        elif column == 1:
            retValue = data.df_fileCntByFolder['fileCnt'].iloc[row]
    except:

        if column == 0:
            retValue = 'Error path : ' #+ 'row:' + str(row) + 'column:' + str(column)
        elif column == 1:
            retValue = 'Error fileCnt : ' #+ 'row:' + str(row) + 'column:' + str(column)

    return retValue

@xlw.func
#   引数で指定されたフォルダ以下のフォルダツリーを作ってサクラエディタで表示する
#       topFolder=None：展開するフォルダツリートップ
#       workTextFile=：表示するサクラエディタのファイルあ（基本的にフルパス指定）
def excelIO_UDF_filrOpen_with_sakuraEditor(topFolder=None, workTextFile=None):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    if topFolder == None or topFolder == "":
        topFolder = r'Y:\旅日記\国内日記\2023'

    #フォルダツリーの生成（結果は、data.lst_generateToAddFolderTree　に保存される）
    folderTreeControl.generateFolderTree(path=topFolder, layer=0, is_last=False,
                                         indent_current=data.indent_tree)

    with open(os.path.join(topFolder, workTextFile), 'w', encoding='utf-8') as f:

        for item in data.lst_generateToAddFolderTree:
            omitChar = item.replace('―', '')
            f.write(omitChar)
            f.write('\n')

    folderTreeControl.textFilrOpen_with_sakuraEditor(topFolder, workTextFile)

def getSearchKeyWord():

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    wb = xlw.Book(data.currentExcel)
    macro = wb.macro(data.xlInterface + '.' + 'getSearchKeyword')
    keyWord = macro()

    return str(keyWord)
    #return str(r'桜')

@xlw.func
def excelIO_UDF_search(srcExcel=None, sheetsList=None):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    data.currentExcel = srcExcel

    df_result = searchItem.searchKeyword_in_folderTree(sheetsList)

    return df_result[data.dfColName_RowCnt]


def excelIO_UDF_search_old(srcExcel=None, sheetsList=None):

    data.currentExcel = srcExcel
    wb = xlw.Book(data.currentExcel)
    df_wb = pd.ExcelFile(wb.fullname)
    #wb = pd.ExcelFile(r'D:\git\diff_FolderTree_pythonProject\検索_python.xlsm')

    for currSheet in sheetsList:
        df_wb_Sht = pd.DataFrame()
        df_wb_Sht = df_wb.parse(currSheet)

        fllPath_searchData = os.path.join(data.dataFolderToSearch, (r'df_' + currSheet + r'.csv'))
        df_wb_Sht.to_csv(fllPath_searchData, encoding='utf-8')

#この関数は、python内部で呼ばれる前提
def excelIO_OutputSearchResults_to_Excel( list_results, currentSheetName=None ):

    testFunction.outputLogMessage_to_loggerObject(msgType=logging.INFO, prefix=r'[func]',
                                                  message=inspect.currentframe().f_code.co_name)
    wb = xlw.Book(data.currentExcel)
    macro = wb.macro(data.xlInterface + '.' + 'writeSearchResults_to_workSheet')
    macro(list_results, currentSheetName)

@xlw.func
def excelIO_UDF_test(srcExcel,row, col):

    data.currentExcel = srcExcel

    #srcExcelr = 'c:\Users\localuser\PycharmProjects\pythonProject\検索_python.xlsm'
    wb = xlw.Book(data.currentExcel)
    sheet = wb.sheets['debug_log']

    sheet.cells(row, col).value = excelIO_UDF_test.__name__


def debugFunction_list_toCsv(outputList, fullpath):
    f = open(fullpath, encoding='UTF-8', mode='w')
    writer = csv.writer(f)
    writer.writerow(outputList)
    f.close()
