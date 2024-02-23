import os
import pathlib
import glob
import subprocess
import xlwings as xlw

import pandas as pd
import data
import interfaceExcel
import testFunction

expandTopIndex = 0
branchNum = 0

#でバグシートの使用列番号
dbgRow_folderTreeControl = 1
dbgCol_folderTreeControl = 5

########################################################
#   フォルダツリー展開
########################################################
def expandFolderTree(root, col):

    interfaceExcel.excelIO_UDF_appendDataToLasRow(data.shtName_dbgLog, 1, expandFolderTree.__name__)

    #引数で指定された、フォルダを展開する先頭フォルダの位置を検出
    index = 0
    while True:
        if root in data.lst_branch[index] or data.treeTop in data.lst_branch[index]:
            expandTopIndex = index
            break

        index = index + 1
        #配列の末端まで到達したらループを抜ける
        if index >= len(data.lst_branch):
            break

    branchNum = len(data.lst_branch)

    expandFolderTree_from_endOfBranch(root, col)

    #expandFolderTree_from_BranchTop(root)

########################################################
#   フォルダツリーの末端から探す
########################################################
def expandFolderTree_from_endOfBranch(root, col):

    interfaceExcel.excelIO_UDF_appendDataToLasRow(data.shtName_dbgLog, 1, expandFolderTree_from_endOfBranch.__name__)

    rowDownCnt = len(data.lst_branch)  # エクセルシート↑の行番号なので、１から配列要素お個数分。
    lst_end = rowDownCnt - 1            #配列絵インデックスの末尾
    branchIndex = lst_end

    wb = xlw.Book(data.currentExcel)
    macro = wb.macro(data.xlInterface + '.' + 'getFullpathWriteStartRowForFileCount')
    destRow = macro()
    rowOffset = destRow - 1

    #配列のクリア     この関数の実行条件（タイミング）は考えなおさないといけない
    ##interfaceExcel.excelIO_UDF_clear_lst_expand_target()
    #data.lst_expandFolderTreeTarget.clear()

    while branchIndex > 0:

        ret_expandFullPath = expandFolderTree_for_OneBranch(data.lst_branch[branchIndex], branchIndex)

        #生成したフルパスを追加
        if ret_expandFullPath != None and ret_expandFullPath != "":
            data.lst_expandFolderTreeTarget.append(ret_expandFullPath)

        if data.outExpFolderTreeToWrkSht == True:
            interfaceExcel.excelIO_UDF_writeDataOnWorksheet(srcExcel=None, shtName=data.shtName_Expand,
                                                            row=(rowDownCnt + rowOffset), col=col,
                                                            outval=ret_expandFullPath)        #ハイパーリンクを付ける
        #ws = wb.sheets['base']
        #ws.cells(rowDownCnt + rowOffset, col).add_hyperlink(ret_expandFullPath, ws.cells(rowDownCnt + rowOffset,
        # col).value)

        rowDownCnt = rowDownCnt - 1

        #次のブラチのフルパス作成のための変数設定
        branchIndex = branchIndex - 1

    testFunction.test_output_list_to_textFile(None, r'lst_expandFolderTreeTarget.txt',
                                              data.lst_expandFolderTreeTarget)


########################################################
#   フォルダ一つ分のフルパス生成
########################################################
def expandFolderTree_for_OneBranch(TipBranch, branchIndex):
    # ブランチ名
    work_branchIndex = branchIndex

    expandFullPath = ''
    before_branchMarkPos = -1  # 負数で、初期値を表す

    # パスの末尾
    work_lst_branch = data.lst_branch[work_branchIndex]
    curr_ret, curr_mark, curr_branchMarkPos = isIncludeBranchMark(work_lst_branch)

    df = pd.DataFrame(data.lst_branch)
    df.to_csv(r'd:\data_lst_branch.txt', sep='\t', header=False, index=False)

    try:
        expandFullPath = work_lst_branch[curr_branchMarkPos + 1:]
    except:
        return None     #r'except00_expandFolderTree_for_OneBranch'

    #パスの末尾よりも上
    result, mark, pos = getPostionOfHoraizonalBar(expandFullPath)
    if result == True:
        expandFullPath = expandFullPath[pos + 1:]

    before_branchMarkPos = curr_branchMarkPos
    work_branchIndex = work_branchIndex - 1

    # ブランチにパスを付加していく
    while work_branchIndex >= 0:
        try:
            work_lst_branch = data.lst_branch[work_branchIndex]
            curr_ret, curr_mark, curr_branchMarkPos = isIncludeBranchMark(work_lst_branch)

            if curr_branchMarkPos < before_branchMarkPos:
                # 一つ上の階層
                upperBranch = work_lst_branch[curr_branchMarkPos + 1:]

                result, mark, pos = getPostionOfHoraizonalBar(upperBranch)
                if result == True:
                    upperBranch = upperBranch[pos + 1:]

                if upperBranch[-1] == '\\':
                    # 末尾に\がついているかいないかで処理を分ける
                    expandFullPath = upperBranch + expandFullPath
                else:
                    expandFullPath = upperBranch + '\\' + expandFullPath

                before_branchMarkPos = curr_branchMarkPos

            work_branchIndex = work_branchIndex - 1

        except:
            expandFullPath = None   #r'except01_expandFolderTree_for_OneBranch'
            break

    #検索対象の戦闘フォルダ以降を抽出する
    topPos = expandFullPath.find(data.searchRootFolder)
    expandFullPath = expandFullPath[topPos:]

    return expandFullPath

########################################################
#   フォルダツリーのトップから探す
########################################################
def expandFolderTree_from_BranchTop(root):
    print(expandFolderTree_from_BranchTop.__name__)

    index = expandTopIndex
    result = False
    mark = ''
    branchMarkPos = -1
    expandFullPath = data.lst_branch[expandTopIndex]
    curr_ret, curr_mark, curr_branchMarkPos = isIncludeBranchMark(data.lst_branch[expandTopIndex])

    while index < branchNum:
        #次を見る（加算したインデックス値は保存しない）
        next_ret, next_mark, next_branchMarkPos = isIncludeBranchMark(data.lst_branch[index + 1])

        #終端判定
        if curr_mark == '└' :
            if next_branchMarkPos < 0:
                EndOfPath = True
            elif next_branchMarkPos > curr_branchMarkPos:
                EndOfPath = False
            else:
                EndOfPath = True

        elif curr_branchMarkPos == next_branchMarkPos:
            EndOfPath = True
        else:
            EndOfPath = False

        #終端判定に基づく処理
        if EndOfPath == True:

            #data.expandFullPath.append(expandFullPath)
            #初期化
            expandFullPath = data.lst_branch[expandTopIndex]

            interfaceExcel.excelIO_UDF_appendDataToLasRow(r'Expand', 1, expandFullPath)
        else:
            index = index + 1
            expandFullPath = expandFullPath + '\\' + data.lst_branch[index]

        #次の処理の準備
        curr_ret = next_ret
        curr_mark = next_mark
        curr_branchMarkPos = next_branchMarkPos

    return expandTopIndex

########################################################
#   ブランチマークの有無および位置を検出する
#   分岐マーク：'├'　'└'
########################################################
def isIncludeBranchMark(targetBranch):

    try:
        if '├' in targetBranch:
            ret = True
            mark = '├'
        elif '└' in targetBranch:
            ret = True
            mark = '└'
        else:
            ret = False
            mark = ''

    except:
        ret = False
        mark = data.exceptBranchMark

    #ブランチ記号が見つかった場合、その位置を取得する
    if ret == True:
        pos = targetBranch.find(mark)
    else:
        pos = -1

    return ret, mark, pos

########################################################
#   横棒の有無および位置を検出する
#   横棒：'-'　'ー'
########################################################
def getPostionOfHoraizonalBar(targetBranch):

    try:
        if '─' in targetBranch:
            ret = True
            mark = '─'
        elif "―" in targetBranch:
            ret = True
            mark = "―"
        else:
            ret = False
            mark = ''
    except:
        ret = False
        mark = ''

    #ブランチ記号が見つかった場合、その位置を取得する
    if ret == True:

        pos = targetBranch.find(mark)
    else:
        pos = -1

    return ret, mark, pos


########################################################
#
########################################################
def getFolderFullpath(root):
    getFolderTree(path=str(root), layer=0, is_last=False, indent_current=data.indent_tree)

###############################################################################
##  フォルダツリー作成
##
##  Pythonでファイルのツリー構造を出力する
##  https://qiita.com/horisuke/items/389ec60407b3baf45f25#%E7%B5%90%E8%AB%96
###############################################################################
def generateFolderTree(path, layer=0, is_last=False, indent_current=data.indent_tree):

    if not pathlib.Path(path).is_absolute():
        path = str(pathlib.Path(path).resolve())

    current = path.split('\\')[::-1][0]
    if layer == 0:
        data.lst_generateToAddFolderTree.clear()
        # カレントディレクトリの表示
        #print('<'+current+'>')
        data.lst_generateToAddFolderTree.append(str(path))
        #data.lst_generateToAddFolderTree.append('<' + current + '>')
        #interfaceExcel.excelIO_UDF_appendDataToLasRow(shtName, col, setVal):
    else:
        branch = '└―' if is_last else '├―'

        branchWithMark = '{indent}{branch}{dirname}'.format(indent=indent_current, branch=branch, dirname=current)

        if r'/' in branchWithMark:
            branchWithMark = branchWithMark.replace(r'/', '')

        data.lst_generateToAddFolderTree.append(branchWithMark)

    # 下の階層のパスを取得
    paths = [p for p in glob.glob(path+'/*') if os.path.isdir(p)]
    def is_last_path(i):
        return i == len(paths)-1

    # 再帰的に表示
    for i, p in enumerate(paths):

        indent_lower = indent_current

        if layer != 0:
            indent_lower += data.indent_tree if is_last else '│ '

        if os.path.isdir(p):
            generateFolderTree(p, layer=layer+1, is_last=is_last_path(i), indent_current=indent_lower)


def textFilrOpen_with_sakuraEditor(folder=None, file=None):

    if folder == None or file == None or len(folder) == 0 or len(file) == 0:
        fullpath = r'D:\git\testData\sampleFolderTree.txt'
    else:
        fullpath = os.path.join(folder,file)

    subprocess.Popen([r'C:\Program Files (x86)\sakura\sakura.exe', fullpath])