import os
import pathlib
import glob
import xlwings as xlw

import data
import interfaceExcel

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

    expandFolderTree_from_BranchTip(root, col)

    #expandFolderTree_from_BranchTop(root)

########################################################
#   フォルダツリーの末端から探す
########################################################
def expandFolderTree_from_BranchTip(root, col):

    interfaceExcel.excelIO_UDF_appendDataToLasRow(data.shtName_dbgLog, 1, expandFolderTree_from_BranchTip.__name__)

    rowDownCnt = len(data.lst_branch)  # エクセルシート↑の行番号なので、１から配列要素お個数分。
    lst_end = rowDownCnt - 1            #配列絵インデックスの末尾
    branchIndex = lst_end

    wb = xlw.Book(data.currentExcel)
    macro = wb.macro(data.xlInterface + '.' + 'getFullpathWriteStartRowForFileCount')
    destRow = macro()
    rowOffset = destRow - 1
    #if destRow > rowDownCnt:
    #    rowOffset  = rowDownCnt + 1
    #else:
    #    rowOffset = 0

    while branchIndex > 0:

        ret_expandFullPath = expandFolderTree_for_OneBranch(data.lst_branch[branchIndex], branchIndex)

        #生成したフルパスを追加
        data.lst_expandFolderTreeTarget.append(ret_expandFullPath)
        #interfaceExcel.excelIO_UDF_appendDataToLasRow(data.shtName_Expand, 1, expandFullPath)
        interfaceExcel.excelIO_UDF_outputdebugLog(srcExcel=None, shtName=data.shtName_Expand,
                                                  row=(rowDownCnt + rowOffset), col=col, outval=ret_expandFullPath)
        rowDownCnt = rowDownCnt - 1

        #次のブラチのフルパス作成のための変数設定
        branchIndex = branchIndex - 1

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
    expandFullPath = work_lst_branch[curr_branchMarkPos + 1:]
    result, mark, pos = getPostionOfHoraizonalBar(expandFullPath)
    if result == True:
        expandFullPath = expandFullPath[pos + 1:]
    before_branchMarkPos = curr_branchMarkPos
    work_branchIndex = work_branchIndex - 1

    # ブランチにパスを付加していく
    while work_branchIndex >= 0:

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

    return expandFullPath


def _expandFolderTree_for_OneBranch(TipBranch, branchIndex):
    # ブランチ名

    work_branch = TipBranch
    work_branchIndex = branchIndex
    dummyRet, dummyMark, topPos = getPostionOfHoraizonalBar(work_branch)

    expandFullPath = work_branch[topPos + 1:]
    before_branchMarkPos = -1  # 負数で、初期値を表す

    # ブランチにパスを付加していく
    while work_branchIndex >= 0:

        work_lst_branch = data.lst_branch[work_branchIndex]
        curr_ret, curr_mark, curr_branchMarkPos = isIncludeBranchMark(work_lst_branch)

        debg_work_lst_branch = work_lst_branch

        # topPos_lstbrnch = work_lst_branch.find('─')
        dummyRet, dummyMark, topPos = getPostionOfHoraizonalBar(work_lst_branch)
        if topPos > 0:
            work_lst_branch = work_lst_branch[topPos + 1:]

        #if topPos < 0:
        #    topPos = work_lst_branch.find('└')
        #    work_lst_branch = work_lst_branch[topPos + 1:]

        if before_branchMarkPos < 0:
            # 最初の１回目
            expandFullPath = work_lst_branch[curr_branchMarkPos + 1:]
            expandFullPath = expandFullPath + '<' + str(work_branchIndex) + '>' + '[' + str(work_branchIndex) + ']'
            before_branchMarkPos = curr_branchMarkPos

        elif curr_branchMarkPos == before_branchMarkPos:
            currentBranch = data.lst_branch[branchIndex]
            topPos_lstbrnch = currentBranch.find('├')
            if topPos_lstbrnch >= 0:
                # expandFullPath = 'branchIndex = ' + str(branchIndex) + 'work_branchIndex = ' + str(work_branchIndex)
                expandFullPath = currentBranch[topPos_lstbrnch + 1:]
            else:
                topPos_lstbrnch = currentBranch.find('└')
                if topPos_lstbrnch >= 0:
                    expandFullPath = currentBranch[topPos_lstbrnch + 1:]
                else:
                    expandFullPath = '???' + work_lst_branch

            result, dummyMark, topPos = getPostionOfHoraizonalBar(expandFullPath)
            if result == True:
                expandFullPath = expandFullPath[topPos + 1:]

            expandFullPath = expandFullPath + '{' + str(work_branchIndex) + '}'

        elif curr_branchMarkPos < before_branchMarkPos:
            upperBranch = work_lst_branch[curr_branchMarkPos + 1:]
            #upperBranch = work_lst_branch

            if len(upperBranch) <= 0:
                break
            if upperBranch[-1] == '\\':
                # 末尾に\がついているかいないかで処理を分ける
                expandFullPath = upperBranch + expandFullPath
            else:
                expandFullPath = upperBranch + '\\' + expandFullPath

            expandFullPath = expandFullPath + '[[' + debg_work_lst_branch + ' : ' + str(curr_branchMarkPos) + ']]'

            result, dummyMark, topPos = getPostionOfHoraizonalBar(expandFullPath)
            if result == True:
                expandFullPath = expandFullPath[topPos + 1:]

            before_branchMarkPos = curr_branchMarkPos

        work_branchIndex = work_branchIndex - 1

    return expandFullPath

########################################################
#   フォルダツリーのトップから探す
########################################################
def expandFolderTree_from_BranchTop(root):
    print(expandFolderTree.__name__)

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

            data.expandFullPath.append(expandFullPath)
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

    if '├' in targetBranch:
        ret = True
        mark = '├'
    elif '└' in targetBranch:
        ret = True
        mark = '└'
    else:
        ret = False
        mark = ''

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

    if '─' in targetBranch:
        ret = True
        mark = '─'
    elif "―" in targetBranch:
        ret = True
        mark = "―"
    else:
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
        #print('{indent}{branch}{dirname}'.format(indent=indent_current, branch=branch, dirname=current))
        data.lst_generateToAddFolderTree.append(
            '{indent}{branch}{dirname}'.format(indent=indent_current, branch=branch, dirname=current))

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
