import os
import pathlib
import glob

import data
import inerfaceExcel

########################################################
#   フォルダツリー展開
########################################################
def expandFolderTree(root, col):

    inerfaceExcel.excelIO_UDF_appendDataToLasRow(data.shtName_dbgLog, 1, expandFolderTree.__name__)

    #引数で指定された、フォルダを展開する先頭フォルダの位置を検出
    index = 0
    while True:
        if root in data.lst_branch[index] or data.treeTop in data.lst_branch[index]:
            expandTopIndex = index
            break

        index = index + 1

    branchNum = len(data.lst_branch)

    expandFolderTree_from_BranchTip(root, col)

    #expandFolderTree_from_BranchTop(root)

########################################################
#   フォルダツリーの末端から探す
########################################################
def expandFolderTree_from_BranchTip(root, col):

    inerfaceExcel.excelIO_UDF_appendDataToLasRow(data.shtName_dbgLog, 1, expandFolderTree_from_BranchTip.__name__)

    rowDownCnt = len(data.lst_branch)   #エクセルシート↑の行番号なので、１から配列要素お個数分。
    lst_end = rowDownCnt - 1            #配列絵インデックスの末尾
    branchIndex = lst_end
    before_ret = 0
    before_mark = 0
    before_branchMarkPos = -1   #府数設定は、初期状態を意味する

    while branchIndex > 0:
        #ブランチ名
        work_expandFullPath = data.lst_branch[branchIndex]
        dummyRet, dummyMark, topPos = getPostionOfHoraizonalBar(work_expandFullPath)

        expandFullPath = work_expandFullPath[topPos + 1:]
        work_branchIndex = branchIndex  #フォルダの一番下の階層（初期値）
        before_branchMarkPos = -1       #負数で、初期値を表す

        #ブランチにパスを付加していく
        while work_branchIndex >= 0:

            work_lst_branch = data.lst_branch[work_branchIndex]
            curr_ret, curr_mark, curr_branchMarkPos = isIncludeBranchMark(work_lst_branch)

            #フォルダの一番下の階層
            work_expandFullPath = expandFullPath

            r_ret, r_mark, topPos_expnd = isIncludeBranchMark(work_expandFullPath)

            #topPos_lstbrnch = work_lst_branch.find('─')
            dummyRet, dummyMark, topPos_lstbrnch = getPostionOfHoraizonalBar(work_expandFullPath)

            if topPos_lstbrnch < 0:
                topPos_lstbrnch = work_lst_branch.find('└')

            if before_branchMarkPos < 0:
                #最初の１回目
                expandFullPath = work_expandFullPath[topPos_expnd+1:]
                before_branchMarkPos = curr_branchMarkPos

            elif curr_branchMarkPos == before_branchMarkPos:
                currentBranch = data.lst_branch[branchIndex]
                topPos_lstbrnch = currentBranch.find('├')
                if topPos_lstbrnch >= 0:
                    #expandFullPath = 'branchIndex = ' + str(branchIndex) + 'work_branchIndex = ' + str(work_branchIndex)
                    expandFullPath = currentBranch[topPos_lstbrnch+1:]
                else:
                    topPos_lstbrnch = currentBranch.find('└')
                    if topPos_lstbrnch >= 0:
                        expandFullPath =  currentBranch[topPos_lstbrnch+1:]
                    else:
                        expandFullPath = '???' + work_lst_branch

                result, dummyMark, topPos = getPostionOfHoraizonalBar(expandFullPath)
                if result == True:
                    expandFullPath = expandFullPath[topPos+1:]

            elif curr_branchMarkPos < before_branchMarkPos:
                upperBranch = work_lst_branch[topPos_lstbrnch + 1:]

                #末尾に\がついているかいないかで処理を分ける
                if upperBranch[-1] == '\\':
                    expandFullPath = upperBranch + work_expandFullPath[topPos_expnd + 1:]
                else:
                    expandFullPath = upperBranch + '\\' + work_expandFullPath[topPos_expnd+1:]

                expandFullPath = expandFullPath + ' : branchIndex = ' + str(branchIndex)

                result, dummyMark, topPos = getPostionOfHoraizonalBar(expandFullPath)
                if result == True:
                    expandFullPath = expandFullPath[topPos+1:]

                before_branchMarkPos = curr_branchMarkPos

            work_branchIndex = work_branchIndex - 1

        #生成したフルパスを追加
        data.lst_expandFolderTreeTarget.append(expandFullPath)
        #inerfaceExcel.excelIO_UDF_appendDataToLasRow(data.shtName_Expand, 1, expandFullPath)
        inerfaceExcel.excelIO_UDF_outputdebugLog(srcExcel=None, shtName=data.shtName_Expand, row=rowDownCnt, col=col, outval=expandFullPath)
        rowDownCnt = rowDownCnt - 1

        #次のブラチのフルパス作成のための変数設定
        branchIndex = branchIndex -1

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

            inerfaceExcel.excelIO_UDF_appendDataToLasRow(r'Expand', 1, expandFullPath)
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
        #inerfaceExcel.excelIO_UDF_appendDataToLasRow(shtName, col, setVal):
    else:
        branch = '└' if is_last else '├'
        #print('{indent}{branch}{dirname}'.format(indent=indent_current, branch=branch, dirname=current))
        data.lst_generateToAddFolderTree.append('{indent}{branch}{dirname}'.format(indent=indent_current, branch=branch, dirname=current))

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
