import os
import pathlib
import glob

import data


########################################################
#
########################################################
def expandFolderTree(root):
    print(expandFolderTree.__name__)

    #引数で指定された、フォルダを展開する先頭フォルダの位置を検出
    index = 0
    while True:
        if root in data.lst_branch[index]:
            expandTopIndex = index
            break

        index = index + 1

    branchNum = len(data.lst_branch)

    index = expandTopIndex
    result = False
    mark = ''
    branchMarkPos = -1
    expandFullPath = data.lst_branch[expandTopIndex]
    curr_ret, curr_mark, curr_branchMarkPos = isIncludeBranchMark(data.lst_branch[expandTopIndex])

    while index < branchNum:
        #次を見る（加算したインデックス値は保存しない）
        next_ret, next_mark, next_branchMarkPos = isIncludeBranchMark(data.lst_branch[index + 1])

        if curr_mark == '└' and curr_branchMarkPos != 0:
            EndOfPath = True
        elif curr_branchMarkPos == next_branchMarkPos:
            EndOfPath = True
        else:
            EndOfPath = False

        if EndOfPath == True:

            data.expandFullPath.append(expandFullPath)
            #初期化
            expandFullPath = data.lst_branch[expandTopIndex]
        else:
            index = index + 1
            expandFullPath = expandFullPath + r'\\' + data.lst_branch[index]

        curr_ret = next_ret
        curr_mark = next_mark
        curr_branchMarkPos = next_branchMarkPos

    return expandTopIndex

def isIncludeBranchMark(targetBranch):

    if '├' in targetBranch:
        ret = True
        mark = '├'
    elif '└' in targetBranch:
        ret = True
        mark = '└'
    else:
        #本来ここは通らない
        ret = False
        mark = ''

    if ret == True:
        pos = targetBranch.find(mark)
    else:
        pos = -1

    return ret, mark, pos

def getFolderFullpath(root):
    getFolderTree(path=root, layer=0, is_last=False, indent_current=data.indent_tree)

###############################################################################
##  フォルダツリー作成
##
##  Pythonでファイルのツリー構造を出力する
##  https://qiita.com/horisuke/items/389ec60407b3baf45f25#%E7%B5%90%E8%AB%96
###############################################################################
def getFolderTree(path, layer=0, is_last=False, indent_current=data.indent_tree):

    if not pathlib.Path(path).is_absolute():
        path = str(pathlib.Path(path).resolve())

    current = path.split('\\')[::-1][0]
    if layer == 0:
        # カレントディレクトリの表示
        print('<'+current+'>')
    else:
        branch = '└' if is_last else '├'
        print('{indent}{branch}{dirname}'.format(indent=indent_current, branch=branch, dirname=current))

    # 下の階層のパスを取得
    paths = [p for p in glob.glob(path+'/*') if os.path.isdir(p)]
    def is_last_path(i):
        return i == len(paths)-1

    # 再帰的に表示
    for i, p in enumerate(paths):

        indent_lower = indent_current
        if layer != 0:
            indent_lower += data.indent_tree if is_last else '│　'

        if os.path.isdir(p):
            getFolderTree(p, layer=layer+1, is_last=is_last_path(i), indent_current=indent_lower)


