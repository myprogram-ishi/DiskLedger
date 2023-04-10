import os
import pathlib
import glob

from anytree import Node, RenderTree
import data

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





###########################################
#   フォルダフルパスの取得
###########################################
def getFolderTree_test(root):

    print(data.strFunc, 'generateFolderpath : ', root)

    tpl_retWalk = os.walk(root)

    parentFolder = Node(root, parent=None)

    for index, [dir,subDirs,files] in enumerate(tpl_retWalk):
        print("--------------------",index, "--------------------")

        print(dir)
        print(subDirs)

        spltDir = dir.split('\\')

        parentBranch = parentFolder
        for index, branch in enumerate(spltDir[2:]):
            #print(branch)
            nodeBranch = Node(branch, parent=parentBranch)
            parentBranch = nodeBranch
            #parentBranch = Node(root, parent=None)

    #for pre, fill, node in RenderTree(parentFolder):
    #    print("%s%s" % (pre, node.name))