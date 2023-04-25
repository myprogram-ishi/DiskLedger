
###############################
#   文字列
###############################

currentExcel = ''

#エクセルシート名
shtName_main = 'メイン'
shtName_base = '基準'
shtName_dbgLog = 'debug_log'
shtName_pyDbgLog = 'python_log'
shtName_Expand = r'Expand'

treeTop = ':\\'
searchRootFolder = 'SPB_17.2'

xlInterface = 'Module_interfacePython'

getEndRowCout = 'getEndOfRowCount'

strFunc = "[func]"

indent_tree = " "

lst_generateToAddFolderTree = []
lst_branch = []
lst_expandFolderTreeBase = []
lst_expandFolderTreeTarget = []


def dataClear():
    lst_branch.clear()
    lst_expandFolderTreeTarget.clear()
    lst_generateToAddFolderTree.clear()
    lst_expandFolderTreeBase.clear()