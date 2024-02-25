
import pandas as pd

dbg_row = 5

#フォルダツリー展開結果をワークシートに出す：True / 出さない:False
outExpFolderTreeToWrkSht = True
lst_expanfFolderTree = []

###############################
#   文字列
###############################
currentExcel = ''

nanData = r'Empty'

#エクセルシート名
shtName_main = 'メイン'
shtName_base = '基準'
shtName_dbgLog = 'debug_log'
shtName_pyDbgLog = 'python_log'
shtName_Expand = r'Expand'

dfColName_RowCnt = r'shtRowNo'

exceptBranchMark = r'expectBranchMark'

treeTop = ':\\'

outoutFolder_debug = r'D:\git\diff_FolderTree_pythonProject\debug'
#outoutFolder_df_tocsv = r'D:\git\diff_FolderTree_pythonProject\csv'
outoutFolder_df_tocsv = outoutFolder_debug + r'\csv'
outoutFolder_list_txt = outoutFolder_debug + r'\list'

dataFolderToSearch = r'D:\git\diff_FolderTree_pythonProject\dataForSezrch'

searchRootFolder = '旅日記'   #'SPB_17.2'

xlInterface = 'Module_interfacePython'

getEndRowCout = 'getEndOfRowCount'

strFunc = "[func]"

indent_tree = " "

cols_dfFileCnt = ['fileCnt', 'path']
df_fileCntByFolder = pd.DataFrame(columns=cols_dfFileCnt)

lst_generateToAddFolderTree = []
lst_branch = []
lst_expandFolderTreeBase = []
lst_expandFolderTreeTarget = []
lst_searchResults = []

def dataClear():
    lst_branch.clear()
    lst_expandFolderTreeTarget.clear()
    lst_generateToAddFolderTree.clear()
    lst_expandFolderTreeBase.clear()
    lst_searchResults.clear()
