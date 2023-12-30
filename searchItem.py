
import os
import xlwings as xlw
import pandas as pd
import data
import interfaceExcel

########################################################
#   エクセルシートをデータフレームに入れて、
#   そのデータフレームにキーワードが含まれるかどうかを検索する
########################################################
def searchKeyword_in_folderTree(sheetsList):

    wb = xlw.Book(data.currentExcel)
    df_wb = pd.ExcelFile(wb.fullname)
    #wb = pd.ExcelFile(r'D:\git\diff_FolderTree_pythonProject\検索_python.xlsm')

    if type(sheetsList) is list:
        for currSheet in sheetsList:
            df_wb_Sht = pd.DataFrame()
            df_wb_Sht = df_wb.parse(currSheet)

            fllPath_searchData = os.path.join(data.dataFolderToSearch, (r'df_' + currSheet + r'.csv'))
            df_wb_Sht.to_csv(fllPath_searchData, encoding='utf-8')
    else:
        df_wb_Sht = pd.DataFrame()
        df_wb_Sht = df_wb.parse(sheetsList)

        fllPath_searchData = os.path.join(data.dataFolderToSearch, (r'dfdf__' + sheetsList + r'.csv'))
        df_wb_Sht.to_csv(fllPath_searchData, encoding='utf-8')

        #searchKeyWord_in_dataFrame(df_wb_Sht, currSheet)

def searchKeyWord_in_dataFrame( df_toBeSearched, currSheet ):

    ret_keyword = interfaceExcel.getSearchKeyWord

    df_ret_dropna = df_toBeSearched.dropna(how='all')

    df_ret = df_toBeSearched[df_toBeSearched[0].str.contains(ret_keyword, na=False)]
    df_ret = df_ret.dropna(how='all')
    df_ret.to_csv(os.path.join(data.dataFolderToSearch, '___' + currSheet))
