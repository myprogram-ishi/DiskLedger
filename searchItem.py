
import os
import xlwings as xlw
import pandas as pd
import data
import interfaceExcel
import testFunction


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

            #先頭行も項目名ではないデータとして取り込む（先頭列も同じ扱いになっているように見える）
            df_wb_Sht = df_wb.parse(currSheet, header=None)

            fllPath_searchData = os.path.join(data.dataFolderToSearch, (r'df_' + currSheet + r'.csv'))
            df_Formatted = DataFrame_Formatting(df_wb_Sht,fllPath_searchData)

        return pd.DataFrame()

    else:
        df_wb_Sht = pd.DataFrame()
        df_wb_Sht = df_wb.parse(sheetsList)

        fllPath_searchData = os.path.join(data.dataFolderToSearch, (r'df_xlPrseFormat_' + sheetsList + r'.csv'))
        df_Formatted, lst_colName = DataFrame_Formatting(df_wb_Sht,fllPath_searchData)
        #df_Formatted.to_csv(os.path.join(data.dataFolderToSearch, 'df_Formatted.csv'))

        df_searchResults = searchKeyWord_in_dataFrame(df_Formatted, sheetsList)

        df_searchResults.to_csv(os.path.join(data.dataFolderToSearch, 'df_searchResults.csv'))
        df_searchResults[data.dfColName_RowCnt].to_csv(
            os.path.join(data.dataFolderToSearch, 'df_searchResults_dfColName_RowCnt.csv'))

        return df_searchResults


#########################################################################
#   データフレームを整形する
#   引数で渡されたデータフレームに列名、およびエクセルシートの行番号の列を付加する
#########################################################################
def DataFrame_Formatting( df_base=pd.DataFrame(), saveFullpath=None ):

    lst_new_colmName=[]

    df_base.to_csv(r'D:\git\diff_FolderTree_pythonProject\dataForSezrch\df_base_py_.csv')

    #先頭列削除
    #df_format = df_base.drop(df_base.columns[0], axis=1)
    #空白行削除
    df_format = df_base.dropna(how='all')
    #df_format.to_csv(r'D:\git\diff_FolderTree_pythonProject\dataForSezrch\df_format.csv')
    #列名の生成、付加
    lst_colmName = df_format.iloc[0]
    #for item in lst_colmName:
    #    print(item)
    #dfインデックスの0番目が、エクセルシートの１行目に相当するようなので、+2 する
    xlSht_topRow = df_base.shape[0] - df_format.shape[0] + 2

    #列名のリストを作る
    for cnt, item in enumerate(lst_colmName):
        item_split_yen = str(item).split('\\')
        frmt_cnt = '{0:04}'.format(cnt) #4桁固定でインデックス表記

        if item_split_yen[-1] != "":
            key_name = item_split_yen[-1]
        else:
            key_name = item_split_yen[-2]   #末尾が空白の場合は、ひとつ前

        df_Key = frmt_cnt + r'_' + key_name
        lst_new_colmName.append(df_Key)

    #列名を更新
    df_format = df_format.set_axis(lst_new_colmName, axis='columns')

    lst_rowNo = list(range(xlSht_topRow, xlSht_topRow + df_format.shape[0], 1))
    #ser_rowNo = pd.DataFrame(lst_rowNo, columns=['shtRowNo'])
    #シート状の行番号を追加する
    df_format[data.dfColName_RowCnt] = lst_rowNo

    if saveFullpath != None:
        df_format.to_csv(saveFullpath, encoding='utf-8')

    return df_format, lst_new_colmName

#########################################################################
#   データフレームから、キーワードを検索する
#########################################################################
def searchKeyWord_in_dataFrame( df_toBeSearched, currSheet ):

    ret_keyword = interfaceExcel.getSearchKeyWord()

    df_ret_dropna = df_toBeSearched.dropna(how='all')

    lst_colName = df_toBeSearched.columns

    for counter, col in enumerate(lst_colName):
        try:
            df_ret = df_toBeSearched[df_toBeSearched[col].str.contains(ret_keyword, na=False)]

        except AttributeError:
            #AttributeError: Can only use .str accessor with string values 対策の処理
            df_ret[col].to_csv(os.path.join(data.dataFolderToSearch, (r'df_ret_AttributeError' + col + r'.csv')))
            continue

        list_Result = df_ret[data.dfColName_RowCnt].to_list()
        list_Result.insert(0, col)
        interfaceExcel.excelIO_OutputSearchResults_to_Excel(list_Result, currSheet)

    return df_ret