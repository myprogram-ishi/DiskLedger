import os
import pandas as pd
import searchItem

import data

def test_DataFrame_Formatting():

    folder_df_def_in_csv = r'D:\git\diff_FolderTree_pythonProject\dataForSezrch'
    #file_df_def_in_csv = r'df_基準.csv'
    #file_df_def_in_csv = r'dfdf__基準.csv'
    #file_df_def_in_csv = r'df_base.csv'
    file_df_def_in_csv = r'df_base_for_test.csv'

    df_fullPath = os.path.join(folder_df_def_in_csv, file_df_def_in_csv)
    curr_df = pd.read_csv(df_fullPath)
    print(curr_df.head(10))

    df_fullPath = os.path.join(folder_df_def_in_csv, 'df__formated.csv')
    curr_df.to_csv(df_fullPath)

    df_sht, lst_colName = searchItem.DataFrame_Formatting(curr_df, df_fullPath)

    df_sht = df_sht.fillna(data.nanData)
    #print(df_sht[df_sht[lst_colName[0]].str.contains('桜')])

    pass
    #curr_df = curr_df.fillna(data.nanData)
    #curr_df.to_csv(os.path.join(folder_df_def_in_csv, r'dfdf__基準.csv'))
    #searchItem.DataFrame_Formatting(curr_df, df_fullPath)

def test_searchKeyWord_in_dataFrame():

    folder_df_def_in_csv = r'D:\git\diff_FolderTree_pythonProject\dataForSezrch'
    file_df_def_in_csv = r'df_基準.csv'

    df_fullPath = os.path.join(folder_df_def_in_csv, file_df_def_in_csv)
    curr_df = pd.read_csv(df_fullPath)
    curr_df = curr_df.fillna('space')

    print('--------- ilic -------------')
    print(curr_df.iloc[0, 0])
    print(curr_df.iloc[1, 0])

    print(curr_df.iloc[0, 1])
    print(curr_df.iloc[1, 1])

    print(curr_df.iloc[3, 0])
    print(curr_df.iloc[3, 1])

    print('----------------------')
    print(curr_df[0:10])
    print('----------------------')
    print(curr_df.iloc[0:10])

    #curr_df_dna = curr_df.dropna()
    curr_df.to_csv(os.path.join(folder_df_def_in_csv, r'curr_df_dna.csv'))

    #print('****************')
    #print(curr_df.head(10))
    #print('================')
    #print(curr_df_dna.head(10))

    searchItem.searchKeyWord_in_dataFrame(curr_df)

#リストをテキストファイルへ出力する
def test_output_list_to_textFile( folderName, fileName, outputListData ):

    if folderName == None:
        folderName = data.outoutFolder_list_txt

    outpuFulltPath = os.path.join(folderName,fileName)

    with open(outpuFulltPath,'w', encoding='utf-8') as f:
        for d in outputListData:
            f.write("%s\n" % d)
