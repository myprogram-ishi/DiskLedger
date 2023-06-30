# これはサンプルの Python スクリプトです。

# Shift+F10 を押して実行するか、ご自身のコードに置き換えてください。
# Shift を2回押す を押すと、クラス/ファイル/ツールウィンドウ/アクション/設定を検索します。

import xlwings as xlw
import interfaceExcel

#import folderTreeControl as fldrctrl

def print_hi(name):
    # スクリプトをデバッグするには以下のコード行でブレークポイントを使用してください。
    print(f'Hi, {name}')  # Ctrl+F8を押すとブレークポイントを切り替えます。


# ガター内の緑色のボタンを押すとスクリプトを実行します。
if __name__ == '__main__':
    print_hi('PyCharm')

    test = 'c:\\test'

    print(test)

    interfaceExcel.excelIO_UDF_getWorkSheetToDataFrame(srcExcel='検索_python.xlsm', srcSheet='1996～2018', row_colName=9)

    #fldrctrl.getFolderTree(root="c:\\")

    #fldrctrl.getFolderFullpath(root="C:\\sourcrcode")

    xlw.serve()


# PyCharm のヘルプは https://www.jetbrains.com/help/pycharm/ を参照してください
