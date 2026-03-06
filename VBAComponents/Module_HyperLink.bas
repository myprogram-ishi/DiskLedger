Attribute VB_Name = "Module_HyperLink"

'*******************************************************************
'       ハイパーリンクを設定する
'*******************************************************************
Sub SetHyperLink(row As Integer, col As Integer, LinkAddress As String)

    Dim oHyperLink  As Hyperlink
    Dim Address As String
    Dim SeparateAdrs As Variant
Str1 = ActiveSheet.Name
    
    'トップフォルダのフォルダ名（基本的に、「旅日記」フォルダ）を取得
    topFolder = Left(LinkAddress, InStr(1, LinkAddress, "\") - 1)
    'HardDiskTopFolder：ドライブ名から、旅日記フォルダまでの階層のフォルダパス
    HardDiskTopFolder = UserForm_Top.ComboBox_BaseFolderTreeTop.Value
    
    ret = InStr(1, HardDiskTopFolder, topFolder)
    
    If ret = 0 Then
        Address = LinkAddress
    Else
        Address = Left(HardDiskTopFolder, ret - 1) & LinkAddress
    End If
    
    'Address = "C:\Users\masahito\Pictures\2018\0616～0617_トロリーバスと黒四\0616_扇沢から黒四往復"
    'Address = HardiskDrive & LinkAddress
    SeparateAdrs = Split(Address, "\")
    DispText = LinkAddress 'リンク名テキスト
 
 
    '// セルへのハイパーリンク設定
    Set oHyperLink = ActiveSheet.Cells(row, col).Hyperlinks.Add( _
        Anchor:=ActiveSheet.Cells(row, col), _
        Address:=Address, _
        ScreenTip:="g", _
        TextToDisplay:=DispText)

End Sub

