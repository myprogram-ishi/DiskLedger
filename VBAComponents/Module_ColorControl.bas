Attribute VB_Name = "Module_ColorControl"
'*****************************************
'   セルの色を指定する
'*****************************************

Public Const bkClr_initialize As Integer = 0
Public Const bkClr_yew As Integer = 1
Public Const bkClr_org As Integer = 2
Public Const bkClr_redPurpl As Integer = 3
Public Const bkClr_red As Integer = 4
Public Const bkClr_blue As Integer = 5

'*****************************************************
'   セルの背景色、フォント色設定
'*****************************************************
Sub setBackColorAndFontColor(colorSelect As Integer)
Attribute setBackColorAndFontColor.VB_ProcData.VB_Invoke_Func = " \n14"

    Select Case colorSelect

        Case bkClr_yew
            '背景[黄色]　フォント[黒]
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 65535
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        
        Case bkClr_org
            '背景[橙]　フォント[黒]
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 49407
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With

        Case bkClr_redPurpl
            '背景[赤紫]　フォント[黒]
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16738047
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        
        Case bkClr_red
            '背景[赤]　フォント[白]
             With Selection.Font
                .Name = "ＭＳ Ｐゴシック"
                .FontStyle = "標準"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
            
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 255
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
        
        Case bkClr_blue
            '背景[青]　フォント[白]
            With Selection.Font
                .Name = "ＭＳ Ｐゴシック"
                .FontStyle = "標準"
                .Size = 11
                .Strikethrough = False
                .Superscript = False
                .Subscript = False
                .OutlineFont = False
                .Shadow = False
                .Underline = xlUnderlineStyleNone
                .ThemeColor = xlThemeColorDark1
                .TintAndShade = 0
                .ThemeFont = xlThemeFontMinor
            End With
        
            With Selection.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = 16750848
                .TintAndShade = 0
                .PatternTintAndShade = 0
            End With
    
    End Select

End Sub
'**********************************************************
'   セルの背景色、フォント色を初期化（白背景、黒フォント）
'**********************************************************
Sub initializeBackColorAndFontColor()
Attribute initializeBackColorAndFontColor.VB_ProcData.VB_Invoke_Func = " \n14"

    With Selection.Font
        .Name = "ＭＳ Ｐゴシック"
        .FontStyle = "標準"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    Rows("4:4").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 10092390
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("A4").Select
End Sub
