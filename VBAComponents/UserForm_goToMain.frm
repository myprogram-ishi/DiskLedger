VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm_goToMain 
   Caption         =   "ÉÅÉCÉìÇ…ñﬂÇÈ"
   ClientHeight    =   1680
   ClientLeft      =   168
   ClientTop       =   664
   ClientWidth     =   2776
   OleObjectBlob   =   "UserForm_goToMain.frx":0000
End
Attribute VB_Name = "UserForm_goToMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton_gotoMain_Click()

    Sheets("ÉÅÉCÉì").Select
    
    Unload Me

End Sub

Private Sub CommandButton1_Click()

    Unload Me

End Sub
