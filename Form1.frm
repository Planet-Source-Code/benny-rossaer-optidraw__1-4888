VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
    ' Declare variables.
    Dim CX, CY, Limit, Radius   As Integer, Msg As String
    ScaleMode = vbPixels    ' Set scale to pixels.
    AutoRedraw = True ' Turn on AutoRedraw.
    Width = Height  ' Change width to match height.
    CX = ScaleWidth / 2 ' Set X position.
    CY = ScaleHeight / 2    ' Set Y position.
    Limit = CX  ' Limit size of circles.
    For Radius = 0 To Limit ' Set radius.
        Circle (CX, CY), Radius, RGB(Rnd * 255, Rnd * 255, Rnd * 255)

DoEvents    ' Yield for other processing.
    Next Radius
    Msg = "Choose OK to save the graphics from this form "
    Msg = Msg & "to a bitmap file."
    MsgBox Msg
    SavePicture Image, "TEST.BMP"   ' Save picture to file.
End Sub
