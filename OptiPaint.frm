VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.1#0"; "COMCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.1#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.1#0"; "RICHTX32.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000000&
   Caption         =   "OptiDraw"
   ClientHeight    =   8310
   ClientLeft      =   285
   ClientTop       =   165
   ClientWidth     =   11880
   DrawStyle       =   5  'Transparent
   Icon            =   "OptiPaint.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "OptiPaint.frx":030A
   ScaleHeight     =   8310
   ScaleWidth      =   11880
   Begin VB.CommandButton Command1 
      Height          =   375
      Left            =   120
      Picture         =   "OptiPaint.frx":AA1C
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4080
      Width           =   375
   End
   Begin VB.CommandButton cmdTekst 
      Height          =   375
      Left            =   120
      Picture         =   "OptiPaint.frx":B1EA
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   3720
      Width           =   375
   End
   Begin VB.CommandButton cmdDrop 
      Height          =   375
      Left            =   120
      Picture         =   "OptiPaint.frx":B596
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3360
      Width           =   375
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   495
      Left            =   0
      TabIndex        =   12
      Top             =   360
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      _Version        =   327680
      Enabled         =   -1  'True
      TextRTF         =   $"OptiPaint.frx":B966
   End
   Begin ComctlLib.Slider PenSlider 
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   120
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   327680
      BorderStyle     =   1
      LargeChange     =   1
      Min             =   1
      Max             =   50
      SelectRange     =   -1  'True
      SelStart        =   1
      TickStyle       =   1
      Value           =   1
   End
   Begin VB.CommandButton cmdOpvullen 
      Height          =   375
      Left            =   120
      Picture         =   "OptiPaint.frx":BA37
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3000
      Width           =   375
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   240
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   327680
   End
   Begin VB.ComboBox stijl 
      Height          =   315
      Left            =   3480
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   720
      Width           =   1695
   End
   Begin VB.PictureBox pCol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1140
      Left            =   240
      MouseIcon       =   "OptiPaint.frx":BCCE
      MousePointer    =   2  'Cross
      Picture         =   "OptiPaint.frx":BFD8
      ScaleHeight     =   1110
      ScaleWidth      =   2145
      TabIndex        =   6
      Top             =   0
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      DrawStyle       =   3  'Dash-Dot
      FillColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   6735
      Left            =   600
      MousePointer    =   2  'Cross
      ScaleHeight     =   449
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   745
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1200
      Width           =   11175
      Begin VB.Shape SelectShape 
         BorderStyle     =   3  'Dot
         Height          =   735
         Left            =   120
         Top             =   1800
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Shape Shape2 
         Height          =   615
         Left            =   240
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000007&
         Visible         =   0   'False
         X1              =   48
         X2              =   96
         Y1              =   24
         Y2              =   48
      End
   End
   Begin VB.CommandButton cmdCirkel 
      Height          =   375
      Left            =   120
      Picture         =   "OptiPaint.frx":13CFA
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2640
      Width           =   375
   End
   Begin VB.CommandButton cmdRechthoekgevuld 
      Height          =   375
      Left            =   120
      Picture         =   "OptiPaint.frx":13F2D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2280
      Width           =   375
   End
   Begin VB.CommandButton cmdRechthoek 
      Height          =   375
      Left            =   120
      Picture         =   "OptiPaint.frx":1418F
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1920
      Width           =   375
   End
   Begin VB.CommandButton cmdLijn 
      Height          =   375
      Left            =   120
      Picture         =   "OptiPaint.frx":143FB
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1560
      Width           =   375
   End
   Begin VB.CommandButton cmdTekenen 
      Height          =   375
      Left            =   120
      Picture         =   "OptiPaint.frx":1460A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1200
      Width           =   375
   End
   Begin ComctlLib.Slider StraalSlider 
      Height          =   375
      Left            =   8760
      TabIndex        =   15
      Top             =   360
      Visible         =   0   'False
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   327680
      BorderStyle     =   1
      LargeChange     =   1
      Min             =   1
      Max             =   200
      SelectRange     =   -1  'True
      SelStart        =   1
      TickStyle       =   1
      Value           =   1
   End
   Begin VB.Label lblStraal 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11160
      TabIndex        =   17
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Radius:"
      Height          =   255
      Left            =   8760
      TabIndex        =   16
      Top             =   120
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   7080
      Stretch         =   -1  'True
      Top             =   240
      Width           =   1455
   End
   Begin VB.Label lblModus 
      BackStyle       =   0  'Transparent
      Caption         =   "Mode:"
      Height          =   615
      Left            =   5280
      TabIndex        =   11
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      Height          =   375
      Left            =   6360
      TabIndex        =   8
      Top             =   240
      Width           =   975
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      FillStyle       =   0  'Solid
      Height          =   975
      Left            =   2640
      Top             =   120
      Width           =   495
   End
   Begin VB.Menu bestan 
      Caption         =   "&File"
      Begin VB.Menu nieuw 
         Caption         =   "&New image"
         Shortcut        =   ^N
      End
      Begin VB.Menu openen 
         Caption         =   "&Open"
         Begin VB.Menu openbestand 
            Caption         =   "&Image from disk..."
            Shortcut        =   ^O
         End
      End
      Begin VB.Menu saveas 
         Caption         =   "&Save as..."
         Shortcut        =   ^S
      End
      Begin VB.Menu dezdez 
         Caption         =   "-"
      End
      Begin VB.Menu printimage 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu pop 
         Caption         =   "-"
      End
      Begin VB.Menu einde 
         Caption         =   "&Exit"
      End
   End
   Begin VB.Menu bewerken 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "&Cut image"
         Shortcut        =   ^X
      End
      Begin VB.Menu copy 
         Caption         =   "&Copy image"
         Shortcut        =   ^C
      End
      Begin VB.Menu paste 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
   End
   Begin VB.Menu afbeelding 
      Caption         =   "&Image"
      Begin VB.Menu groottewijzigen 
         Caption         =   "&Resize field"
      End
      Begin VB.Menu juyjuy 
         Caption         =   "-"
      End
      Begin VB.Menu achtergrond 
         Caption         =   "&Background..."
      End
      Begin VB.Menu negative 
         Caption         =   "&Negative image"
      End
   End
   Begin VB.Menu helpme 
      Caption         =   "&Help"
      Begin VB.Menu info 
         Caption         =   "&Info..."
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'API CALL VOOR OPVULLEN:
Private Declare Function ExtFloodFill Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long, ByVal wFillType As Long) As Long
'VARIABLEN VOOR OPVULLEN:
Dim X1, Y1
Dim draw
Dim temp
Dim o, p As Integer


Private Xpos(99) As Single, Ypos(99) As Single
Private red(99) As Integer, green(99) As Integer, blue(99) As Integer
Dim drawing As Boolean
Const parts = 40

Private Type pointapi
   x As Double
   y As Double
End Type

Option Explicit

Dim colpressed As Boolean  'same as above but for the color chooser
Dim a, b, modus As Integer
Dim point1 As pointapi
Dim point2 As pointapi
Dim tekenlijn As Integer
Dim tekenrechthoek As Integer
Dim selecteer As Integer
Dim gebiedgeselecteerd As Integer








Private Sub achtergrond_Click()

   ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    
    ' Display the Open dialog box
    CommonDialog1.ShowColor
    
    ' Display name of selected file

   Picture1.BackColor = CommonDialog1.Color
   
    
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub cmdCirkel_Click()
modus = 5

StraalSlider.Visible = True
Label2.Visible = True
lblStraal.Visible = True

lblModus.Caption = "Modus:" & vbCrLf & "DRAW CIRCLE"

End Sub

Private Sub cmdDrop_Click()
modus = 7
lblModus.Caption = "Modus:" & vbCrLf & "DROP TOOL"

StraalSlider.Visible = False
Label2.Visible = False
lblStraal.Visible = False

End Sub

Private Sub cmdLijn_Click()
modus = 2
lblModus.Caption = "Modus:" & vbCrLf & "DRAW LINE"

StraalSlider.Visible = False
Label2.Visible = False
lblStraal.Visible = False
End Sub

Private Sub cmdOpvullen_Click()
modus = 6
lblModus.Caption = "Modus:" & vbCrLf & "FILL"

StraalSlider.Visible = False
Label2.Visible = False
lblStraal.Visible = False

End Sub

Private Sub cmdRechthoek_Click()
modus = 3
lblModus.Caption = "Modus:" & vbCrLf & "DRAW RECTACLE"

StraalSlider.Visible = False
Label2.Visible = False
lblStraal.Visible = False
End Sub

Private Sub cmdRechthoekgevuld_Click()
modus = 4
lblModus.Caption = "Modus:" & vbCrLf & "DRAW FILLED RECTACLE"

StraalSlider.Visible = False
Label2.Visible = False
lblStraal.Visible = False
End Sub

Private Sub cmdSelect_Click()
StraalSlider.Visible = False
Label2.Visible = False
lblStraal.Visible = False

modus = 9
lblModus.Caption = "Modus:" & vbCrLf & "SELECT"

End Sub

Private Sub cmdTekenen_Click()
modus = 1
lblModus.Caption = "Modus:" & vbCrLf & "FREE-HAND"

StraalSlider.Visible = False
Label2.Visible = False
lblStraal.Visible = False

End Sub



Private Sub cmdTekst_Click()

StraalSlider.Visible = False
Label2.Visible = False
lblStraal.Visible = False

modus = 8
lblModus.Caption = "Modus:" & vbCrLf & "ADD TEXT"

End Sub

Private Sub Command1_Click()
StraalSlider.Visible = False
Label2.Visible = False
lblStraal.Visible = False
modus = 10 ' <-------------------------------
lblModus.Caption = "Modus:" & vbCrLf & "INSTANT ART"
End Sub

Private Sub Command2_Click()
modus = 9
End Sub

Private Sub copy_Click()
Clipboard.Clear
Clipboard.SetData Picture1.Image, 2

End Sub

Private Sub cut_Click()

If gebiedgeselecteerd = 0 Then

Clipboard.Clear
Clipboard.SetData Picture1.Image, 2

Picture1.Picture = LoadPicture("")

End If
' __________________________________________________

If gebiedgeselecteerd = 1 Then
Clipboard.Clear


Clipboard.SetData Picture1.Image, 2

'Clipboard.SetData


'Picture1.Picture = LoadPicture("")

End If

End Sub



Private Sub einde_Click()
End

End Sub

Private Sub Form_Load()

' -----------
 Dim rstep As Single, gstep As Single, bstep As Single, z As Integer
 Randomize Timer
  
  Dim r, g, b As Integer
  
  
 r = Rnd * 255: g = Rnd * 255: b = Rnd * 255
 
 rstep = (r - 255) / parts: gstep = (g - 255) / parts: bstep = (b - 255) / parts
 For z = 0 To parts - 1
  red(z) = 255 + z * rstep
  green(z) = 255 + z * gstep
  blue(z) = 255 + z * bstep
 
  Xpos(z) = (Picture1.Width / 15) / 2
  Ypos(z) = (Picture1.Height / 15) / 2
 Next z
 
 ' ------------

'Picture1.hDC

Label1.Caption = PenSlider.Value
WindowState = vbMaximized


stijl.AddItem "Vast"
stijl.AddItem "Lijntjes"
stijl.AddItem "Puntjes"
stijl.AddItem "Lijn - punt"
stijl.AddItem "Lijn - punt - punt"
stijl.AddItem "Transparant"

stijl.ListIndex = "0"


Picture1.DrawStyle = stijl.ListIndex




modus = 1


On Error GoTo fout

If Len(Command$) > 0 Then
 Picture1.Picture = LoadPicture(Command$)
End If

Exit Sub

fout:
    MsgBox "Er is de volgende opdrachtparameter meegegeven:" & vbCrLf & vbCrLf & Command$ & vbCrLf & vbCrLf & "OptiDraw kan deze opdracht niet verwerken.  Het is mogelijk dat het bestand niet bestaat of beschadigd is.", vbCritical, Command$
    Exit Sub
    

End Sub








Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)


'ForeColor = QBColor(Combo2.ListIndex)


If modus = 2 Then

If a <> 0 And b <> 0 Then Line (a, b)-(x, y)
a = 0
b = 0
End If


If modus = 3 Then
If a <> 0 And b <> 0 Then Line (a, b)-(x, y), , B
a = 0
b = 0
End If

If modus = 4 Then
If a <> 0 And b <> 0 Then Line (a, b)-(x, y), , BF

a = 0
b = 0
End If


End Sub




Private Sub groottewijzigen_Click()
Dim lengte, hoogte As Integer

On Error GoTo fout

lengte = InputBox("Field length in pixels:", "Lengte wijzigen", ScaleX(Picture1.Width, 1, 3))
hoogte = InputBox("Field height in pixels:", "Hoogte wijzigen", ScaleX(Picture1.Height, 1, 3))


Picture1.Width = ScaleX(lengte, 3, 1)
Picture1.Height = ScaleX(hoogte, 3, 1)

Exit Sub


fout:
MsgBox "Er is een onbekende fout opgetreden.", vbCritical, "Fout"
Exit Sub


End Sub

Private Sub helpsubj_Click()
Dim filetorun As String
filetorun = App.Path & "\optinet.exe " & App.Path & "\help\optidr\index.htm"
Shell filetorun, 1

End Sub

Private Sub info_Click()
MsgBox "OptiDraw 1.0" & vbCrLf & vbCrLf & "Benny Rossaer" & vbCrLf & "1999" & vbCrLf & vbCrLf & "Part of Opti2000" & vbCrLf & vbCrLf & vbCrLf & "Instant Art:" & vbCrLf & "Sneechy (ICQ: 41829463) " & vbCrLf & vbCrLf & "For more (Dutch) Opti2000 applications, check out my website at http://benny.w3site.com.  OptiDraw is currently the only program that I've translated to English, all the other programs on my site are in Dutch." & vbCrLf & vbCrLf & "Why don't you try my game, 'Puke Invaders: Second Mission'?  It's a really poor Space Invaders / Galaga clone, but it's well worth a look.  Go to Planet Source Code (www.planet-source-code.com) and search for 'puke'.", vbInformation, "Info"



End Sub

Private Sub negative_Click()
Picture1.DrawMode = 6
Dim i, d As Integer

Form1.Caption = "APPLYING FILTER, PLEASE WAIT..."

For i = 0 To Picture1.ScaleWidth - 1 Step 1

    For d = 0 To Picture1.ScaleHeight - 1 Step 1

        SetPixel Picture1.hdc, i, d, GetPixel(Picture1.hdc, i, d)

    Next


Next

Picture1.DrawMode = 13
Picture1.Refresh

Form1.Caption = "OptiDraw"



End Sub

Private Sub nieuw_Click()
Picture1.Picture = LoadPicture()
Picture1.Width = 11175
Picture1.Height = 6735


End Sub

Private Sub openbestand_Click()

   ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
    
    
    CommonDialog1.Filter = "Windows bitmap (*.BMP)|*.bmp|Compuserve GIF" & _
    "(*.gif)|*.gif|JPEG Filter|*.jpg|Alle bestanden|*.*"
        
    
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
    CommonDialog1.ShowOpen
    ' Display name of selected file

   Picture1.Picture = LoadPicture(CommonDialog1.filename)
    
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub


End Sub

Private Sub openclipart_Click()
frmClipart.Show

End Sub


Private Sub paste_Click()



Picture1.Picture = Clipboard.GetData(2)




End Sub

Private Sub PenSlider_Scroll()
'MsgBox PenSlider.Container
 'MsgBox PenSlider.Value

Label1.Caption = PenSlider.Value

Picture1.DrawWidth = PenSlider.Value

If PenSlider.Value > 1 Then stijl.Visible = False
If PenSlider.Value = 1 Then stijl.Visible = True


End Sub




Private Sub Picture1_Click()


Image1.Picture = Picture1.Image

End Sub

Private Sub Picture1_DblClick()

SavePicture Image, "c:\windows\desktop\test.bmp"



End Sub

Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'MsgBox Combo1.ListIndex

If modus = 10 Then
 drawing = True
 Dim r, g, b As Integer
 Dim rstep, gstep, bstep, z As Integer
 
 r = Rnd * 255: g = Rnd * 255: b = Rnd * 255
 rstep = (r - 255) / parts
 gstep = (g - 255) / parts
 bstep = (b - 255) / parts
 
 For z = 0 To parts - 1
  red(z) = 255 + z * rstep
  green(z) = 255 + z * gstep
  blue(z) = 255 + z * bstep
  
  Xpos(z) = x / 15: Ypos(z) = y / 15
 Next z
End If


If modus = 9 Then selecteer = 1

If modus = 8 Then


'Sets Flags and Shows FontSelect
CommonDialog1.Flags = cdlCFBoth Or cdlCFEffects Or 262144
CommonDialog1.ShowFont
'Returns 7 Properties
'AdjustButtons


If Len(CommonDialog1.FontName) > 0 Then Picture1.FontName = CommonDialog1.FontName
'lblFont.Caption = FName

Picture1.FontSize = CommonDialog1.FontSize
Picture1.ForeColor = CommonDialog1.Color

Picture1.FontBold = CommonDialog1.FontBold
Picture1.FontItalic = CommonDialog1.FontItalic
Picture1.FontUnderline = CommonDialog1.FontUnderline
Picture1.FontStrikethru = CommonDialog1.FontStrikethru


Picture1.CurrentX = x
Picture1.CurrentY = y
Picture1.Print InputBox("Type in the text you want to place on the image..")

End If


If modus = 7 Then


Shape1.FillColor = Picture1.Point(x, y)
Picture1.ForeColor = Picture1.Point(x, y)




End If

If modus = 6 Then ' ------- OPVULLEN ----------------------

    Picture1.FillColor = Picture1.ForeColor
    Picture1.FillStyle = 0
    ExtFloodFill Picture1.hdc, x, y, Picture1.Point(x, y), 1
    Picture1.FillStyle = 1
    
End If



a = x
b = y


If modus = 1 Then

' MOUSE DONW
point1.x = x
point1.y = y
Picture1.Line (x, y)-(x, y)

End If

If modus = 2 Then o = x: p = y: tekenlijn = 1
If modus = 3 Or modus = 4 Then o = x: p = y: tekenrechthoek = 1

If modus = 5 Or modus = 9 Then o = x: p = y


End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

If modus = 10 Then
Dim xstep As Single, ystep As Single
 Dim a As Integer, b As Integer, z As Integer 'lusvars
 
 Dim newx, newy As Integer
 
 
 If drawing = True Then
   newx = x / 15: newy = y / 15
   
   If Xpos(0) < newx Then xstep = (newx - Xpos(0)) / 20
   If Xpos(0) > newx Then xstep = (Xpos(0) - newx) / 20
   If Ypos(0) < newy Then ystep = (newy - Ypos(0)) / 20
   If Ypos(0) > newy Then ystep = (Ypos(0) - newy) / 20
  
   For b = 1 To 20
    For z = 1 To parts - 1
     If Xpos(z) > Xpos(z - 1) Then Xpos(z) = Xpos(z) - (Xpos(z) - Xpos(z - 1)) / 4
     If Xpos(z) < Xpos(z - 1) Then Xpos(z) = Xpos(z) + (Xpos(z - 1) - Xpos(z)) / 4
     If Ypos(z) > Ypos(z - 1) Then Ypos(z) = Ypos(z) - (Ypos(z) - Ypos(z - 1)) / 4
     If Ypos(z) < Ypos(z - 1) Then Ypos(z) = Ypos(z) + (Ypos(z - 1) - Ypos(z)) / 4
     Picture1.Line (Xpos(z - 1) * 15, Ypos(z - 1) * 15)-(Xpos(z) * 15, Ypos(z) * 15), RGB(red(z), green(z), blue(z))
    Next z
    If Xpos(0) < newx Then Xpos(0) = Xpos(0) + xstep
    If Xpos(0) > newx Then Xpos(0) = Xpos(0) - xstep
    If Ypos(0) < newy Then Ypos(0) = Ypos(0) + ystep
    If Ypos(0) > newy Then Ypos(0) = Ypos(0) - ystep
   Next b
  End If
End If

If Button = 1 And modus = 1 Then

point2 = point1
point1.x = x
point1.y = y
Picture1.Line (point1.x, point1.y)-(point2.x, point2.y)

End If

If tekenlijn = 1 Then
Line1.Visible = True
Line1.BorderWidth = PenSlider.Value
Line1.BorderColor = Picture1.ForeColor

'Picture1.Line (a, b)-(x, y)
Line1.X1 = o
Line1.Y1 = p
Line1.X2 = x
Line1.Y2 = y

End If


If selecteer = 1 Then
SelectShape.Visible = True

SelectShape.Left = o
SelectShape.Top = p


If x > a Then SelectShape.Width = x - o
If y > b Then SelectShape.Height = y - p

If x < a Or y < b Then
SelectShape.Left = x
SelectShape.Top = y

SelectShape.Width = o
SelectShape.Height = p
End If

End If

If tekenrechthoek = 1 Then
Shape2.Visible = True
Shape2.BorderWidth = PenSlider.Value
Shape2.BorderColor = Picture1.ForeColor

'Picture1.Line (a, b)-(x, y)
Shape2.Left = o
Shape2.Top = p

If x > o Then Shape2.Width = x - o
If y > p Then Shape2.Height = y - p

If x < a Or y < b Then
Shape2.Left = x
Shape2.Top = y

Shape2.Width = o
Shape2.Height = p


End If

End If



End Sub


Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)


If modus = 10 Then drawing = False
'ForeColor = QBColor(Combo2.ListIndex)


'If tekenlijn = 1 Then
'Line1.Visible = False
'tekenlijn = 0

'End If


If modus = 2 Then
Line1.Visible = False
tekenlijn = 0
If o <> 0 And p <> 0 Then Picture1.Line (o, p)-(x, y)



o = 0
p = 0
End If


If modus = 3 Then
Shape2.Visible = False
tekenrechthoek = 0
If o <> 0 And p <> 0 Then Picture1.Line (o, p)-(x, y), , B
o = 0
p = 0
End If

If modus = 9 Then ' - - - - SELECTEREN
selecteer = 0
gebiedgeselecteerd = 1

End If



If modus = 4 Then
Shape2.Visible = False
tekenrechthoek = 0
If o <> 0 And p <> 0 Then Picture1.Line (o, p)-(x, y), , BF

o = 0
p = 0
End If

If modus = 5 Then
If o <> 0 And p <> 0 Then
Picture1.Circle (o, p), StraalSlider.Value



End If
o = 0
p = 0
End If


End Sub


Private Sub pCol_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
colpressed = True                          'this function tells the computer to
Shape1.FillColor = pCol.Point(x, y)        'set the colors for shape one and the pen
Picture1.ForeColor = pCol.Point(x, y)




End Sub


Private Sub pCol_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next
If colpressed Then                         'same as above but this is here
Shape1.FillColor = pCol.Point(x, y)        'so you don't have to keep clicking to
Picture1.ForeColor = pCol.Point(x, y)      'change the color...you can just drag
End If
End Sub

Private Sub pCol_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
colpressed = False          'stops the selecting of the color when the user 'unclicks'
End Sub





Private Sub printimage_Click()

RichTextBox1.Text = ""


'Printer.PaintPicture Picture1.Picture, 1, 1

SavePicture Picture1.Image, "c:\temp.bmp"

RichTextBox1.OLEObjects.Add , , "c:\temp.bmp"

'
On Error Resume Next
Printer.Print ""
RichTextBox1.SelPrint Printer.hdc
Printer.EndDoc


RichTextBox1.Visible = False


End Sub

Private Sub saveas_Click()

   ' Set CancelError is True
    CommonDialog1.CancelError = True
    On Error GoTo ErrHandler
    ' Set flags
    CommonDialog1.Flags = cdlOFNHideReadOnly
    ' Set filters
     
    CommonDialog1.Filter = "Windows bitmap (*.BMP)|*.bmp|Compuserve GIF" & _
    "(*.gif)|*.gif|JPEG Filter|*.jpg|Alle bestanden|*.*"
        
    
    ' Specify default filter
    CommonDialog1.FilterIndex = 1
    ' Display the Open dialog box
        CommonDialog1.ShowSave
    'MsgBox CommonDialog1.filename
    

    SavePicture Picture1.Image, CommonDialog1.filename
    
    
    
    Exit Sub
    
ErrHandler:
    'User pressed the Cancel button
    Exit Sub

End Sub

Private Sub spiegelen_Click()
'object.PaintPicture picture, x1, y1, width1, height1, x2, y2, width2, height2, opcode
'PaintPicture Picture1.Picture, 0, 0, -(Picture1.Width), -(Picture1.Height)

'Picture1.Width = -(Picture1.Width)
'Picture1.Height = -(Picture1.Height)


End Sub

Private Sub StraalSlider_Click()

lblStraal.Caption = StraalSlider.Value


End Sub

Private Sub stijl_Click()
Picture1.DrawStyle = stijl.ListIndex
'MsgBox stijl.ListIndex
End Sub
