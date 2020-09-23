VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Skin Designer"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   615
   ClientWidth     =   8325
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   331
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   555
   StartUpPosition =   2  'CenterScreen
   Begin VB.HScrollBar Hor 
      Height          =   255
      LargeChange     =   20
      Left            =   0
      SmallChange     =   5
      TabIndex        =   21
      Top             =   4680
      Width           =   5055
   End
   Begin VB.VScrollBar Ver 
      Height          =   4695
      LargeChange     =   20
      Left            =   5040
      SmallChange     =   5
      TabIndex        =   20
      Top             =   0
      Width           =   255
   End
   Begin VB.Timer tmrBack 
      Interval        =   500
      Left            =   6960
      Top             =   1440
   End
   Begin TabDlg.SSTab tbProps 
      Height          =   4935
      Left            =   5400
      TabIndex        =   2
      Top             =   0
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   8705
      _Version        =   393216
      Style           =   1
      TabHeight       =   529
      TabCaption(0)   =   "Window"
      TabPicture(0)   =   "frmMain.frx":08CA
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label8"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label9"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label10"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label11"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "sldWidth"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "sldHeight"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "clrWBack"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "clrWDark"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "clrWLight"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Title Bar"
      TabPicture(1)   =   "frmMain.frx":08E6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "Label5"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "txtCaption"
      Tab(1).Control(7)=   "sldTitlePer"
      Tab(1).Control(8)=   "clrTCap"
      Tab(1).Control(9)=   "clrTDark"
      Tab(1).Control(10)=   "clrTLight"
      Tab(1).Control(11)=   "clrTBack"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "Buttons"
      TabPicture(2)   =   "frmMain.frx":0902
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "lstButtons"
      Tab(2).ControlCount=   1
      Begin VB.PictureBox clrWLight 
         Height          =   255
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   27
         Top             =   3240
         Width           =   735
      End
      Begin VB.PictureBox clrWDark 
         Height          =   255
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   26
         Top             =   2760
         Width           =   735
      End
      Begin VB.PictureBox clrWBack 
         Height          =   255
         Left            =   1440
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   25
         Top             =   2280
         Width           =   735
      End
      Begin VB.ListBox lstButtons 
         Height          =   2985
         ItemData        =   "frmMain.frx":091E
         Left            =   -74640
         List            =   "frmMain.frx":092B
         Style           =   1  'Checkbox
         TabIndex        =   19
         Top             =   1200
         Width           =   1935
      End
      Begin MSComctlLib.Slider sldHeight 
         Height          =   255
         Left            =   360
         TabIndex        =   18
         Top             =   1560
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   50
         SmallChange     =   10
         Min             =   1
         Max             =   500
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider sldWidth 
         Height          =   255
         Left            =   360
         TabIndex        =   16
         Top             =   840
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   450
         _Version        =   393216
         LargeChange     =   50
         SmallChange     =   10
         Min             =   1
         Max             =   500
         SelStart        =   1
         Value           =   1
      End
      Begin VB.PictureBox clrTBack 
         BackColor       =   &H00C00000&
         Height          =   255
         Left            =   -73800
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   14
         Top             =   2280
         Width           =   735
      End
      Begin VB.PictureBox clrTLight 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   -73800
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   12
         Top             =   3720
         Width           =   735
      End
      Begin VB.PictureBox clrTDark 
         BackColor       =   &H00000000&
         Height          =   255
         Left            =   -73800
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   11
         Top             =   3240
         Width           =   735
      End
      Begin VB.PictureBox clrTCap 
         BackColor       =   &H00FF0000&
         Height          =   255
         Left            =   -73800
         ScaleHeight     =   195
         ScaleWidth      =   675
         TabIndex        =   10
         Top             =   2760
         Width           =   735
      End
      Begin MSComctlLib.Slider sldTitlePer 
         Height          =   375
         Left            =   -74760
         TabIndex        =   8
         Top             =   1680
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   661
         _Version        =   393216
         Min             =   1
         Max             =   5
         SelStart        =   1
         Value           =   1
      End
      Begin VB.TextBox txtCaption 
         Height          =   285
         Left            =   -74760
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Light Color"
         Height          =   255
         Left            =   360
         TabIndex        =   24
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Dark Color"
         Height          =   255
         Left            =   360
         TabIndex        =   23
         Top             =   2760
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Back"
         Height          =   255
         Left            =   360
         TabIndex        =   22
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label8 
         Caption         =   "Height"
         Height          =   255
         Left            =   360
         TabIndex        =   17
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Width"
         Height          =   255
         Left            =   360
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label6 
         Caption         =   "Backcolor"
         Height          =   255
         Left            =   -74760
         TabIndex        =   13
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "Title Height (Percentage)"
         Height          =   255
         Left            =   -74760
         TabIndex        =   9
         Top             =   1440
         Width           =   1815
      End
      Begin VB.Label Label4 
         Caption         =   "Light Color"
         Height          =   255
         Left            =   -74760
         TabIndex        =   7
         Top             =   3720
         Width           =   1695
      End
      Begin VB.Label Label3 
         Caption         =   "Dark Color"
         Height          =   255
         Left            =   -74760
         TabIndex        =   6
         Top             =   3240
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "Caption"
         Height          =   255
         Left            =   -74760
         TabIndex        =   5
         Top             =   2760
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Caption"
         Height          =   255
         Left            =   -74760
         TabIndex        =   4
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.PictureBox Container 
      BackColor       =   &H00808080&
      Enabled         =   0   'False
      Height          =   4695
      Left            =   0
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   333
      TabIndex        =   0
      Top             =   0
      Width           =   5055
      Begin VB.Timer tmrScroll 
         Interval        =   10
         Left            =   2400
         Top             =   1920
      End
      Begin VB.PictureBox Board 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
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
         Left            =   0
         ScaleHeight     =   17
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   17
         TabIndex        =   1
         Top             =   0
         Width           =   255
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuSkin 
         Caption         =   "New Skin"
      End
      Begin VB.Menu mnuLoad 
         Caption         =   "Load Skin"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save Skin"
      End
      Begin VB.Menu mnuBMP 
         Caption         =   "Convert to .BMP"
      End
      Begin VB.Menu mnusep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "About"
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ResizeC As Long

Function NewSkin()
Skin.Title = "New Skin"
txtCaption.Text = "New Skin"

Skin.bntClose = True
Skin.bntMax = False
Skin.bntMin = False

lstButtons.Selected(0) = False
lstButtons.Selected(1) = False
lstButtons.Selected(2) = True

Skin.Height = 200
Skin.Width = 300

sldHeight.Value = 200
sldWidth.Value = 300

Skin.TColors.Back = RGB(150, 0, 0)
Skin.TColors.Caption = vbRed
Skin.TColors.DarkColor = vbBlack
Skin.TColors.LightColor = vbRed

clrTBack.BackColor = RGB(150, 0, 0)
clrTCap.BackColor = vbRed
clrTDark.BackColor = vbBlack
clrTLight.BackColor = vbRed

Skin.WColors.Back = RGB(0, 0, 150)
Skin.WColors.DarkColor = vbBlack
Skin.WColors.LightColor = RGB(0, 0, 255)

clrWBack.BackColor = RGB(0, 0, 150)
clrWDark.BackColor = vbBlack
clrWLight.BackColor = RGB(0, 0, 255)

clrTBack.BackColor = RGB(150, 0, 0)
clrTDark.BackColor = vbBlack
clrTLight.BackColor = RGB(255, 0, 0)
End Function

Function UpdateSkinView()
Dim I As Long

Board.Cls
Board.Width = Skin.Width
Board.Height = Skin.Height

Board.Line (0, 0)-(Board.Width - 2, Board.Height - 2), Skin.WColors.LightColor, BF
Board.Line (1, 1)-(Board.Width - 1, Board.Height - 1), Skin.WColors.DarkColor, BF
Board.Line (1, 1)-(Board.Width - 2, Board.Height - 2), Skin.WColors.Back, BF

Board.Line (0, 0)-(Board.Width - 2, 15 - 2), Skin.TColors.LightColor, BF
Board.Line (1, 1)-(Board.Width - 1, 15 - 1), Skin.TColors.DarkColor, BF
Board.Line (1, 1)-(Board.Width - 2, 15 - 2), Skin.TColors.Back, BF

Board.ForeColor = Skin.TColors.DarkColor
Board.CurrentX = 3
Board.CurrentY = 2
Board.Print Skin.Title
Board.ForeColor = Skin.TColors.Caption
Board.CurrentX = 4
Board.CurrentY = 2
Board.Print Skin.Title

If Skin.bntClose Then
Board.ForeColor = Skin.TColors.DarkColor
Board.CurrentX = Board.ScaleWidth - 13
Board.CurrentY = 2
Board.Print "X"
Board.ForeColor = Skin.TColors.Caption
Board.CurrentX = Board.ScaleWidth - 12
Board.CurrentY = 2
Board.Print "X"
End If

If Skin.bntMax Then
Board.ForeColor = Skin.TColors.DarkColor
Board.CurrentX = Board.ScaleWidth - 28
Board.CurrentY = 3
Board.Line (Board.CurrentX, Board.CurrentY)-(Board.CurrentX + 10, Board.CurrentY + 10), Board.ForeColor, B
Board.ForeColor = Skin.TColors.LightColor
Board.CurrentX = Board.ScaleWidth - 27
Board.CurrentY = 3
Board.Line (Board.CurrentX, Board.CurrentY)-(Board.CurrentX + 10, Board.CurrentY + 10), Board.ForeColor, B
End If

If Skin.bntMin Then
Board.ForeColor = Skin.TColors.DarkColor
Board.CurrentX = Board.ScaleWidth - 43
Board.CurrentY = 3
Board.Line (Board.CurrentX, Board.CurrentY + 10)-(Board.CurrentX + 10, Board.CurrentY + 10), Board.ForeColor, B
Board.ForeColor = Skin.TColors.LightColor
Board.CurrentX = Board.ScaleWidth - 42
Board.CurrentY = 3
Board.Line (Board.CurrentX, Board.CurrentY + 10)-(Board.CurrentX + 10, Board.CurrentY + 10), Board.ForeColor, B
End If
End Function

Function ChangeColor(pic As PictureBox)
On Error GoTo NoColor:
Dim cColor
'32755

Set cColor = CreateObject("MSCOMDLG.CommonDialog")

cColor.CancelError = True
cColor.ShowColor

pic.BackColor = cColor.Color

Set cColor = Nothing 'remove the object from mem

Exit Function
NoColor:

Set cColor = Nothing
End Function

Private Sub clrTBack_Click()
ChangeColor clrTBack
Skin.TColors.Back = clrTBack.BackColor
End Sub

Private Sub clrTCap_Click()
ChangeColor clrTCap
Skin.TColors.Caption = clrTCap.BackColor
End Sub

Private Sub clrTDark_Click()
ChangeColor clrTDark
Skin.TColors.DarkColor = clrTDark.BackColor
End Sub

Private Sub clrTLight_Click()
ChangeColor clrTLight
Skin.TColors.LightColor = clrTLight.BackColor
End Sub

Private Sub clrWBack_Click()
ChangeColor clrWBack
Skin.WColors.Back = clrWBack.BackColor
End Sub

Private Sub clrWDark_Click()
ChangeColor clrWDark
Skin.WColors.DarkColor = clrWDark.BackColor
End Sub

Private Sub clrWLight_Click()
ChangeColor clrWLight
Skin.WColors.LightColor = clrWLight.BackColor
End Sub

Private Sub Form_Load()
NewSkin

If Len(Command) > 2 Then 'if the program has a path to load then load it
LoadSkin Command
End If

Me.Visible = True
Hor.Enabled = False
Ver.Enabled = False
tmrBack_Timer
End Sub

Private Sub Hor_Change()
Board.Left = Hor.Value
End Sub

Private Sub lstButtons_Click()
Select Case lstButtons.ListIndex
Case 0 'min button
Skin.bntMin = lstButtons.Selected(lstButtons.ListIndex)
Case 1 'max button
Skin.bntMax = lstButtons.Selected(lstButtons.ListIndex)
Case 2 'close button
Skin.bntClose = lstButtons.Selected(lstButtons.ListIndex)
End Select
End Sub

Private Sub mnuAbout_Click()
MsgBox (App.Comments & vbCrLf & "By Kevin Fleet")
End Sub

Private Sub mnuBMP_Click()
On Error GoTo NoFile:
Dim cPath
'32755

Set cPath = CreateObject("MSCOMDLG.CommonDialog")

cPath.Filter = "Bitmap Files (*.bmp*)|*.bmp*"
cPath.CancelError = True
cPath.ShowSave

UpdateSkinView 'update the skin
SavePicture Board.Image, Replace(cPath.FileName, ".bmp", "") & ".bmp" 'save the skin

Set cPath = Nothing 'remove the object from mem

Exit Sub
NoFile:

Set cPath = Nothing
End Sub

Private Sub mnuExit_Click()
Unload Me
End Sub

Private Sub mnuLoad_Click()
On Error GoTo NoFile:
Dim cPath
'32755

Set cPath = CreateObject("MSCOMDLG.CommonDialog")

cPath.Filter = "Skin Files (*.skn*)|*.skn*"
cPath.CancelError = True
cPath.ShowOpen

LoadSkin cPath.FileName

Set cPath = Nothing 'remove the object from mem

Exit Sub
NoFile:

Set cPath = Nothing
End Sub

Private Sub mnuSave_Click()
On Error GoTo NoFile:
Dim cPath
'32755

Set cPath = CreateObject("MSCOMDLG.CommonDialog")

cPath.Filter = "Skin Files (*.skn*)|*.skn*"
cPath.CancelError = True
cPath.ShowSave
SaveSkin Replace(cPath.FileName, ".skn", "") & ".skn"

Set cPath = Nothing 'remove the object from mem

Exit Sub
NoFile:

Set cPath = Nothing
End Sub

Private Sub mnuSkin_Click()
Container.Enabled = True
NewSkin
End Sub

Private Sub sldHeight_Click()
Skin.Height = sldHeight.Value
End Sub

Private Sub sldTitlePer_Click()
Skin.TitleHeight = sldTitlePer.Value
End Sub

Private Sub sldWidth_Click()
Skin.Width = sldWidth.Value
End Sub

Private Sub tmrBack_Timer()
UpdateSkinView
End Sub

Private Sub tmrScroll_Timer()
If Board.Height > Container.ScaleHeight Then
Ver.Enabled = True
Ver.Max = (Container.ScaleHeight - Board.Height)
Else
Ver.Enabled = False
Board.Top = 0
End If

If Board.Width > Container.ScaleWidth Then
Hor.Enabled = True
Hor.Max = (Container.ScaleWidth - Board.Width)
Else
Hor.Enabled = False
Board.Left = 0
End If
End Sub

Private Sub txtCaption_Change()
Skin.Title = txtCaption.Text
End Sub

Private Sub Ver_Change()
Board.Top = Ver.Value
End Sub
