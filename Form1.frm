VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   5610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7605
   LinkTopic       =   "Form1"
   ScaleHeight     =   5610
   ScaleWidth      =   7605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "show a stored pic"
      Height          =   1860
      Left            =   5040
      TabIndex        =   5
      Top             =   585
      Width           =   1770
      Begin VB.OptionButton Option1 
         Caption         =   "pic4"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   9
         Top             =   1170
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "pic3"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   8
         Top             =   900
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "pic2"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   630
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "pic1"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   360
         Width           =   1275
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Hide ocx"
      Height          =   420
      Index           =   2
      Left            =   5175
      TabIndex        =   4
      Top             =   4365
      Width           =   1680
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Draw using controls internal blt function"
      Height          =   420
      Index           =   1
      Left            =   5130
      TabIndex        =   3
      Top             =   3780
      Width           =   1680
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Stretch"
      Height          =   240
      Left            =   5175
      TabIndex        =   2
      Top             =   3375
      Width           =   1635
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Download pic from internet"
      Height          =   420
      Index           =   0
      Left            =   5130
      TabIndex        =   1
      Top             =   2835
      Width           =   1680
   End
   Begin projoPicFlex.oPicFlex oPicFlex1 
      Height          =   3345
      Left            =   315
      TabIndex        =   0
      Top             =   540
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   5900
      Picture         =   "Form1.frx":0000
      BorderStyle     =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
  oPicFlex1.Stretch = CBool(Check1.Value)
End Sub

Private Sub Command1_Click(Index As Integer)
Static bShow As Boolean
Dim caption(1) As String
Dim b As Boolean

  caption(0) = "Show ocx"
  caption(1) = "&Hide ocx"
  
  Select Case Index
    Case Is = 0
      oPicFlex1.LoadPicFromInternet _
      "http://www.vb-helper.com/vbcl.jpg"
    Case Is = 1
      b = oPicFlex1.PaintBltFromThis(hdc, 0, 0, Width, Height, , , , , 0)
      Debug.Print b
    Case Is = 2
      oPicFlex1.Visible = bShow
      Command1(2).caption = caption(Abs(CLng(bShow)))
      If Not (bShow) Then
        MsgBox "now click the button above so you can see this " & _
               "control can be used strictly for api drawing " & _
               "operations whether visible or not"
      End If
      bShow = Not (bShow)
  End Select
End Sub

Private Sub Form_Load()
  'load and store some pics
  With oPicFlex1
    .AddToStoredPics App.Path & "\pic1.bmp"
    .AddToStoredPics App.Path & "\pic2.gif"
    .AddToStoredPics App.Path & "\pic3.gif"
    .AddToStoredPics App.Path & "\pic4.jpg"
  End With
End Sub

Private Sub oPicFlex1_PictureDownloadComplete()
   On Error Resume Next
   MsgBox "y dont we store this pic ??,  click YES if you agree (I always loved a democracy)"
   oPicFlex1.AddToStoredPics 'by not specify a picture filepath
                             'we are telling the control to add
                             'the current picture to stored pictures
   Load Option1(4)
   Option1(4).Left = Option1(3).Left
   Option1(4).Top = Option1(3).Top + 300
   Option1(4).caption = "pic5"
   Option1(4).Visible = True
End Sub

Private Sub Option1_Click(Index As Integer)
   oPicFlex1.ActivePicture = Index
End Sub
