VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form fIDE 
   Caption         =   "My Form Designer"
   ClientHeight    =   5700
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "fIDE.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5700
   ScaleWidth      =   7260
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   5385
      Width           =   7260
      _ExtentX        =   12806
      _ExtentY        =   556
      SimpleText      =   ""
      _Version        =   327682
      BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
         NumPanels       =   1
         BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
            Object.Width           =   10583
            MinWidth        =   10583
            Text            =   "Now you can make a Window inside a Window"
            TextSave        =   "Now you can make a Window inside a Window"
            Key             =   ""
            Object.Tag             =   ""
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox pArea 
      Align           =   1  'Align Top
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      Height          =   5475
      Left            =   0
      ScaleHeight     =   5475
      ScaleWidth      =   7260
      TabIndex        =   2
      Top             =   0
      Width           =   7260
      Begin VB.PictureBox pForm 
         BorderStyle     =   0  'None
         Height          =   3255
         Left            =   360
         ScaleHeight     =   3255
         ScaleWidth      =   6015
         TabIndex        =   3
         Top             =   660
         Width           =   6015
         Begin VB.ListBox List1 
            Height          =   1620
            ItemData        =   "fIDE.frx":2CFA
            Left            =   3270
            List            =   "fIDE.frx":2D10
            TabIndex        =   9
            Top             =   120
            Width           =   2595
         End
         Begin VB.ComboBox Combo1 
            Height          =   315
            Left            =   180
            TabIndex        =   8
            Text            =   "Select your age here!!!"
            Top             =   1830
            Width           =   2985
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Open as read only"
            Height          =   225
            Left            =   180
            TabIndex        =   7
            Top             =   2430
            Value           =   -1  'True
            Width           =   1995
         End
         Begin VB.CheckBox Check1 
            Caption         =   "Enable Quick Search"
            Height          =   255
            Left            =   180
            TabIndex        =   6
            Top             =   2190
            Value           =   1  'Checked
            Width           =   2055
         End
         Begin VB.CommandButton Command1 
            Caption         =   "Validate Me!!!"
            Default         =   -1  'True
            Height          =   405
            Left            =   1770
            TabIndex        =   0
            Top             =   1170
            Width           =   1185
         End
         Begin VB.Frame frFrame 
            Caption         =   "This is frame"
            Height          =   1695
            Left            =   180
            TabIndex        =   4
            Top             =   60
            Width           =   2955
            Begin VB.TextBox Text1 
               Height          =   315
               Left            =   210
               TabIndex        =   5
               Text            =   "Please enter your name"
               Top             =   390
               Width           =   2565
            End
         End
      End
   End
End
Attribute VB_Name = "fIDE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Project Name                :  My Form Designer ( a part of my next version of Visual Dialog++ v4.0 )
' Project Version              :  Nope
' Copyright                      :  All goes to me...
' Where did I found this?  :  I don't know :-P


Option Explicit
'
Dim retCaption As Variant

Private Sub Form_Load()
    '
    Call SetWinStyle
    '
End Sub

Private Sub SetWinStyle()
    '
    ' Just in case if the process take long time to finished...
    Screen.MousePointer = vbHourglass
    '
    Dim Style&
    '
    Style = GetWindowLong(pForm.hWnd, GWL_STYLE) ' Get current style
    Style = Style Or WS_THICKFRAME ' Set border style...
    Style = Style Or WS_CAPTION ' Add caption to the form...
    Style = Style Or WS_MINIMIZEBOX ' Show the minimize box...
    Style = Style Or WS_MAXIMIZEBOX ' Show the maximize box...
    Style = Style Or WS_SYSMENU ' Important!! Set this or you'll not see the control box...
    '
    ' Set the new style!!!
    '
    Style = SetWindowLong(pForm.hWnd, GWL_STYLE, Style)
    '
    ' Set the caption...
    '
    retCaption = SetWindowText(pForm.hWnd, "Form1")
    '
    ' Resize the new "FORM" so the "FORM" will display properly...
    '
    pForm.Height = pForm.Height + 30
    '
    ' Revert the mousepointer to the default value
    '
    Screen.MousePointer = vbDefault
    '
End Sub

Private Sub Form_Resize()
    '
    On Error Resume Next
    '
    StatusBar1.Panels(1).Width = Me.ScaleWidth
    pArea.Height = Me.ScaleHeight - StatusBar1.Height
    '
End Sub
