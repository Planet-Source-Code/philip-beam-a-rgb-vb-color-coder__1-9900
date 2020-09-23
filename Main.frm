VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Main 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "RGB and VB Color Coder"
   ClientHeight    =   3360
   ClientLeft      =   4950
   ClientTop       =   2670
   ClientWidth     =   2175
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   2175
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   2280
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00C0C0C0&
      Height          =   3135
      Left            =   2280
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   13
      Text            =   "Main.frx":0000
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2640
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2640
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "EXIT"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3000
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll3 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll2 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   600
      Width           =   1935
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      LargeChange     =   10
      Left            =   120
      Max             =   255
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
   End
   Begin MSComctlLib.Slider Slider3 
      Height          =   255
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider2 
      Height          =   255
      Left            =   120
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   600
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      Max             =   255
      TickStyle       =   3
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   255
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   450
      _Version        =   393216
      LargeChange     =   10
      Max             =   255
      TickStyle       =   3
   End
   Begin VB.CommandButton Command17 
      Caption         =   "Slide"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1935
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Scroll"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
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
      Left            =   120
      TabIndex        =   10
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "By Philip Beam"
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
      Left            =   120
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Blu 
      Alignment       =   2  'Center
      Caption         =   "BLU = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   1320
      Width           =   1935
   End
   Begin VB.Label Grn 
      Alignment       =   2  'Center
      Caption         =   "GRN = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Red 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0C0&
      Caption         =   "RED = 0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   360
      Width           =   1935
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    End
End Sub

Private Sub Command17_Click()
    Main.Command17.Visible = False
    Main.Command4.Visible = True
    Main.Slider1.Visible = True
    Main.Slider2.Visible = True
    Main.Slider3.Visible = True
    Main.HScroll1.Visible = False
    Main.HScroll2.Visible = False
    Main.HScroll3.Visible = False
End Sub

Private Sub Command2_Click()
    Main.Width = 4785
    Main.Command3.Visible = True
    Main.Command2.Visible = False
End Sub

Private Sub Command3_Click()
    Main.Width = 2280
    Main.Command2.Visible = True
    Main.Command3.Visible = False
End Sub

Private Sub Command4_Click()
    Main.Command4.Visible = False
    Main.Command17.Visible = True
    Main.HScroll1.Visible = True
    Main.HScroll2.Visible = True
    Main.HScroll3.Visible = True
    Main.Slider1.Visible = False
    Main.Slider2.Visible = False
    Main.Slider3.Visible = False
End Sub

Private Sub HScroll1_Change()
    Main.Red.Caption = "RED = " & Main.HScroll1.Value & ""
    Main.Label5.ForeColor = RGB(HScroll1, HScroll2, HScroll3)
    Main.Label1.Caption = Main.Label5.ForeColor
End Sub

Private Sub HScroll2_Change()
    Main.Grn.Caption = "GRN = " & Main.HScroll2.Value & ""
    Main.Label5.ForeColor = RGB(HScroll1, HScroll2, HScroll3)
    Main.Label1.Caption = Main.Label5.ForeColor
End Sub

Private Sub HScroll3_Change()
    Main.Blu.Caption = "BLU = " & Main.HScroll3.Value & ""
    Main.Label5.ForeColor = RGB(HScroll1, HScroll2, HScroll3)
    Main.Label1.Caption = Main.Label5.ForeColor
End Sub

Private Sub Label5_Click()
    Main.CommonDialog1.ShowColor
    Main.Label5.ForeColor = Main.CommonDialog1.Color
    Main.Label1.Caption = Main.CommonDialog1.Color
End Sub

Private Sub Slider1_Change()
    Main.Red.Caption = "RED = " & Main.Slider1.Value & ""
    Main.Label5.ForeColor = RGB(Slider1, Slider2, Slider3)
    Main.Label1.Caption = Main.Label5.ForeColor
End Sub

Private Sub Slider2_Change()
    Main.Grn.Caption = "GRN = " & Main.Slider2.Value & ""
    Main.Label5.ForeColor = RGB(Slider1, Slider2, Slider3)
    Main.Label1.Caption = Main.Label5.ForeColor
End Sub

Private Sub Slider3_Change()
    Main.Blu.Caption = "BLU = " & Main.Slider3.Value & ""
    Main.Label5.ForeColor = RGB(Slider1, Slider2, Slider3)
    Main.Label1.Caption = Main.Label5.ForeColor
End Sub
