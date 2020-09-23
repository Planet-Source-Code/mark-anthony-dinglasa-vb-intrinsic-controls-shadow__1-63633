VERSION 5.00
Begin VB.Form frmShadow 
   BackColor       =   &H00FFC0C0&
   Caption         =   "Shadows..."
   ClientHeight    =   5685
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7665
   LinkTopic       =   "Form1"
   ScaleHeight     =   5685
   ScaleWidth      =   7665
   StartUpPosition =   3  'Windows Default
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3480
      TabIndex        =   12
      Top             =   4560
      Width           =   2295
   End
   Begin VB.DirListBox Dir1 
      Height          =   1440
      Left            =   5520
      TabIndex        =   11
      Top             =   2280
      Width           =   1815
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   4320
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   6480
      TabIndex        =   10
      Top             =   3960
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   3840
      TabIndex        =   9
      Text            =   "Combo1"
      Top             =   3480
      Width           =   1335
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0FF&
      Height          =   1455
      Left            =   2400
      ScaleHeight     =   1395
      ScaleWidth      =   1035
      TabIndex        =   8
      Top             =   2760
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   495
      Left            =   2880
      TabIndex        =   7
      Top             =   2040
      Width           =   1215
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   840
      TabIndex        =   6
      Top             =   1680
      Width           =   1575
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   1455
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   255
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Check1"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Option1"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   1560
      Width           =   2895
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   375
      Left            =   2040
      TabIndex        =   1
      Top             =   960
      Width           =   2415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Click Me To give Shadows on Controls"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   5160
      Width           =   3255
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   240
      TabIndex        =   13
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   375
      Left            =   360
      Top             =   720
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   1815
      Left            =   600
      Picture         =   "frmShadow.frx":0000
      Stretch         =   -1  'True
      Top             =   2400
      Width           =   1455
   End
End
Attribute VB_Name = "frmShadow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
  GiveShadow Me, 70, 70
End Sub
