VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4770
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   ScaleHeight     =   4770
   ScaleWidth      =   6570
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   4740
      TabIndex        =   1
      Top             =   1920
      Width           =   1695
   End
   Begin Project1.Fire Fire1 
      Height          =   1635
      Index           =   0
      Left            =   960
      TabIndex        =   0
      Top             =   2280
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   2884
      TimeInterval    =   100
      BackColor       =   64
      ColDecrease     =   10
      DX              =   100
      DY              =   40
      numCopy         =   1
      ToolTipTextString=   "SAEED"
      BorderStyle     =   1
      DX              =   100
      DY              =   40
      Text            =   "SAEED"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
