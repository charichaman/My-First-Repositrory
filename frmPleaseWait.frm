VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPleaseWait 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Please Wait"
   ClientHeight    =   720
   ClientLeft      =   7275
   ClientTop       =   9045
   ClientWidth     =   13800
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   720
   ScaleWidth      =   13800
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox TxtPleaseWait 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   5820
      MaxLength       =   14
      TabIndex        =   0
      Text            =   "Please Wait..."
      Top             =   235
      Width           =   2175
   End
   Begin MSComctlLib.ProgressBar myProgressBar 
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   873
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "frmPleaseWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    TxtPleaseWait.SelStart = Len(TxtPleaseWait)
    
End Sub
