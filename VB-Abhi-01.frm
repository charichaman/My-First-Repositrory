VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form myForm 
   BackColor       =   &H00C0FFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Abhinandan H D"
   ClientHeight    =   10410
   ClientLeft      =   4695
   ClientTop       =   2715
   ClientWidth     =   14850
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10410
   ScaleWidth      =   14850
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog CommonDialogForTally 
      Left            =   12240
      Top             =   8280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0080FFFF&
      Caption         =   "Extract WinMan"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   7200
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Merge Folders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   6600
      Width           =   1815
   End
   Begin VB.ListBox LstClients 
      Height          =   4620
      Left            =   8280
      MultiSelect     =   2  'Extended
      TabIndex        =   19
      Top             =   5400
      Width           =   3855
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Extract Folders"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   12360
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   6000
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H0000FFFF&
      Caption         =   "Open Company in Tally.ERP9"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   13200
      MaskColor       =   &H0000FFFF&
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   2160
      Width           =   1335
   End
   Begin VB.FileListBox FileClientFiles 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   4200
      TabIndex        =   7
      Top             =   7320
      Width           =   3975
   End
   Begin VB.FileListBox FileMyFiles 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2040
      Left            =   120
      TabIndex        =   6
      Top             =   7320
      Width           =   3975
   End
   Begin VB.DirListBox DirClientsFolders 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6645
      Left            =   4200
      TabIndex        =   5
      Top             =   600
      Width           =   3975
   End
   Begin VB.CommandButton CmdBrowse 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   150
      Width           =   975
   End
   Begin VB.DirListBox DirMyFolders 
      BackColor       =   &H00FFFFC0&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6645
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3975
   End
   Begin VB.DriveListBox DrvMyDrives 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label LblCurrentFolderName 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   13635
      TabIndex        =   48
      Top             =   9140
      Width           =   1095
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "Name : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12120
      TabIndex        =   47
      Top             =   9140
      Width           =   1455
   End
   Begin VB.Label LblFolderStatus 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13630
      TabIndex        =   46
      Top             =   10080
      Width           =   1095
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "Status : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12120
      TabIndex        =   45
      Top             =   10080
      Width           =   1455
   End
   Begin VB.Label Label17 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "Remaining : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12120
      TabIndex        =   44
      Top             =   9770
      Width           =   1455
   End
   Begin VB.Label Label16 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      Caption         =   "Total : "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   12120
      TabIndex        =   43
      Top             =   9450
      Width           =   1455
   End
   Begin VB.Label Label15 
      Alignment       =   2  'Center
      BackColor       =   &H00FFC0C0&
      Caption         =   ": : Folders : :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   12120
      TabIndex        =   42
      Top             =   8820
      Width           =   2775
   End
   Begin VB.Label LblFoldersRemains 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   13630
      TabIndex        =   41
      Top             =   9770
      Width           =   1095
   End
   Begin VB.Label LblFoldersTotal 
      Alignment       =   1  'Right Justify
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   13635
      TabIndex        =   40
      Top             =   9450
      Width           =   1095
   End
   Begin VB.Label LblCurrentWorkingFolder 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   120
      TabIndex        =   39
      Top             =   10080
      Width           =   11940
   End
   Begin VB.Label LblTallyPrimePath 
      BackColor       =   &H0080FF80&
      Caption         =   "Click Me to Get ""TallyPrime \ Tally.Exe"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   38
      Top             =   9770
      Width           =   8055
   End
   Begin VB.Label LblTransfered 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10665
      TabIndex        =   37
      Top             =   4400
      Width           =   2295
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Caption         =   "Transfered :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   36
      Top             =   4400
      Width           =   2295
   End
   Begin VB.Label Label13 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      Caption         =   "Net Profit :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   8280
      TabIndex        =   35
      Top             =   3980
      Width           =   2295
   End
   Begin VB.Label LblNetProfit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H000080FF&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   10665
      TabIndex        =   34
      Top             =   3980
      Width           =   2295
   End
   Begin VB.Label Label12 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      Caption         =   "Direct Income :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   33
      Top             =   3435
      Width           =   2295
   End
   Begin VB.Label LblDirectIncome 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0FF&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10665
      TabIndex        =   32
      Top             =   3435
      Width           =   2295
   End
   Begin VB.Label LblYearTo 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   13020
      TabIndex        =   31
      Top             =   1580
      Width           =   1650
   End
   Begin VB.Label Label11 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      Caption         =   ":To:"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12360
      TabIndex        =   30
      Top             =   1580
      Width           =   615
   End
   Begin VB.Label Label10 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "Date From :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   29
      Top             =   1580
      Width           =   2295
   End
   Begin VB.Label LblYearFrom 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10665
      TabIndex        =   28
      Top             =   1580
      Width           =   1650
   End
   Begin VB.Label Label9 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "PAN :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   27
      Top             =   1160
      Width           =   2295
   End
   Begin VB.Label LblCmpPAN 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10665
      TabIndex        =   26
      Top             =   1160
      Width           =   3375
   End
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   "GST :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8280
      TabIndex        =   25
      Top             =   740
      Width           =   2295
   End
   Begin VB.Label LblCmpGST 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10665
      TabIndex        =   24
      Top             =   740
      Width           =   3375
   End
   Begin VB.Label LblParentFolder 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   255
      Left            =   10215
      TabIndex        =   22
      Top             =   5070
      Width           =   3855
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      Caption         =   " Merge Folders :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   300
      Left            =   8280
      TabIndex        =   21
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label LblSalesAccount 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   15
      Top             =   3015
      Width           =   2295
   End
   Begin VB.Label LblClosingStock 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   14
      Top             =   2595
      Width           =   2295
   End
   Begin VB.Label LblGrossProfit 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10680
      TabIndex        =   13
      Top             =   2175
      Width           =   2295
   End
   Begin VB.Label LblCmpName 
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   10680
      TabIndex        =   12
      Top             =   75
      Width           =   3975
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFC0C0&
      Caption         =   "Sales A/c :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8295
      TabIndex        =   11
      Top             =   3015
      Width           =   2295
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0E0FF&
      Caption         =   "Closing Stock :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8295
      TabIndex        =   10
      Top             =   2595
      Width           =   2295
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0C0FF&
      Caption         =   " Gross Profit :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8295
      TabIndex        =   9
      Top             =   2175
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFC0&
      Caption         =   " Company Name :"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8300
      TabIndex        =   8
      Top             =   80
      Width           =   2295
   End
   Begin VB.Label LblTallyERP9Path 
      BackColor       =   &H0080FF80&
      Caption         =   "Click Me to Get ""Tally.ERP9 \ Tally.Exe"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   9450
      Width           =   8055
   End
   Begin VB.Label LblClientsFolderPath 
      BackColor       =   &H0080FF80&
      Caption         =   "Click Me to Get Clients Folder Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   3
      Top             =   240
      Width           =   3495
   End
   Begin VB.Label Label5 
      BackColor       =   &H0080FF80&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4860
      Left            =   8220
      TabIndex        =   16
      Top             =   0
      Width           =   6525
   End
End
Attribute VB_Name = "myForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim myDirControl As DirListBox
Private Sub CmdBrowse_Click()
    Dim ExplorPath As String
    
    'If myDirControl.ListCount > 0 Then
        myDirControl.ListIndex = -1
        ExplorPath = myDirControl.List(myDirControl.ListIndex)
        
        Shell "C:\WINDOWS\EXPLORER.EXE " & ExplorPath, vbNormalFocus
    'End If
    
End Sub

'Private Sub ListFolders(ByVal Folder As Folder)

'    Dim SubFolder As Folder
'    Dim folderList As String
'
'    ' Display the current folder
'    folderList = Folder.Path & vbCrLf
'    Debug.Print folderList  ' You can replace this with a different output method
'
'    ' Recursively list subfolders
'    For Each SubFolder In Folder.SubFolders
'        ListFolders SubFolder
'    Next SubFolder
'
'End Sub

Private Sub Command1_Click()

    Dim fso As Scripting.FileSystemObject
        Dim myParentFolder As Scripting.folder
        Dim myChildFolder As Scripting.folder
        
        Dim merFolders() As String
        Dim LstCnt As Integer
        Dim myArraySize As Integer
    
    Set fso = New Scripting.FileSystemObject
    
    myArraySize = 0
    For LstCnt = 0 To LstClients.ListCount - 1
        If LstClients.Selected(LstCnt) Then
            ReDim merFolders(0 To myArraySize)
            myArraySize = myArraySize + 1
        End If
    Next
    
    If myArraySize > 0 Then
        myArraySize = 0
        For LstCnt = 0 To LstClients.ListCount - 1
            If LstClients.Selected(LstCnt) Then
                merFolders(myArraySize) = LstClients.List(LstCnt)
                myArraySize = myArraySize + 1
            End If
        Next
        
        Set myParentFolder = fso.GetFolder(DirClientsFolders & "\" & LblParentFolder.Caption)
        
        For LstCnt = LBound(merFolders) To UBound(merFolders)
            If Not fso.GetBaseName(myParentFolder) Like merFolders(LstCnt) Then
                Set myChildFolder = fso.GetFolder(DirClientsFolders & "\" & merFolders(LstCnt))
                Call MergeFolders(myChildFolder, myParentFolder)
                
                fso.DeleteFolder myChildFolder
                DirClientsFolders.Refresh
                Call LoadToListBox(myForm)
            End If
        Next
        
    End If
    
    MsgBox "Selected Folders are Successfully Merged...!", vbOKOnly + vbInformation, AdminInfo
    
End Sub

Sub MergeFolders(srcFolder As Scripting.folder, destFolder As Scripting.folder)
    Dim File As Object
    Dim subFolder As Scripting.folder
    Dim NewSubFolder As Scripting.folder
    
    Set fso = New Scripting.FileSystemObject
    
    ' First, copy files without overwriting existing ones
    For Each File In srcFolder.Files
        If Not fso.FileExists(destFolder.Path & "\" & File.Name) Then
            File.Copy destFolder.Path & "\" & File.Name
        End If
    Next File
    
    For Each subFolder In srcFolder.SubFolders
    
        Set NewSubFolder = subFolder
'MsgBox destFolder.Path & "\" & SubFolder.Name

        If Not fso.FolderExists(destFolder.Path & "\" & subFolder.Name) Then

            fso.CreateFolder destFolder.Path & "\" & subFolder.Name
        Else
            If IsNumeric(subFolder.Name) = True Then
                Set NewSubFolder = fso.GetFolder(destFolder.Path & "\" & subFolder.Name)
Back:
                If fso.FolderExists(destFolder.Path & "\" & NewSubFolder.Name) Then
                    If fso.FolderExists(destFolder.Path & "\" & Val(NewSubFolder.Name) + 1) Then
                        Set NewSubFolder = fso.GetFolder(destFolder.Path & "\" & Val(NewSubFolder.Name) + 1)
                        GoTo Back:
                    End If
                End If
                Set NewSubFolder = fso.CreateFolder(destFolder.Path & "\" & Val(NewSubFolder.Name) + 1)
            End If
            
        End If
        
        Call MergeFolders(subFolder, fso.GetFolder(destFolder.Path & "\" & NewSubFolder.Name))
        
    Next subFolder
    
End Sub



Private Sub Command3_Click()
    
    Set fso = New Scripting.FileSystemObject
    
    Dim CmpNamePath As String
    Dim CmpName As String
    
    Dim cmpOpenString As String
    
    Dim CmpDetailsTxtFileName As String
    
    Dim TallyERP9orTallyPrime As String
    
    CmpNamePath = DirClientsFolders.List(DirClientsFolders.ListIndex)
    CmpNamePath = Trim(Mid(CmpNamePath, 1, InStrRev(CmpNamePath, "\")))
    
    CmpName = DirClientsFolders.List(DirClientsFolders.ListIndex)
    CmpName = Trim(Mid(CmpName, InStrRev(CmpName, "\") + 1))

'Debug.Print Dir(CmpNamePath & "\" & CmpName & "\Company.*")

    If IsNumeric(CmpName) Then
         If fso.GetExtensionName(Dir(CmpNamePath & "\" & CmpName & "\Company.*")) = "900" Then
            cmpOpenString = myForm.LblTallyERP9Path.Caption & " /TDL:" & App.Path & "\RefreshTDLTallyERP9.txt /DATA:""" & CmpNamePath & """ /LOAD:" & CmpName
            Shell cmpOpenString, vbNormalFocus
        ElseIf fso.GetExtensionName(Dir(CmpNamePath & "\" & CmpName & "\Company.*")) = "1800" Then
            cmpOpenString = myForm.LblTallyPrimePath.Caption & " /TDL:" & App.Path & "\RefreshTDLTallyPrime.txt /DATA:""" & CmpNamePath & """ /LOAD:" & CmpName
            Shell cmpOpenString, vbNormalFocus
        End If
    End If

    Set fso = Nothing
    
End Sub

Private Sub Command4_Click()
    
    Dim myDirPath As String
    
    'TotalNumOfTallyDataFolders = 0
    
    myDirPath = DirMyFolders.Path
        'TotalNumOfTallyDataFolders = CntTallyDataFoders(myDirPath)
        
        Call myMainPathToCreateCompany999File(myDirPath)
    
    'Dim fso As Scripting.FileSystemObject
    'Set fso = New Scripting.FileSystemObject
    'fso.MoveFolder "G:\01. WD Cloud Data\Boss-Sys-Bkup\Old\Boss-System-Backup\D\Boss", "G:\01. WD Cloud Data\Boss-Sys-Bkup\Old\Boss-System-Backup\D\Boss-1"
    
End Sub

Private Sub Command5_Click()

    Dim fso As Scripting.FileSystemObject
    Dim File As Scripting.File
    
    Dim Cnt As Long
    Dim ClientName As String
    Dim FinYearFromFileName As String
    Dim FinYear As String
    Dim WrongFileNames() As String
    
    Dim sFile As String
    Dim dFile As String
    Dim dFileBase As String
    
    Dim FileNumber As Integer
    
    Set fso = New Scripting.FileSystemObject
 
    TotalNumOfWinmanFiles = 0
    ReDim WinmanFileArray(TotalNumOfWinmanFiles)
    
    Call CntWinmanFiles(DirMyFolders)
        ReDim WinmanFileArray(TotalNumOfWinmanFiles)
    
    TotalNumOfWinmanFiles = 0
    Call CntWinmanFiles(DirMyFolders)
                        
    
    ReDim WrongFileNames(UBound(WinmanFileArray))
    'MsgBox "Total No. of Winman File Found are : " & TotalNumOfWinmanFiles
    
    For Cnt = LBound(WinmanFileArray) To UBound(WinmanFileArray) - 1
        
        Set File = fso.GetFile(WinmanFileArray(Cnt))
                Call SelectedFile(File.Name)
            sFile = File
            ClientName = Mid(File.Name, 1, Len(File.Name) - 10)
            ClientName = MakeCompanyNameProper(ClientName)
        
        FinYear = Mid(File.Name, (InStrRev(File.Name, ".") - 5), 5)
            FinYear = "20" & FinYear
        
        If FinYear Like "####-##" Then
        
            If Not fso.FolderExists(DirClientsFolders & "\" & ClientName) Then
                fso.CreateFolder DirClientsFolders & "\" & ClientName
            End If
            
            If Not fso.FolderExists(DirClientsFolders & "\" & ClientName & "\" & FinYear) Then
                fso.CreateFolder DirClientsFolders & "\" & ClientName & "\" & FinYear
            End If
            
            dFile = DirClientsFolders & "\" & ClientName & "\" & FinYear & "\" & File.Name
            dFileBase = Mid(dFile, 1, InStrRev(dFile, ".") - 1)
            
                FileNumber = 1
                    Do While Dir(dFile) <> ""
                        dFile = dFileBase & "(" & FileNumber & ")" & ".tax"
                        FileNumber = FileNumber + 1
                    Loop
                            
            fso.CopyFile sFile, dFile
            
            DirClientsFolders.Refresh
            
        End If
            
        Set File = Nothing
    Next
    
End Sub

Private Sub DirClientsFolders_Change()

    Call LoadToListBox(myForm)
    FileClientFiles.Path = DirClientsFolders.Path
    
End Sub


Private Sub DirClientsFolders_Click()
    Dim fso As Scripting.FileSystemObject
    Dim CurrentDataFolder As Scripting.folder
    
    Set fso = New Scripting.FileSystemObject
    
    'Call FillCompanyDataToControls(Dir(DirClientsFolders.List(DirClientsFolders.ListIndex), vbDirectory))
    'MsgBox DirClientsFolders.List(DirClientsFolders.ListIndex)
    
    Set CurrentDataFolder = fso.GetFolder(DirClientsFolders.List(DirClientsFolders.ListIndex))
    Call FillCompanyDataToControls(CurrentDataFolder)
    
End Sub

Private Sub DirClientsFolders_GotFocus()
    
    Set myDirControl = DirClientsFolders
    CmdBrowse.BackColor = DirClientsFolders.BackColor
    
End Sub

Private Sub DirClientsFolders_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim fso As Scripting.FileSystemObject
    
    Dim SelectedFolder As Scripting.folder
        Dim FolderPath As String
        Dim FolderNameOld As String
        Dim FolderNameNew As String
    Dim Retry As Variant
    
    Set fso = New Scripting.FileSystemObject
    
    If KeyCode = vbKeyF2 And myDirControl.ListCount > 0 Then
        
        Set SelectedFolder = fso.GetFolder(DirClientsFolders.List(DirClientsFolders.ListIndex))
            FolderPath = SelectedFolder.ParentFolder.Path
            FolderNameOld = Dir(SelectedFolder, vbDirectory)
Back:
        FolderNameNew = InputBox("Enter New Name for this Folder", AdminInfo, FolderNameOld)
        FolderNameNew = FolderPath & "\" & FolderNameNew
        
            If fso.FolderExists(FolderNameNew) Then
                Retry = MsgBox("Folder Already Exist...! Try with Different Name", vbRetryCancel + vbCritical, AdminInfo)
                    If Retry = vbRetry Then
                        GoTo Back:
                    Else
                        Exit Sub
                    End If
            End If
            
            fso.MoveFolder SelectedFolder, FolderNameNew
                DirClientsFolders.Refresh
                Call LoadToListBox(myForm)
    End If
    
End Sub

Private Sub DirClientsFolders_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    DirClientsFolders.ToolTipText = Dir(DirClientsFolders.List(DirClientsFolders.ListIndex), vbDirectory)
    
End Sub


Private Sub DirClientsFolders_OLEDragOver(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single, State As Integer)

    Effect = vbDropEffectCopy

End Sub

Private Sub DirMyFolders_Change()
    FileMyFiles.Path = DirMyFolders.Path
End Sub

Private Sub DirMyFolders_Click()

    DirClientsFolders.Path = LblClientsFolderPath.Caption
    Call ClearAllControls
    
End Sub

Private Sub DirMyFolders_GotFocus()

    Set myDirControl = DirMyFolders
    CmdBrowse.BackColor = DirMyFolders.BackColor
    
End Sub

Private Sub DrvMyDrives_Change()
On Error Resume Next
   Call LoadFolders(Mid(DrvMyDrives, 1, 2) & "\", DirMyFolders)
End Sub

Private Sub Form_Activate()

    Dim i As Integer
    
    For i = 1 To 10
        Debug.Print
    Next
    
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)

    If KeyAscii = 27 Then
        Call LoadUnLoadConfig("Write")
        End
    End If
    
End Sub

Private Sub Form_Load()

    Call InitializeAll(Me)
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Me.ActiveControl = Button Then
        MsgBox "Yes"
    End If
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    Call LoadUnLoadConfig("Write")
        Call LoadUnLoadConfig("Write")
        End
End Sub

Private Sub LblClientsFolderPath_Change()

    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    
        If fso.FolderExists(myForm.LblClientsFolderPath.Caption) Then
            myForm.DirClientsFolders.Path = myForm.LblClientsFolderPath.Caption
        End If
        
    Set fso = Nothing
End Sub

Private Sub LblClientsFolderPath_DblClick()

    Dim fso As Scripting.FileSystemObject
    
    Set fso = New Scripting.FileSystemObject
    
        If fso.FolderExists(myForm.LblClientsFolderPath.Caption) Then
            myForm.DirClientsFolders.Path = myForm.LblClientsFolderPath.Caption
        End If
        
    DirClientsFolders.Refresh
    Set fso = Nothing
    
End Sub

Private Sub LblClientsFolderPath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    Dim oShell As Object, oFolder As Object
 
    If Button = 1 And Shift = 2 Then
        Set oShell = CreateObject("Shell.Application")
        Set oFolder = oShell.BrowseForFolder(0, "Select a folder", 0)
        
            If Not oFolder Is Nothing Then
                LblClientsFolderPath.Caption = oFolder.Items.Item.Path
                    'Call LoadConfig("Write", LblClientsFolderPath.Caption)
            End If
        
        Set oFolder = Nothing
        Set oShell = Nothing
    End If
    
End Sub


Private Sub LblCmpName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    LblCmpName.ToolTipText = LblCmpName.Caption
    
End Sub

Private Sub LblParentFolder_DblClick()
    
    LblParentFolder.Caption = ""
    
End Sub

Private Sub LblTallyERP9Path_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 And Shift = 2 Then
        CommonDialogForTally.FileName = ""
        
        CommonDialogForTally.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
        CommonDialogForTally.DialogTitle = "Select an Tally.Exe (ERP9) File"
    
        CommonDialogForTally.ShowOpen
        
        If CommonDialogForTally.FileName <> "" Then
            MsgBox "You selected: " & CommonDialogForTally.FileName, vbOKOnly + vbInformation, AdminInfo
            LblTallyERP9Path.Caption = CommonDialogForTally.FileName
    
                'Call LoadConfig("Write", LblTallyPath.Caption)
        Else
            MsgBox "No File Selected.", vbOKOnly + vbExclamation, AdminInfo
            LblTallyERP9Path.Caption = "Click Me to Get Tally.Exe"
        End If
    End If
    
End Sub

Private Sub LblTallyPrimePath_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 And Shift = 2 Then
        CommonDialogForTally.FileName = ""
        
        CommonDialogForTally.Filter = "Executable Files (*.exe)|*.exe|All Files (*.*)|*.*"
        CommonDialogForTally.DialogTitle = "Select an Tally.Exe (Prime) File"
    
        CommonDialogForTally.ShowOpen
        
        If CommonDialogForTally.FileName <> "" Then
            MsgBox "You selected: " & CommonDialogForTally.FileName, vbOKOnly + vbInformation, AdminInfo
            LblTallyPrimePath.Caption = CommonDialogForTally.FileName
    
                'Call LoadConfig("Write", LblTallyPath.Caption)
        Else
            MsgBox "No File Selected.", vbOKOnly + vbExclamation, AdminInfo
            LblTallyPrimePath.Caption = "Click Me to Get Tally.Exe"
        End If
    End If
    
End Sub


Private Sub LstClients_DblClick()

    If LstClients.ListCount > 1 Then
        LblParentFolder.Caption = LstClients.Text
        LstClients.Selected(LstClients.ListIndex) = False
    Else
        LblParentFolder.Caption = ""
    End If
    
End Sub
