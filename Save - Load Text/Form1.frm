VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "'SaveText' - 'LoadText'  Functions For TextBoxes."
   ClientHeight    =   4305
   ClientLeft      =   75
   ClientTop       =   420
   ClientWidth     =   6915
   LinkTopic       =   "Form1"
   ScaleHeight     =   4305
   ScaleWidth      =   6915
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   495
      Left            =   5400
      TabIndex        =   3
      Top             =   3600
      Width           =   1335
   End
   Begin MSComDlg.CommonDialog cdlg 
      Left            =   5760
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "LoadText"
      Height          =   495
      Left            =   5400
      TabIndex        =   2
      Top             =   2160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "SaveText"
      Height          =   495
      Left            =   5400
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   4095
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    cdlg.ShowSave
    SaveText Text1, cdlg.FileName & ".txt"
End Sub
Private Sub Command2_Click()
    cdlg.ShowOpen
    Text1 = LoadText(cdlg.FileName)
End Sub
Private Sub Command3_Click()
    Text1 = ""
End Sub
Private Sub Form_Unload(Cancel As Integer)
    Clipboard.SetText ("http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=57313&lngWId=1")
    If MsgBox("Is it Satisfactory?", vbQuestion + vbYesNo, "?") = vbYes Then
        MsgBox "Please Rate my code,The site address is already copied to your clipboard", vbInformation, "ThankYou"
    Else
        MsgBox "Please give FeedBack,The site address is already copied to your clipboard", vbInformation, "Please Give FeedBack"
    End If
End Sub
