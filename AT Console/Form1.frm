VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "GSM Modem AT Command Console"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cbobaud 
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   3600
      List            =   "Form1.frx":0019
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   120
      Width           =   1335
   End
   Begin VB.CheckBox chkpref 
      Caption         =   "Use ""AT"" prefix"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4560
      Width           =   1575
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   120
      Top             =   3720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      RTSEnable       =   -1  'True
   End
   Begin VB.CommandButton cmdconnect 
      Caption         =   "Connect"
      Height          =   315
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cboport 
      Height          =   315
      ItemData        =   "Form1.frx":004F
      Left            =   1200
      List            =   "Form1.frx":0083
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtenter 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   6255
   End
   Begin VB.TextBox txtout 
      Height          =   3735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   "Baud Rate:"
      Height          =   255
      Index           =   1
      Left            =   2640
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Port Number:"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()

End Sub

Private Sub cmdconnect_Click()

If cboport.ListIndex = -1 Or cbobaud.ListIndex = -1 Then
Exit Sub
End If


If cmdconnect.Caption = "Connect" Then

ConnectModem MSComm1, Replace(cboport.Text, "COM", ""), cbobaud.Text
cmdconnect.Caption = "Disconnect"

ElseIf cmdconnect.Caption = "Disconnect" Then

MSComm1.PortOpen = False

cmdconnect.Caption = "Connect"
End If

End Sub

Private Sub cmdexe_Click()
If chkpref.Value = Unchecked Then
ExecuteModem txtenter & vbCrLf, MSComm1
ElseIf chkpref.Value = Checked Then
ExecuteModem "AT" & txtenter & vbCrLf, MSComm1
End If

txtenter.Text = ""
txtenter.SetFocus
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub MSComm1_OnComm()
txtout = txtout & MSComm1.Input
txtout.SelStart = Len(txtout)
End Sub

Private Sub txtenter_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdexe_Click
End If
End Sub
