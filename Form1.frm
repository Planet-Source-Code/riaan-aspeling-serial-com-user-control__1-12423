VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Serial Control Demo"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6630
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   6630
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkHEX 
      Alignment       =   1  'Right Justify
      Caption         =   "Return data as HEX values"
      Height          =   195
      Left            =   4080
      TabIndex        =   11
      Top             =   1740
      Width           =   2355
   End
   Begin VB.TextBox txtSend 
      Enabled         =   0   'False
      Height          =   285
      Left            =   180
      TabIndex        =   4
      Tag             =   "OPEN"
      Top             =   1320
      Width           =   4815
   End
   Begin VB.TextBox txtRec 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3795
      Left            =   180
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Tag             =   "OPEN"
      Top             =   1980
      Width           =   6255
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Send String"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   315
      Index           =   2
      Left            =   5100
      TabIndex        =   5
      Tag             =   "OPEN"
      Top             =   1320
      Width           =   1335
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Close Port"
      Enabled         =   0   'False
      Height          =   495
      Index           =   1
      Left            =   5220
      TabIndex        =   3
      Tag             =   "OPEN"
      Top             =   300
      Width           =   1275
   End
   Begin VB.TextBox txtComSet 
      Height          =   315
      Left            =   1800
      TabIndex        =   1
      Tag             =   "CLOSE"
      Text            =   "9600,n,8,1"
      Top             =   420
      Width           =   1695
   End
   Begin VB.ComboBox cmbPort 
      Height          =   315
      Left            =   480
      TabIndex        =   0
      Tag             =   "CLOSE"
      Top             =   420
      Width           =   1215
   End
   Begin VB.CommandButton Cmd 
      Caption         =   "Open Port"
      Height          =   495
      Index           =   0
      Left            =   3840
      TabIndex        =   2
      Tag             =   "CLOSE"
      Top             =   300
      Width           =   1275
   End
   Begin ARCCommsDemo.ARCComm ARCComm1 
      Left            =   4200
      Top             =   840
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.Label Label4 
      Caption         =   "Text received from Com port:"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   1740
      Width           =   2235
   End
   Begin VB.Line Line1 
      X1              =   6480
      X2              =   120
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label3 
      Caption         =   "Text to send to Com port:"
      Height          =   195
      Left            =   180
      TabIndex        =   9
      Top             =   1080
      Width           =   1875
   End
   Begin VB.Label Label2 
      Caption         =   "Com Settings:"
      Height          =   195
      Left            =   1800
      TabIndex        =   8
      Top             =   180
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Com Port:"
      Height          =   195
      Left            =   480
      TabIndex        =   7
      Top             =   180
      Width           =   795
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ARCComm1_DataIn()
    Dim ans As String
    ans = ARCComm1.GetData
    If chkHEX.Value <> 0 Then
        ans = ReturnHEX(ans)
    End If
    txtRec.Text = txtRec.Text & ans & vbCrLf
    txtRec.SelStart = Len(txtRec)
End Sub

Private Sub Cmd_Click(Index As Integer)
On Error GoTo handelcmd
    Select Case Index
        Case 0 'Open Com Port
            ARCComm1.ComPort = cmbPort.List(cmbPort.ListIndex)
            ARCComm1.ComSettings = txtComSet.Text
            If ARCComm1.InitCom() Then
                SetControllsLockStatus "CLOSE", False
                SetControllsLockStatus "OPEN", True
            End If
        Case 1 'Close Com Port
            ARCComm1.FinCom
            SetControllsLockStatus "CLOSE", True
            SetControllsLockStatus "OPEN", False
        Case 2 'Send to com port
            If Not ARCComm1.WriteData(txtSend) Then
                MsgBox "Error sending data!"
            End If
            txtSend.SetFocus
    End Select
    Exit Sub
handelcmd:
    MsgBox Err.Description, 16, "Error#" & Err.Number
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo handelload
    cmbPort.AddItem "COM1:"
    cmbPort.AddItem "COM2:"
    cmbPort.AddItem "COM3:"
    cmbPort.AddItem "COM4:"
    cmbPort.ListIndex = 0
    Exit Sub
handelload:
    MsgBox Err.Description, 16, "Error#" & Err.Number
    Exit Sub
End Sub

Sub SetControllsLockStatus(TheTAGValue As String, TheLockStatus As Boolean)
On Error GoTo handelSetControllsLockStatus
    Dim xs As Control
    For Each xs In Me
        If xs.Tag = TheTAGValue Then
            xs.Enabled = TheLockStatus
        End If
    Next
    Exit Sub
handelSetControllsLockStatus:
    MsgBox Err.Description, 16, "Error#" & Err.Number
    Exit Sub
End Sub

Function ReturnHEX(TheString As String) As String
    Dim i As Integer, RtnStr As String
    RtnStr = ""
    For i = 1 To Len(TheString)
        RtnStr = RtnStr & Right$("00" & Hex$(Asc(Mid$(TheString, i, 1))), 2) & " "
    Next
    ReturnHEX = RtnStr
End Function
