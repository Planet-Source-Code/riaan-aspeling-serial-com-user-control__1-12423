VERSION 5.00
Begin VB.UserControl ARCComm 
   ClientHeight    =   735
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1110
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   735
   ScaleWidth      =   1110
   ToolboxBitmap   =   "ARCComm.ctx":0000
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   480
      Left            =   0
      Picture         =   "ARCComm.ctx":00FA
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   0
      Top             =   0
      Width           =   480
   End
End
Attribute VB_Name = "ARCComm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ComNum As Long
Private ComName As String
Private ComSett As String
Private IntervalValue As Integer
Private LastData As String
Private ComIsOpen As Boolean
Private DataInRead As Boolean
Private DataSend As String, DataToWaitFor As String, SendAndWait As Boolean, SendAndWaitRetries As Integer
Private MaxChars As Integer
Public Event SendAndReceived(DataSend As String, DataReceived As String, DataToWaitFor As String)
Public Event DataIn()

Private Type COMSTAT
        fCtsHold As Long
        fDsrHold As Long
        fRlsdHold As Long
        fXoffHold As Long
        fXoffSent As Long
        fEof As Long
        fTxim As Long
        fReserved As Long
        cbInQue As Long
        cbOutQue As Long
End Type
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_FLAG_NO_BUFFERING = &H20000000
Private Const FILE_FLAG_WRITE_THROUGH = &H80000000
Private Type COMMTIMEOUTS
        ReadIntervalTimeout As Long
        ReadTotalTimeoutMultiplier As Long
        ReadTotalTimeoutConstant As Long
        WriteTotalTimeoutMultiplier As Long
        WriteTotalTimeoutConstant As Long
End Type
Private Type DCB
        DCBlength As Long
        BaudRate As Long
        fBinary As Long
        fParity As Long
        fOutxCtsFlow As Long
        fOutxDsrFlow As Long
        fDtrControl As Long
        fDsrSensitivity As Long
        fTXContinueOnXoff As Long
        fOutX As Long
        fInX As Long
        fErrorChar As Long
        fNull As Long
        fRtsControl As Long
        fAbortOnError As Long
        fDummy2 As Long
        wReserved As Integer
        XonLim As Integer
        XoffLim As Integer
        ByteSize As Byte
        Parity As Byte
        StopBits As Byte
        XonChar As Byte
        XoffChar As Byte
        ErrorChar As Byte
        EofChar As Byte
        EvtChar As Byte
End Type
Private Declare Function ClearCommError Lib "kernel32" (ByVal hFile As Long, lpErrors As Long, lpStat As COMSTAT) As Long
Private Declare Function BuildCommDCBAndTimeouts Lib "kernel32" Alias "BuildCommDCBAndTimeoutsA" (ByVal lpDef As String, lpDCB As DCB, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function apiBuildCommDCB Lib "kernel32" Alias "BuildCommDCBA" (ByVal lpDef As String, lpDCB As DCB) As Long
Private Declare Function SetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function GetCommTimeouts Lib "kernel32" (ByVal hFile As Long, lpCommTimeouts As COMMTIMEOUTS) As Long
Private Declare Function apiSetCommState Lib "kernel32" Alias "SetCommState" (ByVal hCommDev As Long, lpDCB As DCB) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Private Declare Function PurgeComm Lib "kernel32" (ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Private Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Private Declare Function CreateFile& Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName$, ByVal dwDesiredAccess&, ByVal dwShareMode&, ByVal lpSecurityAttributes&, ByVal dwCreationDisposition&, ByVal dwFlagsAndAttributes&, ByVal hTemplateFile&)
Private Declare Function CloseHandle& Lib "kernel32" (ByVal hObject&)
Private Declare Function FlushFileBuffers& Lib "kernel32" (ByVal hFile&)

Private WithEvents CommTime As XTimer
Attribute CommTime.VB_VarHelpID = -1

Public Function GetData() As String
    GetData = LastData
    LastData = ""
End Function

Public Function WriteData(TextWrite As String) As Boolean
    WriteData = WriteComm32(ComNum, TextWrite)
End Function

Public Function WriteAndWaitFor(TextToWrite As String, TextToWaitFor As String) As Boolean
    Dim RtnVal As Boolean
    DataSend = TextToWrite
    DataToWaitFor = TextToWaitFor
    If Not WriteComm32(ComNum, DataSend) Then
        SendAndWait = False
        DataSend = ""
        DataToWaitFor = ""
        RaiseEvent SendAndReceived("", "", TextToWaitFor)
        Exit Function
    End If
    SendAndWait = True
End Function

Public Function InitCom(Optional ComNumber As String = "ComX:", Optional ComSettings As String = "X") As Boolean
    If ComIsOpen Then
        FinCom
    End If
    If ComNumber = "ComX:" Then ComNumber = ComName
    If ComSettings = "X" Then ComSettings = ComSett
    Dim ComSetup As DCB, Answer, Stat As COMSTAT
    Dim retval As Long
    Dim CtimeOut As COMMTIMEOUTS, BarDCB As DCB
    ' Open the communications port for read/write (&HC0000000).
    ' Must specify existing file (3).
    ComNum = CreateFile(ComNumber, &HC0000000, 0, 0, 3, 0, 0)
    If ComNum = -1 Then
        MsgBox "Com Port " & ComNumber & " not available. Use Serial settings (on the main menu) to setup your ports.", 48
        InitCom = False
        Exit Function
    End If
    'Setup Time Outs for com port
    CtimeOut.ReadIntervalTimeout = 200
    CtimeOut.ReadTotalTimeoutConstant = 1
    CtimeOut.ReadTotalTimeoutMultiplier = 1
    CtimeOut.WriteTotalTimeoutConstant = 1
    CtimeOut.WriteTotalTimeoutMultiplier = 1
    retval = SetCommTimeouts(ComNum, CtimeOut)
    If retval = -1 Then
        retval = GetLastError()
        MsgBox "Unable to set timeouts for port " & ComNumber & " Error: " & retval
        retval = CloseHandle(ComNum)
        InitCom = False
        Exit Function
    End If
    retval = apiBuildCommDCB(ComSettings, BarDCB)
    If retval = -1 Then
        retval = GetLastError()
        MsgBox "Unable to build Comm DCB " & ComSettings & " Error: " & retval
        retval = CloseHandle(ComNum)
        InitCom = False
        Exit Function
    End If
    retval = apiSetCommState(ComNum, BarDCB)
    If retval = -1 Then
        retval = GetLastError()
        MsgBox "Unable to set Comm DCB " & ComSettings & " Error: " & retval
        retval = CloseHandle(ComNum)
        InitCom = False
        Exit Function
    End If
    
    InitCom = True
    ComIsOpen = True
    CommTime.Interval = IntervalValue
    CommTime.Enabled = True
End Function

Private Sub CommTime_Tick()
    Dim Barst As String
    Dim Stat As COMSTAT
    Dim CharLen As Boolean
    Static Retries As Integer
    Barst = ReadComm32(ComNum)
    DoEvents
    If MaxChars > 0 Then
        If Len(LastData) < MaxChars Then
            CharLen = True
        Else
            CharLen = False
        End If
    Else
        CharLen = True
    End If
    If Len(Barst) > 0 And CharLen Then
        LastData = LastData & Barst
        DataInRead = True
       Else
        If DataInRead Then
            FlushComm
            DataInRead = False
            If SendAndWait Then
                If InStr(1, LastData, DataToWaitFor) = 0 Then
                    Retries = Retries + 1
                    'Debug.Print "Retries: " & Retries & " : " & DataSend & " " & DataToWaitFor & " -> " & LastData
                    WriteAndWaitFor DataSend, DataToWaitFor
                    If Retries > SendAndWaitRetries Then
                        Retries = 0
                        SendAndWait = False
                        RaiseEvent SendAndReceived("", "", DataToWaitFor)
                    End If
                Else
                    Retries = 0
                    SendAndWait = False
                    RaiseEvent SendAndReceived(DataSend, LastData, DataToWaitFor)
                    LastData = ""
                End If
            Else
                RaiseEvent DataIn
            End If
        Else
            'There's no reply and it's a send and wait for
            If SendAndWait Then
                Retries = Retries + 1
                'Debug.Print "Retries: " & Retries & " : " & DataSend & " " & DataToWaitFor & " -> " & LastData
                WriteAndWaitFor DataSend, DataToWaitFor
                If Retries > SendAndWaitRetries Then
                    Retries = 0
                    SendAndWait = False
                    RaiseEvent SendAndReceived("", "", DataToWaitFor)
                End If
            End If
        End If
    End If
End Sub

Private Sub UserControl_Initialize()
    Set CommTime = New XTimer
End Sub

Private Sub UserControl_InitProperties()
    ComIsOpen = False
    LastData = ""
    DataInRead = False
    SendAndWaitRetries = 10
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    ComName = PropBag.ReadProperty("ComName", "Com1:")
    ComSett = PropBag.ReadProperty("ComSettings", "9600,N,8,1")
    IntervalValue = PropBag.ReadProperty("IntervalValue", 500)
    SendAndWaitRetries = PropBag.ReadProperty("SendAndWaitRetries", 10)
    MaxChars = PropBag.ReadProperty("MaxChars", 255)
End Sub

Public Property Get MaximumChars() As Integer
    MaximumChars = MaxChars
End Property

Public Property Let MaximumChars(ByVal NewVal As Integer)
    MaxChars = NewVal
End Property


Public Property Get WriteAndWaitForRetries() As Integer
    WriteAndWaitForRetries = SendAndWaitRetries
End Property

Public Property Let WriteAndWaitForRetries(ByVal NewVal As Integer)
    SendAndWaitRetries = NewVal
End Property

Public Property Get ComPort() As String
    ComPort = ComName
End Property

Public Property Let ComPort(ByVal NewComPort As String)
    ComName = NewComPort
    PropertyChanged ComName
End Property

Public Property Get ComSettings() As String
    ComSettings = ComSett
End Property

Public Property Let ComSettings(ByVal NewComSettings As String)
    ComSett = NewComSettings
    PropertyChanged ComSett
End Property

Public Property Get IntervalVal() As Integer
    IntervalVal = IntervalValue
End Property

Public Property Let IntervalVal(ByVal NewInterval As Integer)
    IntervalValue = NewInterval
    PropertyChanged IntervalValue
End Property

Private Sub UserControl_Resize()
    Call UserControl_Show
End Sub

Private Sub UserControl_Show()
    UserControl.Height = Pic.Height
    UserControl.Width = Pic.Width
End Sub

Private Sub UserControl_Terminate()
    FinCom
End Sub

Public Function FinCom()
On Error Resume Next
    CommTime.Enabled = False
    Set CommTime = Nothing
    FinCom = CloseHandle(ComNum)
    ComIsOpen = False
    If Err.Number <> 0 Then Err.Clear
End Function

Private Function ReadComm32(PortHwnd As Long) As String
    Dim RetBytes As Long, i As Integer, ReadStr As String, bRead(256) As Byte, retval As Long
    retval = ReadFile(PortHwnd, bRead(0), 256, RetBytes, 0)
    ReadStr = ""
    If (RetBytes > 0) Then  ' And (RetBytes < 256)
        For i = 0 To RetBytes - 1
            ReadStr = ReadStr & Chr(bRead(i))
        Next i
    End If
    ReadComm32 = ReadStr
End Function

Private Function WriteComm32(PortHwnd As Long, TextValue As String) As Boolean
On Error GoTo handelwritelpt
    Dim RetBytes As Long, LenVal As Long
    Dim retval As Long, bRead(256) As Byte
    
    CommTime.Enabled = False
    CommTime.Interval = 0
    
    If Len(TextValue) > 255 Then
        WriteComm32 PortHwnd, Left$(TextValue, 255)
        WriteComm32 PortHwnd, Right$(TextValue, Len(TextValue) - 255)
        Exit Function
    End If
    
    For LenVal = 0 To Len(TextValue) - 1
        bRead(LenVal) = Asc(Mid$(TextValue, LenVal + 1, 1))
    Next LenVal
    
    retval = WriteFile(PortHwnd, bRead(0), Len(TextValue), RetBytes, 0)
    
    If RetBytes = Len(TextValue) Then
        WriteComm32 = True
       Else
        WriteComm32 = False
    End If
    
    CommTime.Interval = IntervalValue
    CommTime.Enabled = True
    
handelwritelpt:
    Exit Function
End Function

Private Function FlushComm()
    FlushFileBuffers (ComNum)
End Function

Public Property Get Enabled() As Boolean
    Enabled = CommTime.Enabled
End Property

Public Property Let Enabled(NewValue As Boolean)
    CommTime.Enabled = NewValue
End Property
