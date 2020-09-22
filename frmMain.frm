VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00008000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "INTERNET"
   ClientHeight    =   360
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   1620
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   360
   ScaleWidth      =   1620
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Left            =   990
      Top             =   0
   End
   Begin VB.Label lblSTATE 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DISCONNECTED"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   1455
   End
   Begin VB.Menu mnuSettings 
      Caption         =   "Settings"
      Visible         =   0   'False
      Begin VB.Menu mnuSettingsAlwaysOnTop 
         Caption         =   "Always on &Top When Connected"
         Checked         =   -1  'True
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' © D R Lambert 2002 - All Rights Reserved
' DRL Development - Bespoke Software & Web Solutions for the Business Community

' This small app will notify you if you are being connected to the internet (with or without
' your knowledge). The small form can be positioned wherever you like, and as soon as an
' internet connection is made the form will appear and stay on top so you are alerted to
' the connection state.
'
' I find this very useful as a safety precaution when looking at other peoples source code
' or trying out software that I'm not familiar with. As it is possible to make an internet
' connection without notifying the user I think this is pretty essential.
'
'


Private Declare Function InternetGetConnectedState _
    Lib "wininet.dll" (ByRef lpdwFlags As Long, _
    ByVal dwReserved As Long) As Long

Private Declare Function SetWindowPos Lib "user32" _
                (ByVal hwnd As Long, _
                 ByVal hWndInsertAfter As Long, _
                 ByVal x As Long, _
                 ByVal y As Long, _
                 ByVal cx As Long, _
                 ByVal cy As Long, _
                 ByVal wFlags As Long) As Long
                 
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40

Private Const INTERVAL_CONNECTED = 5000             ' Five seconds seems about right
Private Const INTERVAL_DISCONNECTED = 5000          ' Tried 30000 here but prefer 5000
Private Const COLOUR_CONNECTED = &HFF&              ' UK spelling of colour 8-)
Private Const COLOUR_DISCONNECTED = &H8000&

Private bAbort As Boolean
Private bOnTop As Boolean

Public Sub AlwaysOnTop(Enable As Boolean)
  Dim lFlag As Long
  
  If Enable Then
    lFlag = HWND_TOPMOST
  Else
    lFlag = HWND_NOTOPMOST
  End If
    
  SetWindowPos Me.hwnd, lFlag, _
    Me.Left / Screen.TwipsPerPixelX, _
    Me.Top / Screen.TwipsPerPixelY, _
    Me.Width / Screen.TwipsPerPixelX, _
    Me.Height / Screen.TwipsPerPixelY, _
    SWP_NOACTIVATE Or SWP_SHOWWINDOW
End Sub

Private Sub Form_Load()
  Dim bScreenResChange As Boolean
  Dim lngLeft As Long
  Dim lngTop As Long
  Dim lngSCRW As Long
  Dim lngSCRH As Long
  
  lngSCRW = GetSetting(App.Title, "Screen", "Width", Screen.Width)
  lngSCRH = GetSetting(App.Title, "Screen", "Height", Screen.Height)
  
  bScreenResChange = (lngSCRW <> Screen.Width)
  bScreenResChange = bScreenResChange Or (lngSCRH <> Screen.Height)
  
  If bScreenResChange Then ' position the form in the center of the screen
    lngLeft = (Screen.Width \ 2) - (Me.Width \ 2)
    lngTop = (Screen.Height \ 2) - (Me.Height \ 2)
  Else
    ' Reposition the form to the users last preference
    Me.Move GetSetting(App.Title, "ScrnPos", "Left", (Screen.Width \ 2) - (Me.Width \ 2)), _
          GetSetting(App.Title, "ScrnPos", "Top", (Screen.Height \ 2) - (Me.Height \ 2))
  End If
  
  bOnTop = GetSetting(App.Title, "ScrnPos", "OnTop", True)
  
  CheckConnectionState True ' call to ensure the status is correct as the program starts,
                            ' adjust the timer1 interval accordingly, and set the always
                            ' on top condition
End Sub

Private Sub CheckConnectionState(Optional ForceCheck As Boolean = False)
  Dim B As Boolean
  Static bPrevious As Boolean
  
  If Not bAbort Then
    B = (InternetGetConnectedState(0&, 0&) <> 0)
    DoEvents
    If B <> bPrevious Or ForceCheck Then  ' Only if the connection state has changed...
      If B Then
        Me.lblSTATE.Caption = "CONNECTED"
        Me.BackColor = COLOUR_CONNECTED
        Me.Timer1.Enabled = False
        Me.Timer1.Interval = INTERVAL_CONNECTED
        Me.Timer1.Enabled = True
        AlwaysOnTop bOnTop
        If bOnTop Then
          Me.Refresh
        End If
        Beep
      Else
        Me.lblSTATE.Caption = "DISCONNECTED"
        Me.BackColor = COLOUR_DISCONNECTED
        Me.Timer1.Enabled = False
        Me.Timer1.Interval = INTERVAL_DISCONNECTED
        Me.Timer1.Enabled = True
        AlwaysOnTop False ' don't stay on top if not connected to the internet
      End If
      bPrevious = B
    End If
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  bAbort = True
  SaveSetting App.Title, "ScrnPos", "OnTop", bOnTop
  SaveSetting App.Title, "ScrnPos", "Left", Me.Left
  SaveSetting App.Title, "ScrnPos", "Top", Me.Top
  SaveSetting App.Title, "Screen", "Width", Screen.Width
  SaveSetting App.Title, "Screen", "Height", Screen.Height
End Sub

Private Sub lblSTATE_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  DoPopUpMenu Button
End Sub

Private Sub mnuSettingsAlwaysOnTop_Click()
  On Error GoTo ErrHandler
  
  bOnTop = Not bOnTop                         ' invert the current bOnTop value
  AlwaysOnTop bOnTop                          ' make it so...
  Me.mnuSettingsAlwaysOnTop.Checked = bOnTop  ' change the menu item checked state to match
  SaveSetting App.Title, "ScrnPos", "OnTop", bOnTop ' save the current setting
  
ResHandler:
  Exit Sub

ErrHandler:
  MsgBox Err.Description, vbCritical, "Error: " & Err.Number
  bOnTop = GetSetting(App.Title, "ScrnPos", "OnTop", True) ' restore the original setting
  Me.mnuSettingsAlwaysOnTop.Checked = bOnTop  ' change the menu item checked state to match
  Resume ResHandler
End Sub

Private Sub Timer1_Timer()
  CheckConnectionState
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
  DoPopUpMenu Button
End Sub

Private Sub DoPopUpMenu(Button As Integer)
  If Button = 2 Then
    PopupMenu mnuSettings
  End If
End Sub
