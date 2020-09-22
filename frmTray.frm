VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Tray Example by Mischa Balen"
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   720
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label1 
      Caption         =   "System Tray Example"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'System Tray Example, by Mischa Balen
'Submitted to PSC
'Placed into the public domain

Private Sub Form_Load()
    Me.Show 'form must be fully visible
    Me.Refresh
        
        With nid 'with system tray
            .cbSize = Len(nid)
            .hwnd = Me.hwnd
            .uId = vbNull
            .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
            .uCallBackMessage = WM_MOUSEMOVE
            .hIcon = Me.Icon 'use form's icon in tray
            .szTip = "System Tray Example" & vbNullChar 'tooltip text
        End With
        
    Shell_NotifyIcon NIM_ADD, nid 'add to tray
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim Result, Action As Long
    
    'there are two display modes and we need to find out
    'which one the application is using
    
    If Me.ScaleMode = vbPixels Then
        Action = X
    Else
        Action = X / Screen.TwipsPerPixelX
    End If
    
Select Case Action

    Case WM_LBUTTONDBLCLK 'Left Button Double Click
        Me.WindowState = vbNormal 'put into taskbar
            Result = SetForegroundWindow(Me.hwnd)
        Me.Show 'show form
    
    Case WM_RBUTTONUP 'Right Button Up
        Result = SetForegroundWindow(Me.hwnd)
        PopupMenu mnuFile 'popup menu, cool eh?
    
    End Select
    
End Sub

Private Sub Form_Unload(Cancel As Integer) 'on form unload
    Shell_NotifyIcon NIM_DELETE, nid 'remove from tray
End Sub

Private Sub mnuExit_Click() 'exit
    Unload Me: End
End Sub
