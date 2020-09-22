VERSION 5.00
Begin VB.Form frmHotKeys 
   Caption         =   "Daz HotKeys"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   Icon            =   "HotKeys.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu mnu 
      Caption         =   "mnu"
      Visible         =   0   'False
      Begin VB.Menu mnuCustomize 
         Caption         =   "Customize"
      End
      Begin VB.Menu Bar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frmHotKeys"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    With nid
        .cbSize = Len(nid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .uCallBackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = "Darrens HotKeys" & vbNullChar
    End With
    Shell_NotifyIcon NIM_ADD, nid
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'this procedure receives the callbacks from the System Tray icon.
    Dim Result As Long
    Dim msg As Long
    'the value of X will vary depending upon the scalemode setting
    If Me.ScaleMode = vbPixels Then
        msg = X
    Else
        msg = X / Screen.TwipsPerPixelX
    End If
    Select Case msg
        Case WM_RBUTTONUP
            Result = SetForegroundWindow(Me.hwnd)
            Me.PopupMenu Me.mnu, , , , mnuExit
    End Select
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ModMain.UnloadHotKeys
    Shell_NotifyIcon NIM_DELETE, nid
End Sub

Private Sub mnuCustomize_Click()
    frmEdit.Show
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub
