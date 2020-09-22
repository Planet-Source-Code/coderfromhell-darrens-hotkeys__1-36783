Attribute VB_Name = "ModMain"
Option Explicit

Public Declare Function RegisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long, ByVal fsModifiers As Long, ByVal vk As Long) As Long
Public Declare Function UnregisterHotKey Lib "user32" (ByVal hwnd As Long, ByVal id As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Sub CopyMemoryH Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByVal Source As Any, ByVal Length As Long)
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'user defined type required by Shell_NotifyIcon API call
Public Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    uCallBackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type

Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4
Public Const WM_MOUSEMOVE = &H200
Public Const WM_LBUTTONDOWN = &H201     'Button down
Public Const WM_LBUTTONUP = &H202       'Button up
Public Const WM_LBUTTONDBLCLK = &H203   'Double-click
Public Const WM_RBUTTONDOWN = &H204     'Button down
Public Const WM_RBUTTONUP = &H205       'Button up
Public Const WM_RBUTTONDBLCLK = &H206   'Double-click

Public Const GWL_WNDPROC = (-4)
Public Const SW_HIDE = 0
Public Const WM_HOTKEY = &H312
Public Const MOD_SHIFT = &H4
Public Const MOD_WIN = &H8
Public Const VK_5 = &H35
Public Const VK_6 = &H36
Public Const VK_1 = &H31

Public Const VB5_HOTKEY = &H5F      '
Public Const VB6_HOTKEY = &H6F      '
Public Const IE_HOTKEY = &H7F       '

Public nid As NOTIFYICONDATA
Public lngOldWindowProc As Long


Public HotKeys As Collection


Public Function SubProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If wMsg = WM_HOTKEY Then
        Dim i As Long
        If Not HotKeys Is Nothing Then
            For i = 1 To HotKeys.Count
                If HotKeys(i).Registered = True Then
                    If LoWord(lParam) = MOD_WIN And HiWord(lParam) = HotKeys(i).KeyCode_ And wParam = HotKeys(i).HotKey Then
                        If Not HotKeys(i).FileName = "" Then
                            If Not Dir(HotKeys(i).FileName) = "" Then
                                RunFile HotKeys(i).FileName
                            End If
                        End If
                    End If
                End If
            Next i
        End If
    End If
    SubProc = CallWindowProc(lngOldWindowProc, hwnd, wMsg, wParam, lParam)
End Function

Public Sub RunFile(ByVal FileName As String)
   Dim success As Long
   success = ShellExecute(0&, vbNullString, FileName, vbNullString, vbNullString, vbNormalFocus)
End Sub

Public Function LoWord(ByVal dw As Long) As Integer
    On Error GoTo Err_LoWord
    CopyMemoryH LoWord, ByVal VarPtr(dw), 2
Exit_LoWord:
    Exit Function
Err_LoWord:
    GoTo Exit_LoWord
End Function

Public Function HiWord(ByVal dw As Long) As Integer
    On Error GoTo Err_HiWord
    CopyMemoryH HiWord, ByVal VarPtr(dw) + 2, 2
Exit_HiWord:
    Exit Function
    
Err_HiWord:
    GoTo Exit_HiWord
End Function

Public Sub LoadHotKeys()
    Dim i As Integer, j As Long, k As Integer
    Dim NewItem As clsHotKey
    RemoveHotKeys
    Set HotKeys = New Collection
    Const Chars = "abcdefghijklmnopqrstuvwxyz"
    j = 48
    k = 1
    For i = 0 To 9  '(48-57)
        Set NewItem = New clsHotKey
        With NewItem
            .Enabled = False
            .FileName = ""
            .KeyCode_ = j
            .Index = k - 1
        End With
        HotKeys.Add NewItem
        j = j + 1
        k = k + 1
    Next i
    j = 65
    For i = 0 To 25 '(65-90)
        Set NewItem = New clsHotKey
        With NewItem
            .Enabled = False
            .FileName = ""
            .KeyCode_ = j
            .Index = k - 1
        End With
        HotKeys.Add NewItem
        j = j + 1
        k = k + 1
    Next i
    j = 112
    For i = 1 To 12 '(112-127)
        Set NewItem = New clsHotKey
        With NewItem
            .Enabled = False
            .FileName = ""
            .KeyCode_ = j
            .Index = k - 1
        End With
        HotKeys.Add NewItem
        j = j + 1
        k = k + 1
    Next i
    Dim KeyFile As String, FileName As String, Stg1 As String, Stg2 As String
    KeyFile = App.Path & "\Keys.dll"
    If Not Dir(KeyFile) = "" Then
        Close #1
        Open KeyFile For Input As #1
            Do Until EOF(1)
                Input #1, Stg1, Stg2, FileName
                With HotKeys(CLng(Stg1) + 1)
                    .Enabled = CBool(Stg2)
                    .FileName = FileName
                End With
            Loop
        Close #1
    End If
    RegisterHotKeys
End Sub

Private Sub RemoveHotKeys()
    Dim i As Long
    If Not HotKeys Is Nothing Then
        UnregisterHotKeys
        For i = HotKeys.Count To 1 Step -1
            HotKeys.Remove i
        Next i
        Set HotKeys = Nothing
    End If
End Sub

Private Sub RegisterHotKeys()
    If Not HotKeys Is Nothing Then
        Dim i As Long
        For i = 1 To HotKeys.Count
            If HotKeys(i).FileName <> "" And HotKeys(i).Enabled = True Then
                RegisterHotKey frmHotKeys.hwnd, HotKeys(i).HotKey, MOD_WIN, HotKeys(i).KeyCode_
                HotKeys(i).Registered = True
            End If
        Next i
    End If
    lngOldWindowProc = SetWindowLong(frmHotKeys.hwnd, GWL_WNDPROC, AddressOf SubProc)
End Sub

Private Sub UnregisterHotKeys()
    If Not HotKeys Is Nothing Then
        Dim i As Long
        For i = 1 To HotKeys.Count
            If HotKeys(i).Registered = True Then
                UnregisterHotKey frmHotKeys.hwnd, HotKeys(i).HotKey
                HotKeys(i).Registered = False
            End If
        Next i
    End If
    Call SetWindowLong(frmHotKeys.hwnd, GWL_WNDPROC, lngOldWindowProc)
End Sub

Public Sub Main()
    Load frmHotKeys
    LoadHotKeys
End Sub

Public Sub UnloadHotKeys()
    RemoveHotKeys
End Sub

