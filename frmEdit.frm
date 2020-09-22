VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmEdit 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Customize"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "frmEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   4710
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD 
      Left            =   600
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton CancelButton 
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   285
      Left            =   3720
      TabIndex        =   2
      Top             =   2040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   3495
   End
   Begin VB.ListBox lstKeys 
      Appearance      =   0  'Flat
      Height          =   1830
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
End
Attribute VB_Name = "frmEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DontUpdate As Boolean
Dim Update As Boolean

Private Sub CancelButton_Click()
    Update = False
    Unload Me
End Sub

Private Sub Command1_Click()
    On Error GoTo XS
    CD.Flags = cdlOFNExplorer + cdlOFNFileMustExist
    CD.Filter = "Applications (*.exe)|*.exe|All Files (*.*)|*.*"
    CD.DialogTitle = "Select Application"
    CD.ShowOpen
    Me.Text1.Text = CD.FileName
XS:
End Sub

Private Sub Form_Load()
    Dim i As Integer, j As Long
    Const Chars = "abcdefghijklmnopqrstuvwxyz"
    Me.Show
    DoEvents
    j = 48
    For i = 0 To 9  '(48-57)
        lstKeys.AddItem "Key " & i
        lstKeys.ItemData(lstKeys.NewIndex) = j
        j = j + 1
    Next i
    j = 65
    For i = 0 To 25 '(65-90)
        lstKeys.AddItem "Key " & UCase(Mid(Chars, i + 1, 1))
        lstKeys.ItemData(lstKeys.NewIndex) = j
        j = j + 1
    Next i
    j = 112
    For i = 1 To 12 '(112-127)
        lstKeys.AddItem "Key F" & i
        lstKeys.ItemData(lstKeys.NewIndex) = j
        j = j + 1
    Next i
    LoadHotKeys
End Sub

Private Sub SaveHotKeys()
    Dim KeyFile As String, i As Long
    KeyFile = App.Path & "\Keys.dll"
    If Dir(KeyFile) <> "" Then Kill KeyFile
    Close #1
    Open KeyFile For Output As #1
        For i = 1 To HotKeys.Count
            If Not Trim(HotKeys(i).FileName) = "" Then
                Print #1, HotKeys(i).Index & ","; HotKeys(i).Enabled & "," & HotKeys(i).FileName
            End If
        Next i
    Close #1
    ModMain.UnloadHotKeys
    ModMain.LoadHotKeys
End Sub

Private Sub LoadHotKeys()
    Dim i As Long
    If Not HotKeys Is Nothing Then
        For i = 1 To HotKeys.Count
            If HotKeys(i).Enabled = True Then lstKeys.Selected(i - 1) = True
        Next i
    End If
    lstKeys.ListIndex = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    If Update Then SaveHotKeys
End Sub

Private Sub lstKeys_Click()
    DontUpdate = True
    Text1.Text = HotKeys(lstKeys.ListIndex + 1).FileName
    DontUpdate = False
End Sub

Private Sub lstKeys_ItemCheck(Item As Integer)
    If lstKeys.Selected(Item) = True Then
        HotKeys(Item + 1).Enabled = True
    Else
        HotKeys(Item + 1).Enabled = False
    End If
End Sub

Private Sub OKButton_Click()
    Update = True
    Unload Me
End Sub

Private Sub Text1_Change()
    If Not DontUpdate Then HotKeys(lstKeys.ListIndex + 1).FileName = Text1.Text
End Sub
