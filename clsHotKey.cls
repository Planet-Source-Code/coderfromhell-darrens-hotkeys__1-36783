VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsHotKey"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public FileName As String
Public Index As Integer
Public KeyCode_ As Integer
Public Enabled As Boolean
Public Registered As Boolean

Public Property Get HotKey() As Long
    HotKey = Index + 90
End Property
