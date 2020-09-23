VERSION 5.00
Begin VB.UserControl ctlSizeRestrictor 
   CanGetFocus     =   0   'False
   ClientHeight    =   435
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   435
   ScaleWidth      =   435
   ToolboxBitmap   =   "SizeRestrictor.ctx":0000
   Begin SizeRestrictor.Subclass Subclass 
      Left            =   480
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
   End
   Begin VB.Image imgIcon 
      Height          =   420
      Left            =   0
      Picture         =   "SizeRestrictor.ctx":0312
      Top             =   0
      Width           =   420
   End
End
Attribute VB_Name = "ctlSizeRestrictor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'SizeRestrictor - Visual Basic Control
'Copyright (c) 2002 by Keith Hall
'sizerestrictor@khall.cjb.net

Option Explicit

' Subclass stuff (to stop the window from flashing when resizing)
'----------------
'Windows declarations
Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (lpDest As Any, lpSource As Any, ByVal nCount As Long)

'Windows constants
Private Const WM_GETMINMAXINFO = &H24

'Windows data types
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type MINMAXINFO
    ptReserved As POINTAPI
    ptMaxSize As POINTAPI
    ptMaxPosition As POINTAPI
    ptMinTrackSize As POINTAPI
    ptMaxTrackSize As POINTAPI
End Type

'Declarations
Dim lMaxWidth As Long
Dim lMinWidth As Long
Dim lMaxHeight As Long
Dim lMinHeight As Long
Dim bHasMaxSize As Boolean

Private Sub Subclass_WndProc(Msg As Long, wParam As Long, lParam As Long, Result As Long)
    Dim MinMax As MINMAXINFO, Status As String

    If Msg = WM_GETMINMAXINFO Then
        'Copy to our local MinMax variable
        CopyMemory MinMax, ByVal lParam, Len(MinMax)
        'Set minimum/maximum tracking size
        MinMax.ptMinTrackSize.X = lMinWidth
        MinMax.ptMinTrackSize.Y = lMinHeight
        MinMax.ptMaxTrackSize.X = lMaxWidth
        MinMax.ptMaxTrackSize.Y = lMaxHeight
        'Copy data back to Windows
        CopyMemory ByVal lParam, MinMax, Len(MinMax)
        Result = 0
    End If
End Sub

Private Sub UserControl_Initialize()
    If lMaxHeight = 0 Then lMaxHeight = Screen.Height
    If lMaxWidth = 0 Then lMaxWidth = Screen.Width
    If lMinWidth = 0 Then lMinWidth = 200
    If lMinHeight = 0 Then lMinHeight = 200
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    lMaxHeight = PropBag.ReadProperty("MaxHeight", Screen.Height)
    lMaxWidth = PropBag.ReadProperty("MaxWidth", Screen.Width)
    lMinHeight = PropBag.ReadProperty("MinHeight", 200)
    lMinWidth = PropBag.ReadProperty("MinWidth", 200)
    bHasMaxSize = PropBag.ReadProperty("HasMaxSize", False)
    
    If bHasMaxSize = False Then
        lMaxWidth = Screen.Width
        lMaxHeight = Screen.Height
    End If
       
    If lMaxHeight < lMinHeight Then
        lMaxHeight = lMinHeight
    End If
    If lMaxWidth < lMinWidth Then
        lMaxWidth = lMinWidth
    End If
    
    If Not UserControl.Ambient.UserMode Then Exit Sub
    On Error Resume Next
    Subclass.hWnd = UserControl.Parent.hWnd
    Subclass.Messages(WM_GETMINMAXINFO) = True
End Sub

Private Sub UserControl_Resize()
    Size imgIcon.Width, imgIcon.Height
End Sub

Public Property Let HasMaxSize(bNewValue As Boolean)
    If bNewValue = False Then
        lMaxHeight = Screen.Height
        lMaxWidth = Screen.Width
    End If
    bHasMaxSize = bNewValue
End Property

Public Property Get HasMaxSize() As Boolean
    HasMaxSize = bHasMaxSize
End Property

Public Property Let MinWidth(lNewValue As Long)
    lMinWidth = lNewValue
End Property

Public Property Get MinWidth() As Long
    MinWidth = lMinWidth
End Property

Public Property Let MinHeight(lNewValue As Long)
    lMinHeight = lNewValue
End Property

Public Property Get MinHeight() As Long
    MinHeight = lMinHeight
End Property

Public Property Let MaxWidth(lNewValue As Long)
    If bHasMaxSize = True Then
        If lNewValue < lMinWidth Then lNewValue = lMinWidth
        lMaxWidth = lNewValue
    Else
        lMaxWidth = Screen.Width
    End If
End Property

Public Property Get MaxWidth() As Long
    MaxWidth = lMaxWidth
End Property

Public Property Let MaxHeight(lNewValue As Long)
    If bHasMaxSize = True Then
        If lNewValue < lMinHeight Then lNewValue = lMinHeight
        lMaxHeight = lNewValue
    Else
        lMaxHeight = Screen.Height
    End If
End Property

Public Property Get MaxHeight() As Long
    MaxHeight = lMaxHeight
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("MaxHeight", lMaxHeight, Screen.Height)
    Call PropBag.WriteProperty("MaxWidth", lMaxWidth, Screen.Width)
    Call PropBag.WriteProperty("MinHeight", lMinHeight, 200)
    Call PropBag.WriteProperty("MinWidth", lMinWidth, 200)
    Call PropBag.WriteProperty("HasMaxSize", bHasMaxSize, False)
End Sub
