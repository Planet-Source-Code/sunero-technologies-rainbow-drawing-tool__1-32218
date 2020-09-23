VERSION 5.00
Begin VB.Form frmDraw 
   Caption         =   "Rainbow Draw"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6225
   LinkTopic       =   "Form1"
   ScaleHeight     =   412
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   3  'Windows Default
   Begin VB.HScrollBar hBar 
      Height          =   255
      Left            =   60
      Max             =   15
      Min             =   1
      TabIndex        =   1
      Top             =   300
      Value           =   15
      Width           =   5955
   End
   Begin VB.PictureBox picDraw 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   4335
      Left            =   60
      MouseIcon       =   "frmDraw.frx":0000
      MousePointer    =   99  'Custom
      ScaleHeight     =   285
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   401
      TabIndex        =   0
      Top             =   600
      Width           =   6075
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   255
         Left            =   0
         TabIndex        =   3
         Top             =   0
         Width           =   735
      End
   End
   Begin VB.Label lblPen 
      Caption         =   "Line Width ( 1 to 15 ):"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   60
      Width           =   1815
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private R As Long, G As Long, B As Long

Private lColours(360) As Long

Private mCount As Long

Private mDown As Boolean

Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Const PS_SOLID = 0
Private Type POINTAPI
        x As Long
        y As Long
End Type

Private lpAPI As POINTAPI

Public Sub SetHLS(H As Integer, L As Integer, S As Integer)
    Dim MyR     As Single
    Dim MyG     As Single
    Dim MyB     As Single
    Dim MyH     As Single
    Dim MyL     As Single
    Dim MyS     As Single
    Dim Min     As Single
    Dim Max     As Single
    Dim Delta   As Single
    
    MyH = (H / 60) - 1: MyL = L / 100: MyS = S / 100
    If MyS = 0 Then
        MyR = MyL: MyG = MyL: MyB = MyL
    Else
        If MyL <= 0.5 Then
            Min = MyL * (1 - MyS)
        Else
            Min = MyL - MyS * (1 - MyL)
        End If
        Max = 2 * MyL - Min
        Delta = Max - Min
        
        Select Case MyH
        Case Is < 1
            MyR = Max
            If MyH < 0 Then
                MyG = Min
                MyB = MyG - MyH * Delta
            Else
                MyB = Min
                MyG = MyH * Delta + MyB
            End If
        Case Is < 3
            MyG = Max
            If MyH < 2 Then
                MyB = Min
                MyR = MyB - (MyH - 2) * Delta
            Else
                MyR = Min
                MyB = (MyH - 2) * Delta + MyR
            End If
        Case Else
            MyB = Max
            If MyH < 4 Then
                MyR = Min
                MyG = MyR - (MyH - 4) * Delta
            Else
                MyG = Min
                MyR = (MyH - 4) * Delta + MyG
            End If
        End Select
    End If
    
    R = MyR * 255: G = MyG * 255: B = MyB * 255
End Sub

Private Sub cmdClear_Click()
    picDraw.Cls
End Sub

Private Sub Form_Load()
    Dim iLoop As Integer
    
    For iLoop = 0 To 360
        SetHLS iLoop, 50, 100
        lColours(iLoop) = RGB(R, G, B)
    Next iLoop
End Sub

Private Sub Form_Resize()
    hBar.Move 0, hBar.Top, ScaleWidth
    picDraw.Move 0, hBar.Top + hBar.Height + 2, ScaleWidth, ScaleHeight - (hBar.Top + hBar.Height + 2)
End Sub

Private Sub picDraw_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mDown = True
    MoveToEx picDraw.hdc, x, y, lpAPI
End Sub

Private Sub picDraw_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim hPen As Long
Dim hOldPen As Long
If mDown = True Then
    hPen = CreatePen(PS_SOLID, hBar.Value, lColours(mCount))
    hOldPen = SelectObject(picDraw.hdc, hPen)
    LineTo picDraw.hdc, x, y
    hPen = SelectObject(picDraw.hdc, hOldPen)
    DeleteObject hPen
    picDraw.Refresh
    mCount = mCount + 1
    If mCount > 360 Then mCount = 0
End If
End Sub

Private Sub picDraw_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mDown = False
End Sub
