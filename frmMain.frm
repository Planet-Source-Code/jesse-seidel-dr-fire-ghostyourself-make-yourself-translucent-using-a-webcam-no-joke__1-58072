VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   Caption         =   "GhostYourself - By Jesse Seidel"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5970
   ForeColor       =   &H8000000F&
   HasDC           =   0   'False
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   313
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   398
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   5235
      ScaleHeight     =   3840
      ScaleWidth      =   735
      TabIndex        =   8
      Top             =   0
      Width           =   735
      Begin ComctlLib.Slider barAmount 
         Height          =   2235
         Left            =   0
         TabIndex        =   9
         Top             =   840
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   3942
         _Version        =   327682
         Orientation     =   1
         LargeChange     =   51
         Max             =   255
         SelStart        =   180
         TickStyle       =   2
         TickFrequency   =   51
         Value           =   180
      End
      Begin VB.Label Label1 
         Caption         =   "180"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3240
         Width           =   615
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   5970
      TabIndex        =   3
      Top             =   3840
      Width           =   5970
      Begin VB.CommandButton Command1 
         Caption         =   "Start"
         Height          =   495
         Left            =   240
         TabIndex        =   7
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Stop"
         Enabled         =   0   'False
         Height          =   495
         Left            =   1560
         TabIndex        =   6
         Top             =   120
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Save (C:\temp.bmp)"
         Height          =   495
         Left            =   3000
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Set BG"
         Enabled         =   0   'False
         Height          =   495
         Left            =   4560
         TabIndex        =   4
         Top             =   120
         Width           =   1095
      End
      Begin VB.Timer tmrMain 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   3360
         Top             =   120
      End
   End
   Begin VB.PictureBox picDestination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   120
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   323
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   120
      Width           =   4845
   End
   Begin VB.PictureBox picSource 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   3600
      Left            =   120
      Picture         =   "frmMain.frx":000C
      ScaleHeight     =   240
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Visible         =   0   'False
      Width           =   4800
      Begin VB.CommandButton Command5 
         Caption         =   "Command5"
         Height          =   255
         Left            =   4200
         TabIndex        =   2
         Top             =   3600
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function AlphaBlend Lib "msimg32" ( _
ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, _
ByVal xSrc As Long, ByVal ySrc As Long, ByVal widthSrc As Long, _
ByVal heightSrc As Long, ByVal blendFunct As Long) As Boolean

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" ( _
Destination As Any, Source As Any, ByVal Length As Long)

Private Type typeBlendProperties
    tBlendOp As Byte
    tBlendOptions As Byte
    tBlendAmount As Byte
    tAlphaType As Byte
End Type

Private Sub barAmount_Scroll()
    Dim tProperties As typeBlendProperties
    Dim lngBlend As Long
    picDestination.Cls
    tProperties.tBlendAmount = 255 - barAmount
    CopyMemory lngBlend, tProperties, 4
    AlphaBlend picDestination.hDC, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, _
    picSource.hDC, 0, 0, picSource.ScaleWidth, picSource.ScaleHeight, lngBlend
    picDestination.Refresh
    Label1.Caption = barAmount.Value
End Sub

Private Sub Command1_Click()
mCapHwnd = capCreateCaptureWindow("WebcamCapture", 0, 0, 0, 320, 240, Me.hwnd, 0)
DoEvents: SendMessage mCapHwnd, CONNECT, 0, 0
tmrMain.Enabled = True
Command2.Enabled = True
Command1.Enabled = False
Command4.Enabled = True
End Sub

Private Sub Command2_Click()
tmrMain.Enabled = False
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
Command1.Enabled = True
Command2.Enabled = False
Command4.Enabled = False
End Sub

Private Sub Command3_Click()
SavePicture picDestination.Image, "C:\temp.bmp"
End Sub

Private Sub Command4_Click()
picSource.Picture = picDestination.Image
End Sub

Private Sub Form_Load()
barAmount_Scroll
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
DoEvents: SendMessage mCapHwnd, DISCONNECT, 0, 0
End Sub

Private Sub Timer1_Timer()
SavePicture picDestination.Image, "C:\temp.bmp"
End Sub

Private Sub tmrMain_Timer()
On Error Resume Next
SendMessage mCapHwnd, GET_FRAME, 0, 0
SendMessage mCapHwnd, COPY, 0, 0
picDestination.Picture = Clipboard.GetData
Clipboard.Clear
barAmount_Scroll
End Sub
