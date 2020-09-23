VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sinus"
   ClientHeight    =   3210
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6780
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   214
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   452
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3255
      Left            =   0
      TabIndex        =   2
      Top             =   -45
      Width           =   6780
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   3840
         Top             =   2400
      End
      Begin VB.PictureBox PicDisplay 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   2085
         Left            =   120
         ScaleHeight     =   135
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   432
         TabIndex        =   13
         Top             =   240
         Width           =   6540
      End
      Begin VB.CommandButton CmdStart 
         Caption         =   "Start"
         Height          =   375
         Left            =   5280
         TabIndex        =   8
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TxtLenght 
         Height          =   285
         Left            =   1200
         TabIndex        =   7
         Text            =   "30"
         Top             =   2505
         Width           =   375
      End
      Begin VB.TextBox TxtHeight 
         Height          =   285
         Left            =   1200
         TabIndex        =   6
         Text            =   "20"
         Top             =   2865
         Width           =   375
      End
      Begin VB.CommandButton CmdApply 
         Caption         =   "Apply changes"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3600
         TabIndex        =   5
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox TxtSpeed 
         Height          =   285
         Left            =   2760
         TabIndex        =   4
         Text            =   "3"
         Top             =   2505
         Width           =   375
      End
      Begin VB.TextBox TxtScrollSpeed 
         Height          =   285
         Left            =   2760
         TabIndex        =   3
         Text            =   "1"
         Top             =   2865
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Wavelenght"
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Waveheight"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label3 
         Caption         =   "Wavespeed"
         Height          =   255
         Left            =   1680
         TabIndex        =   10
         Top             =   2520
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Scrollspeed"
         Height          =   255
         Left            =   1680
         TabIndex        =   9
         Top             =   2880
         Width           =   975
      End
   End
   Begin VB.PictureBox PicBuffer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2085
      Left            =   480
      ScaleHeight     =   139
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   436
      TabIndex        =   1
      Top             =   4440
      Width           =   6540
      Visible         =   0   'False
   End
   Begin VB.PictureBox PicSrc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1125
      Left            =   1800
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   75
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1000
      TabIndex        =   0
      Top             =   3840
      Width           =   15000
      Visible         =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Dim t, pos, w_lenght, w_height, w_speed, oldval, fps, p As Long
Dim scrollspeed As Double
Dim Cancel As Boolean

Private Sub sinus()
Timer1.Enabled = True
Do While Cancel = False
DoEvents
t = t + w_speed '+ 1
pos = pos + scrollspeed
p = t
For x = 0 To PicBuffer.ScaleWidth
    p = p + 1
    y = PicSrc.ScaleHeight / 2 + Sin((p) / w_lenght) * w_height
    Call StretchBlt(PicBuffer.hdc, x, y, 1, PicSrc.ScaleHeight, PicSrc.hdc, x + pos, 1, 1, PicSrc.ScaleHeight, vbSrcCopy)
Next x
Call BitBlt(PicDisplay.hdc, 0, 0, PicBuffer.ScaleWidth, PicBuffer.ScaleHeight, PicBuffer.hdc, 0, 0, vbSrcCopy)
PicBuffer.Cls
'PicDisplay.Refresh
If pos > PicSrc.ScaleWidth + 100 Then pos = -PicBuffer.ScaleWidth
Loop
End Sub

Private Sub CmdApply_Click()
w_lenght = TxtLenght
w_height = TxtHeight
w_speed = TxtSpeed
scrollspeed = TxtScrollSpeed
End Sub

Private Sub CmdStart_Click()
If CmdStart.Caption = "Start" Then
    Cancel = False
    CmdStart.Caption = "Stop"
    Call sinus
Else
    CmdStart.Caption = "Start"
    Cancel = True
End If
End Sub

Private Sub Form_Load()
w_lenght = 30
w_height = 20
w_speed = 3
scrollspeed = 1
pos = -PicBuffer.ScaleWidth
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
Unload Me
End
End Sub

Private Sub Timer1_Timer()
fps = p - oldval
oldval = p
Me.Caption = "Sinus - " & fps & " fps"
End Sub

Private Sub TxtHeight_Change()
CmdApply.Enabled = True
End Sub

Private Sub TxtLenght_Change()
CmdApply.Enabled = True
End Sub

Private Sub TxtScrollSpeed_Change()
CmdApply.Enabled = True
End Sub

Private Sub TxtSpeed_Change()
CmdApply.Enabled = True
End Sub
