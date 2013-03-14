VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CD開閉プログラム"
   ClientHeight    =   1290
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   1290
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows の既定値
   Begin VB.Timer Timer1 
      Left            =   480
      Top             =   360
   End
   Begin VB.Label Label1 
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   480
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim flag As Boolean

Private Sub Form_Load()
    
    Timer1.Interval = 1500
    Timer1.Enabled = True
    
End Sub

Private Sub Timer1_Timer()
        
    flag = Not flag
    
    If flag Then
        
        Label1.Caption = "ひらく"
        Me.Refresh
        
        '開く
        mciSendStringA "Set CDAudio Door Open", 0&, 0, 0
    
    Else
        
        Label1.Caption = "とじる"
        Me.Refresh
    
        '閉じる
        mciSendStringA "Set CDAudio Door Closed", 0&, 0, 0
        
    End If
    
End Sub
