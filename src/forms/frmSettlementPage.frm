VERSION 5.00
Begin VB.Form frmSettlementPage 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4560
   Icon            =   "frmSettlementPage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer tmrTimer 
      Left            =   4080
      Top             =   2520
   End
   Begin VB.PictureBox picPanel 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   120
      ScaleHeight     =   2415
      ScaleWidth      =   4335
      TabIndex        =   0
      Top             =   480
      Width           =   4335
   End
   Begin VB.Label lblWinner 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "赢家是……"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmSettlementPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Copyright 2025 QZhi Studio
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
'     http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.

Option Explicit

Dim lngFrame As Long

Private Sub Form_Initialize()
    If frmMain.lngRedScore > frmMain.lngGreenScore Then
        lblWinner.Caption = "赢家是红方！"
    ElseIf frmMain.lngRedScore < frmMain.lngGreenScore Then
        lblWinner.Caption = "赢家是绿方！"
    Else
        lblWinner.Caption = "平局！"
    End If
    
    Me.Caption = lblWinner.Caption
    
    tmrTimer.Interval = 25
End Sub

Private Sub tmrTimer_Timer()
    
    Dim pic As IPictureDisp
    
    If frmMain.lngRedScore > frmMain.lngGreenScore Then
        Set pic = frmMain.imgPlayer(1).Picture
    ElseIf frmMain.lngRedScore < frmMain.lngGreenScore Then
        Set pic = frmMain.imgPlayer(2).Picture
    Else
        Set pic = frmMain.imgPlayer(lngFrame Mod 2 + 1).Picture
    End If
    
    With pic
        ' 采用 .Render，以绘制透明图像
        .Render picPanel.hDC, _
            (lngFrame Mod (picPanel.ScaleWidth \ (4 * Screen.TwipsPerPixelX))) * 4& + (lngFrame \ (picPanel.ScaleWidth \ (4 * Screen.TwipsPerPixelX))) * 32, _
            Abs(Sin((lngFrame Mod (picPanel.ScaleWidth \ (4 * Screen.TwipsPerPixelY))) * 3.14159265358979 / 180 * 4)) * picPanel.ScaleHeight / Screen.TwipsPerPixelY, _
            ScaleX(.Width, vbHimetric, vbPixels), ScaleY(.Height, vbHimetric, vbPixels), _
            0&, .Height, .Width, -.Height, ByVal 0&
    End With
    
    picPanel.Refresh
    
    lngFrame = lngFrame + 1
    
    DoEvents
    
    If (lngFrame \ (picPanel.ScaleWidth \ (4 * Screen.TwipsPerPixelX))) * 32 * Screen.TwipsPerPixelX > picPanel.ScaleWidth Then tmrTimer.Interval = 0

End Sub
