VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "QZhi Othello"
   ClientHeight    =   3015
   ClientLeft      =   150
   ClientTop       =   795
   ClientWidth     =   4560
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picScore 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   3015
      TabIndex        =   7
      Top             =   120
      Width           =   3015
      Begin VB.Label lblGreenScore 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   2565
         TabIndex        =   9
         Top             =   0
         Width           =   90
      End
      Begin VB.Label lblRedScore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         Height          =   180
         Left            =   360
         TabIndex        =   8
         Top             =   0
         Width           =   90
      End
      Begin VB.Image imgGreen 
         Height          =   240
         Left            =   2760
         Picture         =   "frmMain.frx":058A
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Image imgRed 
         Height          =   240
         Left            =   0
         Picture         =   "frmMain.frx":0B14
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
   End
   Begin VB.PictureBox picImgData 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   3
      Left            =   3840
      Picture         =   "frmMain.frx":109E
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   6
      Top             =   720
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picCurrentPlayer 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   120
      ScaleHeight     =   255
      ScaleWidth      =   2895
      TabIndex        =   4
      Top             =   480
      Width           =   2895
      Begin VB.Image imgPlayer 
         Height          =   240
         Index           =   2
         Left            =   1920
         Picture         =   "frmMain.frx":17A0
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPlayer 
         Height          =   240
         Index           =   1
         Left            =   1560
         Picture         =   "frmMain.frx":1D2A
         Stretch         =   -1  'True
         Top             =   0
         Visible         =   0   'False
         Width           =   240
      End
      Begin VB.Image imgPlayer 
         Height          =   240
         Index           =   0
         Left            =   720
         Picture         =   "frmMain.frx":22B4
         Stretch         =   -1  'True
         Top             =   0
         Width           =   240
      End
      Begin VB.Label lblCurrentPlayer 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前方："
         Height          =   180
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   720
      End
   End
   Begin VB.PictureBox picImgData 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   2
      Left            =   3360
      Picture         =   "frmMain.frx":283E
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   3
      Top             =   720
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picImgData 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   1
      Left            =   3840
      Picture         =   "frmMain.frx":2F40
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picImgData 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   360
      Index           =   0
      Left            =   3360
      Picture         =   "frmMain.frx":3642
      ScaleHeight     =   360
      ScaleWidth      =   360
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   360
   End
   Begin VB.PictureBox picBoard 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   1935
      Left            =   120
      ScaleHeight     =   1935
      ScaleWidth      =   2895
      TabIndex        =   0
      Top             =   960
      Width           =   2895
   End
   Begin VB.Menu mnuGame 
      Caption         =   "游戏(&G)"
      Begin VB.Menu mnuGameNew 
         Caption         =   "开局(&N)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuGameBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameBeginner 
         Caption         =   "初级(&B)"
      End
      Begin VB.Menu mnuGameIntermediate 
         Caption         =   "中级(&I)"
      End
      Begin VB.Menu mnuGameExpert 
         Caption         =   "高级(&E)"
      End
      Begin VB.Menu mnuGameBar2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuGameExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpContent 
         Caption         =   "内容(&C)..."
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpBar1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于 QZhi Othello(&A)..."
      End
   End
End
Attribute VB_Name = "frmMain"
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

Private Declare Function ShellAboutA Lib "shell32.dll" (ByVal hwnd As Long, ByVal szApp As String, ByVal szOtherStuff As String, ByVal hIcon As Long) As Long

Private Declare Function midiOutOpen Lib "winmm.dll" (lphMidiOut As Long, ByVal uDeviceID As Long, ByVal dwCallback As Long, ByVal dwInstance As Long, ByVal dwFlags As Long) As Long
Private Declare Function midiOutClose Lib "winmm.dll" (ByVal hMidiOut As Long) As Long
Private Declare Function midiOutShortMsg Lib "winmm.dll" (ByVal hMidiOut As Long, ByVal dwMsg As Long) As Long

Private Const ZOOM_RATE = 2
Private Const BORDER_WIDTH = 120

Private lngBoardWidth As Long
Private lngBoardHeight As Long

Private lngPieceWidth As Long
Private lngPieceHeight As Long

Public Enum tagBlockStatus
    Blank = 0
    Red = 1
    Green = 2
End Enum

Private bsBoard() As tagBlockStatus
Private bsCurrentPlayer As tagBlockStatus

Public lngRedScore As Long
Public lngGreenScore As Long

Private lngSkipCount As Long

Private hMidiOut As Long

Private Sub Form_Initialize()
    ' 初始化 Midi 设备
    midiOutOpen hMidiOut, -1, 0, 0, 0
    If hMidiOut <> 0 Then midiOutShortMsg hMidiOut, 117 * &H100 + &HC0
End Sub

Private Sub Form_Load()
    
    lngPieceWidth = picImgData(0).ScaleWidth
    lngPieceHeight = picImgData(0).ScaleHeight

    lblCurrentPlayer.Top = (picCurrentPlayer.ScaleHeight - lblCurrentPlayer.Height) / 2 - 15

    mnuGameIntermediate_Click
End Sub

Private Sub Form_Terminate()
    If hMidiOut <> 0 Then midiOutClose hMidiOut
End Sub

Private Sub mnuGameBeginner_Click()

    If mnuGameBeginner.Checked = True Then Exit Sub

    mnuGameBeginner.Checked = Not mnuGameBeginner.Checked
    mnuGameIntermediate.Checked = Not mnuGameBeginner.Checked
    mnuGameExpert.Checked = Not mnuGameBeginner.Checked
    
    lngBoardWidth = 6
    lngBoardHeight = 6
    
    InitGame
End Sub

Private Sub mnuGameExit_Click()
    Unload Me
End Sub

Private Sub mnuGameExpert_Click()
    
    If mnuGameExpert.Checked = True Then Exit Sub

    mnuGameExpert.Checked = Not mnuGameExpert.Checked
    mnuGameBeginner.Checked = Not mnuGameExpert.Checked
    mnuGameIntermediate.Checked = Not mnuGameExpert.Checked
    
    lngBoardWidth = 12
    lngBoardHeight = 12
    
    InitGame
End Sub

Private Sub mnuGameIntermediate_Click()

    If mnuGameIntermediate.Checked = True Then Exit Sub

    mnuGameIntermediate.Checked = Not mnuGameIntermediate.Checked
    mnuGameBeginner.Checked = Not mnuGameIntermediate.Checked
    mnuGameExpert.Checked = Not mnuGameIntermediate.Checked
    
    lngBoardWidth = 8
    lngBoardHeight = 8
    
    InitGame
End Sub

Private Sub mnuGameNew_Click()
    InitGame
End Sub

Private Sub mnuHelpAbout_Click()
    ShellAboutA Me.hwnd, App.ProductName, "Version " & App.Major & "." & App.Minor & "." & App.Revision, Me.Icon
End Sub

Private Sub mnuHelpContent_Click()
    Dim f As New frmHelpContent
    
    f.Show vbModal, Me
End Sub

Private Sub picBoard_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim idxX As Long
    Dim idxY As Long
    
    If Button = 1 Then
        
        If (x < BORDER_WIDTH) Or (y < BORDER_WIDTH) Or (picBoard.ScaleWidth - x < BORDER_WIDTH) Or (picBoard.ScaleHeight - y < BORDER_WIDTH) Then Exit Sub
        
        idxX = CLng(x - BORDER_WIDTH) \ lngPieceWidth \ ZOOM_RATE
        idxY = CLng(y - BORDER_WIDTH) \ lngPieceHeight \ ZOOM_RATE

        If (idxX >= 0) And (idxX < lngBoardWidth) And (idxY >= 0) And (idxY < lngBoardHeight) Then
            If CanBePlaced(idxX, idxY) = True Then
            
                lngSkipCount = 0
            
                Place idxX, idxY
                If hMidiOut <> 0 Then midiOutShortMsg hMidiOut, &H90 + ((40 + 20) * &H100) + (127 * &H10000) + 0
                CalcScore
                UpdateScore
                
                NextPlayer
            End If
        End If
        
        UpdateBoard
        
    End If
End Sub

Private Sub picBoard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim idxX As Long
    Dim idxY As Long
    
    picBoard.MousePointer = vbNoDrop
    
    If (x < BORDER_WIDTH) Or (y < BORDER_WIDTH) Or (picBoard.ScaleWidth - x < BORDER_WIDTH) Or (picBoard.ScaleHeight - y < BORDER_WIDTH) Then Exit Sub
    
    idxX = CLng(x - BORDER_WIDTH) \ lngPieceWidth \ ZOOM_RATE
    idxY = CLng(y - BORDER_WIDTH) \ lngPieceHeight \ ZOOM_RATE

    If (idxX >= 0) And (idxX < lngBoardWidth) And (idxY >= 0) And (idxY < lngBoardHeight) Then
        If CanBePlaced(idxX, idxY) = True Then
            picBoard.MousePointer = vbDefault
        End If
    End If

End Sub


' ------------------------------------------------------------------------------------------------
' 游戏核心函数
' ------------------------------------------------------------------------------------------------

' 初始化游戏
Private Function InitGame()

    ReDim bsBoard(lngBoardWidth * lngBoardHeight - 1)
    lngSkipCount = 0
    
    picBoard.Cls

    picBoard.Move 120, _
        picCurrentPlayer.Top + picCurrentPlayer.Height + 120, _
        (picBoard.Width - picBoard.ScaleWidth) + lngBoardWidth * lngPieceWidth * ZOOM_RATE + BORDER_WIDTH * 2, _
        (picBoard.Height - picBoard.ScaleHeight) + lngBoardHeight * lngPieceHeight * ZOOM_RATE + BORDER_WIDTH * 2
    
    Me.Width = Me.Width - Me.ScaleWidth + picBoard.Width + 240
    Me.Height = Me.Height - Me.ScaleHeight + picBoard.Top + picBoard.Height + 120
    
    picScore.Width = Me.ScaleWidth - 240
    imgGreen.Left = picScore.ScaleWidth - imgGreen.Width
    
    lblRedScore.Move imgRed.Width + 60, (picScore.ScaleHeight - lblRedScore.Height) / 2 - 15
    lblGreenScore.Move picScore.ScaleWidth - imgGreen.Width - lblGreenScore.Width - 60, (picScore.ScaleHeight - lblGreenScore.Height) / 2 - 15
    
    picBoard.BackColor = picImgData(0).Point(0, 0)
    
    bsBoard(((lngBoardHeight / 2) - 1) * lngBoardWidth + (lngBoardWidth / 2) - 1) = Red
    bsBoard(((lngBoardHeight / 2) - 1) * lngBoardWidth + (lngBoardWidth / 2)) = Green
    bsBoard((lngBoardHeight / 2) * lngBoardWidth + (lngBoardWidth / 2)) = Red
    bsBoard((lngBoardHeight / 2) * lngBoardWidth + (lngBoardWidth / 2) - 1) = Green
    
    bsCurrentPlayer = Red
    Set imgPlayer(0).Picture = imgPlayer(bsCurrentPlayer).Picture
    
    UpdateBoard
    
    CalcScore
    UpdateScore
    
End Function

' 更新棋盘
Private Function UpdateBoard()

    Dim x As Long
    Dim y As Long
    
    For x = 0 To lngBoardWidth - 1
        For y = 0 To lngBoardHeight - 1
            picBoard.PaintPicture picImgData(GetBlock(x, y)).Picture, _
                x * lngPieceWidth * ZOOM_RATE + BORDER_WIDTH, _
                y * lngPieceHeight * ZOOM_RATE + BORDER_WIDTH, _
                lngPieceWidth * ZOOM_RATE, _
                lngPieceHeight * ZOOM_RATE
        Next y
    Next x
    
    For x = 0 To lngBoardWidth - 1
        For y = 0 To lngBoardHeight - 1
            If CanBePlaced(x, y) = True Then
                picBoard.PaintPicture picImgData(3).Picture, _
                    x * lngPieceWidth * ZOOM_RATE + BORDER_WIDTH, _
                    y * lngPieceHeight * ZOOM_RATE + BORDER_WIDTH, _
                    lngPieceWidth * ZOOM_RATE, _
                    lngPieceHeight * ZOOM_RATE
            End If
        Next y
    Next x
    
End Function

' 切换到下一个玩家
Private Function NextPlayer()
    If bsCurrentPlayer = Red Then
        bsCurrentPlayer = Green
    Else
        bsCurrentPlayer = Red
    End If
    
    Dim x As Long
    Dim y As Long
    Dim lngRightPositionCount As Long ' 可放置的棋子的坐标数量
    
    Set imgPlayer(0).Picture = imgPlayer(bsCurrentPlayer).Picture
    
    lngRightPositionCount = 0
    
    For x = 0 To lngBoardWidth - 1
        For y = 0 To lngBoardHeight - 1
            If CanBePlaced(x, y) = True Then
                lngRightPositionCount = lngRightPositionCount + 1
            End If
        Next y
    Next x
    
    If lngRightPositionCount = 0 Then
        If lngSkipCount = 0 Then
            lngSkipCount = 1
            
            If hMidiOut <> 0 Then midiOutShortMsg hMidiOut, 0 * &H100 + &HC0
            If hMidiOut <> 0 Then midiOutShortMsg hMidiOut, &H90 + (40 * &H100) + (127 * &H10000) + 0
            If hMidiOut <> 0 Then midiOutShortMsg hMidiOut, 117 * &H100 + &HC0
            
            NextPlayer
        Else
            EndGame
        End If
    End If
    
End Function

' 获取下一个玩家
Private Function GetNextPlayer() As tagBlockStatus
    If bsCurrentPlayer = Red Then
        GetNextPlayer = Green
    Else
        GetNextPlayer = Red
    End If
End Function

' 检查某个格子是否可以被放置
Private Function CanBePlaced(ByVal idxX As Long, ByVal idxY As Long) As Boolean

    Dim x As Long
    Dim y As Long
    
    Dim lngStepX(7) As Long ' x 的变化量
    Dim lngStepY(7) As Long ' y 的变化量
    
    Dim bsStatus As tagBlockStatus ' 状态变量
    
    Dim i As Long

    CanBePlaced = False

    ' 如果非空，则无法放置
    If GetBlock(idxX, idxY) <> Blank Then Exit Function
    
    lngStepX(0) = -1
    lngStepY(0) = -1
    
    lngStepX(1) = -1
    lngStepY(1) = 0
    
    lngStepX(2) = -1
    lngStepY(2) = 1
    
    lngStepX(3) = 0
    lngStepY(3) = -1
    
    lngStepX(4) = 0
    lngStepY(4) = 1
    
    lngStepX(5) = 1
    lngStepY(5) = -1
    
    lngStepX(6) = 1
    lngStepY(6) = 0
    
    lngStepX(7) = 1
    lngStepY(7) = 1
    
    For i = 0 To 7
    
        ' 状态变量初始化
        bsStatus = Blank
        
        ' 先步进一次，避免处理自身
        x = idxX + lngStepX(i)
        y = idxY + lngStepY(i)
        
        Do While (x >= 0) And (x < lngBoardWidth) And (y >= 0) And (y < lngBoardHeight)

            Select Case GetBlock(x, y)

                ' 如果是对方玩家的棋子，则设置状态为对方玩家
                Case GetNextPlayer
                    bsStatus = GetNextPlayer

                ' 如果是当前玩家的棋子
                Case bsCurrentPlayer
                    ' 且状态为对方玩家，即已经夹住了对方玩家的棋子，则退出函数
                    If bsStatus = GetNextPlayer Then
                        CanBePlaced = True
                        Exit Function
                    End If
                    ' 跳过本次循环
                    Exit Do

                ' 空格，跳过本次循环
                Case Blank
                    Exit Do

            End Select

            ' 再步进一次，更新索引
            x = x + lngStepX(i)
            y = y + lngStepY(i)

        Loop
        
    Next i

End Function

' 放置棋子
Private Function Place(ByVal idxX As Long, ByVal idxY As Long)

    Dim x As Long
    Dim y As Long
    
    Dim lngStepX(7) As Long ' x 的变化量
    Dim lngStepY(7) As Long ' y 的变化量
    
    Dim bsStatus As tagBlockStatus ' 状态变量
    Dim bsBuffer() As tagBlockStatus ' 棋盘缓冲区
    
    Dim i As Long

    ' 如果非空，则无法放置
    If GetBlock(idxX, idxY) <> Blank Then Exit Function
    
    lngStepX(0) = -1
    lngStepY(0) = -1
    
    lngStepX(1) = -1
    lngStepY(1) = 0
    
    lngStepX(2) = -1
    lngStepY(2) = 1
    
    lngStepX(3) = 0
    lngStepY(3) = -1
    
    lngStepX(4) = 0
    lngStepY(4) = 1
    
    lngStepX(5) = 1
    lngStepY(5) = -1
    
    lngStepX(6) = 1
    lngStepY(6) = 0
    
    lngStepX(7) = 1
    lngStepY(7) = 1
    
    ReDim bsBuffer(lngBoardWidth * lngBoardHeight - 1)
    
    For i = 0 To 7
    
        ' 状态变量初始化
        bsStatus = Blank
        
        ' 先步进一次，避免处理自身
        x = idxX + lngStepX(i)
        y = idxY + lngStepY(i)
        
        Do While (x >= 0) And (x < lngBoardWidth) And (y >= 0) And (y < lngBoardHeight)

            Select Case GetBlock(x, y)

                ' 如果是对方玩家的棋子，则设置状态为对方玩家
                Case GetNextPlayer
                    bsStatus = GetNextPlayer

                ' 如果是当前玩家的棋子
                Case bsCurrentPlayer
                    ' 且状态为对方玩家，即已经夹住了对方玩家的棋子，则设置状态为当前玩家
                    If bsStatus = GetNextPlayer Then
                        bsStatus = bsCurrentPlayer
                    End If
                    ' 跳过本次循环
                    Exit Do

                ' 空格，跳过本次循环
                Case Blank
                    Exit Do

            End Select

            ' 再步进一次，更新索引
            x = x + lngStepX(i)
            y = y + lngStepY(i)

        Loop
        
        ' 重置索引并步进一次，避免处理自身
        x = idxX + lngStepX(i)
        y = idxY + lngStepY(i)
        
        ' 状态为当前玩家，即该方向可以翻转
        If bsStatus = bsCurrentPlayer Then
            
            Do While (x >= 0) And (x < lngBoardWidth) And (y >= 0) And (y < lngBoardHeight)
    
                Select Case GetBlock(x, y)
    
                    ' 如果是对方玩家的棋子，则翻转
                    Case GetNextPlayer
                        bsBuffer(y * lngBoardWidth + x) = bsCurrentPlayer
                        bsStatus = GetNextPlayer
    
                    ' 如果是当前玩家的棋子，跳过本次循环
                    Case bsCurrentPlayer
                        Exit Do
    
                    ' 空格，跳过本次循环
                    Case Blank
                        Exit Do
    
                End Select
    
                ' 再步进一次，更新索引
                x = x + lngStepX(i)
                y = y + lngStepY(i)
    
            Loop
            
        End If
        
    Next i
    
    ' 更新棋盘
    For i = 0 To UBound(bsBuffer)
        If bsBuffer(i) <> Blank Then bsBoard(i) = bsBuffer(i)
    Next i
    
    bsBoard(idxY * lngBoardWidth + idxX) = bsCurrentPlayer

End Function

' 获取棋盘状态
Private Function GetBlock(ByVal idxX As Long, ByVal idxY As Long) As tagBlockStatus
    GetBlock = bsBoard(idxY * lngBoardWidth + idxX)
End Function

' 计算分数
Private Function CalcScore()
    Dim i As Long
    Dim lngRedScoreTemp As Long
    Dim lngGreenScoreTemp As Long
    
    lngRedScoreTemp = 0
    lngGreenScoreTemp = 0
    
    For i = 0 To UBound(bsBoard)
        If bsBoard(i) = Red Then
            lngRedScoreTemp = lngRedScoreTemp + 1
        ElseIf bsBoard(i) = Green Then
            lngGreenScoreTemp = lngGreenScoreTemp + 1
        End If
    Next i
    
    lngRedScore = lngRedScoreTemp
    lngGreenScore = lngGreenScoreTemp
    
End Function

' 更新分数
Private Function UpdateScore()
    lblRedScore.Caption = lngRedScore
    lblGreenScore.Caption = lngGreenScore
    
    If lngRedScore > lngGreenScore Then
        lblRedScore.Font.Bold = True
        lblGreenScore.Font.Bold = False
    ElseIf lngRedScore < lngGreenScore Then
        lblRedScore.Font.Bold = False
        lblGreenScore.Font.Bold = True
    Else
        lblRedScore.Font.Bold = False
        lblGreenScore.Font.Bold = False
    End If
    
    lblRedScore.Move imgRed.Width + 60, (picScore.ScaleHeight - lblRedScore.Height) / 2 - 15
    lblGreenScore.Move picScore.ScaleWidth - imgGreen.Width - lblGreenScore.Width - 60, (picScore.ScaleHeight - lblGreenScore.Height) / 2 - 15
End Function

' 结束游戏
Private Function EndGame()
    CalcScore
    
    UpdateBoard
    
    Dim f As New frmSettlementPage
    
    f.Show vbModal, Me
    
    Set f = Nothing
End Function
