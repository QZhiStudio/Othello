VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form frmHelpContent 
   Caption         =   "QZhi Othello 帮助"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360
   Icon            =   "frmHelpContent.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6615
   ScaleWidth      =   9360
   StartUpPosition =   1  '所有者中心
   Begin SHDocVwCtl.WebBrowser brwHelp 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9135
      ExtentX         =   16113
      ExtentY         =   11245
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "frmHelpContent"
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

Private blnIsInitialized As Boolean

Private Sub brwHelp_TitleChange(ByVal Text As String)
    Me.Caption = Text
End Sub

Private Sub Form_Load()
    brwHelp.Navigate "res://" & brwHelp.FullName & "/help.html"
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    brwHelp.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight
End Sub
