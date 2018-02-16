VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Input System"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8055
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   8055
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblWheel 
      Caption         =   "Mouse Wheel:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   4695
   End
   Begin VB.Label lblMiddle 
      Caption         =   "Mouse Middle:"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   4695
   End
   Begin VB.Label lblRight 
      Caption         =   "Mouse Right:"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label lblLeft 
      Caption         =   "Mouse Left:"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   4695
   End
   Begin VB.Label lblCoord 
      Caption         =   "Mouse Coord: "
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   4815
   End
   Begin VB.Label lblChar 
      Caption         =   "Char: "
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   4695
   End
   Begin VB.Label lblRaw 
      Caption         =   "Raw: "
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Unload(Cancel As Integer)
    Terminate
End Sub
