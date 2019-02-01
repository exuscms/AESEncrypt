VERSION 5.00
Begin VB.Form FrmZip 
   BorderStyle     =   1  '단일 고정
   ClientHeight    =   2175
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9255
   Icon            =   "FrmZip.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2175
   ScaleWidth      =   9255
   StartUpPosition =   2  '화면 가운데
   Begin Project1.Xp_ProgressBar PBFile2 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
      ProgressLook    =   2
   End
   Begin Project1.Xp_ProgressBar PBFile 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   661
   End
   Begin VB.Label LabPer 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9015
   End
   Begin VB.Label LabBuffer 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   9015
   End
   Begin VB.Label LabName 
      Appearance      =   0  '평면
      BackColor       =   &H80000005&
      BackStyle       =   0  '투명
      BeginProperty Font 
         Name            =   "돋움"
         Size            =   8.25
         Charset         =   129
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   9015
   End
End
Attribute VB_Name = "FrmZip"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
