VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin Project1.TxtBoxBorder TxtBoxBorder3 
      Height          =   2595
      Left            =   2655
      TabIndex        =   2
      Top             =   300
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   4577
      Text            =   "Color changes"
      FontSize        =   8.25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.TxtBoxBorder TxtBoxBorder2 
      Height          =   870
      Left            =   165
      TabIndex        =   1
      Top             =   2025
      Width           =   2115
      _ExtentX        =   3731
      _ExtentY        =   1535
      NonFocusColor   =   16711680
      Text            =   "No Color Change on Focus"
      FontSize        =   8.25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.TxtBoxBorder TxtBoxBorder1 
      Height          =   1440
      Left            =   195
      TabIndex        =   0
      Top             =   285
      Width           =   2100
      _ExtentX        =   3704
      _ExtentY        =   2540
      Text            =   "Changes border color on focus"
      FontSize        =   8.25
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

