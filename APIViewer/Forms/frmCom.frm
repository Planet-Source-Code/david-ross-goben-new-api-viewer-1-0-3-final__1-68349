VERSION 5.00
Begin VB.Form frmCom 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Common Control Store"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3855
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   3855
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstEnum 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   2880
      Sorted          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.TextBox txtSelectedItems 
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1635
      HideSelection   =   0   'False
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   1920
      Width           =   3615
   End
   Begin VB.ListBox lstType 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   1950
      Sorted          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstDecl 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   1035
      Sorted          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.ListBox lstConst 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1620
      Left            =   120
      Sorted          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   180
      Visible         =   0   'False
      Width           =   855
   End
End
Attribute VB_Name = "frmCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'-------------------------------------------------------------------------------------------------
'This form store common data that would normally be stored in/on the frmAPIViewer form. Though
'placing these items and controls on frmAPIViewer will work in the stand-alone mode, as an
'add-in, the frmAPIViewer form is simply a class definition, and an instance of it is instantiated
'in the connect object as mfrmAddIn. Since mfrmAddIn cannot be referenced from the other forms,
'an intermediate form was required. Although one of the other existing forms could have been used
'for this purpose, their functionality did not bode well for persistent existence, as would be
'required, and so a Communication form, frmCom, is used. It is never displayed, only loaded.
'-------------------------------------------------------------------------------------------------

