VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MenuAPIs"
   ClientHeight    =   3105
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   207
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00404080&
      Height          =   3135
      Left            =   0
      ScaleHeight     =   3075
      ScaleWidth      =   4635
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdOK 
         BackColor       =   &H00FFFFFF&
         Caption         =   "&OK"
         Default         =   -1  'True
         Height          =   375
         Left            =   1080
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2520
         UseMaskColor    =   -1  'True
         Width           =   2535
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmAbout.frx":000C
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   1215
         Left            =   120
         TabIndex        =   3
         Top             =   1320
         Width           =   4215
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "E-mail :  Jahanzaib2k4@hotmail.com                Jahanzaib2004@yahoo.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   3855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "By :  HAFIZ MUHAMMAD JAHANZAIB"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
Unload Me
End Sub
