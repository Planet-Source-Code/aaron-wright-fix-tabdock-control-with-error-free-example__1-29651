VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDebug 
   BorderStyle     =   0  'None
   Caption         =   "Debug"
   ClientHeight    =   2310
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   4200
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin RichTextLib.RichTextBox txtDebug 
      Height          =   735
      Left            =   1080
      TabIndex        =   2
      Top             =   0
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   1296
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      TextRTF         =   $"frmDebug.frx":0000
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
      Height          =   255
      Left            =   350
      TabIndex        =   1
      Top             =   0
      Width           =   615
   End
   Begin VB.PictureBox picBar 
      BackColor       =   &H00FF0000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   0
      ScaleHeight     =   3195
      ScaleWidth      =   210
      TabIndex        =   0
      Top             =   0
      Width           =   275
      Begin VB.Image imgDebug 
         Height          =   645
         Left            =   -10
         Picture         =   "frmDebug.frx":00D6
         Top             =   10
         Width           =   225
      End
   End
End
Attribute VB_Name = "frmDebug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    Open App.Path & "\..\config\debug.dll" For Append As 1
    Print #1, txtDebug.Text & vbCrLf
    Close #1
    
End Sub

Private Sub Form_Load()
    Form_Resize
End Sub

Private Sub Form_Resize()
On Error GoTo 1
    picBar.Left = 5
    picBar.Height = Me.Height
    imgDebug.Top = 0
    
    cmdSave.Move 350!, 10!
    txtDebug.Move 1050!, 10!, Me.ScaleWidth - 15, Me.ScaleHeight
    
1
    Exit Sub
End Sub
