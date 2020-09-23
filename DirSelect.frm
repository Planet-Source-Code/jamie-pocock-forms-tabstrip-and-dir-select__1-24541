VERSION 5.00
Begin VB.Form DirectorySort 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirectorySort"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8925
   FillStyle       =   0  'Solid
   ForeColor       =   &H00400000&
   Icon            =   "DirSelect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   8925
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox FileLocation 
      Alignment       =   2  'Center
      ForeColor       =   &H00000000&
      Height          =   285
      Left            =   60
      TabIndex        =   3
      Top             =   3450
      Width           =   8835
   End
   Begin VB.DirListBox Dir1 
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   45
      TabIndex        =   2
      ToolTipText     =   "Your computers directory."
      Top             =   405
      Width           =   4290
   End
   Begin VB.FileListBox File1 
      ForeColor       =   &H00000000&
      Height          =   3015
      Left            =   4365
      TabIndex        =   1
      ToolTipText     =   "Select files to be registerd or unregisterd."
      Top             =   405
      Width           =   4515
   End
   Begin VB.DriveListBox Drive1 
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   2265
   End
End
Attribute VB_Name = "DirectorySort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Dir1_Change()
    File1 = Dir1
    ChDir Dir1
End Sub

Private Sub Drive1_Change()
  On Error GoTo 10
    Dir1 = Drive1
    ChDrive Drive1
10: Exit Sub
End Sub

Private Sub File1_Click()
    FileLocation = File1.Path & "\" & File1
End Sub
