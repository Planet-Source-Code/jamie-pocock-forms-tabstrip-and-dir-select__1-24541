VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form TabForm 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   7170
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11385
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   11385
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3195
      Index           =   5
      Left            =   7800
      ScaleHeight     =   3135
      ScaleWidth      =   3390
      TabIndex        =   6
      Top             =   3870
      Width           =   3450
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00800000&
      Height          =   3750
      Index           =   4
      Left            =   7200
      ScaleHeight     =   3690
      ScaleWidth      =   3990
      TabIndex        =   5
      Top             =   3330
      Width           =   4050
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C00000&
      Height          =   4290
      Index           =   3
      Left            =   6540
      ScaleHeight     =   4230
      ScaleWidth      =   4650
      TabIndex        =   4
      Top             =   2790
      Width           =   4710
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF0000&
      Height          =   4770
      Index           =   2
      Left            =   5970
      ScaleHeight     =   4710
      ScaleWidth      =   5220
      TabIndex        =   3
      Top             =   2310
      Width           =   5280
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FF8080&
      Height          =   5190
      Index           =   1
      Left            =   5400
      ScaleHeight     =   5130
      ScaleWidth      =   5790
      TabIndex        =   2
      Top             =   1860
      Width           =   5850
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFC0C0&
      Height          =   5610
      Index           =   0
      Left            =   4905
      ScaleHeight     =   5550
      ScaleWidth      =   6300
      TabIndex        =   1
      Top             =   1470
      Width           =   6360
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   7125
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   11310
      _ExtentX        =   19950
      _ExtentY        =   12568
      MultiRow        =   -1  'True
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   5
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 1"
            Key             =   "picture1p"
            Object.Tag             =   "1"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 2"
            Key             =   "picture2p"
            Object.Tag             =   "2"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 3"
            Key             =   "picture3p"
            Object.Tag             =   "3"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 4"
            Key             =   "picture4p"
            Object.Tag             =   "4"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Page 5"
            Key             =   "picture5p"
            Object.Tag             =   "5"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "TabForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Picture1p()
End Sub
Private Sub Picture2p()
End Sub
Private Sub Picture3p()
End Sub
Private Sub Picture4p()
End Sub
Private Sub Picture5p()
End Sub

Private Sub Form_Load()
TabStrip1.Tabs("picture1p").Selected = True
End Sub

Private Sub TabStrip1_Click()
On Error GoTo 10
Picture1(TabStrip1.SelectedItem.Index - 1).Move TabStrip1.ClientLeft _
, TabStrip1.ClientTop, TabStrip1.ClientWidth, TabStrip1.ClientHeight

Picture1(TabStrip1.SelectedItem.Index - 1).ZOrder
Select Case TabStrip1.SelectedItem.Index

    Case 1
        Picture1p
    Case 2
        Picture2p
    Case 3
        Picture3p
    Case 4
        Picture4p
    Case 5
        Picture5p
End Select
10: Exit Sub
End Sub
