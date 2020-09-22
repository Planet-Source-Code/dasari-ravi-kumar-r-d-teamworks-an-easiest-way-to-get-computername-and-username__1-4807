VERSION 5.00
Begin VB.Form Computername 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ComputerName"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5970
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Computer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3720
      Left            =   210
      TabIndex        =   0
      Top             =   165
      Width           =   5475
      Begin VB.CommandButton Command2 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4440
         TabIndex        =   3
         Top             =   3345
         Width           =   780
      End
      Begin VB.TextBox Text1 
         Height          =   360
         Left            =   1515
         TabIndex        =   2
         Text            =   " "
         Top             =   1275
         Width           =   1875
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get Computer Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   645
         Left            =   1575
         MouseIcon       =   "computername.frx":0000
         MousePointer    =   99  'Custom
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   2025
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Computername"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal sBuffer As String, lSize As Long) As Long
 
 
Private Sub Command1_Click()
Dim PCName As String
Dim P As Long
 P = NameOfPC(PCName)
 Text1.Text = PCName
 End Sub
Public Function NameOfPC(MachineName As String) As Long
    Dim NameSize As Long
    Dim X As Long
    MachineName = Space$(16)
    NameSize = Len(MachineName)
    X = GetComputerName(MachineName, NameSize)
End Function

Private Sub Command2_Click()
UserName.Show
Unload Me
End Sub
