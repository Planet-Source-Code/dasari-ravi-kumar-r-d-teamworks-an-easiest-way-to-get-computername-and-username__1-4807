VERSION 5.00
Begin VB.Form UserName 
   Caption         =   "UserName....."
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5490
   LinkTopic       =   "Form2"
   ScaleHeight     =   3555
   ScaleWidth      =   5490
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   3135
      Left            =   180
      TabIndex        =   0
      Top             =   180
      Width           =   5085
      Begin VB.CommandButton Command2 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4320
         TabIndex        =   3
         Top             =   2745
         Width           =   705
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   1095
         TabIndex        =   2
         Text            =   " "
         Top             =   975
         Width           =   2415
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Get User Name"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   570
         Left            =   1095
         MouseIcon       =   "Username.frx":0000
         MousePointer    =   99  'Custom
         TabIndex        =   1
         Top             =   1785
         Width           =   2445
      End
   End
End
Attribute VB_Name = "UserName"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 

     Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                  (ByVal lpBuffer As String, nSize As Long) As Long
     Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
     Private Declare Function GetCurrentProcessId Lib "kernel32" () As Long

     Public Property Get UserName() As Variant
          Dim sBuffer As String
          Dim lSize As Long
          sBuffer = Space$(255)
          lSize = Len(sBuffer)
          Call GetUserName(sBuffer, lSize)
          UserName = Left$(sBuffer, lSize)
          
     End Property

Public Property Get ThreadID() As Variant
ThreadID = GetCurrentThreadId
End Property
 Public Property Get ProcessID() As Variant
 ProcessID = GetCurrentProcessId
 End Property



Private Sub Command1_Click()
Text1.Text = UserName
End Sub

Private Sub Command2_Click()
Computername.Show
Unload Me
End Sub

Private Sub Text1_Change()
Text1.Text = UCase(Text1.Text)
End Sub
