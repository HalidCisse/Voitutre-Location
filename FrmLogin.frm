VERSION 5.00
Begin VB.Form FrmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "Authentification"
   ClientHeight    =   3045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5790
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form2"
   Picture         =   "FrmLogin.frx":0442
   ScaleHeight     =   3045
   ScaleWidth      =   5790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox P2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      ScaleHeight     =   345
      ScaleWidth      =   5265
      TabIndex        =   9
      Top             =   3360
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.CommandButton CmdQuitter 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      MaskColor       =   &H00000000&
      Picture         =   "FrmLogin.frx":DCB3A
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "Quitter"
      Top             =   2400
      Width           =   1215
   End
   Begin VB.CommandButton CmdLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3840
      MaskColor       =   &H00000000&
      Picture         =   "FrmLogin.frx":DD4CC
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Connecter"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.PictureBox FrameLogin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   240
      Picture         =   "FrmLogin.frx":DE1FB
      ScaleHeight     =   1335
      ScaleWidth      =   5295
      TabIndex        =   1
      Top             =   840
      Width           =   5295
      Begin VB.TextBox TextEmail 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2040
         TabIndex        =   3
         ToolTipText     =   "Email ou Nom d'utilisateur"
         Top             =   120
         Width           =   3135
      End
      Begin VB.TextBox TextPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Code Secret"
         Top             =   720
         Width           =   3135
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mot de passe"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   1575
      End
   End
   Begin VB.PictureBox LoadBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      Picture         =   "FrmLogin.frx":1013D8
      ScaleHeight     =   495
      ScaleWidth      =   5295
      TabIndex        =   8
      Top             =   1320
      Width           =   5295
   End
   Begin VB.Timer TimerLoading 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1560
      Top             =   840
   End
   Begin VB.Timer TimerWait 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1440
      Top             =   1200
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   0
      Picture         =   "FrmLogin.frx":1245B5
      Stretch         =   -1  'True
      Top             =   -120
      Width           =   5775
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Chargement .."
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   2040
      TabIndex        =   10
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Gestion Location de Voitures"
      BeginProperty Font 
         Name            =   "@Arial Unicode MS"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404000&
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Turn As String
Private T As Integer
Private Pas As Integer

Private Sub CmdLogin_Click()
  If ConnectUser(Trim(TextEmail.Text), Trim(TextPass.Text)) Then
    TextEmail.Text = ""
    TextPass.Text = ""
    Loading (30)
    'LoadForms
    Load Main
    EffaceOldEvent (1)
  Else
    MsgBox "Identifiants Incorrect !!", vbInformation + vbOKOnly, "Gestion Location Voitures"
    TextPass.Text = ""
    TextPass.SetFocus
  End If
End Sub

Private Function Loading(Second As Integer)
  T = Second
  Pas = 100
  Turn = "va"
    LoadState (1)
End Function

Private Sub TimerWait_Timer()
    T = T - 1
    If T <= 0 Then
         Call LoadState(0)
         Main.Show
         Me.Hide
    End If
End Sub

Private Sub TimerLoading_Timer()
    Call AnimBar
End Sub

Function LoadState(i As Integer)
  If i = 1 Then
    'Pour Animation
    TimerLoading.Enabled = True
    TimerWait.Enabled = True
    FrameLogin.Visible = False
    CmdLogin.Visible = False
  Else
   'Pour Login
    CmdQuitter.Visible = True
    CmdLogin.Visible = True
    TimerLoading.Enabled = False
    TimerWait.Enabled = False
    FrameLogin.Visible = True
  End If
End Function

Private Function AnimBar()
'Anime chargement des données

  If Turn = "va" Then
    If LoadBar.Width <= P2.Width - Pas Then
       LoadBar.Width = LoadBar.Width + Pas
    Else
       Turn = "vient"
       LoadBar.Width = P2.Width
    End If
 Else
    If LoadBar.Width >= Pas Then
       LoadBar.Width = LoadBar.Width - Pas
    Else
       Turn = "va"
       LoadBar.Width = 0
    End If
 End If
LoadBar.Left = P2.Left + (P2.Width - LoadBar.Width) / 2
End Function

Private Sub CmdQuitter_Click()
   x = MsgBox("Voulez vous Quitter definitivement ?", vbQuestion + vbYesNo, "Quitter")
   If x = vbYes Then
      Unload Me
     End
   End If
End Sub

Private Sub Form_Activate()
   Call LoadState(0)
   FrameLogin.Visible = True
   TextEmail.SetFocus
   Call InitMe
End Sub

Private Sub Form_Load()
    Call ConnectDB
    'Call Design(Me)
    LoadBar.Width = 0
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
      UnloadForms
      Unload Main
       Beep
      Unload Me
End Sub


