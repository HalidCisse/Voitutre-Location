VERSION 5.00
Begin VB.Form FrmUser 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Information Utilisateur"
   ClientHeight    =   4845
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7695
   BeginProperty Font 
      Name            =   "MS Reference Sans Serif"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmUser.frx":0000
   ScaleHeight     =   4845
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdRaz 
      Appearance      =   0  'Flat
      Caption         =   "Fermé"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      ToolTipText     =   "Fermé"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.CommandButton CmdEnreg 
      Appearance      =   0  'Flat
      Caption         =   "Enregistré"
      Height          =   375
      Left            =   6000
      TabIndex        =   12
      ToolTipText     =   "Ajouté"
      Top             =   4320
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.ComboBox ComboType 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         ItemData        =   "FrmUser.frx":75D26
         Left            =   4680
         List            =   "FrmUser.frx":75D30
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   120
         Width           =   2415
      End
      Begin VB.TextBox TextNom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3720
         TabIndex        =   10
         Top             =   600
         Width           =   3375
      End
      Begin VB.TextBox TextComfirmPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3720
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   3000
         Width           =   3375
      End
      Begin VB.TextBox TextPass 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   3720
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   2400
         Width           =   3375
      End
      Begin VB.TextBox TextEmail 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   405
         Left            =   3720
         TabIndex        =   7
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox TextPrenom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3720
         TabIndex        =   6
         Top             =   1200
         Width           =   3375
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prénom"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nom"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Type Utilisateur"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   120
         Width           =   2175
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Confirmé Mot de Pass"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   3120
         Width           =   2655
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Mot de Passe"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   2520
         Width           =   1695
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Adresse Email"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   1920
         Width           =   1815
      End
   End
End
Attribute VB_Name = "FrmUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdEnreg_Click()
  
  If ChampsAddUserOk Then
'------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM PROFILES WHERE Email='" & Trim(TextEmail.Text) & "' "
    rs.Open SQL, CN, adOpenKeyset
'------------------------------------------------------------------
    If rs.EOF Then
         x = MsgBox("Efféctuer l'enregistrement ?", vbYesNoCancel + vbQuestion, "Gestion Location de Voitures : Enregistrement")
          If x = vbYes Then
             Call AjouterUser
             MsgBox "Enregistrement Effectué Avec Succées !!", vbInformation, "Gestion Location de Voitures : Voitures - ENREGISTREMENT"
             Me.Hide
          ElseIf x = vbCancel Then
              Me.Visible = False
          End If
    Else
     If Trim(TextEmail.Text) = UserEmail Then
        x = MsgBox("Efféctuer la modification ?", vbYesNoCancel + vbQuestion, "Gestion Location de Voiture : Modification")
         If x = vbYes Then
            Call ModifierUser(Trim(TextEmail.Text))
            MsgBox "Modification Effectuée Avec Succées !!", vbInformation, "Gestion Location de Voitures : Voiture  - Modification"
            Me.Hide
         ElseIf x = vbCancel Then
              Me.Visible = False
         End If
     Else
        MsgBox "Ce Utilisateur Existe Déja !!", vbCritical + vbOKOnly, "Gestion Location de Voiture"
     End If
   End If
'------------------------------------------------------------------
End If

End Sub

Private Sub CmdRaz_Click()
   Me.Hide
End Sub

Private Sub Form_Load()
    Call Design(Me)
End Sub
