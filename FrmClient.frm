VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmClient 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Information Client"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7740
   BeginProperty Font 
      Name            =   "MS Reference Sans Serif"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmClient.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmClient.frx":0442
   ScaleHeight     =   9015
   ScaleWidth      =   7740
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdPrec1 
      Caption         =   "< Prècedent"
      Height          =   405
      Left            =   240
      TabIndex        =   43
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton CmdSuiv1 
      Caption         =   "Suivant >"
      Height          =   405
      Left            =   5880
      TabIndex        =   42
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton CmdRaz1 
      Caption         =   "Effacer"
      Height          =   375
      Left            =   1920
      TabIndex        =   41
      Top             =   8520
      Width           =   1575
   End
   Begin VB.CommandButton CmdEnreg1 
      Caption         =   "Enrégistrer"
      Height          =   375
      Left            =   4200
      TabIndex        =   40
      Top             =   8520
      Width           =   1575
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Information Client"
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   8520
      TabIndex        =   2
      Top             =   5520
      Width           =   7215
      Begin MSComCtl2.DTPicker DTPickerEnreg 
         Height          =   375
         Left            =   3600
         TabIndex        =   44
         Top             =   1800
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81068033
         CurrentDate     =   41738
      End
      Begin VB.TextBox TextInfrac 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3600
         TabIndex        =   34
         Top             =   1320
         Width           =   3255
      End
      Begin VB.TextBox TextStatut 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3600
         TabIndex        =   33
         Top             =   840
         Width           =   3255
      End
      Begin VB.TextBox TextFid 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   3600
         TabIndex        =   32
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label17 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date d'Enregistrement"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   1920
         Width           =   2655
      End
      Begin VB.Label Label16 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre Total d'Infractions"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label Label15 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Statut"
         ForeColor       =   &H80000008&
         Height          =   495
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label14 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fidélité"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Identité Client"
      ForeColor       =   &H80000008&
      Height          =   8175
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         Caption         =   "Information Civil"
         ForeColor       =   &H80000008&
         Height          =   4935
         Left            =   240
         TabIndex        =   12
         Top             =   3000
         Width           =   6735
         Begin VB.TextBox TextAdress 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   3360
            TabIndex        =   31
            Top             =   4080
            Width           =   3135
         End
         Begin VB.TextBox TextEmail 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   3360
            TabIndex        =   30
            Top             =   3600
            Width           =   3135
         End
         Begin VB.TextBox TextNumTel 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   3360
            TabIndex        =   29
            Top             =   3120
            Width           =   3135
         End
         Begin VB.ComboBox ComboNat 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   4320
            TabIndex        =   28
            Top             =   2640
            Width           =   2175
         End
         Begin VB.ComboBox ComboLieuNaiss 
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   4320
            TabIndex        =   27
            Top             =   2160
            Width           =   2175
         End
         Begin MSComCtl2.DTPicker DTPickerNaiss 
            Height          =   375
            Left            =   4320
            TabIndex        =   26
            Top             =   1680
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            Format          =   81068033
            CurrentDate     =   41738
         End
         Begin VB.TextBox TextPrenom 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   3360
            TabIndex        =   25
            Top             =   1200
            Width           =   3135
         End
         Begin VB.TextBox TextNom 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   3360
            TabIndex        =   24
            Top             =   720
            Width           =   3135
         End
         Begin VB.OptionButton OptionF 
            Appearance      =   0  'Flat
            Caption         =   "Femme"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   5280
            TabIndex        =   23
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton OptionH 
            Appearance      =   0  'Flat
            Caption         =   "Homme"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3360
            TabIndex        =   22
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label13 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Adresse Domicile"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   21
            Top             =   4200
            Width           =   2055
         End
         Begin VB.Label Label12 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Email"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   20
            Top             =   3720
            Width           =   855
         End
         Begin VB.Label Label11 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Numéro de Téléphone"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   19
            Top             =   3240
            Width           =   2655
         End
         Begin VB.Label Label10 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nationalité"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   18
            Top             =   2760
            Width           =   1455
         End
         Begin VB.Label Label9 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Lieu de Naissance"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   2280
            Width           =   2295
         End
         Begin VB.Label Label8 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date de Naissance"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   16
            Top             =   1800
            Width           =   2175
         End
         Begin VB.Label Label7 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Prénom"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   15
            Top             =   1320
            Width           =   1095
         End
         Begin VB.Label Label6 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Nom"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   975
         End
         Begin VB.Label Label5 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Sex"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "ID"
         Height          =   2055
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   6735
         Begin MSComCtl2.DTPicker DTPickerExpID 
            Height          =   375
            Left            =   4320
            TabIndex        =   39
            Top             =   1440
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            Format          =   81068033
            CurrentDate     =   41738
         End
         Begin VB.OptionButton OptionCS 
            Appearance      =   0  'Flat
            Caption         =   "Carte Séjour"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   4800
            TabIndex        =   38
            Top             =   480
            Width           =   1815
         End
         Begin VB.OptionButton OptionPass 
            Appearance      =   0  'Flat
            Caption         =   "Passport"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   3360
            TabIndex        =   37
            Top             =   480
            Width           =   1335
         End
         Begin VB.OptionButton OptionCIN 
            Appearance      =   0  'Flat
            Caption         =   "CIN"
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2400
            TabIndex        =   36
            Top             =   480
            Width           =   735
         End
         Begin VB.TextBox TextNumID 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            Height          =   405
            Left            =   3360
            TabIndex        =   35
            Top             =   840
            Width           =   3135
         End
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date Expiration"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   11
            Top             =   1560
            Width           =   1815
         End
         Begin VB.Label Label3 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Numéro de L'identité"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   10
            Top             =   960
            Width           =   2415
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Type d'identité"
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   240
            TabIndex        =   9
            Top             =   480
            Width           =   1935
         End
      End
      Begin MSComCtl2.DTPicker DTPickerPermis 
         Height          =   375
         Left            =   5040
         TabIndex        =   7
         Top             =   360
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   661
         _Version        =   393216
         Format          =   81068033
         CurrentDate     =   41738
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date d'Optention du Permis"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3135
      End
   End
End
Attribute VB_Name = "FrmClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdEnreg1_Click()

If ChampsAddClientOK Then
        
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM CLIENTS WHERE NumID='" & UCase(FrmClient.TextNumID.Text) & "' "
    rs.Open SQL, CN, adOpenKeyset
'----------------------------------------------------
    If rs.EOF Then
       x = MsgBox("Efféctuer l'enregistrement ?", vbYesNo + vbQuestion, "Gestion Location de Voitures")
       If x = vbYes Then
          Call AjouterClient
          MsgBox "Enregistrement Effectué Avec Succées !!", vbInformation, "Gestion Location de Voitures "
          Me.Hide
       End If
    Else
       x = MsgBox("Efféctuer la modification ?", vbYesNo + vbQuestion, "Gestion Location de Voiture ")
       If x = vbYes Then
          Call ModifierClient(UCase(TextNumID.Text))
          MsgBox "Modification Effectuée Avec Succées !!", vbInformation, "Gestion Location de Voitures"
          Me.Hide
       End If
    End If
'-------------------------------------------------------
End If
rs.Close
Set rs = Nothing
End Sub

Private Sub CmdPrec1_Click()
 ShowPageClient (PageInf - 1)
End Sub

Private Sub CmdSuiv1_Click()
 ShowPageClient (PageInf + 1)
End Sub

Private Sub ComboLieuNaiss_Change()
    Call AutoComplete(ComboLieuNaiss)
End Sub

Private Sub ComboNat_Change()
   Call AutoComplete(ComboNat)
End Sub

Private Sub Form_Activate()
   Call ScaleClient
End Sub

Private Sub Form_Load()
  Call Design(Me)
  Call RemplirLesCombosClient
End Sub
