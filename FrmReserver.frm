VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmReserver 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Nouvelle Reservation"
   ClientHeight    =   7230
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
   Icon            =   "FrmReserver.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "FrmReserver.frx":014A
   ScaleHeight     =   7230
   ScaleWidth      =   7695
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerRes 
      Interval        =   1000
      Left            =   3480
      Top             =   7200
   End
   Begin VB.CommandButton CmdEnreg 
      Appearance      =   0  'Flat
      Caption         =   "Enrégistré"
      Height          =   375
      Left            =   5880
      TabIndex        =   20
      Top             =   6720
      Width           =   1575
   End
   Begin VB.CommandButton CmdRaz 
      Appearance      =   0  'Flat
      Caption         =   "Effacé"
      Height          =   405
      Left            =   240
      TabIndex        =   19
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   6255
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   7215
      Begin VB.TextBox TextNumJour 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4080
         TabIndex        =   25
         Top             =   2640
         Width           =   2895
      End
      Begin VB.Frame Frame2 
         Caption         =   "Statut du Vehicule"
         Height          =   1455
         Left            =   240
         TabIndex        =   16
         Top             =   4560
         Width           =   6735
         Begin VB.OptionButton OptionAuGar 
            Caption         =   "Au Garage"
            Height          =   285
            Left            =   240
            TabIndex        =   23
            Top             =   480
            Width           =   1575
         End
         Begin MSComCtl2.DTPicker DTPickerVoitRendu 
            Height          =   375
            Left            =   3960
            TabIndex        =   22
            Top             =   840
            Visible         =   0   'False
            Width           =   2175
            _ExtentX        =   3836
            _ExtentY        =   661
            _Version        =   393216
            Format          =   88604673
            CurrentDate     =   41744
         End
         Begin VB.OptionButton OptionRendu 
            Caption         =   "Rendue"
            Height          =   285
            Left            =   5040
            TabIndex        =   18
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton OptionAvecClient 
            Caption         =   "Avec le Client"
            Height          =   285
            Left            =   2520
            TabIndex        =   17
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label LabelDateRend 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date de Rendition"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   360
            TabIndex        =   24
            Top             =   960
            Visible         =   0   'False
            Width           =   3375
         End
      End
      Begin VB.TextBox TextPay 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4080
         TabIndex        =   15
         Top             =   3600
         Width           =   2895
      End
      Begin VB.TextBox TextPayTotal 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4080
         TabIndex        =   14
         Top             =   3120
         Width           =   2895
      End
      Begin MSComCtl2.DTPicker DTPickerFinExp 
         Height          =   375
         Left            =   4560
         TabIndex        =   13
         Top             =   2160
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   88604673
         CurrentDate     =   41744
      End
      Begin MSComCtl2.DTPicker DTPickerDebExp 
         Height          =   375
         Left            =   4560
         TabIndex        =   12
         Top             =   1680
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   661
         _Version        =   393216
         Format          =   88604673
         CurrentDate     =   41744
      End
      Begin VB.ComboBox ComboMat 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4560
         TabIndex        =   11
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox ComboNumID 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4560
         TabIndex        =   10
         Top             =   720
         Width           =   2415
      End
      Begin VB.TextBox TextResID 
         Appearance      =   0  'Flat
         Height          =   405
         Left            =   4080
         TabIndex        =   9
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label9 
         Caption         =   "Nombre de Jours"
         Height          =   375
         Left            =   240
         TabIndex        =   26
         Top             =   2760
         Width           =   1935
      End
      Begin VB.Label LabelResPay 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   4080
         TabIndex        =   21
         Top             =   4200
         Width           =   2895
      End
      Begin VB.Label Label8 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Rest a Payé"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   4080
         Width           =   1455
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payé"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   3720
         Width           =   735
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Payment Total"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Fin d'exploitation"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   2055
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Début d'exploitation"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   2295
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Matricule de La Voiture"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   2655
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numéro Identité Client"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   2655
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Numéro de la Réservation"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3015
      End
   End
End
Attribute VB_Name = "FrmReserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Activate()
On Error Resume Next
  TextNumJour.Text = DateDiff("d", DTPickerDebExp.Value, DTPickerFinExp.Value)
  ComboMat.RemoveItem (0)
  ComboMat.SelLength = 0
  ComboNumID.RemoveItem (0)
  ComboNumID.SelLength = 0
  ComboNumID.SetFocus
End Sub

Private Sub Form_Load()
   Randomize
   Call Design(FrmReserver)
   Call RemplirLesCombosRes
   
End Sub

Private Sub CmdEnreg_Click()
  
'------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM RESERVATIONS WHERE ResID='" & UCase(TextResID.Text) & "' "
    rs.Open SQL, CN, adOpenKeyset
'------------------------------------------------------------------
    If rs.EOF Then
      If ChampsAddResOK Then
         x = MsgBox("Efféctuer l'enregistrement ?", vbYesNoCancel + vbQuestion, "Gestion Location de Voitures : Enregistrement")
          If x = vbYes Then
             TextResID.Text = GenNewResID
             Call AjouterReserv(UCase(Trim(ComboMat)))
             MsgBox "Enregistrement Effectué Avec Succées !!", vbInformation, "Gestion Location de Voitures : Voitures - ENREGISTREMENT"
             Me.Hide
          ElseIf x = vbCancel Then
              Me.Visible = False
          End If
      End If
    Else
        x = MsgBox("Efféctuer la modification ?", vbYesNoCancel + vbQuestion, "Gestion Location de Voiture : Modification")
         If x = vbYes Then
            Call ChangeVoitStatut(UCase(Trim(ComboMat.Text)), "Disponible")
            If ChampsAddResOK Then
                Call ModifierReserv(UCase(TextResID.Text))
                If OptionAvecClient.Value = True Then
                   Call ChangeVoitStatut(UCase(Trim(ComboMat.Text)), "En Location")
                Else
                   Call ChangeVoitStatut(UCase(Trim(ComboMat.Text)), "Disponible")
                End If
                MsgBox "Modification Effectuée Avec Succées !!", vbInformation, "Gestion Location de Voitures : Voiture  - Modification"
                Me.Hide
            End If
         ElseIf x = vbCancel Then
              Me.Visible = False
         End If
   End If
'------------------------------------------------------------------
End Sub

Private Sub CmdRaz_Click()
   ComboNumID.Text = ""
   ComboMat.Text = ""
   DTPickerDebExp.Value = Date
   DTPickerFinExp.Value = Date
   TextPayTotal.Text = ""
   TextPay.Text = ""
   OptionAuGar.Value = True
   DTPickerVoitRendu.Value = Date
End Sub

Private Sub ComboMat_Change()
   Call AutoComplete(ComboMat)
End Sub

Private Sub ComboNumID_Change()
   Call AutoComplete(ComboNumID)
End Sub

Private Sub DTPickerDebExp_Change()
  TextNumJour.Text = DateDiff("d", DTPickerDebExp.Value, DTPickerFinExp.Value)
  DTPickerFinExp.MinDate = DTPickerDebExp.Value
End Sub

Private Sub DTPickerFinExp_Change()
  TextNumJour.Text = DateDiff("d", DTPickerDebExp.Value, DTPickerFinExp.Value)
End Sub

Private Sub OptionAuGar_Click()
    LabelDateRend.Visible = False
    DTPickerVoitRendu.Visible = False
End Sub

Private Sub OptionAvecClient_Click()
    LabelDateRend.Caption = "Recuperé par le client le :"
    LabelDateRend.Visible = True
    DTPickerVoitRendu.Visible = True
End Sub

Private Sub OptionRendu_Click()
    LabelDateRend.Caption = "Rendu Le :"
    LabelDateRend.Visible = True
    DTPickerVoitRendu.Visible = True
End Sub

Private Sub TextNumJour_Change()
   DTPickerFinExp.Value = DateAdd("d", DTPickerDebExp.Value, Val(TextNumJour.Text))
   TextPayTotal.Text = Val(GetVoitPrix(ComboMat)) * Val(TextNumJour.Text)
End Sub

Private Sub TimerRes_Timer()
     LabelResPay.Caption = Val(TextPayTotal.Text) - Val(TextPay.Text)
End Sub
