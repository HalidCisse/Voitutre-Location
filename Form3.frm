VERSION 5.00
Begin VB.Form Main 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestion Location De Voitures"
   ClientHeight    =   6375
   ClientLeft      =   150
   ClientTop       =   480
   ClientWidth     =   8880
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form3.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   Picture         =   "Form3.frx":014A
   ScaleHeight     =   6375
   ScaleWidth      =   8880
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer TimerStat 
      Interval        =   10000
      Left            =   120
      Top             =   1680
   End
   Begin VB.CommandButton CmdCherRes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cherché"
      Height          =   1095
      Left            =   720
      Picture         =   "Form3.frx":14198
      Style           =   1  'Graphical
      TabIndex        =   8
      ToolTipText     =   "Cherché Une Réservation Par Numéro"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton CmdListRes 
      BackColor       =   &H00FFFFFF&
      Caption         =   "List"
      Height          =   1095
      Left            =   720
      Picture         =   "Form3.frx":16B56
      Style           =   1  'Graphical
      TabIndex        =   7
      ToolTipText     =   "List de Tous Les Réservations"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CmdListClient 
      Caption         =   "List"
      Height          =   1095
      Left            =   3600
      Picture         =   "Form3.frx":1CB38
      Style           =   1  'Graphical
      TabIndex        =   5
      ToolTipText     =   "List de Tous Les Clients"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CmdCherClient 
      Caption         =   "Chercher"
      Height          =   1095
      Left            =   3600
      Picture         =   "Form3.frx":2527A
      Style           =   1  'Graphical
      TabIndex        =   4
      ToolTipText     =   "Cherché Un Client Par Numéro Identité"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton CmdAddClient 
      Caption         =   "Ajouter Client"
      Height          =   1095
      Left            =   3600
      Picture         =   "Form3.frx":2DB6C
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "Ajouter Nouveau Client"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton CmdListVoit 
      Caption         =   "List"
      Height          =   1095
      Left            =   6480
      Picture         =   "Form3.frx":3602E
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "List de Tous Les Voitures"
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton CmdAddVoit 
      Caption         =   "Ajouter"
      Height          =   1095
      Left            =   6480
      Picture         =   "Form3.frx":3F070
      Style           =   1  'Graphical
      TabIndex        =   1
      ToolTipText     =   "Ajouter Nouvelle Voiture"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton CmdCherVoit 
      Caption         =   "Chercher"
      Height          =   1095
      Left            =   6480
      Picture         =   "Form3.frx":4116E
      Style           =   1  'Graphical
      TabIndex        =   0
      ToolTipText     =   "Cherché Une Voiture Par Numéro Matricule"
      Top             =   2400
      Width           =   1695
   End
   Begin VB.CommandButton CmdRes 
      Caption         =   "Reserver"
      Height          =   1095
      Left            =   720
      Picture         =   "Form3.frx":4243B
      Style           =   1  'Graphical
      TabIndex        =   6
      ToolTipText     =   "Ajouté Nouvelle Réservation"
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label LabelResDelaisPass 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "DELAIS PASSEE"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5400
      TabIndex        =   19
      ToolTipText     =   "Click Pour Voir"
      Top             =   5880
      Width           =   1560
   End
   Begin VB.Label LabelNVoit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voitures Enrégistrées"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   720
      TabIndex        =   18
      Top             =   6600
      Width           =   1875
   End
   Begin VB.Label LabelVoitDisp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Voitures Disponibles"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   720
      TabIndex        =   17
      Top             =   7080
      Width           =   1785
   End
   Begin VB.Label LabelResAjour 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reservations se terminent aujourd'hui"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   720
      TabIndex        =   16
      ToolTipText     =   "Click Pour Voir"
      Top             =   5880
      Width           =   3225
   End
   Begin VB.Label LabelResCours 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reservations en cours"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3000
      TabIndex        =   15
      Top             =   7080
      Width           =   1920
   End
   Begin VB.Label LabelNClients 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Clients Enrégistrées"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   3120
      TabIndex        =   14
      Top             =   6600
      Width           =   1755
   End
   Begin VB.Label LabelNRes 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Réservations Enrégistrées"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5520
      TabIndex        =   13
      Top             =   6600
      Width           =   2265
   End
   Begin VB.Label LabelTListVoit 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   6960
      TabIndex        =   12
      ToolTipText     =   "Total Voitures  Enrégistées"
      Top             =   5520
      Width           =   900
   End
   Begin VB.Label LabelTListClient 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   4080
      TabIndex        =   11
      ToolTipText     =   "Total Clients  Enrégistés"
      Top             =   5520
      Width           =   900
   End
   Begin VB.Label LabelTListRes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   1200
      TabIndex        =   10
      ToolTipText     =   "Total Réservations  Enrégistées"
      Top             =   5520
      Width           =   900
   End
   Begin VB.Label LabelTRes 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   240
      Left            =   1080
      TabIndex        =   9
      ToolTipText     =   "Voitures Disponibles"
      Top             =   1560
      Width           =   900
   End
   Begin VB.Menu mnProfile 
      Caption         =   "&Profile"
      Begin VB.Menu mnSeConnecter 
         Caption         =   "&Se Connecter"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnSeDeconnecter 
         Caption         =   "&Se Deconnecter"
      End
      Begin VB.Menu mnQuitter 
         Caption         =   "&Quitter"
      End
   End
   Begin VB.Menu mnReservation 
      Caption         =   "&Reservation"
      Begin VB.Menu mnListRes 
         Caption         =   "&Liste"
      End
      Begin VB.Menu mnReserver 
         Caption         =   "&Reservé"
      End
      Begin VB.Menu mnCherRes 
         Caption         =   "&Cherché"
      End
      Begin VB.Menu mnModRes 
         Caption         =   "&Modifié"
      End
      Begin VB.Menu mnSupRes 
         Caption         =   "&Supprimé"
      End
   End
   Begin VB.Menu mnClient 
      Caption         =   "&Clients"
      Begin VB.Menu mnListClients 
         Caption         =   "&List"
      End
      Begin VB.Menu mnCherClient 
         Caption         =   "&Cherché"
      End
      Begin VB.Menu mnAjoutClient 
         Caption         =   "&Ajouté Nouveau"
      End
      Begin VB.Menu mnSupClient 
         Caption         =   "&Supprimé"
      End
      Begin VB.Menu mnModClient 
         Caption         =   "&Modifier Informations D'Un Client"
      End
   End
   Begin VB.Menu mnVehicules 
      Caption         =   "&Vehicules"
      Begin VB.Menu mnList 
         Caption         =   "&Liste"
      End
      Begin VB.Menu mnChercher 
         Caption         =   "&Cherché"
      End
      Begin VB.Menu mnAjouterNouvelleVoiture 
         Caption         =   "&Ajouté  Nouvelle Voiture"
      End
      Begin VB.Menu mnSupprimer 
         Caption         =   "&Supprimé"
      End
      Begin VB.Menu mnModifier 
         Caption         =   "&Modifier Informations D'Une Voiture"
      End
   End
   Begin VB.Menu mnStatistiques 
      Caption         =   "&Statistiques"
      Begin VB.Menu MnRapportdactivites 
         Caption         =   "&Rapport D'Activités "
      End
      Begin VB.Menu mndisponibles 
         Caption         =   "&Voitures"
         Begin VB.Menu mnVoitDisp 
            Caption         =   "&Disponibles"
         End
         Begin VB.Menu mnVoitIndisp 
            Caption         =   "&Indisponibles"
         End
         Begin VB.Menu mnEnLoc 
            Caption         =   "&En Location"
         End
         Begin VB.Menu mnVoitPanne 
            Caption         =   "&En Panne"
         End
         Begin VB.Menu mnAssInv 
            Caption         =   "&Assurance Invalide"
         End
         Begin VB.Menu mnVigInv 
            Caption         =   "&Vignette Invalide"
         End
      End
      Begin VB.Menu mnSRes 
         Caption         =   "&Reservations"
         Begin VB.Menu mnResEnCours 
            Caption         =   "&En Cours"
         End
         Begin VB.Menu mnResToday 
            Caption         =   "&Se Termine Aujourd'hui"
         End
         Begin VB.Menu mnResDPAss 
            Caption         =   "&Delais Passé"
         End
         Begin VB.Menu mnResNCpay 
            Caption         =   "&Non Completement Payées"
         End
         Begin VB.Menu mnResTerm 
            Caption         =   "&Terminée"
         End
      End
      Begin VB.Menu mnSClient 
         Caption         =   "&Clients"
         Begin VB.Menu mnClientAvecVoit 
            Caption         =   "&Avec Voitures"
         End
         Begin VB.Menu mnClientFidele 
            Caption         =   "&Les Plus Fidèles"
            Enabled         =   0   'False
         End
         Begin VB.Menu mnClientAncien 
            Caption         =   "&Les Plus Anciens"
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu mnAdvancedRec 
         Caption         =   "&Recherche Avancée"
      End
   End
   Begin VB.Menu mnParam 
      Caption         =   "&Paramètres"
      Begin VB.Menu mnGesProfiles 
         Caption         =   "&Gestion Des Profiles"
         Enabled         =   0   'False
         Visible         =   0   'False
         Begin VB.Menu mnListUser 
            Caption         =   "&List des utilisateurs"
         End
         Begin VB.Menu mnAddUser 
            Caption         =   "&Ajouter un Utilisateur"
         End
         Begin VB.Menu mnModUser 
            Caption         =   "&Modifier Un Utilisateur"
         End
         Begin VB.Menu mnSupUser 
            Caption         =   "&Supprimer Un Utilisateur"
         End
         Begin VB.Menu mnCherUser 
            Caption         =   "&Cherché un utilisateur"
         End
         Begin VB.Menu mnGestLogin 
            Caption         =   "&Gestion des Login"
         End
      End
      Begin VB.Menu mnModPass 
         Caption         =   "&Modifier Mon Mot de Passe"
      End
      Begin VB.Menu mnPreferences 
         Caption         =   "&Préferences"
      End
   End
   Begin VB.Menu mnAide 
      Caption         =   "&?"
      Begin VB.Menu mnLicence 
         Caption         =   "&Licence"
      End
      Begin VB.Menu mnApropos 
         Caption         =   "&A Propos"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
















Private Sub mnLicence_Click()
  FrmLicence.Show 1
End Sub

'##########################################################################
'-----------------------------STATISTIQUES------------------------------


Private Sub TimerStat_Timer()
   Call UpdateMainStat
   Call UpdateTileStat
   Call UpdateData
End Sub

Private Sub MnRapportdactivites_Click()
'Seul les ADMIN peuvent voir les activite des ADMIN
    If UserType = "ADMIN" Then
      Call RemplirGrid(FrmShowEvents, FrmShowEvents.EventsGrid, "SELECT Message,UserEmail,EventTime,EventDate FROM EVENTS ORDER BY Cdate(EventDate) DESC, Cdate(EventTime) DESC")
    Else
      Call RemplirGrid(FrmShowEvents, FrmShowEvents.EventsGrid, "SELECT Message,UserEmail,EventTime,EventDate FROM EVENTS WHERE UserEmail = '" & UserEmail & "' ORDER BY Cdate(EventDate) DESC, Cdate(EventTime) DESC")
    End If
End Sub
                           '----------------------------
Private Sub mnResDPAss_Click()
  Call RemplirGridResS("SELECT * FROM RESERVATIONS WHERE ResStatut = 'Avec Client' AND Cdate(ResFin) < Date() ORDER BY Cdate(ResDate) DESC")
End Sub

Private Sub mnResNCpay_Click()
  Call RemplirGridResS("SELECT * FROM RESERVATIONS WHERE RestAPayer > '0' ORDER BY Cdate(ResDate) DESC ")
End Sub

Private Sub mnResTerm_Click()
  Call RemplirGridResS("SELECT * FROM RESERVATIONS WHERE ResStatut = 'Rendu' AND RestAPayer <= '0' ORDER BY Cdate(ResDate) DESC ")
End Sub

Private Sub mnResToday_Click()
 Call RemplirGridResS("SELECT * FROM RESERVATIONS WHERE ResStatut = 'Avec Client' AND Cdate(ResFin) = Date() ORDER BY Cdate(ResDate) DESC ")
End Sub

Private Sub mnResEnCours_Click()
  Call RemplirGridResS("SELECT * FROM RESERVATIONS WHERE ResStatut = 'Avec Client' ORDER BY Cdate(ResDate) DESC ")
End Sub

'                  ----------------------------------------

Private Sub mnClientAvecVoit_Click()
   Call RemplirGridClientS("SELECT * From ClIENTS INNER JOIN RESERVATIONS ON CLIENTS.NumID=RESERVATIONS.NumID WHERE RESERVATIONS.ResStatut = 'Avec Client' ORDER BY Cdate(RESERVATIONS.ResDate) DESC ")
End Sub

'                ----------------------------------------
Private Sub mnVoitDisp_Click()
    Call RemplirGridVoitS("Disponible")
End Sub

Private Sub mnVoitIndisp_Click()
    Call RemplirGridVoitS("Indisponible")
End Sub

Private Sub mnEnLoc_Click()
   Call RemplirGridVoitS("En Location")
End Sub

Private Sub mnVoitPanne_Click()
   Call RemplirGridVoitS("En Panne")
End Sub

Private Sub mnAssInv_Click()
   Call RemplirGridVoitS("Assurance Invalide")
End Sub

Private Sub mnVigInv_Click()
   Call RemplirGridVoitS("Vignette Invalide")
End Sub

'###################################################################################
'--------------------------------   MAIN  --------------------------------------------
Private Sub Form_Load()
 Call Design(Me)

End Sub

Private Sub Form_Activate()
  Call InitMe
  TimerStat.Enabled = True
  Call UpdateMainStat
  Call UpdateTileStat
  
End Sub

Private Sub mnAdvancedRec_Click()
  FrmRecherche.Show vbModal
End Sub

Private Sub Form_Deactivate()
  TimerStat.Enabled = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  x = MsgBox("Ete Vous Sure de Quitter ? ", vbQuestion + vbYesNo, "Gestion Location de Voitures")
  If x = vbNo Then
    Cancel = 1
  Else
    AddEvent ("Deconnecter: " & UserEmail)
    Call UnloadForms
    Unload FrmLogin
     Beep
    Unload Main
  End If
End Sub

Private Sub mnQuitter_Click()
    x = MsgBox("Voulez Vous Quitter", vbQuestion + vbYesNo, "Gestion De Location Voitures")
    If x = vbYes Then
      Call UnloadForms
      Unload FrmLogin
       Beep
      Unload Me
    End If
End Sub

Private Sub mnSeDeconnecter_Click()
 x = MsgBox("Voulez Vous vous deconnectez ?", vbQuestion + vbYesNo, "Gestion Location de Voitures")
 If x = vbYes Then
    Call DeconnectUser
 End If
End Sub

Private Sub mnApropos_Click()
   frmAbout.Show
   Main.WindowState = 1
End Sub

Private Sub LabelTRes_Click()
  Call RemplirGridVoitS("Disponible")
End Sub

Private Sub LabelResDelaisPass_Click()
  Call RemplirGridResS("SELECT * FROM RESERVATIONS WHERE ResStatut = 'Avec Client' AND Cdate(ResFin) < Date() ORDER BY Cdate(ResDate) DESC ")
End Sub

Private Sub LabelResAjour_Click()
  Call RemplirGridResS("SELECT * FROM RESERVATIONS WHERE ResStatut = 'Avec Client' AND Cdate(ResFin) = Date() ORDER BY Cdate(ResDate) DESC ")
End Sub

'######################################################################
'-----------------------Reserver-------------------------------------
Private Sub CmdRes_Click()
    Unload FrmReserver
    Call OpenAddRes
End Sub
Private Sub CmdCherRes_Click()
   Call OpenCherRes
End Sub
Private Sub CmdListRes_Click()
    Call RemplirGrid(FrmShowRes, FrmShowRes.ResGrid, "select * from RESERVATIONS ORDER BY Cdate(ResDate) DESC")
End Sub

'------------------------Client-----------------------------------------
Private Sub CmdAddClient_Click()
    Unload FrmClient
    Call OpenAddClient
End Sub
Private Sub CmdCherClient_Click()
    Call OpenCherClient
End Sub
Private Sub CmdListClient_Click()
   Call RemplirGrid(FrmShowClient, FrmShowClient.ClientGrid, "select * from CLIENTS ORDER BY DateEnreg DESC")
End Sub

'------------------------Voiture-----------------------------------------
Private Sub CmdAddVoit_Click()
   Unload FrmInformation
   Call OpenAddVoit
End Sub
Private Sub CmdCherVoit_Click()
   Call OpenCherVoit
End Sub
Private Sub CmdListVoit_Click()
   Call RemplirGrid(FrmShowVoit, FrmShowVoit.VoitGrid, "select * from VOITURE ORDER BY STATUT DESC")
End Sub
'##############################################################################

'---------------------------------------------------------------------
                   'Menue Voitures
'--------------------------------------------------------------------

Private Sub mnAjouterNouvelleVoiture_Click()
    Call OpenAddVoit
End Sub
Private Sub mnChercher_Click()
   Call OpenCherVoit
End Sub
Private Sub mnList_Click()
  Call RemplirGrid(FrmShowVoit, FrmShowVoit.VoitGrid, "select * from VOITURE ORDER BY STATUT DESC")
End Sub
Private Sub mnModifier_Click()
  Call OpenModVoit
End Sub
Private Sub mnSupprimer_Click()
  Call OpenSupVoit
End Sub

'----------------------------------------------------------------------
                   'Clients Menue
'----------------------------------------------------------------------

Private Sub mnListClients_Click()
   Call RemplirGrid(FrmShowClient, FrmShowClient.ClientGrid, "select * from CLIENTS ORDER BY DateEnreg DESC")
End Sub
Private Sub mnCherClient_Click()
    Call OpenCherClient
End Sub
Private Sub mnAjoutClient_Click()
   Call OpenAddClient
End Sub
Private Sub mnSupClient_Click()
   Call OpenSupClient
End Sub
Private Sub mnModClient_Click()
   Call OpenModClient
End Sub

'----------------------------------------------------------------------
                    ' Reservations Menue
'------------------------------------------------------------------------

Private Sub mnReserver_Click()
  Call OpenAddRes
End Sub
Private Sub mnModRes_Click()
    Call OpenModRes
End Sub
Private Sub mnSupRes_Click()
    Call OpenSupRes
End Sub
Private Sub mnCherRes_Click()
   Call OpenCherRes
End Sub
Private Sub mnListRes_Click()
  Call RemplirGrid(FrmShowRes, FrmShowRes.ResGrid, "select * from RESERVATIONS ORDER BY Cdate(ResDate) DESC")
End Sub

'--------------------------------------------------------------------
                     'Gestion Profiles
'---------------------------------------------------------------------

Private Sub mnListUser_Click()
   Call RemplirGrid(FrmShowUser, FrmShowUser.UserGrid, "select * from PROFILES")
End Sub
Private Sub mnSeConnecter_Click()
  FrmLogin.Show
End Sub

Private Sub mnAddUser_Click()
   Call OpenAddUser
End Sub

Private Sub mnModUser_Click()
   Call OpenModUser
End Sub

Private Sub mnModPass_Click()
   Call OpenModPass
End Sub

Private Sub mnSupUser_Click()
   Call OpenSupUser
End Sub

Private Sub mnCherUser_Click()
   Call OpenCherUser
End Sub


