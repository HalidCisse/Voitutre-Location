VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmRecherche 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Recherche"
   ClientHeight    =   3975
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7350
   BeginProperty Font 
      Name            =   "MS Reference Sans Serif"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmRecherche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3975
   ScaleWidth      =   7350
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox Combo 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1920
      Sorted          =   -1  'True
      TabIndex        =   6
      Text            =   "Choisir.."
      ToolTipText     =   "Choisir"
      Top             =   1680
      Width           =   2895
   End
   Begin VB.Timer Timer 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin MSComCtl2.DTPicker DTPickerF 
      Height          =   375
      Left            =   4920
      TabIndex        =   4
      Top             =   2280
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94765057
      CurrentDate     =   41755
   End
   Begin MSComCtl2.DTPicker DTPickerD 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2280
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   94765057
      CurrentDate     =   41755
   End
   Begin VB.CommandButton CmdAfficher 
      Appearance      =   0  'Flat
      Caption         =   "Afficher"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3240
      Width           =   2415
   End
   Begin VB.ComboBox ComboRecItem 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3720
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   3135
   End
   Begin VB.ComboBox ComboRecType 
      Appearance      =   0  'Flat
      Height          =   405
      ItemData        =   "FrmRecherche.frx":2AD05
      Left            =   600
      List            =   "FrmRecherche.frx":2AD12
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   480
      Width           =   2415
   End
   Begin VB.Label LabelSign 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Reference Sans Serif"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   3240
      TabIndex        =   5
      Top             =   1080
      Width           =   375
   End
End
Attribute VB_Name = "FrmRecherche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Me.Caption = "Recherche Avancée (Non Complet ~ BETA)"
  Call Design(FrmRecherche)
  Call TypeChamps(0, 0, 0, 0)
End Sub

Private Sub CmdAfficher_Click()
  If ComboRecType = "Voitures" Then
    Call RemplirGrid(FrmShowVoit, FrmShowVoit.VoitGrid, DoSQL)
  ElseIf ComboRecType = "Clients" Then
    Call RemplirGrid(FrmShowClient, FrmShowClient.ClientGrid, DoSQL)
  ElseIf ComboRecType = "Reservations" Then
    Call RemplirGrid(FrmShowRes, FrmShowRes.ResGrid, DoSQL)
  End If
End Sub

Private Sub Combo_Change()
'Autocomplete quand on pendant la saisie
  Call AutoComplete(FrmRecherche.Combo)
End Sub

Private Sub ComboRecItem_Validate(Cancel As Boolean)
'On Error Resume Next
'Populate la combobox des donnée correspondant de la BD
   Call Remplir(Combo, ResTable, ResChamps)
   Combo.RemoveItem (0)
End Sub

Private Sub ComboRecType_Validate(Cancel As Boolean)
  ComboRecItem.Visible = True
  Call RemplirItems
End Sub

Private Sub Timer_Timer()
  Call DoSQL
  If ComboRecType <> "" Then
    ComboRecItem.Visible = True
  End If
End Sub
