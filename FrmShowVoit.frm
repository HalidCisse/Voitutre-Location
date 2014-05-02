VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmShowVoit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "List Des Vehicules"
   ClientHeight    =   5655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13860
   Icon            =   "FrmShowVoit.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5655
   ScaleWidth      =   13860
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid VoitGrid 
      Bindings        =   "FrmShowVoit.frx":014A
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double Click Pour Plus d'Informations"
      Top             =   0
      Width           =   23295
      _ExtentX        =   41090
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   3
      Cols            =   11
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   33023
      BackColorSel    =   14737632
      ForeColorSel    =   16777215
      BackColorBkg    =   16777215
      BackColorUnpopulated=   -2147483643
      GridColorFixed  =   0
      GridColorUnpopulated=   16777215
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   3
      BorderStyle     =   0
      Appearance      =   0
      BandDisplay     =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Reference Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   11
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   22
      _Band(0)._MapCol(0)._Name=   "MAT"
      _Band(0)._MapCol(0)._Caption=   "N° Matricule"
      _Band(0)._MapCol(0)._RSIndex=   1
      _Band(0)._MapCol(1)._Name=   "MARQUE"
      _Band(0)._MapCol(1)._RSIndex=   2
      _Band(0)._MapCol(2)._Name=   "S_MARQUE"
      _Band(0)._MapCol(2)._Caption=   "SOUS MARQUE"
      _Band(0)._MapCol(2)._RSIndex=   3
      _Band(0)._MapCol(3)._Name=   "MODELE"
      _Band(0)._MapCol(3)._RSIndex=   4
      _Band(0)._MapCol(4)._Name=   "T_CARBURANT"
      _Band(0)._MapCol(4)._Caption=   "CARBURANT"
      _Band(0)._MapCol(4)._RSIndex=   6
      _Band(0)._MapCol(5)._Name=   "NB_PLACE"
      _Band(0)._MapCol(5)._Caption=   "TYPE"
      _Band(0)._MapCol(5)._RSIndex=   7
      _Band(0)._MapCol(6)._Name=   "COULEUR"
      _Band(0)._MapCol(6)._RSIndex=   8
      _Band(0)._MapCol(7)._Name=   "DATE_FIN_ASSURANCE"
      _Band(0)._MapCol(7)._Caption=   "FIN ASSURANCE"
      _Band(0)._MapCol(7)._RSIndex=   14
      _Band(0)._MapCol(8)._Name=   "VALIDITE_VIGNETTE"
      _Band(0)._MapCol(8)._Caption=   "VALIDITE VIGNETTE"
      _Band(0)._MapCol(8)._RSIndex=   15
      _Band(0)._MapCol(9)._Name=   "Prix"
      _Band(0)._MapCol(9)._Caption=   "PRIX / JOUR"
      _Band(0)._MapCol(9)._RSIndex=   21
      _Band(0)._MapCol(10)._Name=   "STATUT"
      _Band(0)._MapCol(10)._RSIndex=   12
      _Band(0)._MapCol(11)._Name=   "ANNEE_VIGNETTE"
      _Band(0)._MapCol(11)._RSIndex=   16
      _Band(0)._MapCol(11)._Hidden=   -1  'True
      _Band(0)._MapCol(12)._Name=   "Date_Acquis"
      _Band(0)._MapCol(12)._RSIndex=   9
      _Band(0)._MapCol(12)._Hidden=   -1  'True
      _Band(0)._MapCol(13)._Name=   "NATURE_ACHAT"
      _Band(0)._MapCol(13)._RSIndex=   10
      _Band(0)._MapCol(13)._Hidden=   -1  'True
      _Band(0)._MapCol(14)._Name=   "GPS"
      _Band(0)._MapCol(14)._RSIndex=   17
      _Band(0)._MapCol(14)._Hidden=   -1  'True
      _Band(0)._MapCol(15)._Name=   "INV_VOIT"
      _Band(0)._MapCol(15)._RSIndex=   11
      _Band(0)._MapCol(15)._Hidden=   -1  'True
      _Band(0)._MapCol(16)._Name=   "DMCirc"
      _Band(0)._MapCol(16)._RSIndex=   0
      _Band(0)._MapCol(16)._Hidden=   -1  'True
      _Band(0)._MapCol(17)._Name=   "VALIDITE_ASSURANCE"
      _Band(0)._MapCol(17)._RSIndex=   13
      _Band(0)._MapCol(17)._Hidden=   -1  'True
      _Band(0)._MapCol(18)._Name=   "P_FISCALE"
      _Band(0)._MapCol(18)._RSIndex=   5
      _Band(0)._MapCol(18)._Hidden=   -1  'True
      _Band(0)._MapCol(19)._Name=   "BLUETOOTH"
      _Band(0)._MapCol(19)._RSIndex=   18
      _Band(0)._MapCol(19)._Hidden=   -1  'True
      _Band(0)._MapCol(20)._Name=   "WIFI"
      _Band(0)._MapCol(20)._RSIndex=   19
      _Band(0)._MapCol(20)._Hidden=   -1  'True
      _Band(0)._MapCol(21)._Name=   "ECRAN"
      _Band(0)._MapCol(21)._RSIndex=   20
      _Band(0)._MapCol(21)._Hidden=   -1  'True
   End
End
Attribute VB_Name = "FrmShowVoit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
   Call ScaleGrid
End Sub

Private Sub Form_Activate()
   'VoitGrid.RowSel = 1
End Sub

Private Function ScaleGrid()
  For i = 0 To VoitGrid.Cols
      VoitGrid.ColWidth(i) = 1500
  Next
  FrmShowVoit.Width = VoitGrid.Width - 6500
End Function

Private Sub VoitGrid_DblClick()
    If VoitGrid.Rows <> 0 Then
       Mat = VoitGrid.TextMatrix(VoitGrid.RowSel, 0)
       FillFormVoit (Mat)
       FrmInformation.Show vbModal
       ShowPage (1)
    Else
       MsgBox "Vide !!", vbInformation + vbOKOnly, "Gestion Location de Voitures"
    End If
End Sub

