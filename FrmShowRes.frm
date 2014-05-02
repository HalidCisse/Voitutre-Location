VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmShowRes 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "List des Reservations"
   ClientHeight    =   5670
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17355
   Icon            =   "FrmShowRes.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5670
   ScaleWidth      =   17355
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid ResGrid 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double Click Pour Plus d'Informations"
      Top             =   0
      Width           =   17415
      _ExtentX        =   30718
      _ExtentY        =   9551
      _Version        =   393216
      Rows            =   3
      Cols            =   12
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
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
      _Band(0).ColHeader=   1
      _Band(0)._NumMapCols=   12
      _Band(0)._MapCol(0)._Name=   "ResID"
      _Band(0)._MapCol(0)._Caption=   "Numéro de Reservation"
      _Band(0)._MapCol(0)._RSIndex=   0
      _Band(0)._MapCol(1)._Name=   "NumID"
      _Band(0)._MapCol(1)._Caption=   "Num ID du Client"
      _Band(0)._MapCol(1)._RSIndex=   1
      _Band(0)._MapCol(2)._Name=   "MAT"
      _Band(0)._MapCol(2)._Caption=   "Matricule de la voiture"
      _Band(0)._MapCol(2)._RSIndex=   2
      _Band(0)._MapCol(3)._Name=   "ResDate"
      _Band(0)._MapCol(3)._Caption=   "Date de la Reservation"
      _Band(0)._MapCol(3)._RSIndex=   3
      _Band(0)._MapCol(4)._Name=   "ResDebut"
      _Band(0)._MapCol(4)._Caption=   "Debut de l'exploitation"
      _Band(0)._MapCol(4)._RSIndex=   4
      _Band(0)._MapCol(5)._Name=   "ResFin"
      _Band(0)._MapCol(5)._Caption=   "Fin d'exploitation"
      _Band(0)._MapCol(5)._RSIndex=   5
      _Band(0)._MapCol(6)._Name=   "PayementDue"
      _Band(0)._MapCol(6)._Caption=   "Payment Total"
      _Band(0)._MapCol(6)._RSIndex=   6
      _Band(0)._MapCol(7)._Name=   "Paye"
      _Band(0)._MapCol(7)._Caption=   "Payé"
      _Band(0)._MapCol(7)._RSIndex=   7
      _Band(0)._MapCol(8)._Name=   "RestAPayer"
      _Band(0)._MapCol(8)._Caption=   "Rest à Payer"
      _Band(0)._MapCol(8)._RSIndex=   8
      _Band(0)._MapCol(9)._Name=   "ResStatut"
      _Band(0)._MapCol(9)._Caption=   "Statut de la voiture"
      _Band(0)._MapCol(9)._RSIndex=   9
      _Band(0)._MapCol(10)._Name=   "DateVoitRendue"
      _Band(0)._MapCol(10)._Caption=   "Voiture rendue le :"
      _Band(0)._MapCol(10)._RSIndex=   10
      _Band(0)._MapCol(11)._Name=   "DateVoitDClient"
      _Band(0)._MapCol(11)._Caption=   "Donné au client le :"
      _Band(0)._MapCol(11)._RSIndex=   11
   End
End
Attribute VB_Name = "FrmShowRes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Call ScaleGrid
End Sub

Private Sub ResGrid_DblClick()
 If ResGrid.Rows <> 0 Then
   ResId = ResGrid.TextMatrix(ResGrid.RowSel, 0)
   FillFormReserv (ResId)
   FrmReserver.Show vbModal
 Else
    MsgBox "Vide !!", vbInformation + vbOKOnly, "Gestion Location de Voitures"
 End If
End Sub

Private Sub Form_Activate()
   'ResGrid.RowSel = 1
End Sub
Private Function ScaleGrid()
  ResGrid.ColWidth(0) = 2000
  ResGrid.ColWidth(1) = 1500
  ResGrid.ColWidth(2) = 1500
  ResGrid.ColWidth(3) = 1300
  ResGrid.ColWidth(4) = 1300
  ResGrid.ColWidth(5) = 1300
  ResGrid.ColWidth(6) = 1000
  ResGrid.ColWidth(7) = 1000
  ResGrid.ColWidth(8) = 1000
  ResGrid.ColWidth(9) = 1300
  ResGrid.ColWidth(10) = 1300
  ResGrid.ColWidth(11) = 1500
  
  FrmShowRes.Width = 16500
End Function
