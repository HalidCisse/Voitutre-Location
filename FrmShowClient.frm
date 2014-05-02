VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmShowClient 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "List des Clients"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16590
   BeginProperty Font 
      Name            =   "MS Reference Sans Serif"
      Size            =   11.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmShowClient.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   5685
   ScaleWidth      =   16590
   StartUpPosition =   1  'CenterOwner
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid ClientGrid 
      Bindings        =   "FrmShowClient.frx":4C4A
      Height          =   5535
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "Double Click Pour Plus d'Informations"
      Top             =   0
      Width           =   16455
      _ExtentX        =   29025
      _ExtentY        =   9763
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
      GridLinesFixed  =   0
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
      _Band(0)._NumMapCols=   17
      _Band(0)._MapCol(0)._Name=   "Nom"
      _Band(0)._MapCol(0)._RSIndex=   5
      _Band(0)._MapCol(1)._Name=   "Prenom"
      _Band(0)._MapCol(1)._RSIndex=   6
      _Band(0)._MapCol(2)._Name=   "Sex"
      _Band(0)._MapCol(2)._RSIndex=   4
      _Band(0)._MapCol(3)._Name=   "TypeID"
      _Band(0)._MapCol(3)._RSIndex=   1
      _Band(0)._MapCol(4)._Name=   "NumID"
      _Band(0)._MapCol(4)._Caption=   "Numéro de l'identité"
      _Band(0)._MapCol(4)._RSIndex=   2
      _Band(0)._MapCol(5)._Name=   "Phone"
      _Band(0)._MapCol(5)._Caption=   "Télephone"
      _Band(0)._MapCol(5)._RSIndex=   10
      _Band(0)._MapCol(6)._Name=   "Email"
      _Band(0)._MapCol(6)._RSIndex=   11
      _Band(0)._MapCol(7)._Name=   "Adresse"
      _Band(0)._MapCol(7)._Caption=   "Adresse Domicile"
      _Band(0)._MapCol(7)._RSIndex=   12
      _Band(0)._MapCol(8)._Name=   "Statut"
      _Band(0)._MapCol(8)._RSIndex=   14
      _Band(0)._MapCol(9)._Name=   "NombreTotalReservation"
      _Band(0)._MapCol(9)._Caption=   "Total Reservation"
      _Band(0)._MapCol(9)._RSIndex=   13
      _Band(0)._MapCol(10)._Name=   "DateEnreg"
      _Band(0)._MapCol(10)._Caption=   "Date Enrégistrée"
      _Band(0)._MapCol(10)._RSIndex=   16
      _Band(0)._MapCol(11)._Name=   "NombreInfractions"
      _Band(0)._MapCol(11)._RSIndex=   15
      _Band(0)._MapCol(11)._Hidden=   -1  'True
      _Band(0)._MapCol(12)._Name=   "LieuNaiss"
      _Band(0)._MapCol(12)._RSIndex=   8
      _Band(0)._MapCol(12)._Hidden=   -1  'True
      _Band(0)._MapCol(13)._Name=   "DateExpID"
      _Band(0)._MapCol(13)._RSIndex=   3
      _Band(0)._MapCol(13)._Hidden=   -1  'True
      _Band(0)._MapCol(14)._Name=   "Nationalite"
      _Band(0)._MapCol(14)._RSIndex=   9
      _Band(0)._MapCol(14)._Hidden=   -1  'True
      _Band(0)._MapCol(15)._Name=   "DateOptPermis"
      _Band(0)._MapCol(15)._RSIndex=   0
      _Band(0)._MapCol(15)._Hidden=   -1  'True
      _Band(0)._MapCol(16)._Name=   "DateNaiss"
      _Band(0)._MapCol(16)._RSIndex=   7
      _Band(0)._MapCol(16)._Hidden=   -1  'True
   End
End
Attribute VB_Name = "FrmShowClient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
   Call ScaleGrid
End Sub

Private Sub Form_Activate()
   'ClientGrid.RowSel = 1
End Sub

Private Sub ClientGrid_DblClick()
 If ClientGrid.Rows <> 0 Then
   NumID = ClientGrid.TextMatrix(ClientGrid.RowSel, 4)
   FillFormClient (NumID)
   FrmClient.Show vbModal
   ShowPageClient (1)
 Else
    MsgBox "Vide !!", vbInformation + vbOKOnly, "Gestion Location de Voitures"
 End If
End Sub

Private Function ScaleGrid()

  ClientGrid.ColWidth(0) = 2000  'nom
  ClientGrid.ColWidth(1) = 2000  'prenom
  ClientGrid.ColWidth(2) = 1000  'sex
  ClientGrid.ColWidth(3) = 1200  'typeid
  ClientGrid.ColWidth(4) = 1500  'numid
  ClientGrid.ColWidth(5) = 1500  'phone
  ClientGrid.ColWidth(6) = 2000  'email
  ClientGrid.ColWidth(7) = 2000  'adress
  ClientGrid.ColWidth(8) = 1000  'statut
  ClientGrid.ColWidth(9) = 500   'n res
  ClientGrid.ColWidth(10) = 1500  'date enreg
  
  
  For i = 0 To ClientGrid.Cols - 1
      w = w + ClientGrid.ColWidth(i)
  Next
  FrmShowClient.Width = w + 500
End Function


