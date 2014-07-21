VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LOCATION VOITURE : INFORMATIONS VOITURE"
   ClientHeight    =   6825
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6555
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6825
   ScaleWidth      =   6555
   StartUpPosition =   2  'CenterScreen
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHF1 
      Height          =   2295
      Left            =   240
      TabIndex        =   12
      Top             =   4320
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4048
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   8454016
      ForeColorFixed  =   0
      BackColorBkg    =   16777215
      BackColorUnpopulated=   16777215
      GridColor       =   -2147483635
      GridColorFixed  =   65280
      AllowBigSelection=   0   'False
      GridLinesFixed  =   1
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      RowSizingMode   =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Nirmala UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   6255
      Begin VB.CommandButton CmdQuitter 
         Caption         =   "QUITTER"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   11
         Top             =   840
         Width           =   1455
      End
      Begin VB.CommandButton CmdRaz 
         Caption         =   "RAZ"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton Cmdsuprimer 
         Caption         =   "SUPPRIMER"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3240
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CmdRecherche 
         Caption         =   "RECHERCHE"
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         TabIndex        =   8
         Top             =   600
         Width           =   1455
      End
      Begin VB.CommandButton CmdAjouter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Calisto MT"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H00FFFFFF&
         Picture         =   "Form1.frx":1AC2
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Enregistrer"
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   2
      Top             =   1680
      Width           =   3975
   End
   Begin VB.TextBox Text2 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   3975
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MODELE"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   1800
      Width           =   1095
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MARQUE"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "MATRICULE"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub TrierGRID()
'Trier Flex Grid
With MSHF1
     .Col = 1
     .ColSel = 0
     .Sort = flexSortGenericAscending 'flexSortStringAscending
     .Refresh
     .Col = 0
End With
Text1.Text = Empty
Text2.Text = Empty
Text3.Text = Empty
End Sub

Private Sub Cmdsuprimer_Click()

Y = MsgBox("Supprimer Voiture ?", vbYesNo + vbQuestion, "Gestion de Location Voiture : Suppresion Voiture")
If Y = vbYes Then
    Dim rs As New ADODB.Recordset
    Sql = "select * from VOITURE where ID_MAT='" & Text1.Text & "'"
    rs.Open Sql, cn, adOpenKeyset
If Not rs.EOF Then
    cn.Execute ("DELETE from VOITURE where ID_MAT='" & Text1.Text & "'")
    Text1.Text = Empty
    Text2.Text = Empty
    Text3.Text = Empty
    MsgBox "Suppression réussie", vbOKOnly + vbInformation, "Gestion de Location : Suppression"
Else
     MsgBox "Code Voiture Inéxistant", vbCritical, "Gestion Location de Voitures : Voiture - RECHERCHE"
End If
rs.Close
Set rs = Nothing
End If

End Sub
