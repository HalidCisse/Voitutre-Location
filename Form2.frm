VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form FrmInformation 
   Appearance      =   0  'Flat
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Voitures - INFORMATIONS"
   ClientHeight    =   8880
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   14385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":0442
   ScaleHeight     =   8880
   ScaleWidth      =   14385
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton CmdSuiv 
      Caption         =   "Suivant >"
      Height          =   375
      Left            =   5640
      TabIndex        =   56
      ToolTipText     =   "Page Suivante"
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton CmdPrec 
      Caption         =   "< Prècedent"
      Height          =   375
      Left            =   120
      TabIndex        =   55
      ToolTipText     =   "Page Prècedente"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.CommandButton CmdRAZ 
      Caption         =   "Effaché Tous"
      Height          =   375
      Left            =   1920
      TabIndex        =   21
      ToolTipText     =   "Effacé"
      Top             =   8640
      Width           =   1695
   End
   Begin VB.CommandButton CmdENRG 
      Caption         =   "Enregistré"
      Height          =   375
      Left            =   3840
      TabIndex        =   20
      ToolTipText     =   "Enrégistré"
      Top             =   8640
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Information Assurance Et Vignette"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   7320
      TabIndex        =   14
      Top             =   5640
      Visible         =   0   'False
      Width           =   7215
      Begin VB.OptionButton OptionValVig0 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Non"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   6360
         TabIndex        =   54
         Top             =   1560
         Width           =   735
      End
      Begin VB.OptionButton OptionValVig1 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Oui"
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4200
         TabIndex        =   53
         Top             =   1560
         Width           =   855
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   0  'None
         Caption         =   "Frame10"
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   4080
         TabIndex        =   50
         Top             =   480
         Width           =   3015
         Begin VB.OptionButton OptionValAss0 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Non"
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   2280
            TabIndex        =   52
            Top             =   120
            Width           =   975
         End
         Begin VB.OptionButton OptionValAss1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Oui"
            ForeColor       =   &H00FF0000&
            Height          =   240
            Left            =   120
            TabIndex        =   51
            Top             =   120
            Width           =   735
         End
      End
      Begin MSComCtl2.DTPicker DTPickerDFAss 
         Height          =   375
         Left            =   4200
         TabIndex        =   49
         Top             =   960
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   661
         _Version        =   393216
         Format          =   95420417
         CurrentDate     =   41734
      End
      Begin VB.TextBox TxtAnVign 
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         TabIndex        =   19
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label22 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Année Vignette"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2040
         Width           =   1575
      End
      Begin VB.Label Label21 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Validité vignette"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Label Label20 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Fin Assurance"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   2175
      End
      Begin VB.Label Label19 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Validité Assurance"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   2055
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Information Accesoires"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2535
      Left            =   120
      TabIndex        =   13
      Top             =   5640
      Visible         =   0   'False
      Width           =   7215
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "GPS"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   63
         Top             =   240
         Width           =   6855
         Begin VB.OptionButton OptionGPS0 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Non"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   6000
            TabIndex        =   65
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton OptionGPS1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Oui"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   64
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Ecran DVD"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   28
         Top             =   1800
         Width           =   6855
         Begin VB.OptionButton OptionDVD0 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Non"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   6000
            TabIndex        =   30
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton OptionDVD1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Oui"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   4080
            TabIndex        =   29
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "WIFI"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   25
         Top             =   1320
         Width           =   6855
         Begin VB.OptionButton OptionWIFI0 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Non"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   6000
            TabIndex        =   27
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton OptionWIFI1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Oui"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Bluetooth"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   240
         TabIndex        =   22
         Top             =   840
         Width           =   6855
         Begin VB.OptionButton OptionBlue0 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Non"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   6000
            TabIndex        =   24
            Top             =   240
            Width           =   735
         End
         Begin VB.OptionButton OptionBLue1 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Oui"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   4080
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informations Complementaire"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   7320
      TabIndex        =   8
      Top             =   240
      Visible         =   0   'False
      Width           =   7215
      Begin VB.TextBox TextPrix 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   360
         Left            =   4320
         TabIndex        =   66
         Top             =   240
         Width           =   2655
      End
      Begin VB.Frame Frame8 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Etat de La Voiture"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   120
         TabIndex        =   57
         Top             =   2160
         Width           =   6855
         Begin VB.OptionButton OptionInDisp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Indisponible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   4680
            TabIndex        =   59
            Top             =   480
            Width           =   1455
         End
         Begin VB.OptionButton OptionDisp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disponible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00008000&
            Height          =   255
            Left            =   2040
            TabIndex        =   58
            Top             =   480
            Width           =   1335
         End
      End
      Begin MSComCtl2.DTPicker DTPickerAcquis 
         Height          =   375
         Left            =   4320
         TabIndex        =   48
         Top             =   1200
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   95420417
         CurrentDate     =   41734
      End
      Begin VB.ComboBox ComboCouleur 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Form2.frx":1738DC
         Left            =   4320
         List            =   "Form2.frx":1738DE
         TabIndex        =   46
         Top             =   720
         Width           =   2655
      End
      Begin VB.Frame FrameStatut 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Statut"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1695
         Left            =   120
         TabIndex        =   31
         Top             =   3360
         Width           =   6855
         Begin VB.OptionButton OptionSVigInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Vignette Invalide"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   4680
            TabIndex        =   37
            Top             =   1080
            Width           =   2055
         End
         Begin VB.OptionButton OptionSAssInv 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Assurance Invalide"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   2040
            TabIndex        =   36
            Top             =   1080
            Width           =   2535
         End
         Begin VB.OptionButton OptionSAcc 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Accidentée"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   4680
            TabIndex        =   35
            Top             =   840
            Width           =   1695
         End
         Begin VB.OptionButton OptionSEnPanne 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "En Panne"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   2040
            TabIndex        =   34
            Top             =   840
            Width           =   1335
         End
         Begin VB.OptionButton OptionSEnLocation 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "En Location"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   225
            Left            =   4680
            TabIndex        =   33
            Top             =   480
            Width           =   1695
         End
         Begin VB.OptionButton OptionSDisp 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Disponible"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2040
            TabIndex        =   32
            Top             =   480
            Width           =   1455
         End
      End
      Begin VB.ComboBox ComboNatureAchat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "Form2.frx":1738E0
         Left            =   4320
         List            =   "Form2.frx":1738EA
         TabIndex        =   12
         Top             =   1680
         Width           =   2655
      End
      Begin VB.Label Label7 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Prix de Location Journalière"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   67
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Nature Acquisition"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label Label10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Acquis"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1200
         Width           =   1455
      End
      Begin VB.Label Label9 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Couleur"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informations Services Mines"
      BeginProperty Font 
         Name            =   "Palatino Linotype"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   7215
      Begin VB.Frame FrameTypeCarb 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Type de Carburant"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         TabIndex        =   60
         Top             =   3240
         Width           =   6615
         Begin VB.OptionButton OptionDiesel 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Diesel"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C000C0&
            Height          =   255
            Left            =   5400
            TabIndex        =   62
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton OptionEssence 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Essence"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   3720
            TabIndex        =   61
            Top             =   360
            Width           =   1215
         End
      End
      Begin MSComCtl2.DTPicker DTPickerDMCIR 
         Height          =   375
         Left            =   4200
         TabIndex        =   47
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Century"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   95420417
         CurrentDate     =   41734
         MinDate         =   36526
      End
      Begin VB.Frame FrameNPlace 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Caption         =   "Nombre de Voiture"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   855
         Left            =   240
         TabIndex        =   42
         Top             =   4200
         Width           =   6615
         Begin VB.OptionButton OptionFaml 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Familial"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   255
            Left            =   5400
            TabIndex        =   45
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton OptionEco 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Economique"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   255
            Left            =   3720
            TabIndex        =   44
            Top             =   480
            Width           =   1575
         End
         Begin VB.OptionButton OptionMoy 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "Moyen"
            BeginProperty Font 
               Name            =   "MS Reference Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   225
            Left            =   2280
            TabIndex        =   43
            Top             =   480
            Width           =   1095
         End
      End
      Begin VB.ComboBox ComboPuissanceFiscale 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4200
         TabIndex        =   41
         Top             =   2640
         Width           =   2655
      End
      Begin VB.ComboBox ComboModele 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4200
         TabIndex        =   40
         Top             =   2160
         Width           =   2655
      End
      Begin VB.ComboBox ComboSousMarque 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4200
         TabIndex        =   39
         Top             =   1680
         Width           =   2655
      End
      Begin VB.ComboBox ComboMarque 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         ItemData        =   "Form2.frx":173907
         Left            =   4200
         List            =   "Form2.frx":173909
         TabIndex        =   38
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox TxtMat 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Century"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4200
         TabIndex        =   7
         Top             =   720
         Width           =   2655
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Puissance Fiscale"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Modele"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Sous Marque"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Marque"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Matricule"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Date Mis En Circulation"
         BeginProperty Font 
            Name            =   "MS Reference Sans Serif"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   2775
      End
   End
End
Attribute VB_Name = "FrmInformation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Call Design(Me)
  Call RempirLesCombosVoit
 
End Sub

Private Sub Form_Activate()
   Call ScaleFrmVoit
End Sub

Private Sub CmdENRG_Click()

If ChampsAddVoitOK Then
'------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM VOITURE WHERE MAT='" & UCase(TxtMat.Text) & "' "
    rs.Open SQL, CN, adOpenKeyset
'------------------------------------------------------------------
    If rs.EOF Then
         x = MsgBox("Efféctuer l'enregistrement ?", vbYesNoCancel + vbQuestion, "Gestion Location de Voitures : Enregistrement")
          If x = vbYes Then
             Call AjouterVoit
             MsgBox "Enregistrement Effectué Avec Succées !!", vbInformation, "Gestion Location de Voitures : Voitures - ENREGISTREMENT"
             Unload Me
          ElseIf x = vbCancel Then
              Unload Me
          Else
              ShowPage (1)
          End If
    Else
        x = MsgBox("Efféctuer la modification ?", vbYesNoCancel + vbQuestion, "Gestion Location de Voiture : Modification")
         If x = vbYes Then
            Call ModifierVoit
            MsgBox "Modification Effectuée Avec Succées !!", vbInformation, "Gestion Location de Voitures : Voiture  - Modification"
            Unload Me
         ElseIf x = vbCancel Then
              Unload Me
         Else
              ShowPage (1)
         End If
   End If
'------------------------------------------------------------------
End If
End Sub

Function ModifierVoit()
    Call LoadFormVoitData
    CN.Execute ("UPDATE VOITURE SET  DMCirc ='" & DataTab(0) & "',MARQUE ='" & DataTab(2) & "', S_MARQUE='" & DataTab(3) & "', MODELE='" & DataTab(4) & "', P_FISCALE='" & DataTab(5) & "',T_CARBURANT ='" & DataTab(6) & "', NB_PLACE='" & DataTab(7) & "', COULEUR='" & DataTab(8) & "', Date_Acquis='" & DataTab(9) & "', NATURE_ACHAT='" & DataTab(10) & "', INV_VOIT='" & DataTab(11) & "', STATUT='" & DataTab(12) & "', VALIDITE_ASSURANCE='" & DataTab(13) & "', DATE_FIN_ASSURANCE='" & DataTab(14) & "', VALIDITE_VIGNETTE='" & DataTab(15) & "', ANNEE_VIGNETTE='" & DataTab(16) & "', GPS='" & DataTab(17) & "', BLUETOOTH='" & DataTab(18) & "', WIFI='" & DataTab(19) & "', ECRAN='" & DataTab(20) & "', Prix = '" & DataTab(21) & "'   WHERE MAT='" & UCase(TxtMat.Text) & "' ")
AddEvent (" Voiture Modifiée: " & DataTab(1))
End Function

Private Sub CmdPrec_Click()
ShowPage (PageInf - 1)
End Sub

Private Sub CmdSuiv_Click()
  ShowPage (PageInf + 1)
End Sub

Private Sub ComboCouleur_Change()
   Call AutoComplete(ComboCouleur)
End Sub

Private Sub ComboMarque_Change()
  Call AutoComplete(ComboMarque)
End Sub

Private Sub ComboModele_Change()
  Call AutoComplete(ComboModele)
End Sub

Private Sub ComboPuissanceFiscale_Change()
   Call AutoComplete(ComboPuissanceFiscale)
End Sub

Private Sub ComboSousMarque_Change()
   Call AutoComplete(ComboSousMarque)
End Sub

Private Sub OptionDisp_Click()
'activer les options de status
   OptionSDisp.Value = True
   OptionSDisp.Enabled = True
   OptionSEnLocation.Enabled = True
   OptionSEnPanne.Enabled = True
   OptionSAcc.Enabled = True
   OptionSAssInv.Enabled = True
   OptionSVigInv.Enabled = True
End Sub

Private Sub OptionInDisp_Click()
'si on click sur indisponible desactiver les options status
   OptionSDisp.Value = False
   OptionSEnLocation.Value = False
   OptionSEnPanne.Value = False
   OptionSAcc.Value = False
   OptionSAssInv.Value = False
   OptionSVigInv.Value = False
   
   OptionSDisp.Enabled = False
   OptionSEnLocation.Enabled = False
   OptionSEnPanne.Enabled = False
   OptionSAcc.Enabled = False
   OptionSAssInv.Enabled = False
   OptionSVigInv.Enabled = False
End Sub

Private Sub OptionValVig0_Click()
   TxtAnVign.Text = ""
End Sub

Private Sub OptionValVig1_Click()
    TxtAnVign.Text = Year(Date)
End Sub

Private Sub CmdRaz_Click()
  TxtMat.Text = ""
  ComboMarque = ""
  ComboSousMarque = ""
  ComboModele = ""
  ComboPuissanceFiscale = ""
  DTPickerDMCIR.MaxDate = Date
  OptionEssence.Value = True
  OptionMoy.Value = True
  ComboNatureAchat = "Nouveau"
  OptionDisp = True
  OptionSDisp = True
  OptionSEnLocation.Enabled = False
  OptionValAss0 = True
  OptionValVig0 = True
  OptionGPS0 = True
  OptionBlue0 = True
  OptionWIFI0 = True
  OptionDVD0 = True
  Caption = "Ajouter Nouvelle Voiture"
  TextPrix.Text = ""
  ComboCouleur = ""
  OptionDisp.Value = True
End Sub

Private Function ScaleFrmVoit()
  'On Error Resume Next
    Me.Width = 7785
    Me.Height = 6375
    Me.Top = Main.Top
    Me.Left = Main.Left
    Frame1.Left = 240
    Frame2.Left = Frame1.Left
    Frame2.Top = Frame1.Top
    Frame4.Left = Frame1.Left
    Frame4.Top = Frame1.Top
    Frame3.Left = Frame1.Left
    Frame3.Top = Frame4.Top + Frame4.Height + 100
    CmdPrec.Top = 5520
    CmdPrec.Left = Frame1.Left
    CmdSuiv.Top = 5520
    CmdSuiv.Left = Frame1.Left + Frame1.Width - CmdSuiv.Width
    CmdRaz.Left = CmdPrec.Left
    CmdRaz.Top = CmdPrec.Top
    CmdENRG.Left = CmdSuiv.Left
    CmdENRG.Top = CmdSuiv.Top
    PageInf = 1
    ComboMarque.RemoveItem (0)
    ComboMarque.SelLength = 0
    ComboModele.RemoveItem (0)
    ComboModele.SelLength = 0
    ComboSousMarque.RemoveItem (0)
    ComboSousMarque.SelLength = 0
    ComboPuissanceFiscale.RemoveItem (0)
    ComboPuissanceFiscale.SelLength = 0
    ComboCouleur.RemoveItem (0)
    ComboCouleur.SelLength = 0
End Function
