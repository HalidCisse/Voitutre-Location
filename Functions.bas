Attribute VB_Name = "FunctionsVoit"




'------------ STATUTS UPDATES Functions ----------------------


Function GetVoitPrix(MatID As String) As String
  If LoadDbVoitData(MatID) Then
       GetVoitPrix = DataTab(21)
   Else
       GetVoitPrix = 0
   End If
End Function

Function ChangeVoitPrix(MatID As String, Prix As String) As String
  Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM VOITURE"
    rs.Open SQL, CN, adOpenKeyset
  CN.Execute ("UPDATE VOITURE SET  Prix='" & Prix & "'  WHERE MAT='" & UCase(MatID) & "' ")
rs.Close
Set rs = Nothing
End Function

Function GetVoitStatut(MatID As String) As String
   If LoadDbVoitData(MatID) Then
       GetVoitStatut = DataTab(12)
   Else
       GetVoitStatut = "Null"
   End If
End Function

Function ChangeVoitStatut(MatID As String, Statut As String)
  Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM VOITURE"
    rs.Open SQL, CN, adOpenKeyset
    CN.Execute ("UPDATE VOITURE SET  STATUT='" & Statut & "'  WHERE MAT='" & UCase(MatID) & "' ")
rs.Close
Set rs = Nothing
End Function

'--------------- Fin UPDATES Functions -------------------

Function OpenAddVoit()
   Load FrmInformation
   FrmInformation.DTPickerDMCIR.MaxDate = Date
   FrmInformation.OptionEssence.Value = True
   FrmInformation.OptionMoy.Value = True
   FrmInformation.ComboNatureAchat = "Nouveau"
   FrmInformation.OptionDisp = True
   FrmInformation.OptionSDisp = True
   FrmInformation.OptionSEnLocation.Enabled = False
   FrmInformation.OptionValAss0 = True
   FrmInformation.OptionValVig0 = True
   FrmInformation.OptionGPS0 = True
   FrmInformation.OptionBlue0 = True
   FrmInformation.OptionWIFI0 = True
   FrmInformation.OptionDVD0 = True
   FrmInformation.Caption = "Ajouter Nouvelle Voiture"
   FrmInformation.Show (vbModal)
    ShowPage (1)
End Function

Function OpenSupVoit()
DebutSup:
   Mats = InputBox("Entrer Le Matricule de la Voiture a Supprimer ")
   If Trim(Mats) <> "" Then
    If LoadDbVoitData(CStr(Mats)) Then
       y = MsgBox("Supprimer Cette Voiture ?  " & " " & DataTab(2) & " " & DataTab(3) & " " & DataTab(4), vbYesNo + vbQuestion, "Suppression Voiture" & ": " & DataTab(1))
       If y = vbYes Then
         Call SupprimerVoit(CStr(Mats))
         MsgBox "Suppression réussie", vbOKOnly + vbInformation, "Gestion Voiture : Suppression"
       End If
    Else
      y = MsgBox("Matricule Voiture Inéxistant", vbCritical + vbRetryCancel, "Gestion Voiture : Voiture - Suppression")
      If y = vbRetry Then
        GoTo DebutSup
      End If
    End If
   End If
End Function

Function OpenModVoit()
debut:
  Matm = InputBox("Entrer Le Matricule de la Voiture a Modifier ")
  If Trim(Matm) <> "" Then
    If FillFormVoit(CStr(Matm)) Then
       FrmInformation.Caption = "Modifiez les Informations d'une Voiture"
       FrmInformation.Show vbModal
    Else
        x = MsgBox("Voiture Inexistant !!!", vbCritical + vbRetryCancel, "Location Voitures")
        If x = vbRetry Then
             GoTo debut
        End If
    End If
 End If
End Function

Function OpenCherVoit()
debut:
    matc = InputBox("Entrez Le Matricule de la Voiture a Chercher :")
    If matc <> "" Then
        If FillFormVoit(CStr(matc)) Then
          FrmInformation.Caption = "Informations Voiture"
          FrmInformation.Show vbModal
        Else
          x = MsgBox("Matricule Inexistant !!!", vbCritical + vbRetryCancel, "Chercher")
          If x = vbRetry Then
             GoTo debut
          End If
        End If
    End If
End Function

Function AjouterVoit()
    Call LoadFormVoitData
    CN.Execute (" INSERT INTO VOITURE VALUES ( '" & DataTab(0) & "','" & DataTab(1) & "','" & DataTab(2) & "','" & DataTab(3) & "','" & DataTab(4) & "','" & DataTab(5) & "','" & DataTab(6) & "','" & DataTab(7) & "','" & DataTab(8) & "','" & DataTab(9) & "','" & DataTab(10) & "','" & DataTab(11) & "','" & DataTab(12) & "','" & DataTab(13) & "','" & DataTab(14) & "','" & DataTab(15) & "','" & DataTab(16) & "','" & DataTab(17) & "','" & DataTab(18) & "','" & DataTab(19) & "','" & DataTab(20) & "','" & DataTab(21) & "' ) ")
Call SaveNewVoitDATA
AddEvent ("Voiture Ajoutée: " & DataTab(1))
End Function

Function SupprimerVoit(Mats As String)
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM VOITURE WHERE MAT='" & UCase(Mats) & "' "
    rs.Open SQL, CN, adOpenKeyset
        If Not rs.EOF Then
           CN.Execute ("DELETE from VOITURE where Mat='" & UCase(Mats) & "'")
        End If
 rs.Close
 Set rs = Nothing
 AddEvent (" Voiture Supprimée: " & Mats)
End Function

Function LoadFormVoitData()
   DataTab(0) = FrmInformation.DTPickerDMCIR.Value
   DataTab(1) = UCase(FrmInformation.TxtMat.Text)
   DataTab(2) = UCase(FrmInformation.ComboMarque)
   DataTab(3) = UCase(FrmInformation.ComboSousMarque)
   DataTab(4) = UCase(FrmInformation.ComboModele)
   DataTab(5) = FrmInformation.ComboPuissanceFiscale
   '---------------------------------
    If FrmInformation.OptionDiesel.Value = True Then
           DataTab(6) = "Diesel"
        Else
           DataTab(6) = "Essence"
    End If
   '-----------------------------
    If FrmInformation.OptionMoy.Value = True Then
          DataTab(7) = "Moyen"
        ElseIf FrmInformation.OptionEco.Value = True Then
          DataTab(7) = "Economique"
        Else
          DataTab(7) = "Familial"
    End If
   '-------------------------------
   DataTab(8) = FrmInformation.ComboCouleur
   DataTab(9) = FrmInformation.DTPickerAcquis.Value
   DataTab(10) = FrmInformation.ComboNatureAchat
   '-------------------------------
    If FrmInformation.OptionDisp.Value = True Then
             DataTab(11) = "Disponible"
             
            If FrmInformation.OptionSDisp.Value = True Then
                  DataTab(12) = "Disponible"
            ElseIf FrmInformation.OptionSEnLocation.Value = True Then
                  DataTab(12) = "En Location"
            ElseIf FrmInformation.OptionSEnPanne.Value = True Then
                  DataTab(12) = "En Panne"
            ElseIf FrmInformation.OptionSAcc.Value = True Then
                  DataTab(12) = "Accidentée"
            ElseIf FrmInformation.OptionSAssInv.Value = True Then
                  DataTab(12) = "Assurance Invalide"
            ElseIf FrmInformation.OptionSVigInv.Value = True Then
                  DataTab(12) = "Vignette Invalide"
            Else
                  DataTab(12) = "Non Connu"
            End If
    Else
             DataTab(11) = "Indisponible"
             DataTab(12) = "Indisponible"
    End If
    '-----------------------------
    If FrmInformation.OptionValAss1.Value = True Then
           DataTab(13) = "Oui"
        Else
           DataTab(13) = "Non"
    End If
    '-----------------------------
   DataTab(14) = FrmInformation.DTPickerDFAss.Value
   
    If FrmInformation.OptionValVig1.Value = True Then
           DataTab(15) = "Oui"
        Else
           DataTab(15) = "Non"
    End If
   '------------------------------
   DataTab(16) = FrmInformation.TxtAnVign.Text
   '-------------------------------
    If FrmInformation.OptionGPS1.Value = True Then
              DataTab(17) = "Oui"
        Else
              DataTab(17) = "Non"
    End If
    '------------------------------
    If FrmInformation.OptionBLue1.Value = True Then
               DataTab(18) = "Oui"
        Else
               DataTab(18) = "Non"
    End If
    '-------------------------------
    If FrmInformation.OptionWIFI1.Value = True Then
              DataTab(19) = "Oui"
        Else
              DataTab(19) = "Non"
    End If
    '-------------------------------
    If FrmInformation.OptionDVD1.Value = True Then
              DataTab(20) = "Oui"
        Else
              DataTab(20) = "Non"
    End If
    '-------------------------------
    DataTab(21) = Val(FrmInformation.TextPrix)
    
End Function

Function FillFormVoit(Matf As String) As Boolean
If LoadDbVoitData(Matf) Then
   FrmInformation.DTPickerDMCIR.Value = DataTab(0)
   FrmInformation.TxtMat.Text = DataTab(1)
   FrmInformation.ComboMarque = DataTab(2)
   FrmInformation.ComboSousMarque = DataTab(3)
   FrmInformation.ComboModele = DataTab(4)
   FrmInformation.ComboPuissanceFiscale = DataTab(5)
   '---------------------------------
    If DataTab(6) = "Diesel" Then
           FrmInformation.OptionDiesel.Value = True
        Else
           FrmInformation.OptionEssence.Value = True
    End If
   '-----------------------------
    If DataTab(7) = "Familial" Then
          FrmInformation.OptionFaml.Value = True
        ElseIf DataTab(7) = "Economique" Then
          FrmInformation.OptionEco.Value = True
        Else
          FrmInformation.OptionMoy.Value = True
    End If
   '-------------------------------
   FrmInformation.ComboCouleur = DataTab(8)
   FrmInformation.DTPickerAcquis.Value = DataTab(9)
   FrmInformation.ComboNatureAchat = DataTab(10)
   '-------------------------------
    If DataTab(11) = "Oui" Then
             FrmInformation.OptionDisp.Value = True
            If FrmInformation.OptionSDisp.Value = True Then
                ElseIf DataTab(12) = "En Location" Then
                      FrmInformation.OptionSEnLocation.Value = True
                ElseIf DataTab(12) = "En Panne" Then
                      FrmInformation.OptionSEnPanne.Value = True
                ElseIf DataTab(12) = "Accidentée" Then
                      FrmInformation.OptionSAcc.Value = True
                ElseIf DataTab(12) = "Assurance Invalide" Then
                      FrmInformation.OptionSAssInv.Value = True
                ElseIf DataTab(12) = "Vignette Invalide" Then
                      FrmInformation.OptionSVigInv.Value = True
            End If
    Else
             FrmInformation.OptionInDisp.Value = True
    End If
   '------------------------------
    '-----------------------------
    If DataTab(13) = "Oui" Then
           FrmInformation.OptionValAss1.Value = True
        Else
           FrmInformation.OptionValAss0.Value = True
    End If
    '-----------------------------
   FrmInformation.DTPickerDFAss.Value = DataTab(14)
   '------------------------------
    If DataTab(15) = "Oui" Then
           FrmInformation.OptionValVig1.Value = True
        Else
           FrmInformation.OptionValVig0.Value = True
    End If
   '------------------------------
   FrmInformation.TxtAnVign.Text = DataTab(16)
   '-------------------------------
    If DataTab(17) = "Oui" Then
              FrmInformation.OptionGPS1.Value = True
        Else
              FrmInformation.OptionGPS0.Value = True
    End If
    '------------------------------
    If DataTab(18) = "Oui" Then
               FrmInformation.OptionBLue1.Value = True
        Else
               FrmInformation.OptionBlue0.Value = True
    End If
    '-------------------------------
    If DataTab(19) = "Oui" Then
              FrmInformation.OptionWIFI1.Value = True
        Else
              FrmInformation.OptionWIFI0.Value = True
    End If
    '-------------------------------
    If DataTab(20) = "Oui" Then
              FrmInformation.OptionDVD1.Value = True
        Else
              FrmInformation.OptionDVD0.Value = True
    End If
    FrmInformation.TextPrix.Text = DataTab(21)
    '------------------------------
    FillFormVoit = True
Else
    FillFormVoit = False
End If
End Function

Function LoadDbVoitData(Matricule As String) As Boolean
'On Error Resume Next
Dim rs As New ADODB.Recordset
SQL = "select * from VOITURE where MAT='" & UCase(Matricule) & "'"
rs.Open SQL, CN, adOpenKeyset
    If Not rs.EOF Then
        For i = 0 To 21
          DataTab(i) = rs(i)
        Next
        LoadDbVoitData = True
    Else
        LoadDbVoitData = False
    End If
rs.Close
Set rs = Nothing
End Function

Function ShowPage(Page As Integer)
'On Error Resume Next
If Page = 1 Then
    FrmInformation.Frame1.Visible = True
    FrmInformation.Frame2.Visible = False
    FrmInformation.Frame3.Visible = False
    FrmInformation.Frame4.Visible = False
    
    FrmInformation.CmdPrec.Visible = False
    FrmInformation.CmdSuiv.Visible = True
    FrmInformation.CmdRaz.Left = FrmInformation.CmdPrec.Left
    FrmInformation.CmdENRG.Left = FrmInformation.CmdSuiv.Left
    'FrmInformation.DTPickerDMCIR.SetFocus
    
    PageInf = 1
     
ElseIf Page = 2 Then
    FrmInformation.CmdPrec.Visible = True
    FrmInformation.Frame1.Visible = False
    FrmInformation.Frame2.Visible = True
    FrmInformation.Frame3.Visible = False
    FrmInformation.Frame4.Visible = False
    
    FrmInformation.CmdPrec.Visible = True
    FrmInformation.CmdSuiv.Visible = True
    FrmInformation.TextPrix.SetFocus
    PageInf = 2
ElseIf Page = 3 Then
    FrmInformation.Frame1.Visible = False
    FrmInformation.Frame2.Visible = False
    FrmInformation.Frame3.Visible = True
    FrmInformation.Frame4.Visible = True
  
    FrmInformation.CmdENRG.Visible = True
    FrmInformation.CmdPrec.Visible = True
    FrmInformation.CmdSuiv.Visible = False
    FrmInformation.DTPickerDFAss.SetFocus
    PageInf = 3
End If
End Function

Function ChampsAddVoitOK() As Boolean
'Conditions Pour accepter l'ajout d'une nouvelle voit
ChampsAddVoitOK = False
AgeVoit = (DateDiff("yyyy", FrmInformation.DTPickerDMCIR.Value, Date))
    
       If Trim(FrmInformation.TxtMat.Text) = "" Then
           MsgBox "Matricule Invalide", vbOKOnly + vbInformation, "Gestion Location de Voiture"
           ShowPage (1)
           FrmInformation.TxtMat.SetFocus
       ElseIf Trim(FrmInformation.ComboMarque) = "" Then
           MsgBox "Marque Invalide", vbOKOnly + vbInformation, "Gestion Location de Voiture"
           ShowPage (1)
           FrmInformation.ComboMarque.SetFocus
       ElseIf Trim(FrmInformation.ComboCouleur) = "" Then
           MsgBox "Couleur Invalide", vbOKOnly + vbInformation, "Gestion Location de Voiture"
           ShowPage (2)
           FrmInformation.ComboCouleur.SetFocus
       ElseIf Trim(FrmInformation.ComboNatureAchat) = "" Then
           MsgBox "Veiller Précisez La Nature d'acquisition", vbOKOnly + vbInformation, "Gestion Location de Voiture"
           ShowPage (2)
           FrmInformation.ComboNatureAchat.SetFocus
       ElseIf Trim(FrmInformation.TxtAnVign.Text) = "" Then
           MsgBox "Veillez Precisez l'année de la vignette", vbOKOnly + vbInformation, "Gestion Location de Voiture"
           ShowPage (3)
           FrmInformation.TxtAnVign.SetFocus
       ElseIf AgeVoit > 5 Then
           If FrmInformation.OptionDisp.Value = True Then
                x = MsgBox("Cette Voiture Sera Indisponible Pour Location Car Elle Est Vieille De " & AgeVoit & " ans " & "Voulez Vous l'enregistré ?", vbYesNo + vbInformation, "Gestion Location de Voiture")
                If x = vbYes Then
                   FrmInformation.OptionInDisp = True
                   ChampsAddVoitOK = True
                Else
                   x = MsgBox("Voulez vous Quitté ?", vbQuestion + vbYesNo, "Gestion Location de Voiture")
                   If x = vbYes Then
                      Unload FrmInformation
                   Else
                      ShowPage (1)
                      FrmInformation.DTPickerDMCIR.SetFocus
                   End If
                End If
           Else
              ChampsAddVoitOK = True
           End If
      Else
        ChampsAddVoitOK = True
      End If
              
End Function

