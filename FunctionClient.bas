Attribute VB_Name = "Client"

'--------------------INTERFACES--------------------------------
Function ChampsAddClientOK() As Boolean
ChampsAddClientOK = False
  
  v = DateDiff("yyyy", FrmClient.DTPickerPermis.Value, Date)
  If v > 2 Then
     MsgBox "Permis agé de " & c & " ans", vbCritical + vbOKOnly, "Gestion Location de Voitures"
  End If
       If Trim(FrmClient.TextNumID.Text) = "" Then
             MsgBox "Veillez entrer le numéro de l'identité du client", vbCritical + vbOKOnly, "Gestion Location de Voitures"
             ShowPageClient (1)
        ElseIf DateDiff("yyyy", FrmClient.DTPickerExpID.Value, Date) > 0 Then
             MsgBox "L'identité du client est expiré !!", vbCritical + vbOKOnly, "Gestion Location de Voitures"
             ShowPageClient (1)
        ElseIf Trim(FrmClient.TextNom.Text) = "" Then
             MsgBox "Veillez entrez le nom du client !", vbCritical + vbOKOnly, "Gestion de Location de Voitures"
             ShowPageClient (1)
             FrmClient.TextNom.SetFocus
        ElseIf Trim(FrmClient.TextPrenom.Text) = "" Then
             MsgBox "Veillez entrez le prénom du client !", vbCritical + vbOKOnly, "Gestion de Location de Voitures"
             ShowPageClient (1)
             FrmClient.TextPrenom.SetFocus
        ElseIf Trim(FrmClient.ComboNat) = "" Then
             MsgBox "Veillez entrez la nationalité du client !", vbCritical + vbOKOnly, "Gestion de Location de Voitures"
             ShowPageClient (1)
             FrmClient.ComboNat.SetFocus
        ElseIf Trim(FrmClient.ComboLieuNaiss) = "" Then
             MsgBox "Veillez entrez le lieu de naissance du client !", vbCritical + vbOKOnly, "Gestion de Location de Voitures"
             ShowPageClient (1)
             FrmClient.ComboLieuNaiss.SetFocus
        ElseIf Trim(FrmClient.TextNumTel.Text) = "" Then
             MsgBox "Veillez entrez le numéro de téléphone du client !", vbCritical + vbOKOnly, "Gestion de Location de Voitures"
             ShowPageClient (1)
             FrmClient.TextNumTel.SetFocus
        ElseIf Trim(FrmClient.TextAdress.Text) = "" Then
             MsgBox "Veillez entrez l'adresse du client !", vbCritical + vbOKOnly, "Gestion de Location de Voitures"
             ShowPageClient (1)
             FrmClient.TextAdress.SetFocus
        Else
           ChampsAddClientOK = True
        End If
  
End Function

Function OpenModClient()
debut:
    idm = InputBox("Entrer Le Matricule du Client a Modifier ")
    If Trim(idm) <> "" Then
        If FillFormClient(CStr(idm)) Then
            FrmClient.Caption = "Modifier Des Informations de " & DataTab(5) & " " & DataTab(6)
            FrmClient.Show vbModal
            FrmClient.Hide
        Else
            x = MsgBox("Client Inexistant !!!", vbCritical + vbRetryCancel, "Location Voitures")
            If x = vbRetry Then
                 GoTo debut
            End If
        End If
    End If
End Function

Function OpenSupClient()
DebutSup:
   clientid = InputBox("Entrer Le numéro ID du client a Supprimer ")
   If Trim(clientid) <> "" Then
    If LoadDbClientData(CStr(clientid)) Then
       y = MsgBox("Supprimer Les Informations de ce Clients ?  " & " " & DataTab(6) & " " & DataTab(5), vbYesNo + vbQuestion, "Location Voiture" & ": " & DataTab(4))
       If y = vbYes Then
         Call SupprimerClient(CStr(clientid))
         MsgBox "Suppression réussie", vbOKOnly + vbInformation, "Gestion Voiture : Suppression"
       End If
    Else
      y = MsgBox("Client Inéxistant", vbCritical + vbRetryCancel, "Gestion Voiture : Voiture - Suppression")
      If y = vbRetry Then
        GoTo DebutSup
      End If
    End If
   End If
   
End Function

Function OpenAddClient()
    Load FrmClient
    FrmClient.Caption = "Informations Nouveau Client"
    FrmClient.OptionCIN.Value = True
    FrmClient.OptionH.Value = True
    FrmClient.TextFid.Text = 0
    FrmClient.TextStatut.Text = "Nouveau"
    FrmClient.TextInfrac.Text = 0
    FrmClient.Show vbModal
End Function

Function OpenCherClient()
debut:
    NumID = InputBox("Entrez Le Numero Identite du Client a Chercher :")
    If NumID <> "" Then
        If FillFormClient(CStr(NumID)) Then
          FrmClient.Caption = "Informations de " & DataTab(5) & " " & DataTab(6)
          FrmClient.Show vbModal
        Else
          x = MsgBox("Client Inexistant !!!", vbCritical + vbRetryCancel, "Location Voitures")
          If x = vbRetry Then
             GoTo debut
          End If
        End If
    End If
End Function
'-----------------------Ops -------------------------------------------

Function GetClientStatut(NumID As String) As String
   If LoadDbClientData(NumID) Then
     GetClientStatut = DataTab(14)
   Else
     GetClientStatut = "NULL"
   End If
End Function

Function AddClientFid(NumID As String)
  If LoadDbClientData(NumID) Then
       AddFid = CStr(Val(DataTab(13)) + 1)
    Dim rs As New ADODB.Recordset
    SQL = "SELECT NombreTotalReservation FROM CLIENTS"
    rs.Open SQL, CN, adOpenKeyset
       CN.Execute ("UPDATE CLIENTS SET  NombreTotalReservation='" & AddFid & "'  WHERE NumID='" & NumID & "' ")
       AddEvent ("[AUTO] [" & NumID & "] Fidelité + 1")
  End If
 rs.Close
 Set rs = Nothing
End Function

Function AddClientInfrac(NumID As String)
  If LoadDbClientData(NumID) Then
       AddIfrac = CStr(Val(DataTab(15)) + 1)
    Dim rs As New ADODB.Recordset
    SQL = "SELECT NombreInfractions FROM CLIENTS"
    rs.Open SQL, CN, adOpenKeyset
       CN.Execute ("UPDATE CLIENTS SET  NombreInfractions='" & AddIfrac & "'  WHERE NumID='" & NumID & "' ")
       AddEvent ("[AUTO] [" & NumID & "] Infraction + 1")
  End If
 rs.Close
 Set rs = Nothing
End Function

Function AjouterClient()
    Call LoadFilledClientData
    CN.Execute (" INSERT INTO CLIENTS VALUES ( '" & DataTab(0) & "','" & DataTab(1) & "','" & DataTab(2) & "','" & DataTab(3) & "','" & DataTab(4) & "','" & DataTab(5) & "','" & DataTab(6) & "','" & DataTab(7) & "','" & DataTab(8) & "','" & DataTab(9) & "','" & DataTab(10) & "','" & DataTab(11) & "','" & DataTab(12) & "','" & DataTab(13) & "','" & DataTab(14) & "','" & DataTab(15) & "','" & DataTab(16) & "' ) ")
SaveNewClientDATA
AddEvent ("Client Ajouter: " & DataTab(5) & " " & DataTab(6) & " " & DataTab(2))
End Function

Function ModifierClient(ID As String)
    Call LoadFilledClientData
    CN.Execute ("UPDATE CLIENTS SET  DateOptPermis ='" & DataTab(0) & "',TypeID ='" & DataTab(1) & "', DateExpID='" & DataTab(3) & "', Sex='" & DataTab(4) & "', Nom='" & DataTab(5) & "',Prenom ='" & DataTab(6) & "', DateNaiss='" & DataTab(7) & "', LieuNaiss='" & DataTab(8) & "', Nationalite='" & DataTab(9) & "', Phone='" & DataTab(10) & "', Email='" & DataTab(11) & "', Adresse='" & DataTab(12) & "', NombreTotalReservation='" & DataTab(13) & "', Statut='" & DataTab(14) & "', NombreInfractions='" & DataTab(15) & "', DateEnreg='" & DataTab(16) & "' WHERE NumID='" & UCase(ID) & "'")
AddEvent ("Client Modifier: " & DataTab(5) & " " & DataTab(6) & " " & DataTab(2))
End Function

Function SupprimerClient(IDS As String)
    Call LoadDbClientData(CStr(IDS))
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM CLIENTS WHERE NumID='" & UCase(IDS) & "' "
    rs.Open SQL, CN, adOpenKeyset
        CN.Execute ("DELETE from CLIENTS where NumID='" & UCase(IDS) & "'")
rs.Close
Set rs = Nothing
AddEvent ("Client Supprimer: " & DataTab(5) & " " & DataTab(6) & " " & DataTab(2))
End Function
'---------------------------------Base donnees----------------------
Function LoadFilledClientData()

   DataTab(0) = FrmClient.DTPickerPermis.Value
'-----------------------------------------------
   If FrmClient.OptionCS.Value = True Then
           DataTab(1) = "Carte Sejour"
        ElseIf FrmClient.OptionPass.Value = True Then
           DataTab(1) = "Passport"
        Else
           DataTab(1) = "CIN"
   End If
'------------------------------------------------
   DataTab(2) = UCase(FrmClient.TextNumID)
'-----------------------------------------------
   DataTab(3) = FrmClient.DTPickerExpID.Value
'-------------------------------------------------
   If FrmClient.OptionF = True Then
           DataTab(4) = "Femme"
        Else
           DataTab(4) = "Homme"
   End If
'------------------------------------------------
   DataTab(5) = UCase(FrmClient.TextNom.Text)
'-----------------------------------------------
   DataTab(6) = UCase(FrmClient.TextPrenom.Text)
'-----------------------------------------------
    DataTab(7) = FrmClient.DTPickerNaiss.Value
'------------------------------------------------
   DataTab(8) = UCase(FrmClient.ComboLieuNaiss)
'------------------------------------------------
   DataTab(9) = UCase(FrmClient.ComboNat)
'--------------------------------------------------
   DataTab(10) = FrmClient.TextNumTel.Text
'------------------------------------------------
   DataTab(11) = FrmClient.TextEmail.Text
'-------------------------------------------------
   DataTab(12) = FrmClient.TextAdress.Text
'------------------------------------------------
   DataTab(13) = FrmClient.TextFid.Text
'------------------------------------------------
   DataTab(14) = FrmClient.TextStatut.Text
'------------------------------------------------
   DataTab(15) = FrmClient.TextInfrac.Text
'----------------------------------------------
   DataTab(16) = FrmClient.DTPickerEnreg
'---------------------------------------------

End Function

Function LoadDbClientData(NumID As String) As Boolean
LoadDbClientData = False
Dim rs As New ADODB.Recordset
SQL = "SELECT * FROM CLIENTS WHERE NumID = '" & UCase(NumID) & "'"
rs.Open SQL, CN, adOpenKeyset
    If Not rs.EOF Then
        DataTab(0) = rs(0)
        DataTab(1) = rs(1)
        DataTab(2) = rs(2)
        DataTab(3) = rs(3)
        DataTab(4) = rs(4)
        DataTab(5) = rs(5)
        DataTab(6) = rs(6)
        DataTab(7) = rs(7)
        DataTab(8) = rs(8)
        DataTab(9) = rs(9)
        DataTab(10) = rs(10)
        DataTab(11) = rs(11)
        DataTab(12) = rs(12)
        DataTab(13) = rs(13)
        DataTab(14) = rs(14)
        DataTab(15) = rs(15)
        DataTab(16) = rs(16)
        LoadDbClientData = True
   End If
rs.Close
Set rs = Nothing
End Function

Function FillFormClient(NumID As String) As Boolean
On Error GoTo Enfer
FillFormClient = False
If LoadDbClientData(CStr(NumID)) = True Then
   'FrmClient.DTPickerPermis.Value = DataTab(0)
   '--------------------------------------------
   If CStr(DataTab(1)) = "CIN" Then
          FrmClient.OptionCIN.Value = True
        ElseIf CStr(DataTab(1)) = "Passport" Then
          FrmClient.OptionPass.Value = True
        Else
          FrmClient.OptionCS.Value = True
    End If
   '--------------------------------------------
   FrmClient.TextNumID.Text = DataTab(2)
   '----------------------------------------------
   FrmClient.DTPickerExpID.Value = DataTab(3)
   '----------------------------------------------
   If DataTab(4) = "Homme" Then
             FrmClient.OptionH.Value = True
        Else
             FrmClient.OptionF.Value = True
   End If
   '------------------------------------------------
   FrmClient.TextNom.Text = DataTab(5)
   '------------------------------------------------
   FrmClient.TextPrenom.Text = DataTab(6)
   '-------------------------------------------------
    FrmClient.DTPickerNaiss.Value = DataTab(7)
   '-------------------------------------------------
    FrmClient.ComboLieuNaiss = DataTab(8)
   '--------------------------------------------------
   FrmClient.ComboNat = DataTab(9)
   '----------------------------------------------
   FrmClient.TextNumTel.Text = DataTab(10)
   '--------------------------------------------------
   FrmClient.TextEmail.Text = DataTab(11)
   '-------------------------------------------------
    FrmClient.TextAdress.Text = DataTab(12)
   '--------------------------------------------------
    FrmClient.TextFid.Text = DataTab(13)
    '-----------------------------------------------
    FrmClient.TextStatut.Text = DataTab(14)
    '------------------------------------------------
    FrmClient.TextInfrac.Text = DataTab(15)
   '--------------------------------------------------
    FrmClient.DTPickerEnreg.Value = DataTab(16)
    FillFormClient = True
End If
Enfer:
End Function

Function ScaleClient()
   'On Error Resume Next
    PageInf = 1
    FrmClient.Width = 7740
    FrmClient.Height = 9450
    FrmClient.CmdPrec1.Top = 8520
    FrmClient.CmdPrec1.Left = FrmClient.Frame1.Left
    FrmClient.CmdSuiv1.Top = FrmClient.CmdPrec1.Top
    FrmClient.CmdSuiv1.Left = 5880
    FrmClient.CmdRaz1.Left = FrmClient.CmdPrec1.Left
    FrmClient.CmdRaz1.Top = FrmClient.CmdPrec1.Top
    FrmClient.CmdEnreg1.Left = FrmClient.CmdSuiv1.Left
    FrmClient.CmdEnreg1.Top = FrmClient.CmdSuiv1.Top
    FrmClient.Frame2.Left = FrmClient.Frame1.Left
    FrmClient.Frame2.Top = FrmClient.Frame1.Top
  FrmClient.ComboLieuNaiss.RemoveItem (0)
  FrmClient.ComboLieuNaiss.SelLength = 0
  FrmClient.ComboNat.RemoveItem (0)
  FrmClient.ComboNat.SelLength = 0
   ShowPageClient (1)
End Function

Function ShowPageClient(Page As Integer)
    'On Error Resume Next
    If Page = 1 Then
        FrmClient.Frame1.Visible = True
        FrmClient.Frame2.Visible = False
        
        FrmClient.CmdPrec1.Visible = False
        FrmClient.CmdSuiv1.Visible = True
        FrmClient.CmdRaz1.Left = FrmClient.CmdPrec1.Left
        FrmClient.CmdEnreg1.Left = FrmClient.CmdSuiv1.Left
        
        PageInf = 1
    ElseIf Page = 2 Then
        FrmClient.Frame1.Visible = False
        FrmClient.Frame2.Visible = True
      
        FrmClient.CmdEnreg1.Visible = True
        FrmClient.CmdPrec1.Visible = True
        FrmClient.CmdSuiv1.Visible = False
        PageInf = 2
    End If
End Function
