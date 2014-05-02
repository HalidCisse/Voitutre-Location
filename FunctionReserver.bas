Attribute VB_Name = "FunctionReserver"


Function ChampsAddResOK() As Boolean
           'verifie
'Si la voiture existe et que elle n est pas en location
'si le client est enregistrer
ChampsAddResOK = False

If LoadDbVoitData(Trim(FrmReserver.ComboMat)) Then  'si la voit est enregistré
    If GetVoitStatut(Trim(FrmReserver.ComboMat)) = "Disponible" Then  'si la voiture est disponible
          If Not LoadDbClientData(Trim(FrmReserver.ComboNumID)) Then
              x = MsgBox("Ce Client n'est pas enrégistrer ! Voulez Vous l'ajouter ? ", vbQuestion + vbYesNo, "Gestion Location Voitures")
              If x = vbYes Then  'Si client n exist pas l ajouter
                 FrmClient.TextNumID.Text = FrmReserver.ComboNumID
                 Call OpenAddClient
                 Call RemplirLesCombosRes
              End If
          Else
              ChampsAddResOK = True
          End If
     Else
       MsgBox "Cette Voiture n'est pas disponible pour locattion : " & GetVoitStatut(Trim(FrmReserver.ComboMat)), vbCritical + vbOKOnly, "Gestion Location de Voitures"
     End If
Else         'si la voiture n existe pas l'ajouter
        x = MsgBox("Cette voiture n'est pas Enrégistrer! Voulez Vous l'ajouter ? ", vbQuestion + vbYesNo, "Gestion Location Voitures")
        If x = vbYes Then
            FrmInformation.TxtMat.Text = FrmReserver.ComboMat
            Call OpenAddVoit
            Call RemplirLesCombosRes
        End If
 End If

End Function

Function OpenModRes()
debut:
  ResMod = InputBox("Entrer le numéro de la réservation a Modifier ")
  If ResMod <> "" Then
    If FillFormReserv(CStr(ResMod)) Then
        FrmClient.Caption = "Modifier Des Informations de " & DataTab(0)
        FrmReserver.Show vbModal
        FrmReserver.Hide
    Else
        x = MsgBox("Réservation Inexistant !!!", vbCritical + vbRetryCancel, "Location Voitures")
        If x = vbRetry Then
             GoTo debut
        End If
    End If
 End If
  
End Function

Function OpenSupRes()
DebutSup:
   ResId = Trim(InputBox("Entrer Le numéro de la reservation a Supprimer "))
   If ResId <> "" Then
    If LoadDbReservData(CStr(ResId)) Then
       y = MsgBox("Supprimer Les Informations de cette reservation ?  " & " " & DataTab(0), vbYesNo + vbQuestion, "Gestion Location Voiture" & ": " & DataTab(0))
       If y = vbYes Then
         Call SupprimerReserv(CStr(ResId))
         MsgBox "Suppression réussie", vbOKOnly + vbInformation, "Gestion Voiture : Suppression"
       End If
    Else
      y = MsgBox("Reservation Inéxistant", vbCritical + vbRetryCancel, "Gestion Voiture : Voiture - Suppression")
      If y = vbRetry Then
        GoTo DebutSup
      End If
    End If
   End If
End Function

Function OpenAddRes()
   Load FrmReserver
   FrmReserver.Caption = "Informations Nouvelle Réservation"
   FrmReserver.OptionAuGar.Value = True
   FrmReserver.TextResID.Text = GenNewResID
   FrmReserver.Show vbModal
End Function

Function OpenCherRes()
debut:
    ResId = InputBox("Entrez Le Numéro de la reservation a Chercher :")
    If ResId <> "" Then
        If FillFormReserv(CStr(ResId)) Then
          FrmReserver.Caption = "Informations de Reservation" & DataTab(0)
          FrmReserver.Show vbModal
        Else
          x = MsgBox("Réservation Inexistant !!!", vbCritical + vbRetryCancel, "Location Voitures")
          If x = vbRetry Then
             GoTo debut
          End If
        End If
    End If
End Function

'-----------------------------------------------------------------------------
Function AjouterReserv(MatID As String)
   Call LoadFilledReservData
   CN.Execute (" INSERT INTO RESERVATIONS VALUES ( '" & DataTab(0) & "','" & DataTab(1) & "','" & DataTab(2) & "','" & DataTab(3) & "','" & DataTab(4) & "','" & DataTab(5) & "','" & DataTab(6) & "','" & DataTab(7) & "','" & DataTab(8) & "','" & DataTab(9) & "','" & DataTab(10) & "','" & DataTab(11) & "' ) ")
   Call ChangeVoitStatut(MatID, "En Location")
AddClientFid (DataTab(1))
AddEvent ("Réservation Ajouter: " & DataTab(0) & ", ClientID = " & DataTab(1) & ", Mat ID = " & DataTab(2))
End Function

Function ModifierReserv(ResId As String)
    Call LoadFilledReservData
    CN.Execute ("UPDATE RESERVATIONS SET NumID ='" & DataTab(1) & "', MAT='" & DataTab(2) & "', ResDebut='" & DataTab(4) & "',ResFin ='" & DataTab(5) & "', PayementDue='" & DataTab(6) & "', Paye='" & DataTab(7) & "', RestAPayer='" & DataTab(8) & "', ResStatut='" & DataTab(9) & "', DateVoitRendue='" & DataTab(10) & "', DateVoitDClient='" & DataTab(11) & "' WHERE ResID='" & UCase(FrmReserver.TextResID.Text) & "' ")
AddEvent ("Réservation Modifier: " & DataTab(0) & ", ClientID = " & DataTab(1) & ", Mat ID = " & DataTab(2))
End Function

Function SupprimerReserv(ResId As String)
LoadDbReservData (CStr(ResId))
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM RESERVATIONS WHERE ResID='" & UCase(ResId) & "' "
    rs.Open SQL, CN, adOpenKeyset
        If Not rs.EOF Then
           CN.Execute ("DELETE FROM RESERVATIONS WHERE ResID='" & UCase(ResId) & "'")
        End If
 rs.Close
 Set rs = Nothing
AddEvent ("Réservation Supprimer: " & DataTab(0) & ", ClientID = " & DataTab(1) & ", Mat ID = " & DataTab(2))
End Function

'----------------Data conection-----------------------------

Function LoadFilledReservData()
'remplir la table datatab des données de la form reserver
Call LoadDbReservData(UCase(FrmReserver.TextResID.Text))
    DataTab(0) = UCase(FrmReserver.TextResID.Text)
    DataTab(1) = UCase(FrmReserver.ComboNumID)
    DataTab(2) = UCase(FrmReserver.ComboMat)
    DataTab(3) = Date
    DataTab(4) = FrmReserver.DTPickerDebExp.Value
    DataTab(5) = FrmReserver.DTPickerFinExp.Value
    DataTab(6) = FrmReserver.TextPayTotal.Text
    DataTab(7) = FrmReserver.TextPay.Text
    DataTab(8) = FrmReserver.LabelResPay.Caption
    If FrmReserver.OptionAuGar.Value = True Then
            DataTab(9) = "Au Garage"
            DataTab(10) = ""
            DataTab(11) = ""
        ElseIf FrmReserver.OptionAvecClient.Value = True Then
            DataTab(9) = "Avec Client"
            'DataTab(10) = ""
            DataTab(11) = FrmReserver.DTPickerVoitRendu.Value
        ElseIf FrmReserver.OptionRendu.Value = True Then
            DataTab(9) = "Rendu"
            DataTab(10) = FrmReserver.DTPickerVoitRendu.Value
            'DataTab(11) = ""
    End If
End Function

Function FillFormReserv(ResId As String)
'On Error Resume Next
'Remplir la form "reserver" des données de la base de donnée
FillFormReserv = False
If LoadDbReservData(ResId) Then
    FrmReserver.TextResID.Text = DataTab(0)
    FrmReserver.ComboNumID = DataTab(1)
    FrmReserver.ComboMat = DataTab(2)
    FrmReserver.DTPickerDebExp.Value = DataTab(4)
    FrmReserver.DTPickerFinExp.Value = DataTab(5)
    FrmReserver.TextPayTotal.Text = DataTab(6)
    FrmReserver.TextPay.Text = DataTab(7)
    FrmReserver.LabelResPay.Caption = DataTab(8)
    If DataTab(9) = "Au Garage" Then
            FrmReserver.OptionAuGar.Value = True
            'FrmReserver.DTPickerVoitRendu.Value = Date
        ElseIf DataTab(9) = "Avec Client" Then
            FrmReserver.OptionAvecClient.Value = True
            FrmReserver.DTPickerVoitRendu.Value = DataTab(11)
        ElseIf DataTab(9) = "Rendu" Then
            FrmReserver.OptionRendu.Value = True
            FrmReserver.DTPickerVoitRendu.Value = DataTab(10)
    End If
    FillFormReserv = True
End If
End Function

Function LoadDbReservData(ResId As String) As Boolean
'Mettre les données de la base de donnée sur la table datatab
On Error Resume Next
LoadDbReservData = False
Dim rs As New ADODB.Recordset
SQL = "SELECT * FROM RESERVATIONS WHERE ResID='" & UCase(ResId) & "'"
rs.Open SQL, CN, adOpenKeyset
    If Not rs.EOF Then
        For i = 0 To 11
          DataTab(i) = rs(i)
        Next
        LoadDbReservData = True
    End If
rs.Close
Set rs = Nothing
End Function

Function GenNewResID() As String
'Generez un numéro de reservation
rep:
    random = Int((100 - 1 + 1) * Rnd + 1)
   NewResID = "RES" & Mid(Year(Date), 3, 2) & Month(Date) & Mid(Trim(FrmReserver.ComboNumID), 1, 2) & Mid(Trim(FrmReserver.ComboMat), 1, 3) & random
   If LoadDbReservData(CStr(NewResID)) Then
      GoTo rep
   Else
      GenNewResID = UCase(NewResID)
   End If
End Function
