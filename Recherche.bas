Attribute VB_Name = "Recherche"
Public ResChamps As String
Public ResTable As String

Function RemplirItems()
'Remplir les combos des données necessaires
FrmRecherche.ComboRecItem.Clear

  If FrmRecherche.ComboRecType = "Voitures" Then
  '----------------------------------------------------------------
     FrmRecherche.ComboRecItem.AddItem ("Tous")
     FrmRecherche.ComboRecItem.AddItem ("Statut")
     FrmRecherche.ComboRecItem.AddItem ("Matricule")
     FrmRecherche.ComboRecItem.AddItem ("Prix")
     FrmRecherche.ComboRecItem.AddItem ("Type Voiture")
     FrmRecherche.ComboRecItem.AddItem ("Marque")
     FrmRecherche.ComboRecItem.AddItem ("Sous Marque")
     FrmRecherche.ComboRecItem.AddItem ("Modele")
     FrmRecherche.ComboRecItem.AddItem ("Couleur")
     FrmRecherche.ComboRecItem.AddItem ("Type Carburant")
     FrmRecherche.ComboRecItem.AddItem ("Puissance Fiscale")
     FrmRecherche.ComboRecItem.AddItem ("Date Acquis")
     FrmRecherche.ComboRecItem.AddItem ("Date de Mise en Circulation")
     FrmRecherche.ComboRecItem.AddItem ("Nature Achat")
  '-----------------------------------------------------------------
  ElseIf FrmRecherche.ComboRecType = "Clients" Then
    FrmRecherche.ComboRecItem.AddItem ("Tous")
    FrmRecherche.ComboRecItem.AddItem ("Numéro D'identité")
    FrmRecherche.ComboRecItem.AddItem ("Prenom")
    FrmRecherche.ComboRecItem.AddItem ("Nom")
    FrmRecherche.ComboRecItem.AddItem ("Sex")
    FrmRecherche.ComboRecItem.AddItem ("Nombre Total de Reservations")
    FrmRecherche.ComboRecItem.AddItem ("Type Identité")
    FrmRecherche.ComboRecItem.AddItem ("Statut")
    FrmRecherche.ComboRecItem.AddItem ("Date de Naissance")
    FrmRecherche.ComboRecItem.AddItem ("Lieu de Naissance")
    FrmRecherche.ComboRecItem.AddItem ("Nationalité")
    FrmRecherche.ComboRecItem.AddItem ("Numéro de Télephone")
    FrmRecherche.ComboRecItem.AddItem ("Adresse Email")
    FrmRecherche.ComboRecItem.AddItem ("Adresse Domicile")
    FrmRecherche.ComboRecItem.AddItem ("Date Obtention du permis")
    FrmRecherche.ComboRecItem.AddItem ("Nombre d'Infractions")
    FrmRecherche.ComboRecItem.AddItem ("Date Expiration de L'identité")
    FrmRecherche.ComboRecItem.AddItem ("Date de la Prémière Visite")
  '--------------------------------------------------------------------
  ElseIf FrmRecherche.ComboRecType = "Reservations" Then
    FrmRecherche.ComboRecItem.AddItem ("Tous")
    FrmRecherche.ComboRecItem.AddItem ("Numéro De La Reservation")
    FrmRecherche.ComboRecItem.AddItem ("Statut de la Réservation")
    FrmRecherche.ComboRecItem.AddItem ("Numéro ID du Client")
    FrmRecherche.ComboRecItem.AddItem ("Matricule de la Voiture reservée")
    FrmRecherche.ComboRecItem.AddItem ("Date de la Réservation")
    FrmRecherche.ComboRecItem.AddItem ("Debut d'exploitation")
    FrmRecherche.ComboRecItem.AddItem ("Fin d'exploitation")
    FrmRecherche.ComboRecItem.AddItem ("Payment Due")
    FrmRecherche.ComboRecItem.AddItem ("Deja Payé")
    FrmRecherche.ComboRecItem.AddItem ("Rest a payer")
    FrmRecherche.ComboRecItem.AddItem ("Date la Voiture est rendue")
    FrmRecherche.ComboRecItem.AddItem ("Date la voiture est donnée au client")
  End If
FrmRecherche.ComboRecItem.Refresh
End Function

Function DoSQL() As String
'Calculer le SQL coresspondant aux criteres
Call TypeChamps(1)

'------------------------------------------------
  If FrmRecherche.ComboRecType.Text = "Voitures" Then
'------------------------------------------------
     ResTable = "VOITURE"
    If FrmRecherche.ComboRecItem = "Tous" Then
       ResChamps = "MAT"
       Call TypeChamps(0, 1)
       FrmRecherche.LabelSign.Visible = False
       FrmRecherche.Combo.Visible = False
       DoSQL = "SELECT * FROM " & ResTable & ""
    Else
     If FrmRecherche.ComboRecItem = "Matricule" Then
       ResChamps = "MAT"
       Call TypeChamps(1, 1)
     ElseIf FrmRecherche.ComboRecItem = "Date de Mise en Circulation" Then
       ResChamps = "DMCirc"
       Call TypeChamps(1, 1)
     ElseIf FrmRecherche.ComboRecItem = "Marque" Then
       ResChamps = "MARQUE"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Sous Marque" Then
       ResChamps = "S_MARQUE"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Modele" Then
       ResChamps = "MODELE"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Puissance Fiscale" Then
       ResChamps = "P_FISCALE"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Type Carburant" Then
       ResChamps = "T_CARBURANT"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Type Voiture" Then
       ResChamps = "NB_PLACE"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Couleur" Then
       ResChamps = "COULEUR"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Date Acquis" Then
       ResChamps = "Date_Acquis"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Nature Achat" Then
       ResChamps = "NATURE_ACHAT"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Statut" Then
       ResChamps = "STATUT"
       Call TypeChamps(1)
     ElseIf FrmRecherche.ComboRecItem = "Prix" Then
       ResChamps = "Prix"
       Call TypeChamps(1)
     End If
     DoSQL = "SELECT * FROM VOITURE WHERE " & ResChamps & " = '" & UCase(Trim(FrmRecherche.Combo)) & "' "
   End If
'-------------------------------------------------------
  ElseIf FrmRecherche.ComboRecType = "Clients" Then
'-------------------------------------------------------
    ResTable = "CLIENTS"
  If FrmRecherche.ComboRecItem = "Tous" Then
       ResChamps = "NumID"
       Call TypeChamps(0, 1)
       FrmRecherche.LabelSign.Visible = False
       FrmRecherche.Combo.Visible = False
       DoSQL = "SELECT * FROM " & ResTable & ""
  Else
    If FrmRecherche.ComboRecItem = "Date Obtention du permis" Then
       ResChamps = "DateOptPermis"
    ElseIf FrmRecherche.ComboRecItem = "Type Identité" Then
       ResChamps = "TypeID"
    ElseIf FrmRecherche.ComboRecItem = "Numéro D'identité" Then
       ResChamps = "NumID"
    ElseIf FrmRecherche.ComboRecItem = "Date Expiration de L'identité" Then
       ResChamps = "DateExpID"
    ElseIf FrmRecherche.ComboRecItem = "Sex" Then
       ResChamps = "Sex"
    ElseIf FrmRecherche.ComboRecItem = "Prenom" Then
       ResChamps = "Prenom"
    ElseIf FrmRecherche.ComboRecItem = "Nom" Then
       ResChamps = "Nom"
    ElseIf FrmRecherche.ComboRecItem = "Date de Naissance" Then
       ResChamps = "DateNaiss"
    ElseIf FrmRecherche.ComboRecItem = "Lieu de Naissance" Then
       ResChamps = "LieuNaiss"
    ElseIf FrmRecherche.ComboRecItem = "Nationalité" Then
       ResChamps = "Nationalite"
    ElseIf FrmRecherche.ComboRecItem = "Numéro de Télephone" Then
       ResChamps = "Phone"
    ElseIf FrmRecherche.ComboRecItem = "Adresse Email" Then
       ResChamps = "Email"
    ElseIf FrmRecherche.ComboRecItem = "Adresse Domicile" Then
       ResChamps = "Adresse"
    ElseIf FrmRecherche.ComboRecItem = "Nombre Total de Reservations" Then
       ResChamps = "NombreTotalReservation"
    ElseIf FrmRecherche.ComboRecItem = "Statut" Then
       ResChamps = "Statut"
    ElseIf FrmRecherche.ComboRecItem = "Nombre d'Infractions" Then
       ResChamps = "NombreInfractions"
    ElseIf FrmRecherche.ComboRecItem = "Date de la Prémière Visite" Then
       ResChamps = "DateEnreg"
    End If
    DoSQL = "SELECT * FROM CLIENTS WHERE " & ResChamps & " = '" & UCase(Trim(FrmRecherche.Combo.Text)) & "'"
  End If
'----------------------------------------------------------
  ElseIf FrmRecherche.ComboRecType = "Reservations" Then
'----------------------------------------------------------
   ResTable = "RESERVATIONS"
  If FrmRecherche.ComboRecItem = "Tous" Then
       ResChamps = "ResID"
       Call TypeChamps
       FrmRecherche.LabelSign.Visible = False
       FrmRecherche.Combo.Visible = False
       DoSQL = "SELECT * FROM " & ResTable & ""
   Else
    If FrmRecherche.ComboRecItem = "Numéro De La Reservation" Then
       ResChamps = "ResID"
    ElseIf FrmRecherche.ComboRecItem = "Numéro ID du Client" Then
       ResChamps = "NumID"
    ElseIf FrmRecherche.ComboRecItem = "Matricule de la Voiture reservée" Then
       ResChamps = "MAT"
    ElseIf FrmRecherche.ComboRecItem = "Date de la Réservation" Then
       ResChamps = "ResDate"
    ElseIf FrmRecherche.ComboRecItem = "Debut d'exploitation" Then
       ResChamps = "ResDebut"
    ElseIf FrmRecherche.ComboRecItem = "Fin d'exploitation" Then
       ResChamps = "ResFin"
    ElseIf FrmRecherche.ComboRecItem = "Payment Due" Then
       ResChamps = "PayementDue"
    ElseIf FrmRecherche.ComboRecItem = "Deja Payé" Then
       ResChamps = "Paye"
    ElseIf FrmRecherche.ComboRecItem = "Rest a payer" Then
       ResChamps = "RestAPayer"
    ElseIf FrmRecherche.ComboRecItem = "Statut de la Réservation" Then
       ResChamps = "ResStatut"
    ElseIf FrmRecherche.ComboRecItem = "Date la Voiture est rendue" Then
       ResChamps = "DateVoitRendue"
    ElseIf FrmRecherche.ComboRecItem = "Date la voiture est donnée au client" Then
       ResChamps = "DateVoitDClient"
     
    End If
    DoSQL = "SELECT * FROM RESERVATIONS WHERE " & ResChamps & " = '" & UCase(Trim(FrmRecherche.Combo.Text)) & "'"
   End If
  End If
End Function

Function TypeChamps(Optional Sign As Integer = 0, Optional Comb As Integer = 1, Optional DT1 As Integer = 0, Optional DT2 As Integer = 0)
'Changer Dynamique les champs de la form
'-----------------------------------------------
    FrmRecherche.LabelSign.Visible = False
    FrmRecherche.Combo.Visible = True
    FrmRecherche.DTPickerD.Visible = True
    FrmRecherche.DTPickerF.Visible = True
'-----------------------------------------------
  If Sign = 1 Then
    FrmRecherche.LabelSign.Visible = True
    FrmRecherche.LabelSign.Caption = "="
  ElseIf Sign = 2 Then
    FrmRecherche.LabelSign.Visible = True
    FrmRecherche.LabelSign.Caption = "Entre"
  End If
  
  If Comb = 0 Then
    FrmRecherche.Combo.Visible = False
  End If
    
  If DT1 = 0 Then
    FrmRecherche.DTPickerD.Visible = False
  End If
  If DT2 = 0 Then
    FrmRecherche.DTPickerF.Visible = False
  End If
End Function
