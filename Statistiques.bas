Attribute VB_Name = "Statistiques"



'----------------------------- AUTOMATIONS ------------------------
Function UpdateData()
  Call VerifieVIGNETTE
  Call VerifieASSURANCE
  Call VerifieAge
End Function
Function UpdateTileStat()
  Main.LabelTRes.Caption = GetNbreVoitDisp
  Main.LabelTListRes.Caption = GetNbreResEnreg
  Main.LabelTListClient.Caption = GetNbreClientEnreg
  Main.LabelTListVoit.Caption = GetNbreVoitEnreg
End Function
Function UpdateMainStat()
  x = GetResAjourD
  Main.LabelResAjour.Caption = x & " Reservations se terminent aujourd'hui"
  If x = 0 Then
    Main.LabelResAjour.Visible = False
  Else
    Main.LabelResAjour.Visible = True
    Main.LabelResAjour.ForeColor = vbRed
  End If

  x = GetNbreResDPass
  Main.LabelResDelaisPass.Caption = x & " Délais Passée ,Voiture Non Rendue"
  If x = 0 Then
    Main.LabelResAjour.Visible = False
  Else
    Main.LabelResDelaisPass.Visible = True
    Main.LabelResDelaisPass.ForeColor = vbRed
  End If
  
  
  'Main.LabelVoitDisp.Caption = GetNbreVoitDisp & " Voitures Disponibles"
  'Main.LabelResCours.Caption = GetNbreResEnCours & " Réservations En Cours"
  
  'Main.LabelNVoit.Caption = GetNbreVoitEnreg & " Voitures Enrégistrées"
  'Main.LabelNClients.Caption = GetNbreClientEnreg & " Clients Enrégistrées"
  'Main.LabelNRes.Caption = GetNbreResEnreg & " Réservations Enrégistrées"
  
End Function

'----------------------- Get Statistiques ----------------------

Function GetNbreResEnreg() As Integer
  GetNbreResEnreg = NombreEnreg("RESERVATIONS", "", "Tous")
End Function
Function GetNbreResDPass() As Integer
  GetNbreResDPass = NombreEnreg("RESERVATIONS", "SQL", "SELECT COUNT(*) AS Nombre FROM RESERVATIONS WHERE ResStatut = 'Avec Client' AND Cdate(ResFin) < Date()")
End Function
Function GetNbreResEnCours() As Integer
  GetNbreResEnCours = NombreEnreg("RESERVATIONS", "SQL", "SELECT COUNT(*) AS Nombre FROM RESERVATIONS WHERE ResStatut = 'Avec Client'")
End Function
Function GetResAjourD() As Integer
  GetResAjourD = NombreEnreg("RESERVATIONS", "SQL", "SELECT COUNT(*) AS Nombre FROM RESERVATIONS WHERE ResStatut = 'Avec Client' AND ResFin ='" & Date & "'")
End Function
Function GetNbreVoitEnreg() As Integer
  GetNbreVoitEnreg = NombreEnreg("VOITURE", "", "Tous")
End Function
Function GetNbreVoitDisp() As Integer
  GetNbreVoitDisp = NombreEnreg("VOITURE", "STATUT", "Disponible")
End Function
Function GetNbreVoitEnPossession(NumID As String) As Integer
  GetNbreVoitEnPossession = NombreEnreg("RESERVATIONS", "SQL", "SELECT COUNT(NumID) AS Nombre FROM RESERVATIONS WHERE NumID = '" & NumID & "' And ResStatut = 'Avec Client'")
End Function
Function GetNbreClientEnreg() As Integer
  GetNbreClientEnreg = NombreEnreg("CLIENTS", "", "Tous")
End Function

'-------------------------- Remplir Grid -------------------------------------

Function RemplirGridResS(SQL As String)
   Call RemplirGrid(FrmShowRes, FrmShowRes.ResGrid, SQL)
End Function
Function RemplirGridClientS(SQL As String)
   Call RemplirGrid(FrmShowClient, FrmShowClient.ClientGrid, SQL)
End Function
Function RemplirGridVoitS(Statut As String)
   Call RemplirGrid(FrmShowVoit, FrmShowVoit.VoitGrid, "SELECT * FROM VOITURE WHERE STATUT='" & Statut & "'")
End Function

