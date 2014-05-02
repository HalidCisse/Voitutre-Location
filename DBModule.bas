Attribute VB_Name = "DBModule"
Global DataTab(21) As String
Global PageInf As Integer
Global CN As New ADODB.Connection

Function OpenCN(fdb As String) As Boolean
Set CN = New ADODB.Connection
'CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fdb
CN.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & fdb & ";Persist Security Info=False;Jet OLEDB:Database Password= " & "halid" & ""
CN.Open
    If CN.State = adStateOpen Then
       OpenCN = True
    Else
       OpenCN = False
    End If
End Function

Function ConnectDB()
    Dim DB As String
      DB = App.Path
    If Right(DB, 1) <> "\" Then
      DB = DB & "\"
    End If
      DB = DB & "LOC_VT.mdb"
    If OpenCN(DB) = False Then
      MsgBox "Impossible d'établir une connexion avec Gestion Location de Voitures", vbCritical, "Gestion Location de Voitures"
      End
    Else
      'MsgBox "Connexion - Gestion Location de Voitures - Effectuée ", vbInformation, "Gestion Location de Voitures"
    End If
    currentdb = App.Path & "\LOC_VT.mdb"
End Function
'------------------------########################-------------------------
'------------------------########################-------------------------





Function AddEvent(Message As String)
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM EVENTS"
    rs.Open SQL, CN, adOpenKeyset
    CN.Execute ("INSERT INTO EVENTS (UserEmail,Message,EventTime,EventDate) VALUES ('" & UserEmail & "','" & Message & "','" & Time & "','" & Date & "' ) ")
 rs.Close
 Set rs = Nothing
End Function

Function NombreEnreg(Table As String, Champs As String, Statut As String) As Integer
'On Error Resume Next
Dim rs As New ADODB.Recordset
    If Champs = "SQL" Then
      SQL = Statut
    ElseIf Statut = "Tous" Then
      SQL = "SELECT COUNT(*) AS Nombre FROM " & Table
    Else
      SQL = "SELECT COUNT(*) AS Nombre FROM " & Table & " WHERE " & Champs & " = '" & Statut & "'"
    End If
rs.Open SQL, CN, adOpenKeyset
        NombreEnreg = Val(rs(0))
rs.Close
Set rs = Nothing
End Function

'#################### AUTOCOMPLETE DATASAVE ##########################################

Function SaveNewVoitDATA()
  Call AddNewDATA("MARQUE", UCase(FrmInformation.ComboMarque))
  Call AddNewDATA("SOUSMARQUES", UCase(FrmInformation.ComboSousMarque))
  Call AddNewDATA("MODELE", UCase(FrmInformation.ComboModele))
  Call AddNewDATA("PuissanceFiscale", UCase(FrmInformation.ComboPuissanceFiscale))
  Call AddNewDATA("COULEUR", UCase(FrmInformation.ComboCouleur))
End Function
Function SaveNewClientDATA()
  Call AddNewDATA("NATIONALITES", UCase(FrmClient.ComboNat))
  Call AddNewDATA("VILLES", UCase(FrmClient.ComboLieuNaiss))
End Function
Function AddNewDATA(Table As String, Value As String)
'permet d'ajouter un enregistrement au donnée pour autocompleter pendant la saisie
Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM " & Table & " WHERE ( " & Table & " = '" & Value & "')"
    rs.Open SQL, CN, adOpenKeyset
    If rs.EOF Then
      CN.Execute ("INSERT INTO " & Table & " VALUES ( '" & Value & "' ) ")
    End If
rs.Close
Set rs = Nothing
End Function

'###################### MIS A JOUR AUTOMATIQUE ################################################################

Function VerifieVIGNETTE()
'cette fonction verifie si la vignette est valide sinon elle modifie le statut en vignette invalide
    Dim rs As New ADODB.Recordset
    SQL = "SELECT MAT FROM VOITURE WHERE STATUT = 'Disponible' and ANNEE_VIGNETTE < '" & Year(Date) & "'"
    rs.Open SQL, CN, adOpenKeyset
      For i = 0 To rs.RecordCount - 1
        Call ChangeVoitStatut(rs(i), "Vignette Invalide")
      Next i
    rs.Close
    Set rs = Nothing
End Function
Function VerifieASSURANCE()
'cette fonction verifie si l'assurance est valide sinon elle modifie le statut en assurance invalide
    Dim rs As New ADODB.Recordset
    SQL = "SELECT MAT FROM VOITURE WHERE STATUT = 'Disponible' and Cdate(DATE_FIN_ASSURANCE) < Date()"     ''" & Date & "'"
    rs.Open SQL, CN, adOpenKeyset
      For i = 0 To rs.RecordCount - 1
        Call ChangeVoitStatut(rs(i), "Assurance Invalide")
      Next i
    rs.Close
    Set rs = Nothing
End Function
Function VerifieAge()
'cette fonction verifie si la age est valide sinon elle modifie le statut en indisponile
  Dim rs As New ADODB.Recordset
    SQL = "SELECT MAT FROM VOITURE WHERE STATUT = 'Disponible' and DateDiff('yyyy',DMCirc,Date()) > 5 "
    rs.Open SQL, CN, adOpenKeyset
      For i = 0 To rs.RecordCount - 1
        Call ChangeVoitStatut(rs(i), "Indisponible")
      Next
    rs.Close
    Set rs = Nothing
End Function

'###################### MSHFGRID ###############################

Function RemplirGrid(Form As Form, Grid As MSHFlexGrid, SQL As String)
On Error Resume Next
     Load Form
     Dim rs As New ADODB.Recordset
     rs.Open SQL, CN, adOpenKeyset
     Set Grid.DataSource = rs
     Form.Show vbModal
End Function

'####################### Combos ########################################

Function RempirLesCombosVoit()
    Call Remplir(FrmInformation.ComboMarque, "MARQUE", , 0)
    Call Remplir(FrmInformation.ComboSousMarque, "SOUSMARQUES", , 0)
    Call Remplir(FrmInformation.ComboModele, "MODELE", , 0)
    Call Remplir(FrmInformation.ComboPuissanceFiscale, "PuissanceFiscale", , 0)
    Call Remplir(FrmInformation.ComboCouleur, "COULEUR", , 0)
End Function
Function RemplirLesCombosClient()
     Call Remplir(FrmClient.ComboNat, "NATIONALITES", , 0)
     Call Remplir(FrmClient.ComboLieuNaiss, "VILLES", , 0)
End Function
Function RemplirLesCombosRes()
     Call Remplir(FrmReserver.ComboNumID, "CLIENTS", , 2)
     Call RemplirComboVoitDisp
End Function

Function RemplirComboVoitDisp()
FrmReserver.ComboMat.AddItem ("")
    Dim rs As New ADODB.Recordset
    SQL = "SELECT MAT FROM VOITURE WHERE STATUT= '" & "Disponible" & "' "
    rs.Open SQL, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not rs.EOF
        FrmReserver.ComboMat.AddItem (rs.Fields(0))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Function
 Function Remplir(Combo As ComboBox, Table As String, Optional NomChamps As String = "", Optional NumChamps As Integer = 0)
'On Error Resume Next
Combo.Clear
Combo.AddItem ("")
    Dim rs As New ADODB.Recordset
    If NomChamps <> "" Then
      SQL = "SELECT DISTINCT " & NomChamps & " FROM " & Table & ""
      NumChamps = 0
    Else
      SQL = "SELECT DISTINCT * FROM " & Table & ""
    End If
    rs.Open SQL, CN, adOpenDynamic, adLockOptimistic, adCmdText
    Do While Not rs.EOF
        Combo.AddItem (rs.Fields(NumChamps))
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
End Function

 Function AutoComplete(cboName As ComboBox)
On Error Resume Next
    Dim strPart As String, iLoop As Integer, iStart As Integer, strItem As String
        'don't do if no text or if change was made by autocomplete coding
        If Not blnAuto And cboName.Text <> "" Then
            'save the selection start point (cursor position)
            iStart = cboName.SelStart
            'get the part the user has typed (not selected)
            strPart = Left$(cboName.Text, iStart)
            For iLoop = 0 To cboName.ListCount - 1
                'compare each item to the part the user has typed,
                '"complete" with the first good match
                strItem = UCase$(cboName.List(iLoop))
                If strItem Like UCase$(strPart & "*") And _
                        strItem <> UCase$(cboName.Text) Then
                    'partial match but not the whole thing.
                    '(if whole thing, nothing to complete!)
                    blnAuto = True
                    cboName.SelText = Mid$(cboName.List(iLoop), iStart + 1) 'add on the new ending
                    cboName.SelStart = iStart   'reset the selection
                    cboName.SelLength = Len(cboName.Text) - iStart
                    blnAuto = False
                    Exit For
                End If
            Next iLoop
            'Add statement here like FilterRecord or whatever :)
        End If
End Function

'###################### Forms ##################################################3

Function LoadForms()
    Load Main
    Load FrmInformation
    Load FrmReserver
    Load FrmClient
    Load FrmUser
    Load FrmRecherche
    
    Load FrmShowVoit
    Load FrmShowUser
    Load FrmShowClient
    Load FrmShowRes
End Function

Function UnloadForms()
   Unload FrmClient
   Unload FrmInformation
   Unload FrmReserver
   Unload FrmUser
   Unload FrmRecherche
   
   Unload FrmShowClient
   Unload FrmShowRes
   Unload FrmShowVoit
   Unload FrmShowUser
End Function

