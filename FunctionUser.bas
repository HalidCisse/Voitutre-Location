Attribute VB_Name = "FunctionProfile"
Global UserEmail As String
Global UserPass As String
Global UserType As String
Global UserName As String
Global UserPreNom As String


Function ChampsAddUserOk() As Boolean
'-----------------------------------------
   ChampsAddUserOk = True
End Function

'-----------------------Interfaces---------------------------------
Function OpenCherUser()
   If UserType = "ADMIN" Then
debut:
    UserCher = Trim(InputBox("Entrez L'email de l'utilisateur a Chercher :"))
    If UserCher <> "" Then
        If FillFormUser(CStr(UserCher)) Then
            FrmUser.Caption = "Informations de " & DataTab(0)
           'FrmUser.TextEmail.Enabled = True
           FrmUser.ComboType.Enabled = True
           FrmUser.Show vbModal
        Else
          x = MsgBox("Utilisateur Inexistant !!!", vbCritical + vbRetryCancel, "Location Voitures")
          If x = vbRetry Then
             GoTo debut
          End If
        End If
    End If
Else
  MsgBox "Vous devez etre Administrateur pour voir les infos d'un utilisateur", vbInformation + vbOKOnly, "Gestion Location Voiture"
End If
End Function

Function OpenSupUser()
   If UserType = "ADMIN" Then
DebutSup:
        UserSup = Trim(InputBox("Entrer L'email de l'utilisateur a Supprimer "))
        If UserSup <> "" Then
            If LoadDbUserData(CStr(UserSup)) Then
                If UserSup <> UserEmail Then
                   y = MsgBox("Supprimer Cet utilisateur ?  " & " " & DataTab(3) & " " & DataTab(4), vbYesNo + vbQuestion, "Location Voiture" & ": " & DataTab(1))
                   If y = vbYes Then
                     Call SupprimerUser(CStr(UserSup))
                     MsgBox "Suppression réussie", vbOKOnly + vbInformation, "Gestion Voiture : Suppression"
                   End If
                Else
                    MsgBox "Vous Ne Pouvez pas vous supprimer !!", vbCritical + vbOKOnly, "Gestion Location de Voitures"
                End If
            Else
                y = MsgBox("Utilisateur Inéxistant", vbCritical + vbRetryCancel, "Gestion Voiture : Voiture - Suppression")
                If y = vbRetry Then
                  GoTo DebutSup
                End If
           End If
        End If
Else
   MsgBox "Vous devez etre Administrateur pour pouvoir supprimer", vbInformation + vbOKOnly, "Gestion Location Voiture"
End If
End Function

Function OpenModPass()
debut:
     UserCurrentPass = Trim(InputBox("Entrer votre mot de passe actuel : "))
        If UserCurrentPass <> "" Then
             Call LoadDbUserData(UserEmail)
             If UserCurrentPass = UserPass Then
DebutNewPass:
                     NewPass = Trim(InputBox("Entrez votre nouveau mot de passe : "))
                    If NewPass <> "" Then
                           VerNewPass = Trim(InputBox("Reentrez votre nouveau mot de passe : "))
                           If VerNewPass = NewPass Then
                              Call ModUserPass(UserEmail, CStr(NewPass))
                              MsgBox "votre mot de passe a été changé avec succès", vbInformation + vbOKOnly, "Gestion Location de Voitures"
                           Else
                              MsgBox "Incorrect !!", vbInformation + vbOKOnly, "Gestion Location de Voitures"
                              GoTo DebutNewPass
                           End If
                   End If
            Else
                x = MsgBox("Mot de passe Incorrect !!!", vbCritical + vbRetryCancel, "Location Voitures")
                If x = vbRetry Then
                     GoTo debut
                End If
           End If
       End If
End Function

Function OpenModUser()
  If UserType = "ADMIN" Then
debut:
     UserMod = Trim(InputBox("Entrer l'email de l'utilisateur a Modifier "))
     If UserMod <> "" Then
         If FillFormUser(CStr(UserMod)) Then
             FrmUser.Caption = "Modifier Des Informations de " & DataTab(0)
             FrmUser.ComboType.Enabled = True
             FrmUser.Show vbModal
             FrmUser.Hide
        Else
            x = MsgBox("Utilisateur Inexistant !!!", vbCritical + vbRetryCancel, "Location Voitures")
            If x = vbRetry Then
                 GoTo debut
            End If
        End If
    End If
 Else
    Call FillFormUser(CStr(UserEmail))
        FrmUser.Caption = "Modifier Des Informations de " & DataTab(0)
        FrmUser.TextEmail.Enabled = False
        FrmUser.ComboType.Enabled = False
        FrmUser.Show vbModal
        FrmUser.Hide
 End If
End Function

Function OpenAddUser()
   If UserType = "ADMIN" Then
   FrmUser.Caption = "Informations Nouvel utilisateur"
   FrmUser.TextEmail.Enabled = True
   FrmUser.ComboType.Enabled = True
   FrmUser.Show vbModal
Else
  MsgBox "Vous devez etre Administrateur pour pouvoir Ajouter un utilisateur ! ", vbInformation + vbOKOnly, "Gestion Location Voiture"
End If
End Function
'---------------Operations--------------------------------------------
Function AjouterUser()
   Call LoadFilledUserData
   CN.Execute (" INSERT INTO PROFILES VALUES ( '" & DataTab(0) & "','" & DataTab(1) & "','" & DataTab(2) & "','" & DataTab(3) & "','" & DataTab(4) & "' ) ")
AddEvent ("Utilisateur Ajouter: " & DataTab(0) & ", " & DataTab(3) & " " & DataTab(4))
End Function

Function ModifierUser(UserID As String)
    Call LoadFilledUserData
    CN.Execute ("UPDATE PROFILES SET Pass ='" & DataTab(1) & "', UserType='" & DataTab(2) & "', Nom='" & DataTab(3) & "', Prenom='" & DataTab(4) & "'  WHERE Email='" & UserID & "' ")
AddEvent ("Utilisateur Modifier: " & DataTab(0) & ", " & DataTab(3) & " " & DataTab(4))
End Function

Function ModUserPass(UserID As String, NewPass As String)
    CN.Execute ("UPDATE PROFILES SET Pass ='" & NewPass & "'  WHERE Email='" & UserID & "' ")
AddEvent ("Utilisateur A Changer Mot de Pass: " & DataTab(0) & ", " & DataTab(3) & " " & DataTab(4))
End Function

Function SupprimerUser(UserID As String)
LoadDbUserData (CStr(UserID))
    Dim rs As New ADODB.Recordset
    SQL = "SELECT * FROM PROFILES WHERE Email='" & UserID & "' "
    rs.Open SQL, CN, adOpenKeyset
        If Not rs.EOF Then
           CN.Execute ("DELETE FROM PROFILES WHERE Email='" & UserID & "'")
        End If
 rs.Close
 Set rs = Nothing
 AddEvent ("Utilisateur Supprimer: " & DataTab(0) & ", " & DataTab(3) & " " & DataTab(4))
End Function
'-------------------------Base de donnees---------------------------
Function LoadFilledUserData()
    DataTab(0) = FrmUser.TextEmail.Text
    DataTab(1) = FrmUser.TextPass.Text
    DataTab(2) = UCase(FrmUser.ComboType)
    DataTab(3) = UCase(FrmUser.TextNom.Text)
    DataTab(4) = UCase(FrmUser.TextPrenom.Text)
End Function

Function FillFormUser(UserID As String)
FillFormUser = False
If LoadDbUserData(UserID) Then
    FrmUser.TextEmail = DataTab(0)
    FrmUser.TextPass.Text = ""
    FrmUser.TextComfirmPass.Text = ""
    FrmUser.ComboType = DataTab(2)
    FrmUser.TextNom.Text = DataTab(3)
    FrmUser.TextPrenom.Text = DataTab(4)
    FillFormUser = True
End If
End Function

Function LoadDbUserData(UserID As String) As Boolean
On Error Resume Next
LoadDbUserData = False
Dim rs As New ADODB.Recordset
SQL = "SELECT * FROM PROFILES WHERE Email='" & UserID & "'"
rs.Open SQL, CN, adOpenKeyset
    If Not rs.EOF Then
        DataTab(0) = rs(0)
        DataTab(1) = rs(1)
        DataTab(2) = rs(2)
        DataTab(3) = rs(3)
        DataTab(4) = rs(4)
        LoadDbUserData = True
    End If
rs.Close
Set rs = Nothing
End Function
'-----------------------------Connection------------------------------------
Function ConnectUser(UserID As String, Pass As String) As Boolean
ConnectUser = False
Dim rs As New ADODB.Recordset
SQL = "SELECT * FROM PROFILES WHERE Email='" & UserID & "'"
rs.Open SQL, CN, adOpenKeyset
    If Not rs.EOF Then
        If Pass = rs(1) Then
           UserEmail = rs(0)
           UserPass = rs(1)
           UserType = rs(2)
           UserName = rs(3)
           UserPreNom = rs(4)
           Main.mnSeConnecter.Caption = UserEmail
           Main.mnSeConnecter.Enabled = False
           Main.mnSeDeconnecter.Enabled = True
           ConnectUser = True
            If UserType = "ADMIN" Then
               Main.mnGesProfiles.Enabled = True
               Main.mnGesProfiles.Visible = True
            End If
         AddEvent ("Connecter: " & UserEmail)
         Else
           AddEvent (CStr("Erreur de Connection : " & UserID & " "))
         End If
    End If
rs.Close
Set rs = Nothing
End Function

Function DeconnectUser()
AddEvent ("Deconnecter: " & UserEmail)
    UserEmail = ""
    UserType = ""
    UserName = ""
    UserPreNom = ""
    Main.mnSeConnecter.Caption = "Se Connecter"
    Main.mnSeConnecter.Enabled = True
    Main.mnSeDeconnecter.Enabled = False
    Main.mnGesProfiles.Enabled = False
    Main.mnGesProfiles.Visible = False
    Call UnloadForms
    FrmLogin.Show
    Main.Hide
End Function

