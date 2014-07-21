Attribute VB_Name = "DataReport"

Function OpenReportRES(ResId As String)
  
  Dim rs As New ADODB.Recordset
  SQL = "SELECT NumID FROM RESERVATIONS WHERE ResID = '" & ResId & "'"
  rs.Open SQL, CN, 3, 3
  NumID = rs(0)
  Set DRRes.DataSource = rs
  rs.Close
  SQL = "SELECT MAT FROM RESERVATIONS WHERE ResID = '" & ResId & "'"
  rs.Open SQL, CN, 3, 3
  Mat = rs(0)
  rs.Close
  SQL = "SELECT * FROM EntrepriseINFO "
  rs.Open SQL, CN, 3, 3
  
  '-------- - Entreprise Data  ---------------------------
  DNameEnt = CStr(rs(0))  '"OSIS Location De Voiture"
  DTelEnt = CStr(rs(1))        '"N° Tel : 06854785 / Fax : 06525845"
  DAdressEnt = CStr(rs(2))    '"Rue 652 ,Imm X , Agdal , Rabat"
  TTVA = Val(rs(5))
  
  DRRes.Sections("Section4").Controls("NameEnt").Caption = DNameEnt
  DRRes.Sections("Section4").Controls("TelEnt").Caption = Format(DTelEnt, "###-##-##-###")
  DRRes.Sections("Section4").Controls("AdressEnt").Caption = DAdressEnt
  
  '------------- Reservation Data -------------------------
  LoadDbReservData (ResId)
  DResID = ResId '   'DataTab(1)
  DVoitMat = DataTab(2)
  DResDeb = CDate(DataTab(4))
  DResfin = CDate(DataTab(5))
  DNumJour = DateDiff("d", DResDeb, DResfin)
  
  DRRes.Sections("Section4").Controls("LabelResID").Caption = "Numéro de Reservation : " & DResID
  DRRes.Sections("Section1").Controls("VoitMat").Caption = DVoitMat
  DRRes.Sections("Section1").Controls("ResDeb").Caption = Format(DResDeb, "ddd dd-mmm-yyyy")
  DRRes.Sections("Section1").Controls("Resfin").Caption = Format(DResfin, "ddd  dd-mmm-yyyy")
  DRRes.Sections("Section1").Controls("NumJour").Caption = DNumJour
  DRRes.Sections("Section5").Controls("LabelDate").Caption = Format(Date, "ddd dd-mmm-yyyy") & "  " & Format(Time, "h:nn")
  DRRes.Sections("Section5").Controls("LabelResIDFooter").Caption = DResID
  
  '---------------- Client data ---------------------------
  LoadDbClientData (NumID)
  DClientName = DataTab(5)
  DClientPrenom = DataTab(6)
  DClientNat = DataTab(9)
  DClientTel = DataTab(10)
  DClientAdress = DataTab(12)
  DClientNPS = DClientName & " " & DClientPrenom
  
  DRRes.Sections("Section1").Controls("ClientName").Caption = DClientName
  DRRes.Sections("Section1").Controls("ClientPrenom").Caption = DClientPrenom
  DRRes.Sections("Section1").Controls("ClientNat").Caption = DClientNat
  DRRes.Sections("Section1").Controls("ClientTel").Caption = Format(DClientTel, "########-##-##-###")
  DRRes.Sections("Section1").Controls("ClientAdress").Caption = DClientAdress
  DRRes.Sections("Section1").Controls("ClientNPS").Caption = DClientNPS
  
  '---------------- Voiture data  --------------------------
  LoadDbVoitData (Mat)
  DVoitPrixJour = DataTab(21)
  DResTHT = Val(DNumJour) * Val(DVoitPrixJour)
  DTVA = Val(DResTHT * TTVA / 100)
  DTOTAL = Val(DResTHT) + Val(DTVA)
  DUserNomPrenomS = UserName & " " & UserPreNom
  
  DRRes.Sections("Section1").Controls("VoitPrixJour").Caption = DVoitPrixJour
  DRRes.Sections("Section1").Controls("ResTHT").Caption = DResTHT
  DRRes.Sections("Section1").Controls("LabelTVA").Caption = "TVA à " & TTVA & " %"
  DRRes.Sections("Section1").Controls("TVA").Caption = DTVA
  DRRes.Sections("Section1").Controls("TOTAL").Caption = DTOTAL
  DRRes.Sections("Section1").Controls("UserNomPrenomS").Caption = DUserNomPrenomS
  DRRes.Show vbModal
  
End Function

