Option Compare Database

' Tabelle öffnen, sich auf den ersten Datensatz setzen und dann alle Datensätze durchlaufen
Public Sub DatensaetzeVereinzeln()

' Virtuelle und temporäre Datenbank, sowie 2 tabellen (meine gefüllte TEMP und meine neu zu
' erstellende TEMP_NEU erzeugen
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim rstneu As DAO.Recordset

' Temporäre Variable erzeugen
Dim iAnzahlDS, x As Integer

Set db = CurrentDb
Set rst = db.OpenRecordset("TEMP", dbOpenDynaset)
Set rstneu = db.OpenRecordset("TEMP_NEU", dbOpenDynaset)

' Auf den ersten Datensatz der Tabelle TEMP setzen
rst.MoveFirst

' Alle Datensätze der neuen Tabelle TEMP_NEU löschen (also Tabelle leeren)
DoCmd.RunSQL ("DELETE * from TEMP_NEU")

' Jetzt durchlaufen wir jeden DS von TEMP, bis es keinen mehr gibt
Do While Not rst.EOF
  
  ' Ab hier lesen wir einen Datensatz aus TEMP und schreiben Ihn X-mal neu in TEMP_NEU
  iAnzahlDS = rst!Anzahl.Value ' Dies ist das Feld mit der Anzahl der neu zu erzeugenden
                               ' Datensätze

  ' Schleife wird sooft durchlaufen, wie die Anzahl groß ist
  For x = 1 To iAnzahlDS
    rstneu.AddNew ' Ein neuer Datensatz wird in TEMP_NEU angelegt
    rstneu!lfdnr = rst!lfdnr
    rstneu!lfdnrNEU = x
    rstneu!Name = rst!Name
    rstneu!STRASSEneu = rst!STRASSEneu
    rstneu!PLZORT = rst!PLZORT
    rstneu!Land = rst!Land
    rstneu!Anzahl = rst!Anzahl
    ' Bundende erst wenn letzte lfdnr erreicht
    If x = iAnzahlDS Then
      rstneu!Bundende = "#"
    Else
      rstneu!Bundende = Null
    End If
    rstneu.Update ' Der neue Datensatz wird final in TEMP_NEU geschrieben
  Next x ' Nächsten Datensatz erzeugen bis Menge Anzanzahl erreicht ist
  
  ' Nächster Datensatz in TEMP
  rst.MoveNext

Loop

' Temporäre Database und recordsets wieder löschen
Set rst = Nothing
Set rstneu = Nothing
Set db = Nothing

' MessageBox, dass das Vereinzeln beendet ist
MsgBox ("Habe fertisch!")

End Sub
