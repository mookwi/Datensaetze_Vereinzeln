Option Compare Database

' Tabelle öffnen, sich auf den ersten Datensatz setzen und dann alle Datensätze durchlaufen
Public Sub DatensaetzeVereinzeln()

' Virtuelle und temporäre Datenbank, sowie 2 Tabellen (meine gefüllte TEMP und meine neu zu
' erstellende TEMP_NEU erzeugen.
' Tabelle TEMP
' FELD            BEMERKUNG
' lfdnr           - laufende, eindeutige Nummer
' Name            - Name
' STRAASEneu      - Strasse
' PLZORT          - beinhaltet PLZ und Ort in einem Feld
' Land            - Bei DE leer, ansonsten Landesbezeichnung
' Anzahl          - Menge, wie oft der Datensatz dupliziert werden soll
'
' Ich habe die bestehende Tabelle TEMP kopiert und in TEMP_NEU umbenannt. Dann habe ich noch
' folgende Felder hinzugefügt:
' lfdnrNEU        - Ist eine fortlaufende Nummerierung des duplizierten Datensatzes und beginnt immer mit 1
' Bundende        - wird mit einem '#' gefüllt, wenn es sich um den letzten duplizierten Datensatz handelt

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
  iAnzahlDS = rst!Anzahl.Value             ' Dies ist das Feld mit der Anzahl der neu zu erzeugenden
                                           ' Datensätze

  ' Schleife wird sooft durchlaufen, wie die Menge im Feld 'Anzahl' groß ist
  For x = 1 To iAnzahlDS
    rstneu.AddNew                          ' Ein neuer Datensatz wird in TEMP_NEU angelegt
    rstneu!lfdnr = rst!lfdnr               ' Bestehendes Feld in TEMP_NEU kopieren
    rstneu!lfdnrNEU = x                    ' Neues Feld mit aktuell duplizierten Datensatz, zum besseren sortieren
    rstneu!Name = rst!Name                 ' Bestehendes Feld in TEMP_NEU kopieren
    rstneu!STRASSEneu = rst!STRASSEneu     ' Bestehendes Feld in TEMP_NEU kopieren
    rstneu!PLZORT = rst!PLZORT             ' Bestehendes Feld in TEMP_NEU kopieren
    rstneu!Land = rst!Land                 ' Bestehendes Feld in TEMP_NEU kopieren
    rstneu!Anzahl = rst!Anzahl             ' Bestehendes Feld in TEMP_NEU kopieren
   
    ' Bundende '#' erst wenn letzte lfdnr erreicht. Wenn wir ein Bundanfang '#' benötigen,
    ' dann müssen wir im Code iAnzahlDS durch eine '1' ersetzen!
    If x = iAnzahlDS Then
      rstneu!Bundende = "#"
    Else
      rstneu!Bundende = Null
    End If
 
    rstneu.Update                           ' Der neue Datensatz wird final in TEMP_NEU geschrieben
  Next x                                    ' Nächsten Datensatz erzeugen bis Menge Anzahl erreicht ist
  
  ' Nächster Datensatz in TEMP aufrufen
  rst.MoveNext

Loop

' Temporäre Database und recordsets wieder löschen
Set rst = Nothing
Set rstneu = Nothing
Set db = Nothing

' MessageBox, dass das Vereinzeln beendet ist
MsgBox ("Habe fertisch!")

' Danach habe ich eine neue Abfrage auf die Tabelle 'TEMP_NEU' erstellt, bei der ich zuerst nach dem
' Feld 'lfdnr' (aufsteigend) und dann nach dem Feld 'lfdnrNEU' (aufsteigend) sortiert habe, damit das
' Bundende Zeichen '#' auch garantiert auf dem letzten duplizierten Datensatz liegt.

End Sub
