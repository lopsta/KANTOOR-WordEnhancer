Attribute VB_Name = "lopstaTextGenderToggler"
' ================================================================
' Copyright (c) 2021
'
' Nils Kassebohm
' Rechtsanwalt und Fachanwalt f�r Strafrecht
' Oxfordstr. 4, 53111 Bonn
'
' Jedem, der eine Kopie dieser Software und der zugeh�rigen Dokumentationsdateien (die "Software") erh�lt, wird hiermit kostenlos die Erlaubnis erteilt,
' ohne Einschr�nkung mit der Software zu handeln, einschlie�lich und ohne Einschr�nkung der Rechte zur Nutzung, zum Kopieren, �ndern, Zusammenf�hren,
' Ver�ffentlichen, Verteilen, Unterlizenzieren und/oder Verkaufen von Kopien der Software, und Personen, denen die Software zur Verf�gung gestellt wird,
' dies unter den folgenden Bedingungen zu gestatten:
'
' Der obige Urheberrechtshinweis und dieser Genehmigungshinweis m�ssen in allen Kopien oder wesentlichen Teilen der Software enthalten sein.
'
' DIE SOFTWARE WIRD OHNE M�NGELGEW�HR UND OHNE JEGLICHE AUSDR�CKLICHE ODER STILLSCHWEIGENDE GEW�HRLEISTUNG, EINSCHLIE�LICH,
' ABER NICHT BESCHR�NKT AUF DIE GEW�HRLEISTUNG DER MARKTG�NGIGKEIT, DER EIGNUNG F�R EINEN BESTIMMTEN ZWECK UND DER NICHTVERLETZUNG VON RECHTEN DRITTER,
' ZUR VERF�GUNG GESTELLT. DIE AUTOREN ODER URHEBERRECHTSINHABER SIND IN KEINEM FALL HAFTBAR F�R ANSPR�CHE, SCH�DEN ODER ANDERE VERPFLICHTUNGEN,
' OB IN EINER VERTRAGS- ODER HAFTUNGSKLAGE, EINER UNERLAUBTEN HANDLUNG ODER ANDERWEITIG, DIE SICH AUS, AUS ODER IN VERBINDUNG MIT DER SOFTWARE
' ODER DER NUTZUNG ODER ANDEREN GESCH�FTEN MIT DER SOFTWARE ERGEBEN.
'
' ================================================================

' ................................................................
' Verwendung:
'
' Das WORD-VBA-Modul dient dazu, Textinhalte vereinfacht an eine
' m�nnliche oder weibliche Formulierung anzupassen.
'
' Beispiel aus der Anwaltspraxis:
'
' [...] zeige ich an, dass mich Herr XXX mit seiner Verteidigung beauftragt hat.
' [...] zeige ich an, dass mich Frau YYY mit ihrer Verteidigung beauftragt hat.
'
' Routineaufgaben in einer forensisch ausgerichteten Strafverteidigerkanzlei
' lassen sich mit verschiedenen Musterschreiben (z.B. Bestellung als Verteidiger,
' Absage einer Beschuldigtenvorladung bei der Polizei, Akteneinsichtsantrag,
' Einspruch gegen einen Strafbefehl, Einspruch gegen einen Bu�gelbescheid, Haftpr�fungsantrag
' usw.) effizent und schnell erledigen. In diesen Schreiben ist der Inhalt im wesentlichen
' gleich un bedarf keiner oder nur geringer Anpassungen. Formulierungsunterschiede
' im Sinne des oben aufgezeigten Beispiels k�nnen sich in einem Musterschreiben
' jedoch an vielen Stellen ergeben. Das Musterschreiben bei seiner Verwendung
' per Hand anzupassen, ist zeitaufwendig, fehleranf�llig und h�lt auf.
'
' Word verf�gt leider �ber keine Funktion die es erm�glicht, Textvorlagen so zu erstellen,
' dass sich Anpassungen an m�nnliche und weibliche Formulierungen automatisieren lassen.
'
' lopstaKANTOOR[WordEnhancer] f�gt diese Funktion Word hinzu. Dazu muss lediglich die Datei
' lopstaKANTOOR[WordEnhancer]v0-0-0.dotm in das Verzeichnis
'     c:\Users\<IhrUserName>\AppData\Roaming\Microsoft\Word\STARTUP
' kopiert werden.
'
' Oder es kann der nachstehende Quelltext in ein VBA-Modul �bernommen werden.
'
' �ber einen Word-Ribbon-Button wird die Methode toggleTextGender(Optional control As IRibbonControl)
' aufgerufen. Der aufrufende Button muss als Control �bergeben werden. Der Button f�r die
' m�nnliche Form muss die ID lopstaButton201 und der f�r die weibliche Form die
' Id lopstaButton202 haben.
'
' Zun�chst wird gepr�ft, ob alle anzupassenden Textstellen bereits in einer Liste
' erfasst wurden. Dies ist deshalb erforderlich, weil andernfalls nach einer
' einmaligen Anpassung die Markierungen verloren sind und eine erneute Anpassung -
' z.B. bei versehentlicher Falschauswahl - nicht mehr m�glich w�re. Dies wird in der Mehtode
' TextGenderBookmarksGrabber() erledigt. Die Liste der Textstellen werden in der Collection
' lopstaTextGenderBookmarks gespeichert. Da die Textstellen bei der �nderung verloren gehen,
' wird die Liste nur beim erstmaligen Aufruf erstellt und speicher den Urzustand der Felder.
'
' Die Anpassung erfolgt in der Methode TextGenderToggler(Optional control As IRibbonControl).
' An die Methode muss das Button Control weitergereicht werden. Die Unterscheidung zwischen
' m�nnlicher und weiblicher Form findet anhand der Button Id statt. Nachdem der Text angepasst
' (ausgelesen aus der Liste) wurde, wird der Textverweis neu gesetzt um weitere
' Anpassungen zu erm�glichen.
'
' In der Word-Vorlage (oder Word-Datei) die angepasst werden soll, werden die Textstellen
' wie folgt eingef�gt:
'
'      m�nnliche Form|weibliche Form
'
' Damit die Stelle im text erkannt und ge�ndert werden kann, muss sie markiert und als Textmarke
' (Ribbon Einf�gen/Links/Textmarke) unter dem Namen
'
'      Gender001
'
' gespeichert werden.
'
' ................................................................


' ================================
' Collection Objekt in dem
' alle gefundenen Gender-Bookmarks
' gespeichert werden. Die Textanpassungen
' erfolgen aus den gespeicherten Objekten.
' ================================
Private lopstaTextGenderBookmarks As Collection


' ================================
' Flag mit der festgehakten wird,
' ob die Collection der Bookmarks
' bereits erstellt wurde.
' Die Collection soll nur einmal
' erstellt werden, um den Ausgangszustand
' der Textinhalte zu speichern.
' ================================
Private lopstaTextGenderBookmarksGrabbedState As Boolean


' ================================
' Mehtode zu Einstieg in die Textanpassung.
' Nur diese Methode kann und soll
' aufgerufen werden.
' ================================
Public Sub toggleTextGender(Optional control As IRibbonControl)
    
    On Error Resume Next
    
    If Not lopstaTextGenderBookmarksGrabbedState Then
        Call TextGenderBookmarksGrabber
    End If
    
    Call TextGenderToggler(control)
    
End Sub


' ================================
' Methode liest alle Gender-Bookmarks
' aus und speichert sie in der Collection.
' ================================
Private Sub TextGenderBookmarksGrabber()

    On Error Resume Next
    
    Set lopstaTextGenderBookmarks = New Collection
    
    lopstaTextGenderBookmarksGrabbedState = False

    ' .....................................
    ' RegEx Pattern f�r die Gender Bookmarks erzeugen
    ' .....................................
    
    Dim regex As New RegExp
    regex.Pattern = "Gender\d*"
    
    ' .....................................
    ' Zun�chst m�ssen alle vorhandenen Bookmarks in einer Liste gespeichert werden.
    ' Die im Text gesetzten Bookmarks gehen verloren, wenn die Text.Range �berarbeitet wird.
    ' Mit der Liste k�nnen die Bookmarks neu gesetzt werden.
    ' .....................................
    
    Dim bkm As Bookmark
    
    For Each bkm In ActiveDocument.Bookmarks
        
        ' Initialisieren eines neues Speicher-Objektes
        Dim h As lopstaClassTextGenderConserver
        Set h = New lopstaClassTextGenderConserver
        
        ' Zerlegen der gefundenen Bookmarks in die Gender-Bestandteile
        If regex.Test(bkm.Name) Then
        
            Dim bkmTextParts() As String
            bkmTextParts = Split(bkm.Range.Text, "|") ' Teilt den Text der Bookmark in einen vorderen und hinteren Teil
            bkmTextParts(0) = Replace(bkmTextParts(0), "�", "") ' Herausl�schen der Anf�hrungszeichen
            bkmTextParts(0) = Replace(bkmTextParts(0), "�", "")
            bkmTextParts(1) = Replace(bkmTextParts(1), "�", "")
            bkmTextParts(1) = Replace(bkmTextParts(1), "�", "")
            
            With h
                .Bookmark = bkm.Name
                .OriginalBookmarkText = bkm.Range.Text
                .Male = bkmTextParts(0)
                .Female = bkmTextParts(1)
            End With
            
            ' Speichern des neuen Objektes in der Collection
            lopstaTextGenderBookmarks.Add Item:=h

        End If
                
    Next
    
    ' Flag setzen, die verhindert, dass Gender-Bookmarks erneut ausgelsen
    ' und die Collection mit dem Urzustand �berschrieben wird.
    lopstaTextGenderBookmarksGrabbedState = True

End Sub


' ================================
' Einf�gen der angepassten Form
' aus der Collection
' ================================
Private Sub TextGenderToggler(Optional control As IRibbonControl)

    On Error Resume Next
    
    ' .....................................
    ' Bearbeitung der Bookmarks
    ' .....................................
    
    Dim i As lopstaClassTextGenderConserver
    
    For Each i In lopstaTextGenderBookmarks
    
        On Error Resume Next
        
        If ActiveDocument.Bookmarks.Exists(i.Bookmark) Then
            
            Dim rng As Range
            Set rng = ActiveDocument.Bookmarks(i.Bookmark).Range
                
            Select Case control.ID
                Case Is = "lopstaButton201"
                    rng.Text = i.Male
                Case Is = "lopstaButton202"
                    rng.Text = i.Female
                End Select
                
            ActiveDocument.Bookmarks.Add Name:=i.Bookmark, Range:=rng
            
        End If
    
    Next
    
End Sub
