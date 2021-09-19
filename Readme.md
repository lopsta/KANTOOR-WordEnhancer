# Copyright (c) 2021

Nils Kassebohm
Rechtsanwalt und Fachanwalt für Strafrecht
Oxfordstr. 4, 53111 Bonn

Jedem, der eine Kopie dieser Software und der zugehörigen Dokumentationsdateien (die "Software") erhält, wird hiermit kostenlos die Erlaubnis erteilt, ohne Einschränkung mit der Software zu handeln, einschließlich und ohne Einschränkung der Rechte zur Nutzung, zum Kopieren, Ändern, Zusammenführen, Veröffentlichen, Verteilen, Unterlizenzieren und/oder Verkaufen von Kopien der Software, und Personen, denen die Software zur Verfügung gestellt wird, dies unter den folgenden Bedingungen zu gestatten:

Der obige Urheberrechtshinweis und dieser Genehmigungshinweis müssen in allen Kopien oder wesentlichen Teilen der Software enthalten sein.

DIE SOFTWARE WIRD OHNE MÄNGELGEWÄHR UND OHNE JEGLICHE AUSDRÜCKLICHE ODER STILLSCHWEIGENDE GEWÄHRLEISTUNG, EINSCHLIEßLICH, ABER NICHT BESCHRÄNKT AUF DIE GEWÄHRLEISTUNG DER MARKTGÄNGIGKEIT, DER EIGNUNG FÜR EINEN BESTIMMTEN ZWECK UND DER NICHTVERLETZUNG VON RECHTEN DRITTER, ZUR VERFÜGUNG GESTELLT. DIE AUTOREN ODER URHEBERRECHTSINHABER SIND IN KEINEM FALL HAFTBAR FÜR ANSPRÜCHE, SCHÄDEN ODER ANDERE VERPFLICHTUNGEN, OB IN EINER VERTRAGS- ODER HAFTUNGSKLAGE, EINER UNERLAUBTEN HANDLUNG ODER ANDERWEITIG, DIE SICH AUS, AUS ODER IN VERBINDUNG MIT DER SOFTWARE ODER DER NUTZUNG ODER ANDEREN GESCHÄFTEN MIT DER SOFTWARE ERGEBEN.

---

# Verwendung:

Das WORD-VBA-Modul dient dazu, Textinhalte vereinfacht an eine männliche oder weibliche Formulierung anzupassen.

Beispiel aus der Anwaltspraxis:

    [...] zeige ich an, dass mich Herr XXX mit seiner Verteidigung beauftragt hat.
    [...] zeige ich an, dass mich Frau YYY mit ihrer Verteidigung beauftragt hat.

Routineaufgaben in einer forensisch ausgerichteten Strafverteidigerkanzlei lassen sich mit verschiedenen Musterschreiben (z.B. Bestellung als Verteidiger, Absage einer Beschuldigtenvorladung bei der Polizei, Akteneinsichtsantrag, Einspruch gegen einen Strafbefehl, Einspruch gegen einen Bußgelbescheid, Haftprüfungsantrag usw.) effizent und schnell erledigen. In diesen Schreiben ist der Inhalt im wesentlichen gleich un bedarf keiner oder nur geringer Anpassungen. Formulierungsunterschiede im Sinne des oben aufgezeigten Beispiels können sich in einem Musterschreiben jedoch an vielen Stellen ergeben. Das Musterschreiben bei seiner Verwendung per Hand anzupassen, ist zeitaufwendig, fehleranfällig und hält auf.

Word verfügt leider über keine Funktion die es ermöglicht, Textvorlagen so zu erstellen, dass sich Anpassungen an männliche und weibliche Formulierungen automatisieren lassen.

lopstaKANTOOR[WordEnhancer] fügt diese Funktion Word hinzu. Dazu muss lediglich die Datei 
    
    lopstaKANTOOR[WordEnhancer]v0-0-0.dotm 

in das Verzeichnis

    c:\Users\<IhrUserName>\AppData\Roaming\Microsoft\Word\STARTUP
    
kopiert werden.

Oder es kann der nachstehende Quelltext in ein VBA-Modul übernommen werden.

Über einen Word-Ribbon-Button wird die Methode toggleTextGender(Optional control As IRibbonControl) aufgerufen. Der aufrufende Button muss als Control übergeben werden. Der Button für die männliche Form muss die ID lopstaButton201 und der für die weibliche Form die Id lopstaButton202 haben.

Zunächst wird geprüft, ob alle anzupassenden Textstellen bereits in einer Liste erfasst wurden. Dies ist deshalb erforderlich, weil andernfalls nach einer einmaligen Anpassung die Markierungen verloren sind und eine erneute Anpassung - z.B. bei versehentlicher Falschauswahl - nicht mehr möglich wäre. Dies wird in der Mehtode TextGenderBookmarksGrabber() erledigt. Die Liste der Textstellen werden in der Collection lopstaTextGenderBookmarks gespeichert. Da die Textstellen bei der Änderung verloren gehen, wird die Liste nur beim erstmaligen Aufruf erstellt und speicher den Urzustand der Felder.

Die Anpassung erfolgt in der Methode TextGenderToggler(Optional control As IRibbonControl). An die Methode muss das Button Control weitergereicht werden. Die Unterscheidung zwischen männlicher und weiblicher Form findet anhand der Button Id statt. Nachdem der Text angepasst (ausgelesen aus der Liste) wurde, wird der Textverweis neu gesetzt um weitere Anpassungen zu ermöglichen.

In der Word-Vorlage (oder Word-Datei) die angepasst werden soll, werden die Textstellen wie folgt eingefügt:

    «männliche Form|weibliche Form»

Damit die Stelle im text erkannt und geändert werden kann, muss sie markiert und als Textmarke (Ribbon Einfügen/Links/Textmarke) unter dem Namen

    Gender001

gespeichert werden.

# Ausprobieren - Testen

### Datei nach AppData kopieren
Um den lopstaKANTOOR[WordEnhancer] auszuprobieren und zu test, kopieren Sie bitte die Datei

    lopstaKANTOOR[WordEnhancer]beta_v0-0-2.dotm

in Ihr Verzeichnis

    c:\Users\<IhrUserName>\AppData\Roaming\Microsoft\Word\STARTUP

Der »AppData« Ordner wird von Windows standardmäßig dem Benutzer nicht angezeigt. Ggfls. müssen Sie die Anzeige versteckter Ordner und Dateien erst einschalten. Zu Ihrem »AppData« Ordner gelangen Sie jedoch sehr einfach.

Drücken Sie die Tastenkombination [Windows]+[R] und geben Sie %appdata% ein und bestätigen mit der Entertaste. Oder Sie geben den Suchbegriff einfach in das Suchfeld von Windows 10 ein. Danach navigieren Sie im Ordner »Roaming« über den Ordner »Microsoft« zum Ordner »Word«. Ist dort noch kein Ordner »STARTUP« vorhanden, erstellen Sie ihn bitte.

Sie können auch eine Verknüpfung erstellen und anstelle der Datei in den Ordner »c:\Users\<IhrUserName>\AppData\Roaming\Microsoft\Word\STARTUP« verschieben.

### Word starten
Danach starten Sie Ihr Word Programm. Nach dem Start erscheint in der Ribbon-Auswahl (Menü oben) ein neuer Eintrag »lopstaKANTOOR«. Dort finden Sie die Button mit den hinzugefügten lopsta Funktionen.

### Briefbogen öffnen
Doppelklicken Sie auf die Datei

    musterBriefbogen.dotx

um ein neues Dokument mit der Wordvorlage »musterBriefbogen.dotx« zu erstellen.

### Mustertext einfügen
Wählen Sie den im Ribbon den Reiter »Einfügen«. Dort finden Sie einen Auswahlbutton »Objekt«. Wählen Sie »Text aus Datei« und fügen Sie bitte die Datei

    MusterTextbaustein.dotx

ein. Im Text sehen Sie nun Einfügungen wie z.B. «mein Mandant|meine Mandantin».

### Form an männliche oder weibliche Formulierung anpassen
Um den Text an die gewünschte männlich oder weibliche Form anzupassen, wählen Sie im Ribbon den Reiter »lopstaKANTOOR« und dort im Feld »Anpassen« den auf die gewünschte Form zutreffenden Button. Nach dem Klick erscheint im Text nur noch die männliche oder die weibliche Form. Haben Sie versehentlich den falschen Button geklickt, klicken Sie einfach erneut auf den Button der richtigen Form und der Text wird erneut angepasst.

Ich wünsche Ihnen gutes Gelingen und hoffe, dass Sie Gefallen an meiner Programmierung haben.

Ihr
Nils Kassebohm
Rechtsanwalt
Fachanwalt für Strafrecht