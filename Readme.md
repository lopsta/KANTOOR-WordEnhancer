Copyright (c) 2021

Nils Kassebohm
Rechtsanwalt und Fachanwalt für Strafrecht
Oxfordstr. 4, 53111 Bonn

Jedem, der eine Kopie dieser Software und der zugehörigen Dokumentationsdateien (die "Software") erhält, wird hiermit kostenlos die Erlaubnis erteilt, ohne Einschränkung mit der Software zu handeln, einschließlich und ohne Einschränkung der Rechte zur Nutzung, zum Kopieren, Ändern, Zusammenführen, Veröffentlichen, Verteilen, Unterlizenzieren und/oder Verkaufen von Kopien der Software, und Personen, denen die Software zur Verfügung gestellt wird, dies unter den folgenden Bedingungen zu gestatten:

Der obige Urheberrechtshinweis und dieser Genehmigungshinweis müssen in allen Kopien oder wesentlichen Teilen der Software enthalten sein.

DIE SOFTWARE WIRD OHNE MÄNGELGEWÄHR UND OHNE JEGLICHE AUSDRÜCKLICHE ODER STILLSCHWEIGENDE GEWÄHRLEISTUNG, EINSCHLIEßLICH, ABER NICHT BESCHRÄNKT AUF DIE GEWÄHRLEISTUNG DER MARKTGÄNGIGKEIT, DER EIGNUNG FÜR EINEN BESTIMMTEN ZWECK UND DER NICHTVERLETZUNG VON RECHTEN DRITTER, ZUR VERFÜGUNG GESTELLT. DIE AUTOREN ODER URHEBERRECHTSINHABER SIND IN KEINEM FALL HAFTBAR FÜR ANSPRÜCHE, SCHÄDEN ODER ANDERE VERPFLICHTUNGEN, OB IN EINER VERTRAGS- ODER HAFTUNGSKLAGE, EINER UNERLAUBTEN HANDLUNG ODER ANDERWEITIG, DIE SICH AUS, AUS ODER IN VERBINDUNG MIT DER SOFTWARE ODER DER NUTZUNG ODER ANDEREN GESCHÄFTEN MIT DER SOFTWARE ERGEBEN.

---

Verwendung:

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