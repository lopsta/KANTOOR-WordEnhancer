VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "lopstaClassTextGenderConserver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' ================================================================
' Copyright (c) 2021
'
' Nils Kassebohm
' Rechtsanwalt und Fachanwalt für Strafrecht
' Oxfordstr. 4, 53111 Bonn
'
' Jedem, der eine Kopie dieser Software und der zugehörigen Dokumentationsdateien (die "Software") erhält, wird hiermit kostenlos die Erlaubnis erteilt,
' ohne Einschränkung mit der Software zu handeln, einschließlich und ohne Einschränkung der Rechte zur Nutzung, zum Kopieren, Ändern, Zusammenführen,
' Veröffentlichen, Verteilen, Unterlizenzieren und/oder Verkaufen von Kopien der Software, und Personen, denen die Software zur Verfügung gestellt wird,
' dies unter den folgenden Bedingungen zu gestatten:
'
' Der obige Urheberrechtshinweis und dieser Genehmigungshinweis müssen in allen Kopien oder wesentlichen Teilen der Software enthalten sein.
'
' DIE SOFTWARE WIRD OHNE MÄNGELGEWÄHR UND OHNE JEGLICHE AUSDRÜCKLICHE ODER STILLSCHWEIGENDE GEWÄHRLEISTUNG, EINSCHLIEßLICH,
' ABER NICHT BESCHRÄNKT AUF DIE GEWÄHRLEISTUNG DER MARKTGÄNGIGKEIT, DER EIGNUNG FÜR EINEN BESTIMMTEN ZWECK UND DER NICHTVERLETZUNG VON RECHTEN DRITTER,
' ZUR VERFÜGUNG GESTELLT. DIE AUTOREN ODER URHEBERRECHTSINHABER SIND IN KEINEM FALL HAFTBAR FÜR ANSPRÜCHE, SCHÄDEN ODER ANDERE VERPFLICHTUNGEN,
' OB IN EINER VERTRAGS- ODER HAFTUNGSKLAGE, EINER UNERLAUBTEN HANDLUNG ODER ANDERWEITIG, DIE SICH AUS, AUS ODER IN VERBINDUNG MIT DER SOFTWARE
' ODER DER NUTZUNG ODER ANDEREN GESCHÄFTEN MIT DER SOFTWARE ERGEBEN.
'
' ================================================================

Private intBookmark As String
Private intOriginalBookmarkText As String
Private intMale As String
Private intFemale As String

Public Property Let Bookmark(v As String)
    intBookmark = v
End Property

Public Property Get Bookmark() As String
    Bookmark = intBookmark
End Property

Public Property Let OriginalBookmarkText(v As String)
    intOriginalBookmarkText = v
End Property

Public Property Get OriginalBookmarkText() As String
    OriginalBookmarkText = intOriginalBookmarkText
End Property

Public Property Let Male(v As String)
    intMale = v
End Property

Public Property Get Male() As String
    Male = intMale
End Property

Public Property Let Female(v As String)
    intFemale = v
End Property

Public Property Get Female() As String
    Female = intFemale
End Property
