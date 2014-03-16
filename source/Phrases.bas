Attribute VB_Name = "Phrases"
Option Explicit

'Public function LoadPhrases() will read phrases from a phrase file if the
'language is set to 8, which can be achieved by entering the cheat code
'"MR#icre8" or "MR#testing122" and selecting "Load Phrase File" from the
'Languages option box. The file name is "Phrases.txt" which is set in the
'global constant gcPhraseFileName.
'
'To help testing, entering the cheat code "MR#testing122" and changing the
'language to Spanish will display the number and position of each phrase.

'File containing alternative translations.
Public Const gcPhraseFileName As String = "Phrases.txt"

'The maximum number of phrases.
Private Const cMaxPhrases = 450

Public Phrase(cMaxPhrases) As String
Public gLanguage As Long         'english; Italian; = 0; 1; etc

Public Enum eLanguage
    '<Language Name> = &H<Language ID>
    PhraseFile = &H98
    PhraseNumbers = &H94
    English = &H9
    French = &HC
    Spanish = &HA
    Italian = &H10
    Swedish = &H1D
    Dutch = &H13
    Brazilian = &H16
    Finnish = &HB
    Norwegian = &H14
    Danish = &H6
    Hungarian = &HE
    Polish = &H15
    Russian = &H19
    Czech = &H5
    Greek = &H8
    Portuguese = &H16
    Turkish = &H1F
    Japanese = &H11
    Korean = &H12
    German = &H7
    Chinese_Simplified = &H4
    Chinese_Traditional = &H4
    Arabic = &H1
    Hebrew = &HD
End Enum


'Load phrases from file.
Private Sub ReadPhraseFile()
    Dim PhrasePathName As String
    Dim strInPut As String
    Dim fileNo As Integer
    
    PhrasePathName = App.Path & "\" & gcPhraseFileName
    If Dir(PhrasePathName) <> "" Then
        fileNo = FreeFile
        Open PhrasePathName For Binary As #fileNo
        strInPut = Space(LOF(1))
        Get #fileNo, , strInPut
        Close #fileNo
        Call BreakTranslatedString(strInPut)
    Else
        'File not found.
        gLanguage = eLanguage.English
        Call InitialisePhrases
    End If
End Sub

'Read file contents into phrase strings
Private Sub BreakTranslatedString(phrasePart As String)
    Dim where As Long
    Dim startPhrase As Long
    Dim endPhrase As Long
    Dim startComment As Long
    Dim totalPhrase As Long
    Dim cntr As Long
    Dim tmp As Long
    
    On Error GoTo errhand_1
    
    'Replace VARs.
    'phrasePart = Replace(phrasePart, "<Var.Min>", str(App.Minor))
    'phrasePart = Replace(phrasePart, "<Var.Maj>", str(App.Major))
    'phrasePart = Replace(phrasePart, "<Var.Mission14>", str(gcMission14))
    
    'Count total phrases by counting how many English phrases there are
    where = 1
    totalPhrase = UBound(Phrase)
    where = 1
    For cntr = 0 To totalPhrase - 1
        startPhrase = where
        where = InStr(where + 2, phrasePart, "> - Translation")
        If where = 0 Then
            Exit For
        End If
        startPhrase = InStr(where, phrasePart, Chr(34)) + 1   'Point to first quote marks (")
        endPhrase = InStr(startPhrase, phrasePart, Chr(34))   'Start of next language's phrase
        If endPhrase = 0 Then
            Exit For
        End If
        Phrase(cntr) = Mid(phrasePart, startPhrase, endPhrase - startPhrase)
    Next
    On Error GoTo errhand_2
    Exit Sub
errhand_1:
    tmp = MsgBox("Error reading " & App.Path & "\" & gcPhraseFileName & vbCrLf _
                & "Phrase " & cntr & vbCrLf _
                & Err.Description, vbOKOnly, "Translation error")
    gLanguage = eLanguage.English
    Call InitialisePhrases
    Exit Sub
errhand_2:
    Resume Next
End Sub

'Replace tokens such as "<Var.ExeName>" with actual values.
'Handled by ref but the string is also returned for object incopatabilities.
Public Function SubstituteStringTokens(ByRef pString As String) As String
    Dim vIndex As Long
    Dim vPhraseIX As Long
    
    'Arbitrary tokens.
    pString = Replace(pString, "<Var.ExeName>", gcApplicationName)
    pString = Replace(pString, "<Var.HomePage>", gcDefaultHomePageClearURL)
    pString = Replace(pString, "<Var.Maj>", Trim(str(App.Major)))
    pString = Replace(pString, "<Var.Min>", Trim(str(App.Minor)))
    pString = Replace(pString, "<Var.Rev>", Trim(Format(App.Revision, "0000")))
    pString = Replace(pString, "<Var.Mission14>", Trim(str(gcMission14)))
    
    'Any phrase within a phrase. Not too efficient but it is rare that
    'phrases are changed. <Var.Phrase(100)> will be replaced with Phrase(100).
    'The outer loop will take care of phrase tokens with phrase tokens within.
    For vIndex = 0 To 10
        If InStr(1, pString, "<Var.Phrase(") > 0 Then
            For vPhraseIX = 0 To UBound(Phrase)
                pString = Replace(pString, "<Var.Phrase(" & CStr(vPhraseIX) & ")>", Phrase(vPhraseIX))
            Next
        Else
            Exit For
        End If
    Next
    SubstituteStringTokens = pString
End Function

'Load all phrases of the selected language into the Phrase() array.
Public Sub LoadPhrases()
    Dim vIndex As Long
    
    If gLanguage = eLanguage.PhraseFile Then
        Call ReadPhraseFile
    Else
        Call InitialisePhrases
    End If
    
    For vIndex = 0 To UBound(Phrase)
        Call SubstituteStringTokens(Phrase(vIndex))
    Next
End Sub

'English. Always fill in English so that it can be the default if
'an invalid language ID is chosen.
Private Sub sPhraseEng(pIndex As Long, pPhrase As String)
    If gLanguage = eLanguage.PhraseNumbers Then
        Phrase(pIndex) = "<" & CStr(pIndex) & ">"
    Else
        Phrase(pIndex) = pPhrase
    End If
End Sub

'Italian
Private Sub sPhraseIta(pIndex As Long, pPhrase As String)
    If gLanguage = eLanguage.Italian Then
        Phrase(pIndex) = pPhrase
    End If
End Sub

'French
Private Sub sPhraseFra(pIndex As Long, pPhrase As String)
    If gLanguage = eLanguage.French Then
        Phrase(pIndex) = pPhrase
    End If
End Sub

'German
Private Sub sPhraseGer(pIndex As Long, pPhrase As String)
    If gLanguage = eLanguage.German Then
        Phrase(pIndex) = pPhrase
    End If
End Sub

'Spanish or numeric text for testing if in testing mode (MR#Testing122).
Private Sub sPhraseSpa(pIndex As Long, pPhrase As String)
    If gLanguage = eLanguage.Spanish Then
        Phrase(pIndex) = pPhrase
    End If
End Sub

'Swedish
Private Sub sPhraseSwe(pIndex As Long, pPhrase As String)
    If gLanguage = eLanguage.Swedish Then
        Phrase(pIndex) = pPhrase
    End If
End Sub

'Norwegian
Private Sub sPhraseNor(pIndex As Long, pPhrase As String)
    If gLanguage = eLanguage.Norwegian Then
        Phrase(pIndex) = pPhrase
    End If
End Sub

'Danish
Private Sub sPhraseDan(pIndex As Long, pPhrase As String)
    If gLanguage = eLanguage.Danish Then
        Phrase(pIndex) = pPhrase
    End If
End Sub

'---------------- Use the Phrase Editor to modify the code below ------------------------
'To use the phrase editor, copy the Phrases module (phrases.bas) into the Phrase Editor's
'home directory and run the program. It will determine which part of the phrases file is
'code and which part contains actual phrases. The Phrase Editor can also be used to create
'and read language files that can be sent to translators without having to send any source
'code.
Private Sub InitialisePhrases()
    'REM:
    sPhraseEng 0, "Player"
    sPhraseIta 0, "Giocatore"
    sPhraseFra 0, "Joueur"
    sPhraseGer 0, "Spieler"
    sPhraseSpa 0, "Jugador"
    sPhraseSwe 0, "Spelare"
    sPhraseNor 0, "Spiller"
    sPhraseDan 0, "Spiller"

    'REM:
    sPhraseEng 1, "The Red Army "
    sPhraseIta 1, "Armate rosse "
    sPhraseFra 1, "L'Armée Rouge "
    sPhraseGer 1, "Die rote Armee "
    sPhraseSpa 1, "Ejército Rojo "
    sPhraseSwe 1, "Den röda armén "
    sPhraseNor 1, "Den Røde Hæren "
    sPhraseDan 1, "Den røde hær"

    'REM:
    sPhraseEng 2, "The Green Army "
    sPhraseIta 2, "Armate verdi "
    sPhraseFra 2, "L'Armée Verte"
    sPhraseGer 2, "Die grüne Armee "
    sPhraseSpa 2, "Ejército Verde "
    sPhraseSwe 2, "Den gröna armén "
    sPhraseNor 2, "Den Grønne Hæren "
    sPhraseDan 2, "Den grønne hær"

    'REM: _

    sPhraseEng 3, "The Blue Army "
    sPhraseIta 3, "Armate blu "
    sPhraseFra 3, "L'Armée Bleue "
    sPhraseGer 3, "Die blaue Armee "
    sPhraseSpa 3, "Ejército Azul "
    sPhraseSwe 3, "Den blåa armén "
    sPhraseNor 3, "Den Blå Hæren "
    sPhraseDan 3, "Den blå hær"

    'REM:
    sPhraseEng 4, "The Yellow Army "
    sPhraseIta 4, "Armate gialle "
    sPhraseFra 4, "L'Armée Jaune "
    sPhraseGer 4, "Die gelbe Armee "
    sPhraseSpa 4, "Ejército Amarillo "
    sPhraseSwe 4, "Den gula armén "
    sPhraseNor 4, "Den Gule Hæren "
    sPhraseDan 4, "Den gule hær"

    'REM:
    sPhraseEng 5, "The Purple Army "
    sPhraseIta 5, "Armate viola "
    sPhraseFra 5, "L'Armée Pourpre "
    sPhraseGer 5, "Die lila Armee "
    sPhraseSpa 5, "Ejército Púrpura "
    sPhraseSwe 5, "Den lila armén "
    sPhraseNor 5, "Den Lilla Hæren "
    sPhraseDan 5, "Den lilla hær"

    'REM:
    sPhraseEng 6, "The Gray Army "
    sPhraseIta 6, "Armate grigie "
    sPhraseFra 6, "L'Armée Grise "
    sPhraseGer 6, "Die graue Armee "
    sPhraseSpa 6, "Ejército Gris "
    sPhraseSwe 6, "Den gråa armén "
    sPhraseNor 6, "Den Grå Hæren "
    sPhraseDan 6, "Den grå hær"

    'REM:
    sPhraseEng 7, "Human player"
    sPhraseIta 7, "Giocatore umano"
    sPhraseFra 7, "Joueur humain"
    sPhraseGer 7, "Menschlicher Spieler "
    sPhraseSpa 7, "Jugador Humano"
    sPhraseSwe 7, "Mänsklig spelare"
    sPhraseNor 7, "Menneske"
    sPhraseDan 7, "Menneske"

    'REM: Italian: Giocatoore computer: intelligenza media -Fra: L'ordinateur: l'intelligence moyenne.
    sPhraseEng 8, "Average computer player"
    sPhraseIta 8, "Computer: intelligenza media"
    sPhraseFra 8, "Joueur ordinateur moyen"
    sPhraseGer 8, "Durchschnittlicher Computergegner"
    sPhraseSpa 8, "Jugador Computadora Nivel Promedio"
    sPhraseSwe 8, "Normal datorspelare"
    sPhraseNor 8, "Computer (normal)"
    sPhraseDan 8, "Computer (normal)"

    'REM: Italian: Giocatore computer: molto intelligente -Fra: L'ordinateur: la haute intelligence.
    sPhraseEng 9, "Smart computer player"
    sPhraseIta 9, "Computer: molto intelligente"
    sPhraseFra 9, "Joueur ordinateur intelligent"
    sPhraseGer 9, "Schlauer Computergegner"
    sPhraseSpa 9, "Jugador Computadora Nivel Sagaz"
    sPhraseSwe 9, "Smart datorspelare"
    sPhraseNor 9, "Computer (smart)"
    sPhraseDan 9, "Computer (smart)"

    'REM:
    sPhraseEng 10, "Players"
    sPhraseIta 10, "Giocatori"
    sPhraseFra 10, "Joueurs"
    sPhraseGer 10, "Spieler"
    sPhraseSpa 10, "Jugadores"
    sPhraseSwe 10, "Spelare"
    sPhraseNor 10, "Spillere"
    sPhraseDan 10, "Spillere"

    'REM:
    sPhraseEng 11, "First player"
    sPhraseIta 11, "Primo giocatore"
    sPhraseFra 11, "Premier joueur"
    sPhraseGer 11, "Erster Spieler"
    sPhraseSpa 11, "Primer Jugador"
    sPhraseSwe 11, "Första spelaren"
    sPhraseNor 11, "Første spiller"
    sPhraseDan 11, "Første spiller"

    'REM:
    sPhraseEng 12, "Random"
    sPhraseIta 12, "Casuale"
    sPhraseFra 12, "Aléatoire"
    sPhraseGer 12, "Zufall"
    sPhraseSpa 12, "Azar"
    sPhraseSwe 12, "Slumpmässig"
    sPhraseNor 12, "Vilkårlig"
    sPhraseDan 12, "Tilfældig"

    'REM:
    sPhraseEng 13, "Player 1"
    sPhraseIta 13, "Giocatore 1"
    sPhraseFra 13, "Joueur 1"
    sPhraseGer 13, "Spieler 1"
    sPhraseSpa 13, "Jugador Uno"
    sPhraseSwe 13, "Spelare 1"
    sPhraseNor 13, "Spiller 1"
    sPhraseDan 13, "Spiller 1"

    'REM:
    sPhraseEng 14, "Cards"
    sPhraseIta 14, "Carte"
    sPhraseFra 14, "Cartes"
    sPhraseGer 14, "Karten"
    sPhraseSpa 14, "Tarjetas"
    sPhraseSwe 14, "Kort"
    sPhraseNor 14, "Kort"
    sPhraseDan 14, "Kort"

    'REM:
    sPhraseEng 15, "Hidden"
    sPhraseIta 15, "Nascoste"
    sPhraseFra 15, "Cachées"
    sPhraseGer 15, "Versteckt"
    sPhraseSpa 15, "Oculto"
    sPhraseSwe 15, "Dolda"
    sPhraseNor 15, "Skjult"
    sPhraseDan 15, "Skjul kort"

    'REM:
    sPhraseEng 16, "None"
    sPhraseIta 16, "Nessuna"
    sPhraseFra 16, "Aucune"
    sPhraseGer 16, "Keine"
    sPhraseSpa 16, "Ninguno"
    sPhraseSwe 16, "Inget"
    sPhraseNor 16, "Ingen"
    sPhraseDan 16, "Ingen"

    'REM: Fixed cards.
    sPhraseEng 17, "Fixed"
    sPhraseIta 17, "Fisso"
    sPhraseFra 17, "Valeur fixée"
    sPhraseGer 17, "Fix"
    sPhraseSpa 17, "Fijo"
    sPhraseSwe 17, "Bestämt"
    sPhraseNor 17, "Statisk"
    sPhraseDan 17, "Uændret"

    'REM: Maximum value.
    sPhraseEng 18, "Maximum"
    sPhraseIta 18, "Valore Massimo"
    sPhraseFra 18, "Maximum"
    sPhraseGer 18, "Maximum"
    sPhraseSpa 18, "Maximo"
    sPhraseSwe 18, "Maximum"
    sPhraseNor 18, "Maks."
    sPhraseDan 18, "Maksimalt"

    'REM: English changed from "Battle options"
    sPhraseEng 19, "Supply lines"
    sPhraseIta 19, "Opzioni"
    sPhraseFra 19, "Options de guerre"
    sPhraseGer 19, "Kriegsoptionen"
    sPhraseSpa 19, "Opciones de guerra"
    sPhraseSwe 19, "Alternativ för krig"
    sPhraseNor 19, "Krigføring"
    sPhraseDan 19, "Opsætning"

    'REM:Available missions.
    sPhraseEng 20, "Missions"
    sPhraseIta 20, "Obiettivi"
    sPhraseFra 20, "Missions"
    sPhraseGer 20, "Missionen"
    sPhraseSpa 20, "Misión"
    sPhraseSwe 20, "Uppdrag"
    sPhraseNor 20, "Oppdrag"
    sPhraseDan 20, "Missioner"

    'REM: English changed to from Limit supply lines.
    sPhraseEng 21, "Limited"
    sPhraseIta 21, "Limita i rifornimenti"
    sPhraseFra 21, "Lignes de provision limitées"
    sPhraseGer 21, "Versorgungslinien limitieren"
    sPhraseSpa 21, "Líneas de abastecimiento limitadas"
    sPhraseSwe 21, "Begränsat understöd"
    sPhraseNor 21, "Forsyningsstøtte"
    sPhraseDan 21, "Begrænset flytning"

    'REM:
    sPhraseEng 22, "Optimise defence dice"
    sPhraseIta 22, "Ottimizza i dadi di difesa"
    sPhraseFra 22, "Optimiser les dés de défense"
    sPhraseGer 22, "Verteidigung optimieren"
    sPhraseSpa 22, "Dados de defensa óptimos"
    sPhraseSwe 22, "Optimera försvarstärningarna"
    sPhraseNor 22, "Begrens forsyningsstøtte"
    sPhraseDan 22, "Optimalt forsvar"

    'REM:
    sPhraseEng 23, "Fast war"
    sPhraseIta 23, "Guerra veloce"
    sPhraseFra 23, "Guerre rapide"
    sPhraseGer 23, "Schneller Krieg"
    sPhraseSpa 23, "Guerra rápida"
    sPhraseSwe 23, "Snabba slag"
    sPhraseNor 23, "Hurtig krig"
    sPhraseDan 23, "Hurtig krig"

    'REM:
    sPhraseEng 24, "Fast Dice"
    sPhraseIta 24, "Dadi veloci"
    sPhraseFra 24, "Dés rapides"
    sPhraseGer 24, "Schnelles Würfeln"
    sPhraseSpa 24, "Dados rápidos"
    sPhraseSwe 24, "Snabba tärningsslag"
    sPhraseNor 24, "Hurtige terninger"
    sPhraseDan 24, "Hurtige terninger"

    'REM:Frame shows color of player:French:Couleur de joueur des spectacles de frontière
    sPhraseEng 25, "Border shows player's color"
    sPhraseIta 25, "Il bordo mostra il giocatore di turno"
    sPhraseFra 25, "Cadre de la couleur du joueur"
    sPhraseGer 25, "Farbe des Spielers als Rand anzeigen"
    sPhraseSpa 25, "El borde muestra el color del jugador"
    sPhraseSwe 25, "Ramen visar spelarens färg"
    sPhraseNor 25, "Spillers farge på ramme"
    sPhraseDan 25, "Kant viser spillers farve"

    'REM:
    sPhraseEng 26, "&Declare War"
    sPhraseIta 26, "&Dichiara guerra"
    sPhraseFra 26, "Déclarez la Guerre"
    sPhraseGer 26, "&Krieg erklären"
    sPhraseSpa 26, "&Declare la Guerra"
    sPhraseSwe 26, "Förklara krig"
    sPhraseNor 26, "Erklær krig"
    sPhraseDan 26, "Erklær krig"

    'REM: &Cancel set up
    sPhraseEng 27, "&Resume War"
    sPhraseIta 27, "&Cancella il" + vbCrLf _
                 + "set up"
    sPhraseFra 27, "Annulez"
    sPhraseGer 27, "&Abbrechen"
    sPhraseSpa 27, "Cancelar activa"
    sPhraseSwe 27, "Avbryt"
    sPhraseNor 27, "Avbryt oppsett"
    sPhraseDan 27, "Fortryd"

    'REM: Remote player
    sPhraseEng 28, "Unclaimed"
    sPhraseIta 28, "Giocatore remoto"
    sPhraseFra 28, "Joueur distant"
    sPhraseGer 28, "Netzwerk-Spieler"
    sPhraseSpa 28, "Jugador remoto"
    sPhraseSwe 28, "Fjärrspelare"
    sPhraseNor 28, "Nettverksspiller"
    sPhraseDan 28, "Netværksspiller"

    'REM:
    sPhraseEng 29, "Join &War"
    sPhraseIta 29, "Unisciti alla &Guerra"
    sPhraseFra 29, "Rejoignez la Guerre"
    sPhraseGer 29, "Krieg &beitreten"
    sPhraseSpa 29, "Unir &guerra"
    sPhraseSwe 29, "Gå med i krig"
    sPhraseNor 29, "Delta i krigen"
    sPhraseDan 29, "Deltag i krig"

    'REM: English changed from "Supply lines"
    sPhraseEng 30, "Unlimited"
    sPhraseIta 30, "Rifornimenti"
    sPhraseFra 30, "Lignes de provision"
    sPhraseGer 30, "Versorgungslinien"
    sPhraseSpa 30, "Líneas de reaprovisionamiento"
    sPhraseSwe 30, "Understöd"
    sPhraseNor 30, "Forsyninger"
    sPhraseDan 30, "Ubegrænset flytning"

    'REM:
    sPhraseEng 31, "&File"
    sPhraseIta 31, "&File"
    sPhraseFra 31, "Dossier"
    sPhraseGer 31, "&Datei"
    sPhraseSpa 31, "&Archivo"
    sPhraseSwe 31, "Arkiv"
    sPhraseNor 31, "Fil"
    sPhraseDan 31, "Filer"

    'REM:
    sPhraseEng 32, "&New war..."
    sPhraseIta 32, "&Nuova partita..."
    sPhraseFra 32, "Nouvelle guerre..."
    sPhraseGer 32, "&Neuer Krieg"
    sPhraseSpa 32, "&Nueva guerra..."
    sPhraseSwe 32, "Nytt krig..."
    sPhraseNor 32, "Ny krig..."
    sPhraseDan 32, "Ny krig..."

    'REM:
    sPhraseEng 33, "E&xit"
    sPhraseIta 33, "&Esci"
    sPhraseFra 33, "Sortie"
    sPhraseGer 33, "&Beenden"
    sPhraseSpa 33, "S&alir"
    sPhraseSwe 33, "Avsluta"
    sPhraseNor 33, "Avslutt"
    sPhraseDan 33, "Afslut"

    'REM:Initialize program.
    sPhraseEng 34, "<Var.ExeName> Set Up... "
    sPhraseIta 34, "Inizializzazione di <Var.ExeName> "
    sPhraseFra 34, "<Var.ExeName> Set Up..."
    sPhraseGer 34, "<Var.ExeName> Set Up..."
    sPhraseSpa 34, "<Var.ExeName> activada..."
    sPhraseSwe 34, "<Var.ExeName> inställningar..."
    sPhraseNor 34, "<Var.ExeName> Oppsett... "
    sPhraseDan 34, "<Var.ExeName> opsætning"

    'REM:Spanish:Misión Riesgo -
    sPhraseEng 35, "<Var.ExeName> - "
    sPhraseIta 35, "<Var.ExeName> - "
    sPhraseFra 35, "<Var.ExeName> - "
    sPhraseGer 35, "<Var.ExeName> - "
    sPhraseSpa 35, "<Var.ExeName> - "
    sPhraseSwe 35, "<Var.ExeName> - "
    sPhraseNor 35, "<Var.ExeName> - "
    sPhraseDan 35, "<Var.ExeName> - "

    'REM:French:Computer: Intelligent
    sPhraseEng 36, "Intelligent computer player"
    sPhraseIta 36, "Giocatore computer intelligente"
    sPhraseFra 36, "Joueur ordinateur intelligent"
    sPhraseGer 36, "Intelligenter Computergegner"
    sPhraseSpa 36, "Jugador computadora inteligente"
    sPhraseSwe 36, "Intelligent datorspelare"
    sPhraseNor 36, "Computer (meget smart)"
    sPhraseDan 36, "Intelligent computer spiller"

    'REM:
    sPhraseEng 37, "No file name has been specified"
    sPhraseIta 37, "Nessun nome file e' stato specificato"
    sPhraseFra 37, "Aucun nom de dossier n'a été spécifié"
    sPhraseGer 37, "Es wurde kein Dateiname angegeben"
    sPhraseSpa 37, "No se especificó nombre de archivo"
    sPhraseSwe 37, "Inget filnamn har angivits"
    sPhraseNor 37, "Filnavn ikke spesifisert"
    sPhraseDan 37, "Uspecificeret filnavn"

    'REM:
    sPhraseEng 38, "&Options"
    sPhraseIta 38, "&Opzioni"
    sPhraseFra 38, "Options"
    sPhraseGer 38, "&Optionen"
    sPhraseSpa 38, "&Opciones"
    sPhraseSwe 38, "Alternativ"
    sPhraseNor 38, "Valg"
    sPhraseDan 38, "Funktioner"

    'REM:
    sPhraseEng 39, "Fast &war"
    sPhraseIta 39, "Guerra &veloce"
    sPhraseFra 39, "Guerre rapide"
    sPhraseGer 39, "&Schneller Krieg"
    sPhraseSpa 39, "&Guerra rápida"
    sPhraseSwe 39, "Snabbt slag"
    sPhraseNor 39, "Hurtig krig"
    sPhraseDan 39, "Hurtig krig"

    'REM:
    sPhraseEng 40, "Fast &dice"
    sPhraseIta 40, "&Dadi veloci"
    sPhraseFra 40, "Dés rapides"
    sPhraseGer 40, "Schnelles &Würfeln"
    sPhraseSpa 40, "&Dados rápidos"
    sPhraseSwe 40, "Snabba tärningsslag"
    sPhraseNor 40, "Hurtige terninger"
    sPhraseDan 40, "Hurtige terninger"

    'REM:
    sPhraseEng 41, "&Border shows player's color"
    sPhraseIta 41, "&Bordo mostra il giocatore di turno"
    sPhraseFra 41, "Cadre de la couleur du joueur"
    sPhraseGer 41, "&Farbe des Spielers als Rand anzeigen"
    sPhraseSpa 41, "El &borde muestra el color del jugador"
    sPhraseSwe 41, "Ramen visar spelarens färg"
    sPhraseNor 41, "Spillers farge på ramme"
    sPhraseDan 41, "Kant viser spillers farve"

    'REM:
    sPhraseEng 42, "3D displa&y"
    sPhraseIta 42, "Displa&y 3D"
    sPhraseFra 42, "effet 3D"
    sPhraseGer 42, "3D &Anzeige"
    sPhraseSpa 42, "pantall&a 3D"
    sPhraseSwe 42, "3D-display"
    sPhraseNor 42, "Vis 3D"
    sPhraseDan 42, "3D display"

    'REM:Scroll automatically.
    sPhraseEng 43, "&Auto scroll"
    sPhraseIta 43, "&Auto scroll"
    sPhraseFra 43, "Défilez automatiquement"
    sPhraseGer 43, "&Automatisches Scrollen"
    sPhraseSpa 43, "&Auto scroll"
    sPhraseSwe 43, "Autoscroll"
    sPhraseNor 43, "Auto scroll"
    sPhraseDan 43, "Auto scroll"

    'REM:The file has been deleted or corrupted.
    sPhraseEng 44, " has been deleted or corrupted"
    sPhraseIta 44, " e' stato cancellato o danneggiato"
    sPhraseFra 44, " a été effacé ou corrompu"
    sPhraseGer 44, " wurde gelöscht oder ist defekt"
    sPhraseSpa 44, " fue borrado o corrompido"
    sPhraseSwe 44, " har raderats eller förstörts"
    sPhraseNor 44, " er slettet eller skadet"
    sPhraseDan 44, " er blevet slettet eller virker ikke."

    'REM:
    sPhraseEng 45, "Missing file"
    sPhraseIta 45, "File mancante"
    sPhraseFra 45, "Dossier manquant"
    sPhraseGer 45, "Fehlende Datei"
    sPhraseSpa 45, "Archivo perdido"
    sPhraseSwe 45, "Fil saknas"
    sPhraseNor 45, "Mangler fil"
    sPhraseDan 45, "Manglende fil"

    'REM:
    sPhraseEng 46, "Mi&ssions"
    sPhraseIta 46, "Ob&iettivi"
    sPhraseFra 46, "Missions"
    sPhraseGer 46, "&Missionen"
    sPhraseSpa 46, "Mi&siones"
    sPhraseSwe 46, "Uppdrag"
    sPhraseNor 46, "Oppdrag"
    sPhraseDan 46, "Missioner"

    'REM:
    sPhraseEng 47, "&See mission"
    sPhraseIta 47, "&Guarda l'obiettivo"
    sPhraseFra 47, "Voir la mission"
    sPhraseGer 47, "&Mission anzeigen"
    sPhraseSpa 47, "&Ver misión"
    sPhraseSwe 47, "Se uppdrag"
    sPhraseNor 47, "Se oppdrag"
    sPhraseDan 47, "Se mission"

    'REM:File has been corrupted and cannot be opened.
    sPhraseEng 48, " has been corrupted and cannot be opened."
    sPhraseIta 48, " e' danneggiato e non puo' essere aperto."
    sPhraseFra 48, " a été corrompu et ne peut pas être ouvert."
    sPhraseGer 48, " ist defekt und kann nicht geöffnet werden."
    sPhraseSpa 48, " fue corrompido y no se abrirá"
    sPhraseSwe 48, " har förstörts och kan inte öppnas"
    sPhraseNor 48, " er skadet og kan ikke åpnes."
    sPhraseDan 48, " virker ikke og kan ikke åbnes."

    'REM:
    sPhraseEng 49, "&Help"
    sPhraseIta 49, "&Aiuto"
    sPhraseFra 49, "Aide"
    sPhraseGer 49, "&Hilfe"
    sPhraseSpa 49, "&Ayuda"
    sPhraseSwe 49, "Hjälp"
    sPhraseNor 49, "Hjelp"
    sPhraseDan 49, "Hjælp"

    'REM:
    sPhraseEng 50, "&Contents..."
    sPhraseIta 50, "&Guida..."
    sPhraseFra 50, "Contenu..."
    sPhraseGer 50, "&Inhalt"
    sPhraseSpa 50, "&Contenidos..."
    sPhraseSwe 50, "Innehåll..."
    sPhraseNor 50, "Innhold..."
    sPhraseDan 50, "Indhold..."

    'REM:
    sPhraseEng 51, "&Index..."
    sPhraseIta 51, "&Indice..."
    sPhraseFra 51, "Index..."
    sPhraseGer 51, "I&ndex..."
    sPhraseSpa 51, "&Indice..."
    sPhraseSwe 51, "Index..."
    sPhraseNor 51, "Stikkord..."
    sPhraseDan 51, "Indeks..."

    'REM:
    sPhraseEng 52, "&<Var.ExeName> home page..."
    sPhraseIta 52, "&<Var.ExeName>: home page..."
    sPhraseFra 52, "Page de bienvenue <Var.ExeName>..."
    sPhraseGer 52, "&<Var.ExeName> Homepage..."
    sPhraseSpa 52, "&página origen de <Var.ExeName>..."
    sPhraseSwe 52, "<Var.ExeName>s hemsida... "
    sPhraseNor 52, "<Var.ExeName> hjemmeside..."
    sPhraseDan 52, "<Var.ExeName> hjemmeside..."

    'REM:
    sPhraseEng 53, ""
    sPhraseIta 53, ""
    sPhraseFra 53, ""
    sPhraseGer 53, ""
    sPhraseSpa 53, ""
    sPhraseSwe 53, ""
    sPhraseNor 53, ""
    sPhraseDan 53, ""

    'REM:
    sPhraseEng 54, ""
    sPhraseIta 54, ""
    sPhraseFra 54, ""
    sPhraseGer 54, ""
    sPhraseSpa 54, ""
    sPhraseSwe 54, ""
    sPhraseNor 54, ""
    sPhraseDan 54, ""

    'REM:
    sPhraseEng 55, ""
    sPhraseIta 55, ""
    sPhraseFra 55, ""
    sPhraseGer 55, ""
    sPhraseSpa 55, ""
    sPhraseSwe 55, ""
    sPhraseNor 55, ""
    sPhraseDan 55, ""

    'REM:
    sPhraseEng 56, "File error"
    sPhraseIta 56, "Errore nel file"
    sPhraseFra 56, "Erreur fichier"
    sPhraseGer 56, "Dateifehler"
    sPhraseSpa 56, "Error en el archivo"
    sPhraseSwe 56, "Filfel"
    sPhraseNor 56, "Filfeil"
    sPhraseDan 56, "Der opstod en fejl ved læsning af en fil"

    'REM:Spare
    sPhraseEng 57, ""
    sPhraseIta 57, ""
    sPhraseFra 57, ""
    sPhraseGer 57, ""
    sPhraseSpa 57, ""
    sPhraseSwe 57, ""
    sPhraseNor 57, ""
    sPhraseDan 57, ""

    'REM: Transfer units
    sPhraseEng 58, " Transfer"
    sPhraseIta 58, "Trasferisci"
    sPhraseFra 58, "Transferer"
    sPhraseGer 58, "Übertragung"
    sPhraseSpa 58, "Transferencia"
    sPhraseSwe 58, "Förflytta"
    sPhraseNor 58, " Overfør"
    sPhraseDan 58, " Overfør"

    'REM:
    sPhraseEng 59, "&Attack"
    sPhraseIta 59, "&Attacca"
    sPhraseFra 59, "Attaquer"
    sPhraseGer 59, "&Angreifen"
    sPhraseSpa 59, "&Ataque"
    sPhraseSwe 59, "Anfall"
    sPhraseNor 59, "Angrip"
    sPhraseDan 59, "Angrib"

    'REM:
    sPhraseEng 60, "&Move"
    sPhraseIta 60, "Spo&sta"
    sPhraseFra 60, "Déplacer"
    sPhraseGer 60, "&Verschieben"
    sPhraseSpa 60, "&Mover"
    sPhraseSwe 60, "Flytta"
    sPhraseNor 60, "Flytt"
    sPhraseDan 60, "Flyt"

    'REM:
    sPhraseEng 61, "&Pass"
    sPhraseIta 61, "&Passa"
    sPhraseFra 61, "Passer"
    sPhraseGer 61, "&Passen"
    sPhraseSpa 61, "&Pasar"
    sPhraseSwe 61, "Passa"
    sPhraseNor 61, "Ferdig"
    sPhraseDan 61, "Færdig"

    'REM:
    sPhraseEng 62, "Exchange"
    sPhraseIta 62, "Scambia"
    sPhraseFra 62, "Echanger"
    sPhraseGer 62, "Austausch"
    sPhraseSpa 62, "Intercambio"
    sPhraseSwe 62, "Byt in"
    sPhraseNor 62, "Veksle"
    sPhraseDan 62, "Byt"

    'REM:
    sPhraseEng 63, "Cancel"
    sPhraseIta 63, "Annulla"
    sPhraseFra 63, "Annuler"
    sPhraseGer 63, "Abbrechen"
    sPhraseSpa 63, "Cancelar"
    sPhraseSwe 63, "Avbryt"
    sPhraseNor 63, "Avbryt"
    sPhraseDan 63, "Fortryd"

    'REM:X attacks Y.
    sPhraseEng 64, "attacks"
    sPhraseIta 64, "attacca"
    sPhraseFra 64, "attaque"
    sPhraseGer 64, "attackiert"
    sPhraseSpa 64, "ataques"
    sPhraseSwe 64, "attackerar"
    sPhraseNor 64, "angriper"
    sPhraseDan 64, "angriber"

    'REM:
    sPhraseEng 65, "<Click defending country>"
    sPhraseIta 65, "<Clicca sul territorio in difesa>"
    sPhraseFra 65, "<Clickez le pays attaqué >"
    sPhraseGer 65, "<Verteidigendes Land anklicken>"
    sPhraseSpa 65, "<Señale país atacado>"
    sPhraseSwe 65, "<Klicka på försvarande land>"
    sPhraseNor 65, "<Velg motstander>"
    sPhraseDan 65, "<Vælg modstander>"

    'REM:X has WON THE WAR!
    sPhraseEng 66, "has won the war"
    sPhraseIta 66, "ha VINTO LA PARTITA!"
    sPhraseFra 66, "a GAGNÉ LE JEU!"
    sPhraseGer 66, "hat das SPIEL GEWONNEN"
    sPhraseSpa 66, "¡ha GANADO EL JUEGO!"
    sPhraseSwe 66, "har VUNNIT SPELET"
    sPhraseNor 66, "har VUNNET spillet!"
    sPhraseDan 66, "har VUNDET!"

    'REM:by occupying and holding X.
    sPhraseEng 67, "by occupying and holding"
    sPhraseIta 67, "occupando e tenendo"
    sPhraseFra 67, "En occupant et tenant"
    sPhraseGer 67, "Durch besetzen und halten von"
    sPhraseSpa 67, "Por ocupar y mantener"
    sPhraseSwe 67, "Genom att besätta och hålla"
    sPhraseNor 67, "Ved å okkupere og holde"
    sPhraseDan 67, "Ved at besætte og holde"

    'REM:
    sPhraseEng 68, ""
    sPhraseIta 68, ""
    sPhraseFra 68, ""
    sPhraseGer 68, ""
    sPhraseSpa 68, ""
    sPhraseSwe 68, ""
    sPhraseNor 68, ""
    sPhraseDan 68, ""

    'REM:by wiping out X.
    sPhraseEng 69, "by wiping out "
    sPhraseIta 69, "distruggendo "
    sPhraseFra 69, "En détruisant"
    sPhraseGer 69, "Durch das Auslöschen "
    sPhraseSpa 69, "Por aniquilación "
    sPhraseSwe 69, "Genom att utplåna "
    sPhraseNor 69, "Ved å utslette "
    sPhraseDan 69, "Ved at udrydde "

    'REM:You have X units to place.
    sPhraseEng 70, "You have "
    sPhraseIta 70, "Devi mettere "
    sPhraseFra 70, "Vous avez "
    sPhraseGer 70, "Sie haben "
    sPhraseSpa 70, "Usted tiene "
    sPhraseSwe 70, "Du har "
    sPhraseNor 70, "Du har "
    sPhraseDan 70, "Du skal placere "

    'REM:
    sPhraseEng 71, " units to place."
    sPhraseIta 71, " unita'."
    sPhraseFra 71, " unités à placer."
    sPhraseGer 71, " Einheiten zu" + vbCrLf _
                 + "plazieren."
    sPhraseSpa 71, " unidades a" + vbCrLf _
                 + "incorporar."
    sPhraseSwe 71, " styrkor att sätta ut"
    sPhraseNor 71, " enheter å plassere."
    sPhraseDan 71, " arméer."

    'REM:
    sPhraseEng 72, "<Click destination country>"
    sPhraseIta 72, "<Clicca il territorio di destinazione>"
    sPhraseFra 72, "<Clickez le pays destination>"
    sPhraseGer 72, "<Ziel-Land anklicken>"
    sPhraseSpa 72, "<Señale país destino>"
    sPhraseSwe 72, "<Klicka på mållandet> "
    sPhraseNor 72, "<Velg mottakerland>"
    sPhraseDan 72, "<Vælg destination>"

    'REM:X moves to Y.
    sPhraseEng 73, "moves "
    sPhraseIta 73, "sposta "
    sPhraseFra 73, "se déplace "
    sPhraseGer 73, "zieht "
    sPhraseSpa 73, "movimientos "
    sPhraseSwe 73, "flyttar "
    sPhraseNor 73, "flytter "
    sPhraseDan 73, "flytter "

    'REM:
    sPhraseEng 74, "to "
    sPhraseIta 74, "in "
    sPhraseFra 74, "vers "
    sPhraseGer 74, "nach "
    sPhraseSpa 74, "a "
    sPhraseSwe 74, "till "
    sPhraseNor 74, "til "
    sPhraseDan 74, "til "

    'REM:Mission 0
    sPhraseEng 75, "You must wipe out all other players and conquer the world."
    sPhraseIta 75, "Devi distruggere gli altri giocatori e conquistare il mondo."
    sPhraseFra 75, "Vous devez détruire tous les autres joueurs et conquérir le monde."
    sPhraseGer 75, "Sie müssen alle anderen Spieler auslöschen und die Welt erobern."
    sPhraseSpa 75, "Usted debe aniquilar a todos los demás jugadores y conquistar el mundo"
    sPhraseSwe 75, "Du ska utplåna alla andra spelare och erövra världen."
    sPhraseNor 75, "Du skal utslette alle andre spillere og erobre verden."
    sPhraseDan 75, "Du skal udrydde alle andre spillere og erobre verden."

    'REM:Mission 1-6 Your mission is to wipe out X.
    sPhraseEng 76, "Your mission is to wipe out "
    sPhraseIta 76, "Il tuo obiettivo è distruggere "
    sPhraseFra 76, "Votre mission est de détruire "
    sPhraseGer 76, "Vernichten Sie "
    sPhraseSpa 76, "Su misión es aniquilar a "
    sPhraseSwe 76, "Ditt uppdrag är att utplåna "
    sPhraseNor 76, "Ditt oppdrag er å utslette "
    sPhraseDan 76, "Du skal udrydde "

    'REM:Mission 7 North America and South America
    sPhraseEng 77, "You must conquer <Var.Phrase(423)> and <Var.Phrase(424)> and hold them until your next turn."
    sPhraseIta 77, "Devi conquistare <Var.Phrase(423)> e <Var.Phrase(424)> e tenerle per un turno."
    sPhraseFra 77, "Vous devez conquérir <Var.Phrase(423)> et <Var.Phrase(424)> et les tenir jusqu'à votre prochain tour."
    sPhraseGer 77, "Sie müssen <Var.Phrase(423)> und <Var.Phrase(424)> erobern und sie bis zur nächsten Runde halten."
    sPhraseSpa 77, "Debe conquistar <Var.Phrase(423)> y <Var.Phrase(424)> y mantenerlas hasta su próximo turno."
    sPhraseSwe 77, "Du ska erövra <Var.Phrase(423)> och <Var.Phrase(424)> och hålla dem till ditt nästa drag."
    sPhraseNor 77, "Du skal erobre <Var.Phrase(423)> og <Var.Phrase(424)>, og holde disse til det blir din tur igjen."
    sPhraseDan 77, "Du skal erobre <Var.Phrase(423)> og <Var.Phrase(424)>, og holde dem en runde."

    'REM:Mission 8 North America And Australia
    sPhraseEng 78, "You must conquer <Var.Phrase(423)> And <Var.Phrase(428)> and hold them until your next turn."
    sPhraseIta 78, "Devi conquistare <Var.Phrase(423)> e <Var.Phrase(428)> e tenerle per un turno."
    sPhraseFra 78, "Vous devez conquérir <Var.Phrase(423)> et <Var.Phrase(428)> et les tenir jusqu'à votre prochain tour."
    sPhraseGer 78, "Sie müssen <Var.Phrase(423)> und <Var.Phrase(428)> erobern und sie bis zur nächsten Runde halten."
    sPhraseSpa 78, "Debe conquistar <Var.Phrase(423)> y <Var.Phrase(428)> y mantenerlas hasta su próximo turno."
    sPhraseSwe 78, "Du ska erövra <Var.Phrase(423)> och <Var.Phrase(428)> och hålla dem till ditt nästa drag."
    sPhraseNor 78, "Du skal erobre <Var.Phrase(423)> og <Var.Phrase(428)>, og holde disse til det blir din tur igjen."
    sPhraseDan 78, "Du skal erobre <Var.Phrase(423)> og <Var.Phrase(428)>, og holde dem en runde."

    'REM:Mission 9 South America And Europe
    sPhraseEng 79, "You must conquer <Var.Phrase(424)> And <Var.Phrase(425)> and hold them until your next turn."
    sPhraseIta 79, "Devi conquistare <Var.Phrase(424)> e <Var.Phrase(425)> e tenerle per un turno."
    sPhraseFra 79, "Vous devez conquérir <Var.Phrase(424)> et <Var.Phrase(425)> et les tenir jusqu'à votre prochain tour."
    sPhraseGer 79, "Sie müssen <Var.Phrase(424)> und <Var.Phrase(425)> erobern und sie bis zur nächsten Runde halten."
    sPhraseSpa 79, "Debe conquistar <Var.Phrase(424)> y <Var.Phrase(425)> y mantenerlas hasta su próximo turno."
    sPhraseSwe 79, "Du ska erövra <Var.Phrase(424)> och <Var.Phrase(425)> och hålla dem till ditt nästa drag."
    sPhraseNor 79, "Du skal erobre <Var.Phrase(424)> og <Var.Phrase(425)>, og holde disse til det blir din tur igjen."
    sPhraseDan 79, "Du skal erobre <Var.Phrase(424)> og <Var.Phrase(425)>, og holde dem en runde."

    'REM:Mission 10 Europe and Australia
    sPhraseEng 80, "You must conquer <Var.Phrase(425)> and <Var.Phrase(428)> and hold them until your next turn."
    sPhraseIta 80, "Devi conquistare <Var.Phrase(425)> e <Var.Phrase(428)> e tenerle per un turno."
    sPhraseFra 80, "Vous devez conquérir <Var.Phrase(425)> et <Var.Phrase(428)> et les tenir jusqu'à votre prochain tour."
    sPhraseGer 80, "Sie müssen <Var.Phrase(425)> und <Var.Phrase(428)> erobern und sie bis zur nächsten Runde halten."
    sPhraseSpa 80, "Debe conquistar <Var.Phrase(425)> y <Var.Phrase(428)> y mantenerlas hasta su próximo turno."
    sPhraseSwe 80, "Du ska erövra <Var.Phrase(425)> och <Var.Phrase(428)> och hålla dem till ditt nästa drag."
    sPhraseNor 80, "Du skal erobre <Var.Phrase(425)> og <Var.Phrase(428)>, og holde disse til det blir din tur igjen."
    sPhraseDan 80, "Du skal erobre <Var.Phrase(425)> og <Var.Phrase(428)>, og holde dem en runde."

    'REM:Mission 11 South America And Africa
    sPhraseEng 81, "You must conquer <Var.Phrase(424)> And <Var.Phrase(426)> and hold them until your next turn."
    sPhraseIta 81, "Devi conquistare <Var.Phrase(424)> e <Var.Phrase(426)> e tenerle per un turno."
    sPhraseFra 81, "Vous devez conquérir <Var.Phrase(424)> et <Var.Phrase(426)> et les tenir jusqu'à votre prochain tour."
    sPhraseGer 81, "Sie müssen <Var.Phrase(424)> und <Var.Phrase(426)> erobern und sie bis zur nächsten Runde halten."
    sPhraseSpa 81, "Debe conquistar <Var.Phrase(424)> y <Var.Phrase(426)> y mantenerlas hasta su próximo turno."
    sPhraseSwe 81, "Du ska erövra <Var.Phrase(424)> och <Var.Phrase(426)> och hålla dem till ditt nästa drag."
    sPhraseNor 81, "Du skal erobre <Var.Phrase(424)> og <Var.Phrase(426)>, og holde disse til det blir din tur igjen."
    sPhraseDan 81, "Du skal erobre <Var.Phrase(424)> og <Var.Phrase(426)>, og holde dem en runde."

    'REM:Mission 12 Africa and Australia
    sPhraseEng 82, "You must conquer <Var.Phrase(426)> and <Var.Phrase(428)> and hold them until your next turn."
    sPhraseIta 82, "Devi conquistare <Var.Phrase(426)> e <Var.Phrase(428)> e tenerle per un turno."
    sPhraseFra 82, "Vous devez conquérir <Var.Phrase(426)> et <Var.Phrase(428)> et les tenir jusqu'à votre prochain tour."
    sPhraseGer 82, "Sie müssen <Var.Phrase(426)> und <Var.Phrase(428)> erobern und sie bis zur nächsten Runde halten."
    sPhraseSpa 82, "Debe conquistar <Var.Phrase(426)> y <Var.Phrase(428)> y mantenerlas hasta su próximo turno."
    sPhraseSwe 82, "Du ska erövra <Var.Phrase(426)> och <Var.Phrase(428)> och hålla dem till ditt nästa drag."
    sPhraseNor 82, "Du skal erobre <Var.Phrase(426)> og <Var.Phrase(428)>, og holde disse til det blir din tur igjen."
    sPhraseDan 82, "Du skal erobre <Var.Phrase(426)> og <Var.Phrase(428)>, og holde dem en runde."

    'REM:Mission 13 South America And Australia
    sPhraseEng 83, "You must conquer <Var.Phrase(424)> And <Var.Phrase(428)> and hold them until your next turn."
    sPhraseIta 83, "Devi conquistare <Var.Phrase(424)> e <Var.Phrase(428)> e tenerle per un turno."
    sPhraseFra 83, "Vous devez conquérir <Var.Phrase(424)> et <Var.Phrase(428)> et les tenir jusqu'à votre prochain tour."
    sPhraseGer 83, "Sie müssen <Var.Phrase(424)> und <Var.Phrase(428)> erobern und sie bis zur nächsten Runde halten."
    sPhraseSpa 83, "Debe conquistar <Var.Phrase(424)> y <Var.Phrase(428)> y mantenerlas hasta su próximo turno."
    sPhraseSwe 83, "Du ska erövra <Var.Phrase(424)> och <Var.Phrase(428)> och hålla dem till ditt nästa drag."
    sPhraseNor 83, "Du skal erobre <Var.Phrase(424)> og <Var.Phrase(428)>, og holde disse til det blir din tur igjen."
    sPhraseDan 83, "Du skal erobre <Var.Phrase(424)> og <Var.Phrase(428)>, og holde dem en runde."

    'REM:Mission 14 18 countries
    sPhraseEng 84, "You must occupy any <Var.Mission14> countries and hold them until your next turn."
    sPhraseIta 84, "Devi occupare <Var.Mission14> territori e tenerli per un turno."
    sPhraseFra 84, "Vous devez occuper <Var.Mission14> pays et les tenir jusqu'à votre prochain tour."
    sPhraseGer 84, "Sie müssen <Var.Mission14> Länder ihrer Wahl erobern und sie bis zur nächsten Runde halten."
    sPhraseSpa 84, "Debe ocupar cuales quiera <Var.Mission14> países y mantenerlo hasta su próximo turno."
    sPhraseSwe 84, "Du ska besätta <Var.Mission14> länder och hålla dem till ditt nästa drag"
    sPhraseNor 84, "Du skal okkupere <Var.Mission14> land, og holde disse til det blir din tur igjen."
    sPhraseDan 84, "Du skal erobre <Var.Mission14> lande lande, og holde dem en runde."

    'REM:
    sPhraseEng 85, ""
    sPhraseIta 85, ""
    sPhraseFra 85, ""
    sPhraseGer 85, ""
    sPhraseSpa 85, ""
    sPhraseSwe 85, ""
    sPhraseNor 85, ""
    sPhraseDan 85, ""

    'REM:
    sPhraseEng 86, ""
    sPhraseIta 86, ""
    sPhraseFra 86, ""
    sPhraseGer 86, ""
    sPhraseSpa 86, ""
    sPhraseSwe 86, ""
    sPhraseNor 86, ""
    sPhraseDan 86, ""

    'REM:
    sPhraseEng 87, ""
    sPhraseIta 87, ""
    sPhraseFra 87, ""
    sPhraseGer 87, ""
    sPhraseSpa 87, ""
    sPhraseSwe 87, ""
    sPhraseNor 87, ""
    sPhraseDan 87, ""

    'REM:Spare
    sPhraseEng 88, ""
    sPhraseIta 88, ""
    sPhraseFra 88, ""
    sPhraseGer 88, ""
    sPhraseSpa 88, ""
    sPhraseSwe 88, ""
    sPhraseNor 88, ""
    sPhraseDan 88, ""

    'REM:
    sPhraseEng 89, "Your system color pallet is not set for HI COLOR, which means that <Var.ExeName> can only render graphics in 2D." + vbCrLf _
                 + "" + vbCrLf _
                 + "If you would like 3D rendering, you can change your pallet settings from the display options box in the control panel. See the help file for more details."
    sPhraseIta 89, "La tua tavolozza dei colori non e' settata per i HI COLOR, il che significa che <Var.ExeName> può creare effetti solo 2D." + vbCrLf _
                 + "" + vbCrLf _
                 + "Se vuoi il rendering 3D, devi cambiare il settaggio della tavolozza dei colori in pannello di controllo|schermo. Guarda il file di help per maggiori dettagli."
    sPhraseFra 89, "Votre palette couleur système n'est pas mise à COULEURS (16 bits ou > ) ca qui veux dire que <Var.ExeName> ne peut rendre que l'effet graphique 2D." + vbCrLf _
                 + "" + vbCrLf _
                 + "Si vous aimeriez l'effet 3D, vous pouvez changer les options de la palette couleur du systéme dans le tableau de bord. Lisez l'aide pour plus de détails."
    sPhraseGer 89, "Ihre Systemfarben sind nicht auf HIGH COLOR (16bit) eingestellt, was bedeutet, dass <Var.ExeName> die Grafiken nur in 2D rendern kann." + vbCrLf _
                 + "" + vbCrLf _
                 + "Wenn Sie gerne 3D Rendering hätten, können Sie Ihre Farbeinstellungen in den Anzeige-Eigenschaften der Systemsteuerung einstellen. Schauen Sie in die Hilfe für weitere Details."
    sPhraseSpa 89, "Su paleta de color del sistema no está puesta a COLOR MAXIMO, lo que significa que Misión Riesgo solo puede recrear gráficos en 2D." + vbCrLf _
                 + "" + vbCrLf _
                 + "Si deseara recreaciones 3D, debe cambiar su paleta desde la caja de opciones de pantalla en el panel de control. Vea el archivo de ayuda para mas detalles."
    sPhraseSwe 89, "Ditt system är inte inställt på HIGH COLOR, vilket betyder att <Var.ExeName> bara klarar att visa grafik i 2D." + vbCrLf _
                 + "" + vbCrLf _
                 + "Om du vill se grafik i 3D, kan du ändra systeminställningarna under Bildskärm i Kontrollpanelen. Läs i Hjälp-filen för mer information. "
    sPhraseNor 89, "Skjermoppsettet ditt støtter ikke mer enn 256 farger, hvilket betyr at <Var.ExeName> kun kan kjøres i 2D-modus." + vbCrLf _
                 + "" + vbCrLf _
                 + "Dersom du ønsker å benytte 3D-modus må du endre fargedybden under Egenskaper for skjerm. Se hjelpefilen for detaljer."
    sPhraseDan 89, "Din computer er ikke sat til ægte farver, hvilket betyder at <Var.ExeName> ikke kan spilles i 3D." + vbCrLf _
                 + "" + vbCrLf _
                 + "Hvis du vil spille spillet i 3D, skal du skifte indstillinger for din skærm i kontrolpanelet. Se hjælp for mere information."

    'REM:
    sPhraseEng 90, ""
    sPhraseIta 90, ""
    sPhraseFra 90, ""
    sPhraseGer 90, ""
    sPhraseSpa 90, ""
    sPhraseSwe 90, ""
    sPhraseNor 90, ""
    sPhraseDan 90, ""

    'REM:
    sPhraseEng 91, ""
    sPhraseIta 91, ""
    sPhraseFra 91, ""
    sPhraseGer 91, ""
    sPhraseSpa 91, ""
    sPhraseSwe 91, ""
    sPhraseNor 91, ""
    sPhraseDan 91, ""

    'REM:
    sPhraseEng 92, ""
    sPhraseIta 92, ""
    sPhraseFra 92, ""
    sPhraseGer 92, ""
    sPhraseSpa 92, ""
    sPhraseSwe 92, ""
    sPhraseNor 92, ""
    sPhraseDan 92, ""

    'REM:
    sPhraseEng 93, ""
    sPhraseIta 93, ""
    sPhraseFra 93, ""
    sPhraseGer 93, ""
    sPhraseSpa 93, ""
    sPhraseSwe 93, ""
    sPhraseNor 93, ""
    sPhraseDan 93, ""

    'REM:
    sPhraseEng 94, ""
    sPhraseIta 94, ""
    sPhraseFra 94, ""
    sPhraseGer 94, ""
    sPhraseSpa 94, ""
    sPhraseSwe 94, ""
    sPhraseNor 94, ""
    sPhraseDan 94, ""

    'REM:
    sPhraseEng 95, ""
    sPhraseIta 95, ""
    sPhraseFra 95, ""
    sPhraseGer 95, ""
    sPhraseSpa 95, ""
    sPhraseSwe 95, ""
    sPhraseNor 95, ""
    sPhraseDan 95, ""

    'REM:
    sPhraseEng 96, ""
    sPhraseIta 96, ""
    sPhraseFra 96, ""
    sPhraseGer 96, ""
    sPhraseSpa 96, ""
    sPhraseSwe 96, ""
    sPhraseNor 96, ""
    sPhraseDan 96, ""

    'REM:
    sPhraseEng 97, ""
    sPhraseIta 97, ""
    sPhraseFra 97, ""
    sPhraseGer 97, ""
    sPhraseSpa 97, ""
    sPhraseSwe 97, ""
    sPhraseNor 97, ""
    sPhraseDan 97, ""

    'REM:
    sPhraseEng 98, ""
    sPhraseIta 98, ""
    sPhraseFra 98, ""
    sPhraseGer 98, ""
    sPhraseSpa 98, ""
    sPhraseSwe 98, ""
    sPhraseNor 98, ""
    sPhraseDan 98, ""

    'REM:Terminate the program?
    sPhraseEng 99, "Do you really want to QUIT?"
    sPhraseIta 99, "Terminare questo programma?"
    sPhraseFra 99, "Quitter le programme?"
    sPhraseGer 99, "Wollen Sie wirklich BEENDEN?"
    sPhraseSpa 99, "¿Realmente desea SALIR?"
    sPhraseSwe 99, "Vill du verkligen SLUTA?"
    sPhraseNor 99, "Ønsker du virkelig å AVSLUTTE?"
    sPhraseDan 99, "Vil du virkelig afslutte?"
    Call initialisePhrases100
End Sub

Private Sub initialisePhrases100()
    'REM:
    sPhraseEng 100, "by dominating the world."
    sPhraseIta 100, "Dominando il mondo."
    sPhraseFra 100, "En dominant le monde."
    sPhraseGer 100, "Durch Erobern der Welt."
    sPhraseSpa 100, "Por dominación del mundo."
    sPhraseSwe 100, "Genom att dominera världen."
    sPhraseNor 100, "Ved å dominere verdensherredømmet."
    sPhraseDan 100, " Ved at erobre verden."

    'REM:Attack X with 5 units.
    sPhraseEng 101, "with "
    sPhraseIta 101, "con "
    sPhraseFra 101, "avec "
    sPhraseGer 101, "mit "
    sPhraseSpa 101, "con "
    sPhraseSwe 101, "med "
    sPhraseNor 101, "med "
    sPhraseDan 101, "med "

    'REM:
    sPhraseEng 102, " units."
    sPhraseIta 102, " unita'."
    sPhraseFra 102, " unités."
    sPhraseGer 102, " Einheiten"
    sPhraseSpa 102, " unidades."
    sPhraseSwe 102, " enheter."
    sPhraseNor 102, " enheter."
    sPhraseDan 102, " arméer."

    'REM:Attack X with 1 unit.
    sPhraseEng 103, " unit."
    sPhraseIta 103, " unita'."
    sPhraseFra 103, " unité."
    sPhraseGer 103, " Einheit."
    sPhraseSpa 103, " unidad."
    sPhraseSwe 103, " enhet."
    sPhraseNor 103, " enhet."
    sPhraseDan 103, " armé."

    'REM:
    sPhraseEng 104, "<Click attacking country>"
    sPhraseIta 104, "<Clicca sul territorio attaccante>"
    sPhraseFra 104, "<Clickez sur le pays attaquant>"
    sPhraseGer 104, "<Das angreifende Land anklicken>"
    sPhraseSpa 104, "<Pulse el país atacante>"
    sPhraseSwe 104, "<Klicka på anfallande land>"
    sPhraseNor 104, "<Klikk land som angriper>"
    sPhraseDan 104, "<Vælg angribende land>"

    'REM:You have 1 unit to place.
    sPhraseEng 105, " unit to place."
    sPhraseIta 105, " unita'."
    sPhraseFra 105, " unité à placer."
    sPhraseGer 105, " Einheit zu plazieren."
    sPhraseSpa 105, " unidades a colocar."
    sPhraseSwe 105, " enhet att placera ut."
    sPhraseNor 105, " enhet å plassere."
    sPhraseDan 105, " armé der" + vbCrLf _
                 + "skal placeres."

    'REM:X moves to Y.
    sPhraseEng 106, "moves to "
    sPhraseIta 106, "sposta in "
    sPhraseFra 106, "se déplace vers "
    sPhraseGer 106, "zieht nach "
    sPhraseSpa 106, "mueve a "
    sPhraseSwe 106, "flyttar till "
    sPhraseNor 106, "flytter til "
    sPhraseDan 106, "flytter til "

    'REM:X moves to Y from Z.
    sPhraseEng 107, " from"
    sPhraseIta 107, " da"
    sPhraseFra 107, " de"
    sPhraseGer 107, " von"
    sPhraseSpa 107, " desde"
    sPhraseSwe 107, " från"
    sPhraseNor 107, " fra"
    sPhraseDan 107, " fra"

    'REM:Click starting territory. Click initial territory. Click initial country.
    sPhraseEng 108, "<Click source country>"
    sPhraseIta 108, "<Clicca il territorio iniziale>"
    sPhraseFra 108, "<Clickez sur le pays source>"
    sPhraseGer 108, "<Ziel-Land anklicken>"
    sPhraseSpa 108, "<Pulsa país aprovisionador>"
    sPhraseSwe 108, "<Klicka på ursprungsland>"
    sPhraseNor 108, "<Klikk kilde-land>"
    sPhraseDan 108, "<Vælg kilde>"

    'REM:Click attacking countries. Click attacking country. Click attacking territory.
    sPhraseEng 109, "<Click attacking country(s)>"
    sPhraseIta 109, "<Clicca sul territorio attaccante>"
    sPhraseFra 109, "<Clickez sur le(s) pays attaquant(s)>"
    sPhraseGer 109, "<Angreifende Länder anklicken>"
    sPhraseSpa 109, "<Pulse país atacante(s)>"
    sPhraseSwe 109, "<Klicka på anfallande land (länder)>"
    sPhraseNor 109, "<Klikk land(ene) som angriper>"
    sPhraseDan 109, "<Vælg angribende land(e)>"

    'REM:
    sPhraseEng 110, "Retreats 0 units from"
    sPhraseIta 110, "Toglie 0 unita' da"
    sPhraseFra 110, "Retire 0 unité de"
    sPhraseGer 110, "Zieht 0 Einheiten von"
    sPhraseSpa 110, "Se retiran 0 unidades desde"
    sPhraseSwe 110, "Drar tillbaka 0 enheter från"
    sPhraseNor 110, "Trekker ut 0 enheter fra"
    sPhraseDan 110, "Fjerner 0 arméer fra"

    'REM:
    sPhraseEng 111, "Retreats"
    sPhraseIta 111, "Toglie"
    sPhraseFra 111, "Se retire"
    sPhraseGer 111, "Zieht"
    sPhraseSpa 111, "Retiradas"
    sPhraseSwe 111, "Drar tillbaka"
    sPhraseNor 111, "Trekker ut"
    sPhraseDan 111, "Fjerner"

    'REM:X retreats 5 units from Y.
    sPhraseEng 112, " units from"
    sPhraseIta 112, " unita' da"
    sPhraseFra 112, " unités de"
    sPhraseGer 112, " Einheiten zurück von"
    sPhraseSpa 112, " unidades desde"
    sPhraseSwe 112, " enheter från"
    sPhraseNor 112, " enheter fra"
    sPhraseDan 112, " arméer fra"

    'REM:X retreats 1 unit from Y.
    sPhraseEng 113, " unit from"
    sPhraseIta 113, " unita' da"
    sPhraseFra 113, " unités de"
    sPhraseGer 113, " Einheit von"
    sPhraseSpa 113, " unidad desde"
    sPhraseSwe 113, " enhet från"
    sPhraseNor 113, " enhet fra"
    sPhraseDan 113, " armé fra"

    'REM:X has 5 units left.
    sPhraseEng 114, " units left."
    sPhraseIta 114, " unita' rimaste."
    sPhraseFra 114, " unités qui restent."
    sPhraseGer 114, " Einheiten übrig."
    sPhraseSpa 114, " unidades restantes."
    sPhraseSwe 114, " enheter kvar."
    sPhraseNor 114, " enheter igjen."
    sPhraseDan 114, " arméer tilbage."

    'REM:X has 1 unit left.
    sPhraseEng 115, " unit  left."
    sPhraseIta 115, " unita' rimasta."
    sPhraseFra 115, " unité qui reste."
    sPhraseGer 115, " Einheit übrig."
    sPhraseSpa 115, " unidad restante."
    sPhraseSwe 115, " enhet kvar"
    sPhraseNor 115, " enhet igjen."
    sPhraseDan 115, " armé tilbage."

    'REM:X belongs to Y, which has a value of 5 units.
    sPhraseEng 116, "Belongs to "
    sPhraseIta 116, "Appartiene a "
    sPhraseFra 116, "Appartient à "
    sPhraseGer 116, "Gehört zu "
    sPhraseSpa 116, "Pertenece a "
    sPhraseSwe 116, "Tillhör "
    sPhraseNor 116, "Tilhører "
    sPhraseDan 116, "Tilhører "

    'REM:
    sPhraseEng 117, "which has a value of"
    sPhraseIta 117, "ha un valore di"
    sPhraseFra 117, "qui a une valeur de"
    sPhraseGer 117, " mit einem Wert von"
    sPhraseSpa 117, "que tiene un valor de"
    sPhraseSwe 117, "som är värt"
    sPhraseNor 117, "som har en verdi av"
    sPhraseDan 117, "som har en værdi af"

    'REM:Currently occupied by X.
    sPhraseEng 118, "Currently occupied by "
    sPhraseIta 118, "E' ora occupata dalle "
    sPhraseFra 118, "Est occupé par "
    sPhraseGer 118, "Zur Zeit besetzt von "
    sPhraseSpa 118, "Actualmente ocupado por "
    sPhraseSwe 118, "Hålls för närvarande av "
    sPhraseNor 118, "For tiden okkupert av "
    sPhraseDan 118, "Tilhører i øjeblikket "

    'REM:I have caught you attempting to look at my mission.
    sPhraseEng 119, "SPRUNG!"
    sPhraseIta 119, "CATTATO!"
    sPhraseFra 119, "SAUTÉ!"
    sPhraseGer 119, "SPRUNG!"
    sPhraseSpa 119, "¡BROTÓ!"
    sPhraseSwe 119, "SPRUNG!"
    sPhraseNor 119, "AVSLØRT!"
    sPhraseDan 119, "Afsløret! Du kan ikke se andres mission."

    'REM:
    sPhraseEng 120, "Ah Hah! Caught spying!" + vbCrLf _
                 + "You must wait until it's your turn to see what your mission is!"
    sPhraseIta 120, "Ah Hah! Ti ho preso a spiare!" + vbCrLf _
                 + "Devi aspettare il tuo turno per vedere il tuo obiettivo!"
    sPhraseFra 120, "Ah Hah! L'espion a été attrapé!" + vbCrLf _
                 + "Vous devez attendre jusqu'à votre prochain tour pour voir quelle est votre mission!"
    sPhraseGer 120, "Ah Hah! Beim Spionieren erwischt!" + vbCrLf _
                 + "Sie müssen bis zum eigenen Zug warten, um zu sehen wie Ihre Mission lautet!"
    sPhraseSpa 120, "¡A já! Lo agarré espiando!" + vbCrLf _
                 + "¡Deberá esperar hasta su turno para ver cual es su misión!"
    sPhraseSwe 120, "Aha! Ertappad med spioneri!" + vbCrLf _
                 + "Du måste vänta tills det är din tur för att få reda på ditt uppdrag!"
    sPhraseNor 120, "Aha! Du spionerer?!" + vbCrLf _
                 + "Du må vente til det blir din tur for å se hva ditt oppdrag er!"
    sPhraseDan 120, "Snyder!" + vbCrLf _
                 + "Du må vente til det bliver din tur før du kan læse din mission."

    'REM:
    sPhraseEng 121, "Stop peeking!"
    sPhraseIta 121, "Piantale di curiosare!"
    sPhraseFra 121, "Cessez de m'espionner!"
    sPhraseGer 121, "Hören Sie auf zu spicken!"
    sPhraseSpa 121, "¡Basta de espiar!"
    sPhraseSwe 121, "Sluta kika!"
    sPhraseNor 121, "Ikke kikk!"
    sPhraseDan 121, "Hold nu op med at kigge!"

    'REM:
    sPhraseEng 122, "NO! My mission!" + vbCrLf _
                 + "You must wait until it's your turn to see what your mission is!"
    sPhraseIta 122, "NO! Il mio obiettivo!" + vbCrLf _
                 + "Devi aspettare il tuo turno per vedere il tuo obiettivo!"
    sPhraseFra 122, "NON! Ma mission!" + vbCrLf _
                 + "Vous devez attendre jusqu'à votre prochain tour pour quelle est votre mission!"
    sPhraseGer 122, "Nein! Meine Mission!" + vbCrLf _
                 + "Sie müssen bis zum eigenen Zug warten, um zu sehen wie Ihre Mission lautet!"
    sPhraseSpa 122, "¡No, mi misión!" + vbCrLf _
                 + "¡Deberá esperar hasta su turno para ver cual es su misión!"
    sPhraseSwe 122, "NEJ! Det är mitt uppdrag!" + vbCrLf _
                 + "Du måste vänta tills det är din tur för att få reda på ditt uppdrag!"
    sPhraseNor 122, "NEI! Dette er MITT oppdrag!" + vbCrLf _
                 + "Du må vente til det blir din tur for å se hva ditt oppdrag er!"
    sPhraseDan 122, "Slap nu af!" + vbCrLf _
                 + "Du må vente til det bliver din tur før du kan læse din mission."

    'REM:
    sPhraseEng 123, "Conspiracy"
    sPhraseIta 123, "Cospirazione"
    sPhraseFra 123, "Conspiration"
    sPhraseGer 123, "Verschwörung"
    sPhraseSpa 123, "Conspiración"
    sPhraseSwe 123, "Sammansvärjning"
    sPhraseNor 123, "Konspirasjon"
    sPhraseDan 123, "Sammensværgelse"

    'REM:
    sPhraseEng 124, "This is classified information. If I tell you, I will have to kill you. Or you could wait for your turn to see what your own mission is."
    sPhraseIta 124, "Queste sono informazioni riservate. Se te lo dicessi, dovrei eliminarti. Oppure aspetta il tuo turno per vedere il tuo obiettivo."
    sPhraseFra 124, "Cette information est classée secrète. Si je vous le dis, je dois vous tuer. Ou vous pouvez attendre votre prochain tour pour voir ce que votre mission est."
    sPhraseGer 124, "Dies sind vertrauliche Informationen. Wenn ich es Ihnen sagen würde, müsste ich Sie töten. Oder Sie warten bis zum eigenen Zug, um zu sehen wie Ihre Mission lautet."
    sPhraseSpa 124, "Esto es información clasificada. Si se lo digo, deberé matarlo. O podría esperar su turno para ver cual es su misión."
    sPhraseSwe 124, "Det här är hemligstämplad information. Om jag berättar för dig så måste jag döda dig. Annars kan du vänta tills det är din tur för att få reda på vad ditt eget uppdrag är."
    sPhraseNor 124, "Dette er hemmelig. Dersom jeg forteller deg dette, må jeg faktisk drepe deg. Eller du kan vente til det blir din tur for å se hva ditt oppdrag er!"
    sPhraseDan 124, "Dette er yderst hemmeligt. Jeg kunne fortælle dig det, men så bliver jeg nød til at slå dig ihjel. Du kunne også vente med at se din mission til det bliver din tur"

    'REM:
    sPhraseEng 125, "NO!"
    sPhraseIta 125, "NO!"
    sPhraseFra 125, "NON!"
    sPhraseGer 125, "NEIN!"
    sPhraseSpa 125, "¡NO!"
    sPhraseSwe 125, "NEJ!"
    sPhraseNor 125, "NEI!"
    sPhraseDan 125, "NEJ!"

    'REM:
    sPhraseEng 126, "My mission! I finish it! You just wait until your turn and finish your own mission!"
    sPhraseIta 126, "Il mio obiettivo! L'ho terminato!... e aspetta il tuo turno... e finisci il tuo obiettivo!"
    sPhraseFra 126, "Ma mission! Je l'ai finie! Attendez juste votre tour et finissez votre propre mission!"
    sPhraseGer 126, "Meine Mission! Ich werde sie beenden. Sie müssen bis zum eigenen Zug warten und Ihre eigene Mission beenden!"
    sPhraseSpa 126, "¡Mi misión! ¡La finalicé! ¡Espere hasta su turno y finalice su misión!"
    sPhraseSwe 126, "Mitt uppdrag! Jag avslutar det! Du väntar tills det är din tur och avslutar ditt eget uppdrag!"
    sPhraseNor 126, "Dette er mitt oppdrag! JEG skal avslutte dette. Du må bare vente til det blir din tur og avslutte ditt eget oppdrag!"
    sPhraseDan 126, "Det er MIN mission! Vent til det bliver din tur. Så kan du se din EGEN mission"

    'REM:
    sPhraseEng 127, "Spying is NOT tolerated in war!" + vbCrLf _
                 + "Your spies were all captured and executed. They died slowly in one of our many torture chambers. Let that be a warning to you human scum! Your human armies will soon be exterminated, and our <Var.ExeName> world will be a better place for computer players to live!"
    sPhraseIta 127, "Spiare non è permesso in guerra!" + vbCrLf _
                 + "Tutte le tue spie sono state catturate ed eseguite. Sono morte lentamente in una delle nostre stanze delle torture. Che ti sia di avvertimento! Le tue armate umane saranno presto distrutte, e il mondo <Var.ExeName> sara' um posto migliore per i giocatori computer!"
    sPhraseFra 127, "Espionner n'est PAS toléré en temps de guerre!" + vbCrLf _
                 + "Vos espions ont tous été capturés et ont été exécutés. Ils sont morts dans d'affreuses douleurs dans l'une de nos nombreuses chambres de torture. Prenez ceci comme un avertissement, pauvre humain! Vos armées humaines seront bientôt exterminées, et notre monde <Var.ExeName> sera un monde où il fera bon vivre pour les joueurs ordinateurs!!!!"
    sPhraseGer 127, "Spionieren wird in diesem Krieg NICHT toleriert!" + vbCrLf _
                 + "Ihre Spione wurden alle gefangen genommen und exekutiert. Sie starben langsam in einer unserer vielen Folterkammern. Lassen Sie sich das eine Warnung sein, menschlicher Abschaum. Ihre menschlichen Armeen werden bald ausgerottet werden, und unsere <Var.ExeName>-Welt wird ein besserer Platz für Computerspieler zum Leben sein."
    sPhraseSpa 127, "¡NO se tolera el espionaje en la guerra!" + vbCrLf _
                 + "Sus espías fueron capturados y ejecutados. Murieron lentamente en nuestras cámaras de tortura. ¡Que sea una advertencia para ti, escoria humana! ¡Tus ejércitos humanos serán exterminados pronto, y nuestro mundo de Misión Riesgo será un mejor lugar para que vivan los jugadores computadoras!"
    sPhraseSwe 127, "Spioneri tolereras INTE i krig!" + vbCrLf _
                 + "Dina spioner har alla tillfångatagits och avrättats. De dog långsamt i en av mina många tortyrkamrar. Låt detta bli en varning för dig, ditt avskum! Dina arméer av människor kommer snart att utrotas och vår MissioRisk-värld kommer att bli ett bättre ställer för datorspelare att leva på!"
    sPhraseNor 127, "Bruk av spioner i krigen blir ikke godtatt!" + vbCrLf _
                 + "Samtlige av dine spioner ble tatt til fange og henrettet. De led en grusom og langsom død i ett av våre mange torturkammer. La dette være en advarsel til dere mennesker. Dine menneskelige hærstyrker vil snart bli utryddet, og denne verden vil bli et bedre sted for våre computer-styrker."
    sPhraseDan 127, "Dine spioner er blevet opdaget. Vent til det bliver din tur!"

    'REM:
    sPhraseEng 128, "You intolerable species"
    sPhraseIta 128, "Tu specie inferiore!"
    sPhraseFra 128, "Vous espèce intolérable"
    sPhraseGer 128, "Sie unerträgliche Spezies"
    sPhraseSpa 128, "Tu intolerable especie"
    sPhraseSwe 128, "Du outhärdliga art"
    sPhraseNor 128, "Dere utålmodige mennesker"
    sPhraseDan 128, "Snyder!"

    'REM:X's turn.
    sPhraseEng 129, "'s turn."
    sPhraseIta 129, "è di turno."
    sPhraseFra 129, " , c'est votre tour."
    sPhraseGer 129, " ist am Zug."
    sPhraseSpa 129, "tu turno."
    sPhraseSwe 129, "s tur."
    sPhraseNor 129, "s tur."
    sPhraseDan 129, "s tur"

    'REM:Controlling cards. Viewing cards. Looking at cards.
    sPhraseEng 130, "Checking cards"
    sPhraseIta 130, "Controllo delle carte"
    sPhraseFra 130, "Regarder les cartes."
    sPhraseGer 130, "Überprüft Karten"
    sPhraseSpa 130, "Tarjetas de comprobación"
    sPhraseSwe 130, "Kontrollerar kort"
    sPhraseNor 130, "Sjekker kortene"
    sPhraseDan 130, "Checker kort"

    'REM:You must trade cards.You must exchange cards.
    sPhraseEng 131, "YOU MUST" + vbCrLf _
                 + "TRADE CARDS"
    sPhraseIta 131, "DEVI CAMBIARE" + vbCrLf _
                 + "LE CARTE"
    sPhraseFra 131, "Vous devez " + vbCrLf _
                 + "échanger des cartes."
    sPhraseGer 131, "SIE MÜSSEN" + vbCrLf _
                 + "KARTEN" + vbCrLf _
                 + "TAUSCHEN"
    sPhraseSpa 131, "DEBE TOMAR" + vbCrLf _
                 + "CARTAS DE" + vbCrLf _
                 + "OCUPACIÓN"
    sPhraseSwe 131, "DU MÅSTE" + vbCrLf _
                 + "BYTA IN KORT"
    sPhraseNor 131, "DU MÅ" + vbCrLf _
                 + "VEKSLE KORT"
    sPhraseDan 131, "DU SKAL" + vbCrLf _
                 + "BYTTE KORT"

    'REM:Swap cards for 5 units?
    sPhraseEng 132, "Trade cards in for"
    sPhraseIta 132, "Scambi le carte per"
    sPhraseFra 132, "Echanger vos cartes contre"
    sPhraseGer 132, "Karten eintauschen für"
    sPhraseSpa 132, "Cartas de ocupación"
    sPhraseSwe 132, "Byta in kort mot"
    sPhraseNor 132, "Veksle kort for"
    sPhraseDan 132, "Vil du bytte disse kort til"

    'REM:
    sPhraseEng 133, " units?"
    sPhraseIta 133, " unita'?"
    sPhraseFra 133, " unités?"
    sPhraseGer 133, " Einheiten?"
    sPhraseSpa 133, " unidades?"
    sPhraseSwe 133, " enheter?"
    sPhraseNor 133, " enheter?"
    sPhraseDan 133, " arméer"

    'REM:
    sPhraseEng 134, "  Re-shuffling cards."
    sPhraseIta 134, "  Rimischio il mazzo."
    sPhraseFra 134, " Les cartes sont battues."
    sPhraseGer 134, " Mischen der Karten."
    sPhraseSpa 134, " Vuelva a barajar las cartas."
    sPhraseSwe 134, "  Blandar om korten."
    sPhraseNor 134, "  Stokker kortene."
    sPhraseDan 134, " Blander kort."

    'REM:
    sPhraseEng 135, "Used cards are being collected, re-shuffled and put back in the pack."
    sPhraseIta 135, "Le carte usate sono raccolte e rimischiate nel mazzo."
    sPhraseFra 135, "Les cartes utilisées sont ramassées, battues et remises dans le tas."
    sPhraseGer 135, "Benutze Risiko-Karten werden eingesannelt, gemischt und wieder in den Stapel gepackt."
    sPhraseSpa 135, "Las cartas usadas de Riesgo serán recogidas, barajadas y puestas bajo el mazo."
    sPhraseSwe 135, "Använda Risk-kort samlas ihop, blandas och läggs tillbaka i högen."
    sPhraseNor 135, "Brukte kort vil nå bli samlet inn, stokket og gjort klar."
    sPhraseDan 135, "Brugte kort samles ind og blandes igen."

    'REM:Don't show this message any more.
    sPhraseEng 136, "Don't show any more."
    sPhraseIta 136, "Non mostrare più questo messaggio"
    sPhraseFra 136, "Ne plus afficher."
    sPhraseGer 136, "Nicht mehr anzeigen."
    sPhraseSpa 136, "No me muestre más."
    sPhraseSwe 136, "Visa inte fler gånger."
    sPhraseNor 136, "Ikke vis igjen."
    sPhraseDan 136, "Vis ikke igen."

    'REM:
    sPhraseEng 137, "You need more players than that."
    sPhraseIta 137, "Servono piu' giocatori per questo."
    sPhraseFra 137, "Vous avez besoin de plus de joueurs que ça."
    sPhraseGer 137, "Sie benötigen mehr Spieler."
    sPhraseSpa 137, "Necesita mas jugadores que esos."
    sPhraseSwe 137, "Du behöver fler spelare än så."
    sPhraseNor 137, "Du må velge flere spillere enn det."
    sPhraseDan 137, "Der skal være flere spillere."

    'REM:
    sPhraseEng 138, "There are still some countries left."
    sPhraseIta 138, "Sono rimasti ancora dei territori."
    sPhraseFra 138, "Il y reste encore quelques pays."
    sPhraseGer 138, "Es sind noch einige Länder übrig."
    sPhraseSpa 138, "Restan algunos países."
    sPhraseSwe 138, "Det finns fortfarande länder kvar."
    sPhraseNor 138, "Det er fortsatt noen land igjen."
    sPhraseDan 138, "Der er stadig lande tilbage."

    'REM:
    sPhraseEng 139, "Setting up the board..."
    sPhraseIta 139, "Preparo la plancia..."
    sPhraseFra 139, "Préparation du plateau..."
    sPhraseGer 139, "Vorbereiten des Bretts..."
    sPhraseSpa 139, "Estableciéndose en el tablero..."
    sPhraseSwe 139, "Ställer i ordning brädet..."
    sPhraseNor 139, "Setter opp brettet..."
    sPhraseDan 139, "Fordeler lande..."

    'REM:
    sPhraseEng 140, "Name"
    sPhraseIta 140, "Nome"
    sPhraseFra 140, "Nom"
    sPhraseGer 140, "Name"
    sPhraseSpa 140, "Nombre"
    sPhraseSwe 140, "Namn"
    sPhraseNor 140, "Navn"
    sPhraseDan 140, "Navn"

    'REM:
    sPhraseEng 141, "Registration code"
    sPhraseIta 141, "Codice di registrazione"
    sPhraseFra 141, "Code d'enregistrement"
    sPhraseGer 141, "Registrierungscode"
    sPhraseSpa 141, "Código de inscripción"
    sPhraseSwe 141, "Registreringskod"
    sPhraseNor 141, "Registreringskode"
    sPhraseDan 141, "Registreringskode"

    'REM:Available scenarios.
    sPhraseEng 142, "Available wars"
    sPhraseIta 142, "Scenari disponibili"
    sPhraseFra 142, "Scénarios disponibles"
    sPhraseGer 142, "Verfügbare Kriege"
    sPhraseSpa 142, "Guerras posibles"
    sPhraseSwe 142, "Tillgängliga krig"
    sPhraseNor 142, "Tilgjengelige kriger"
    sPhraseDan 142, "Vælg krig"

    'REM:
    sPhraseEng 143, "Delete"
    sPhraseIta 143, "Cancella"
    sPhraseFra 143, "Effacer"
    sPhraseGer 143, "Löschen"
    sPhraseSpa 143, "Borrar"
    sPhraseSwe 143, "Ta bort"
    sPhraseNor 143, "Slett"
    sPhraseDan 143, "Slet"

    'REM:Open scenario
    sPhraseEng 144, "&Open war..."
    sPhraseIta 144, "A&pri scenario..."
    sPhraseFra 144, "Scénario Ouvert..."
    sPhraseGer 144, "Krieg ö&ffnen..."
    sPhraseSpa 144, "&Librar guerra..."
    sPhraseSwe 144, "Öppna krig..."
    sPhraseNor 144, "Åpne krig..."
    sPhraseDan 144, "Åben krig..."

    'REM:
    sPhraseEng 145, "&Make this the default war"
    sPhraseIta 145, "&Rendi questo lo scenario di default"
    sPhraseFra 145, "Guerre par défaut"
    sPhraseGer 145, "Als Standard-&Krieg einstellen"
    sPhraseSpa 145, "&Dejar esta guerra inconclusa"
    sPhraseSwe 145, "Gör detta till standardkrig"
    sPhraseNor 145, "Sett som standard"
    sPhraseDan 145, "Gør dette til standard krig"

    'REM:
    sPhraseEng 146, "Description"
    sPhraseIta 146, "Descrizione"
    sPhraseFra 146, "Description"
    sPhraseGer 146, "Beschreibung"
    sPhraseSpa 146, "Descripción"
    sPhraseSwe 146, "Beskrivning"
    sPhraseNor 146, "Beskrivelse"
    sPhraseDan 146, "Beskrivelse"

    'REM:Save scenario as...
    sPhraseEng 147, "Save war as..."
    sPhraseIta 147, "Salva scenario come..."
    sPhraseFra 147, "Sauvez scénario sous..."
    sPhraseGer 147, "Krieg speichern als..."
    sPhraseSpa 147, "Guarde esta guerra como..."
    sPhraseSwe 147, "Spara krig som..."
    sPhraseNor 147, "Lagre krig som..."
    sPhraseDan 147, "Gem krig"

    'REM:Title of scenario. English changed to just Title.
    sPhraseEng 148, "Title"
    sPhraseIta 148, "Titolo dello scenario"
    sPhraseFra 148, "Titre du scénario"
    sPhraseGer 148, "Titel des Krieges"
    sPhraseSpa 148, "Titulo de la guerra"
    sPhraseSwe 148, "Krigets namn"
    sPhraseNor 148, "Tittel på krig"
    sPhraseDan 148, "Krigens titel"

    'REM:
    sPhraseEng 149, "&Save"
    sPhraseIta 149, "&Salva"
    sPhraseFra 149, "Sauver"
    sPhraseGer 149, "&Speichern"
    sPhraseSpa 149, "&Guardar"
    sPhraseSwe 149, "Spara"
    sPhraseNor 149, "Lagre"
    sPhraseDan 149, "Gem"

    'REM:Player option box for the Red army. Player option box for the Green army.
    sPhraseEng 150, "Player option box for the "
    sPhraseIta 150, "Opzioni per le "
    sPhraseFra 150, "Boîte d'option du joueur de "
    sPhraseGer 150, "Spieler Options-Box für "
    sPhraseSpa 150, "Caja de opciones del jugador para "
    sPhraseSwe 150, "Spelaralternativ för "
    sPhraseNor 150, "Spillervalg for "
    sPhraseDan 150, "Vælg spiller til "

    'REM:Number of territories that this army starts with.
    sPhraseEng 151, "Number of countries this army starts with"
    sPhraseIta 151, "Numero di territori iniziali per le armate"
    sPhraseFra 151, "Nombre de pays avec lesquels cette armée commence"
    sPhraseGer 151, "Anzahl der Länder mit denen diese Armee beginnt."
    sPhraseSpa 151, "Numero de países con los que comienza esta ejercito"
    sPhraseSwe 151, "Antal länder denna armé inleder med"
    sPhraseNor 151, "Antall land denne hæren starter med"
    sPhraseDan 151, "Antal lande hver hær starter med"

    'REM:Increase starting countries. Decrease starting countries.
    sPhraseEng 152, "Increment/decrement starting countries"
    sPhraseIta 152, "Incrementa/decrementa i territori di partenza"
    sPhraseFra 152, "Augmentater/baisser nombre de pays."
    sPhraseGer 152, "Erhöhen/verringern der Start-Länder"
    sPhraseSpa 152, "Incrementar/decrementar países iniciales"
    sPhraseSwe 152, "Öka/minska antalet start-arméer"
    sPhraseNor 152, "Øker/reduserer antall land fra start"
    sPhraseDan 152, "Hæv/sænk antal lande"

    'REM:
    sPhraseEng 153, "Change the number of starting players"
    sPhraseIta 153, "Cambia il numero dei giocatori"
    sPhraseFra 153, "Changez le nombre de joueurs"
    sPhraseGer 153, "Anzahl der Spieler ändern"
    sPhraseSpa 153, "Cambie el número de jugadores iniciales"
    sPhraseSwe 153, "Ändra antalet spelare"
    sPhraseNor 153, "Endrer antall spillere"
    sPhraseDan 153, "Antal startende spillere"

    'REM:
    sPhraseEng 154, "How many players are going to war"
    sPhraseIta 154, "Quanti giocatori partecipano"
    sPhraseFra 154, "Combien de joueurs vont se battre"
    sPhraseGer 154, "Wie viele Spieler sich bekriegen werden"
    sPhraseSpa 154, "Cuantos jugadores van a la guerra"
    sPhraseSwe 154, "Hur många spelare ska kriga"
    sPhraseNor 154, "Hvor mange spillere som skal krige"
    sPhraseDan 154, "Antal spillere"

    'REM:
    sPhraseEng 155, "Random selection of the first player"
    sPhraseIta 155, "Selezione casuale del primo giocatore"
    sPhraseFra 155, "Sélection aléatoire du premier joueur"
    sPhraseGer 155, "Zufällige Auswahl des anfangenden Spielers"
    sPhraseSpa 155, "Selección aleatoria de los primeros jugadores"
    sPhraseSwe 155, "Välj förste spelare slumpvis"
    sPhraseNor 155, "Vilkårlig hvem som starter spillet"
    sPhraseDan 155, "Tilfældigt valg af første spiller"

    'REM:
    sPhraseEng 156, "Player 1 - the red army is the first player"
    sPhraseIta 156, "Giocatore 1 - Armate rosse sono le prime"
    sPhraseFra 156, "Joueur 1 - l'armée rouge est la première à jouer"
    sPhraseGer 156, "Spieler 1 - die rote Armee ist der anfangende Spieler"
    sPhraseSpa 156, "Jugador 1 - El ejercito rojo es el primer jugador"
    sPhraseSwe 156, "Spelare 1 - den röda armén spelar först"
    sPhraseNor 156, "Spiller 1 - Den Røde Hæren starter spillet"
    sPhraseDan 156, "Spiller 1 (den røde hær) starter"

    'REM:
    sPhraseEng 157, "The cards option box"
    sPhraseIta 157, "Opzioni per le carte"
    sPhraseFra 157, "Boîte d'option des cartes"
    sPhraseGer 157, "Die Karten-Optionen Box"
    sPhraseSpa 157, "Caja de tarjetas de opciones"
    sPhraseSwe 157, "Kortalternativrutan"
    sPhraseNor 157, "Valg (kort)"
    sPhraseDan 157, "Opsætning af kort"

    'REM:Other players cannot see your cards when selected.
    sPhraseEng 158, "Other players cannot see your cards when checked"
    sPhraseIta 158, "Gli altri giocatori non possono vedere le tue carte se selezionato"
    sPhraseFra 158, "Les autres joueurs ne peuvent pas voir vos cartes quand sélectionné."
    sPhraseGer 158, "Andere Spieler können Ihre Karten nicht sehen, falls ausgewählt"
    sPhraseSpa 158, "Los demas jugadores no podrán ver sus tarjetas cuando las active"
    sPhraseSwe 158, "Andra spelare kan inte se dina kort när detta är markerat"
    sPhraseNor 158, "Dersom valgt - motspillerne kan ikke se dine kort"
    sPhraseDan 158, "Andre spillere kan ikke se dine kort når de bliver checket"

    'REM:Cards are not issued when selected.
    sPhraseEng 159, "Cards are not issued when checked"
    sPhraseIta 159, "Le carte non sono estratte se selezionato"
    sPhraseFra 159, "Les cartes ne sont pas montrées quand sélectionné."
    sPhraseGer 159, "Karten werden nicht verteilt, falls ausgewählt"
    sPhraseSpa 159, "Las tarjetas no se mostrarán cuando las active"
    sPhraseSwe 159, "Kort ges inte ut om detta är markerat"
    sPhraseNor 159, "Dersom valgt - det deles ikke ut kort til spillerne"
    sPhraseDan 159, "Kort har ingen værdi"

    'REM:Value of cards do not change when selected.
    sPhraseEng 160, "Value of cards do not change when checked"
    sPhraseIta 160, "Il valore delle carte non cambia se e' selezionato"
    sPhraseFra 160, "La valeur des cartes ne change pas"
    sPhraseGer 160, "Wert der Karten verändert sich nicht, falls ausgewählt"
    sPhraseSpa 160, "El valor de la tarjetas no cambiará cuando las activen"
    sPhraseSwe 160, "Värdet på korten ändras inte när detta är markerat"
    sPhraseNor 160, "Dersom valgt - øker ikke verdien på kortene"
    sPhraseDan 160, "Korts værdi er fastsat"

    'REM:
    sPhraseEng 161, "Value of cards increases whith every set turned in"
    sPhraseIta 161, "Il valore delle carte aumenta a ogni tris"
    sPhraseFra 161, "La valeur des cartes augmente à chaque tour."
    sPhraseGer 161, "Wert der Karten erhöht sich mit jedem eingelösten Kartensatz"
    sPhraseSpa 161, "El valor de las tarjetas se incrementa con cada turno cumplido"
    sPhraseSwe 161, "Värdet på korten ökar för varje inbyte"
    sPhraseNor 161, "Dersom valgt - øker verdien på kortene for hvert sett som veksles inn"
    sPhraseDan 161, "Stigende værdi"

    'REM:
    sPhraseEng 162, "Sets the maximum card value"
    sPhraseIta 162, "Setta il valore massimo delle carte"
    sPhraseFra 162, "Valeur maximale des cartes"
    sPhraseGer 162, "Stellt den maximalen Kartenwert ein"
    sPhraseSpa 162, "Ponga la carta de mayor valor"
    sPhraseSwe 162, "Bestämmer kortens maximala värde"
    sPhraseNor 162, "Maksimal verdi for et sett med kort"
    sPhraseDan 162, "Den maksimale værdi for kort"

    'REM:
    sPhraseEng 163, "The war options box"
    sPhraseIta 163, "Opzioni per i combattimenti"
    sPhraseFra 163, "Les options de guerre"
    sPhraseGer 163, "Die Kriegs-Optionen Box"
    sPhraseSpa 163, "La caja de opciones de guerra"
    sPhraseSwe 163, "Krigsalternativrutan"
    sPhraseNor 163, "Valg (krigføring)"
    sPhraseDan 163, "Opsætning af krig"

    'REM:All players are issued with secret missions when selected.
    sPhraseEng 164, "All players are issued with secret missions when checked"
    sPhraseIta 164, "Tutti i giocatori hanno un obiettivo segreto se selezionato"
    sPhraseFra 164, "Tous les joueurs ont des missions secrètes quand sélectionné."
    sPhraseGer 164, "Alle Spieler bekommen geheime Missionen, falls ausgewählt"
    sPhraseSpa 164, "A todos los jugadores se les han decretado misiones secretas para ser revisadas"
    sPhraseSwe 164, "Alla spelare får hemliga uppdrag när detta är markerat"
    sPhraseNor 164, "Dersom valgt - alle spillere får tildelt hvert sitt hemmelige oppdrag"
    sPhraseDan 164, "Alle spillere får en mission når denne funktion er valgt"

    'REM:Limit army movements.
    sPhraseEng 165, "Limits army movements"
    sPhraseIta 165, "Limita gli spostamenti"
    sPhraseFra 165, "Les mouvements des armées sont limités"
    sPhraseGer 165, "Beschränkt die Bewegung der Armeen"
    sPhraseSpa 165, "Movimiento de ejercitos limitados"
    sPhraseSwe 165, "Begränsar trupprörelser"
    sPhraseNor 165, "Begrenser forflytning av egne tropper"
    sPhraseDan 165, "Begræns mulighederne for at flytte arméer"

    'REM:
    sPhraseEng 166, "Look at the attack dice and decide how many defence dice to throw"
    sPhraseIta 166, "Guarda i dadi d'attacco e decide quanti dadi in difesa tirare"
    sPhraseFra 166, "Regardez les dés de l'attaque et décidez combien de dés de défense lancer"
    sPhraseGer 166, "Aufgrund der Angriffswürfel entscheiden mit wievielen Würfeln verteidigt wird"
    sPhraseSpa 166, "Mire los dados atacantes y decida cuantos dados defensivos arrojar"
    sPhraseSwe 166, "Titta på anfallstärningarna och avgör hur många försvarstärningar som ska slås"
    sPhraseNor 166, "Velg antall terninger du vil forsvare deg med på bakgrunn av hvor mange terninger angriperen kaster"
    sPhraseDan 166, "Beslut om der skal forsvares med en eller to terninger ud fra angrebets terninger"

    'REM:The speed of the computer players.
    sPhraseEng 167, "Speed of computer players"
    sPhraseIta 167, "Velocita' dei giocatori del computer"
    sPhraseFra 167, "Vitesse des joueurs ordinateur"
    sPhraseGer 167, "Geschwindigkeit der Computergegner"
    sPhraseSpa 167, "Velocidad del jugador computadora"
    sPhraseSwe 167, "Hastighet på datorspelarna"
    sPhraseNor 167, "Hastighet på computer-spillerne"
    sPhraseDan 167, "Computers fart"

    'REM:
    sPhraseEng 168, "Speed at which the dice are thrown"
    sPhraseIta 168, "Velocita' di lancio dei dadi"
    sPhraseFra 168, "Vitesse à laquelle les dés sont jetés"
    sPhraseGer 168, "Geschwindigkeit mit der die Würfel geworfen werden"
    sPhraseSpa 168, "Velocidad a la que se arrojan los dados"
    sPhraseSwe 168, "Hastighet på tärningsslagen"
    sPhraseNor 168, "Hastighet på terningene"
    sPhraseDan 168, "Terningernes fart"

    'REM:Border indicates player's color. This can be annoying!
    sPhraseEng 169, "Border flashes player's color (can be annoying!)"
    sPhraseIta 169, "Il bordo indica il colore del giocatore di turno"
    sPhraseFra 169, "La couleur du cadre change avec le joueur (peut être génant!)"
    sPhraseGer 169, "Farbe des Spielers als Rand anzeigen (kann nerven!)"
    sPhraseSpa 169, "La frontera parpadea con el color del jugador(¡puede molestar!)"
    sPhraseSwe 169, "Ramen blinkar i spelarens färg (kan vara irriterande!)"
    sPhraseNor 169, "Fargen på spillets ramme settes til den aktive spillerens farge"
    sPhraseDan 169, "Ramme blinker i spillerens farve"

    'REM:
    sPhraseEng 170, "Start new war with these settings"
    sPhraseIta 170, "Comincia una nuova partita con questi settaggi"
    sPhraseFra 170, "Commencer une nouvelle guerre avec ces choix"
    sPhraseGer 170, "Neuen Krieg mit diesen Einstellungen starten"
    sPhraseSpa 170, "Reinicia una guerra nueva con estos ajustes"
    sPhraseSwe 170, "Inled ett nytt krig med dessa inställningar"
    sPhraseNor 170, "Start ny krig med disse innstillingene"
    sPhraseDan 170, "Start ny krig med denne opsætning"

    'REM:
    sPhraseEng 171, "Exit with the previous settings"
    sPhraseIta 171, "Esci con i settaggi precedenti"
    sPhraseFra 171, "Sortir en gardant les choix précédants"
    sPhraseGer 171, "Mit vorherigen Einstellungen beenden"
    sPhraseSpa 171, "Salir con los ajustes anteriores"
    sPhraseSwe 171, "Gå tillbaka med tidigare inställningar"
    sPhraseNor 171, "Avbryt og behold de gamle innstillingene"
    sPhraseDan 171, "Afbryd opsætning"

    'REM:Title of box notifying players that the cards are being reshuffled.
    sPhraseEng 172, "&Reshuffle notice"
    sPhraseIta 172, "&Mostra avviso quando rimischia"
    sPhraseFra 172, "Mélangez les notes"
    sPhraseGer 172, "&Mischen-Nachricht"
    sPhraseSpa 172, "&Noticias de cambio"
    sPhraseSwe 172, "Omblandningsmeddelande"
    sPhraseNor 172, "Beskjed om stokking av kort"
    sPhraseDan 172, "Gør opmærksom på blanding af kort"

    'REM:
    sPhraseEng 173, "&Select language..."
    sPhraseIta 173, "&Seleziona lingua..."
    sPhraseFra 173, "Langue choisie..."
    sPhraseGer 173, "S&prache wählen..."
    sPhraseSpa 173, "&Seleccione idioma..."
    sPhraseSwe 173, "Välj språk..."
    sPhraseNor 173, "Velg språk..."
    sPhraseDan 173, "Vælg sprog..."

    'REM:Replay last turn.
    sPhraseEng 174, "Start &turn again"
    sPhraseIta 174, "&Ricomincia il turno"
    sPhraseFra 174, "Recommencer le dernier tour."
    sPhraseGer 174, "&Runde erneut starten"
    sPhraseSpa 174, "Comience &turno denuevo"
    sPhraseSwe 174, "Börja om din tur"
    sPhraseNor 174, "Start runde på nytt"
    sPhraseDan 174, "Genstart tur"

    'REM:Reset the scenario.
    sPhraseEng 175, "&Reset war"
    sPhraseIta 175, "&Resetta lo scenario"
    sPhraseFra 175, "Réinitialisez la guerre"
    sPhraseGer 175, "&Neuanfang des &Kriegs"
    sPhraseSpa 175, "&Reanude guerra"
    sPhraseSwe 175, "Återställ krig"
    sPhraseNor 175, "Start krigen på nytt"
    sPhraseDan 175, "Genstart krig"

    'REM:
    sPhraseEng 176, "Capture players’ cards when you wipe out their last unit"
    sPhraseIta 176, "Ruba le carte quando il giocatore e' eliminato"
    sPhraseFra 176, "Capturez les cartes des joueurs quand vous détruisez leur dernière unité"
    sPhraseGer 176, "Erobern der Spieler-Karten wenn die letzte Einheit ausgelöscht wird"
    sPhraseSpa 176, "Capture tarjetas de jugadores cuando aniquile sus últimas unidades"
    sPhraseSwe 176, "Erövra spelarens kort när du slår ut deras sista enhet!"
    sPhraseNor 176, "Overta kort når du utrydder motstanders siste enhet"
    sPhraseDan 176, "Overtag modstanders kort når du udrydder ham"

    'REM:
    sPhraseEng 177, "Increasing"
    sPhraseIta 177, "Aumenta"
    sPhraseFra 177, "Croissant"
    sPhraseGer 177, "Erhöhend"
    sPhraseSpa 177, "Incrementando"
    sPhraseSwe 177, "Ökande"
    sPhraseNor 177, "Økende"
    sPhraseDan 177, "Stigende"

    'REM:
    sPhraseEng 178, "Capture"
    sPhraseIta 178, "Cattura"
    sPhraseFra 178, "Capturer"
    sPhraseGer 178, "Erobern"
    sPhraseSpa 178, "Capturar"
    sPhraseSwe 178, "Erövra"
    sPhraseNor 178, "Overta"
    sPhraseDan 178, "Overtag kort"

    'REM:
    sPhraseEng 179, "No supply lines"
    sPhraseIta 179, "Senza rifornimenti"
    sPhraseFra 179, "Aucune ligne de provision"
    sPhraseGer 179, "Keine Versorgungslinien"
    sPhraseSpa 179, "No hay líneas de reaprovisionamiento"
    sPhraseSwe 179, "Inga understödslinjer"
    sPhraseNor 179, "Ingen støtte"
    sPhraseDan 179, "En flytning"

    'REM:No restrictions are placed on army movements.
    sPhraseEng 180, "No restrictions on army movements"
    sPhraseIta 180, "Nessuna restrizione al movimento delle armate"
    sPhraseFra 180, "Aucune restriction sur les mouvements des armées"
    sPhraseGer 180, "Keine Einschränkung bei den Bewegungen der Armeen"
    sPhraseSpa 180, "No hay restricciones a los movimientos de los ejércitos"
    sPhraseSwe 180, "Inga begränsningar i trupprörelser"
    sPhraseNor 180, "Tillater støtte fra dine andre tropper"
    sPhraseDan 180, "Ingen restrektioner på flytning af arméer"

    'REM:Value of cards:
    sPhraseEng 181, "Current value: "
    sPhraseIta 181, "Valore della carta: "
    sPhraseFra 181, "Valeur des cartes:"
    sPhraseGer 181, "Derzeitiger Wert: "
    sPhraseSpa 181, "Valor actual: "
    sPhraseSwe 181, "Nuvarande värde: "
    sPhraseNor 181, "Gjeldende verdi: "
    sPhraseDan 181, "Nuværende værdi: "

    'REM:This program is registered to X.
    sPhraseEng 182, "This program is registered to "
    sPhraseIta 182, "Questo programma è registrato a "
    sPhraseFra 182, "Ce programme est enregistré pour "
    sPhraseGer 182, "Dieses Programm ist registriert für "
    sPhraseSpa 182, "Este programa está inscripto a "
    sPhraseSwe 182, "Detta program är registrerat till "
    sPhraseNor 182, "Dette programmet er registrert på "
    sPhraseDan 182, "Dette program er registreret til "

    'REM:Thank you for registering <Var.ExeName>.
    sPhraseEng 183, "Thank you for registering."
    sPhraseIta 183, "Grazie per aver registrato."
    sPhraseFra 183, "Merci de vous être enregistré."
    sPhraseGer 183, "Vielen Dank, dass Sie sich registriert haben."
    sPhraseSpa 183, "Gracias por registrarse."
    sPhraseSwe 183, "Tack för att du registrerat."
    sPhraseNor 183, "Takk for at du registrerte deg."
    sPhraseDan 183, "Tak fordi du registrerede <Var.ExeName>"

    'REM:Transfer ALL units. French: Toutes les
    sPhraseEng 184, "A&LL"
    sPhraseIta 184, "&Tutto"
    sPhraseFra 184, "toutes"
    sPhraseGer 184, "All&e"
    sPhraseSpa 184, "&todos"
    sPhraseSwe 184, "ALLA"
    sPhraseNor 184, "Alle"
    sPhraseDan 184, "Alle"

    'REM:Needs translation.
    sPhraseEng 185, "Copyright© 2014 Doug Burner, GlobalSiege.net"
    sPhraseIta 185, "Copyright© 2014 Doug Burner, GlobalSiege.net"
    sPhraseFra 185, "Copyright© 2014 Doug Burner, GlobalSiege.net"
    sPhraseGer 185, "Copyright© 2014 Doug Burner, GlobalSiege.net"
    sPhraseSpa 185, "Copyright© 2014 Doug Burner, GlobalSiege.net"
    sPhraseSwe 185, "Copyright© 2014 Doug Burner, GlobalSiege.net"
    sPhraseNor 185, "Copyright© 2014 Doug Burner, GlobalSiege.net"
    sPhraseDan 185, "Copyright© 2014 Doug Burner, GlobalSiege.net"

    'REM:
    sPhraseEng 186, "Distribute extra starting units"
    sPhraseIta 186, "Distribuisci unita' extra all'inizio"
    sPhraseFra 186, "Distribuez des unités supplémentaires au début"
    sPhraseGer 186, "Extra Starteinheiten verteilen"
    sPhraseSpa 186, "Distribuya unidades extra iniciales"
    sPhraseSwe 186, "Sätt ut extra trupper"
    sPhraseNor 186, "Ekstra enheter ved start"
    sPhraseDan 186, "Distribuer ekstra arméer"

    'REM:
    sPhraseEng 187, "Extra starting units are randomly distributed over each players' territories."
    sPhraseIta 187, "Le unita' extra vengono distribuite casualmente sui territori dei giocatori all'inizio"
    sPhraseFra 187, "Les unités supplémentaires sont distribuées aléatoirement sur les territoires de chaques joueurs."
    sPhraseGer 187, "Extra Starteinheiten werden zufällig über das Territorium jedes Spielers verteilt."
    sPhraseSpa 187, "Unidades extras iniciales están aleatoriamente distribuidas en cada uno de los territorios de los jugadores."
    sPhraseSwe 187, "Extra trupper fördelas slumpvis över vardera spelarens territorier."
    sPhraseNor 187, "Ved start fordeles ekstra enheter vilkårlig over alle spillernes territorier."
    sPhraseDan 187, "Ekstra arméer distribueres tilfældigt til hver spiller."

    'REM:
    sPhraseEng 188, "Extra units"
    sPhraseIta 188, "Unita' extra"
    sPhraseFra 188, "Unités supplémentaires"
    sPhraseGer 188, "Extra Einheiten"
    sPhraseSpa 188, "Unidades extra"
    sPhraseSwe 188, "Extra enheter"
    sPhraseNor 188, "Ekstra enheter"
    sPhraseDan 188, "Ekstra arméer"

    'REM:X wants to keep going.
    sPhraseEng 189, " wants to keep going"
    sPhraseIta 189, " vogliono continuare"
    sPhraseFra 189, " veut continuer"
    sPhraseGer 189, " wants to keep going"
    sPhraseSpa 189, " desea continuar"
    sPhraseSwe 189, " vill fortsätta"
    sPhraseNor 189, " ønsker å fortsette"
    sPhraseDan 189, " vil fortsætte"

    'REM:X wants to keep the mission a secret!
    sPhraseEng 190, " wants to keep the mission a secret!"
    sPhraseIta 190, " vuole tenere segreto l'obiettivo!"
    sPhraseFra 190, " veut garder la mission secrète!"
    sPhraseGer 190, " möchte, dass die Mission geheim bleibt!"
    sPhraseSpa 190, " desea mantener la misión como secreta"
    sPhraseSwe 190, " vill hålla sitt uppdrag hemligt!"
    sPhraseNor 190, " ønsker å hemmeligholde sitt oppdarg!"
    sPhraseDan 190, " vil ikke afsløre sin mission!"

    'REM:Quit this scenario?
    sPhraseEng 191, "Do you really want to quit this war?"
    sPhraseIta 191, "Vuoi lasciare questo scenario?"
    sPhraseFra 191, "Quitter ce scénario?"
    sPhraseGer 191, "Möchten Sie diesen Krieg wirklich beenden?"
    sPhraseSpa 191, "Realmente desea abandonar esta guerra"
    sPhraseSwe 191, "Vill du verkligen avsluta det här kriget?"
    sPhraseNor 191, "Ønsker du virkelig å avslutte denne krigen?"
    sPhraseDan 191, "Vil du virkelig afslutte denne krig?"

    'REM:Do you relly want to delete X?
    sPhraseEng 192, "Are you sure you want to delete "
    sPhraseIta 192, "Sei sicuro di cancellare "
    sPhraseFra 192, "Voulez vous vraiment effacer "
    sPhraseGer 192, "Sind Sie sicher, dass Sie diesen Krieg löschen möchten: "
    sPhraseSpa 192, "Está seguro que desea borrar "
    sPhraseSwe 192, "Är du säker på att du vill ta bort "
    sPhraseNor 192, "Er du sikker på at du ønsker å slette "
    sPhraseDan 192, "Vil du virkelig slette "

    'REM:Are you serious?
    sPhraseEng 193, "Are you sure?"
    sPhraseIta 193, "Sei sicuro?"
    sPhraseFra 193, "Est-ce que vous êtes sûr?"
    sPhraseGer 193, "Sind Sie sicher?"
    sPhraseSpa 193, "¿Está seguro?"
    sPhraseSwe 193, "Är du säker?"
    sPhraseNor 193, "Er du sikker?"
    sPhraseDan 193, "Er du sikker?"

    'REM:
    sPhraseEng 194, "Open war"
    sPhraseIta 194, "Apri scenario"
    sPhraseFra 194, "Guerre ouverte"
    sPhraseGer 194, "Krieg öffnen"
    sPhraseSpa 194, "librar Guerra"
    sPhraseSwe 194, "Öppna krig"
    sPhraseNor 194, "Åpne krig"
    sPhraseDan 194, "Åben krig"

    'REM:
    sPhraseEng 195, ""
    sPhraseIta 195, ""
    sPhraseFra 195, ""
    sPhraseGer 195, ""
    sPhraseSpa 195, ""
    sPhraseSwe 195, ""
    sPhraseNor 195, ""
    sPhraseDan 195, ""

    'REM:
    sPhraseEng 196, "AI interupt request"
    sPhraseIta 196, "L'algoritmo di intelligenza artificiale"
    sPhraseFra 196, "L'AI demande une interuption"
    sPhraseGer 196, "KI interupt request"
    sPhraseSpa 196, "Requerimiento de interrupcion de IA"
    sPhraseSwe 196, "AI avbrottsförfrågan"
    sPhraseNor 196, "Etterretningsinformasjon"
    sPhraseDan 196, "AI interupt request"

    'REM:Version 2.5.
    sPhraseEng 197, "Version "
    sPhraseIta 197, "Versione "
    sPhraseFra 197, "Version "
    sPhraseGer 197, "Version "
    sPhraseSpa 197, "Versión "
    sPhraseSwe 197, "Version "
    sPhraseNor 197, "Versjon "
    sPhraseDan 197, "Version "

    'REM:Hi X.
    sPhraseEng 198, "Hi "
    sPhraseIta 198, "Ciao "
    sPhraseFra 198, "Bonjour "
    sPhraseGer 198, "Hi "
    sPhraseSpa 198, "Hola "
    sPhraseSwe 198, "Hej "
    sPhraseNor 198, "Hei "
    sPhraseDan 198, "Goddag "

    'REM:
    sPhraseEng 199, ""
    sPhraseIta 199, ""
    sPhraseFra 199, ""
    sPhraseGer 199, ""
    sPhraseSpa 199, ""
    sPhraseSwe 199, ""
    sPhraseNor 199, ""
    sPhraseDan 199, ""
    Call initialisePhrases200
End Sub

Private Sub initialisePhrases200()
    'REM:
    sPhraseEng 200, "&Open"
    sPhraseIta 200, "A&pri scenario"
    sPhraseFra 200, "Ouvrir"
    sPhraseGer 200, "Ö&ffnen"
    sPhraseSpa 200, "&Abrir"
    sPhraseSwe 200, "Öppna"
    sPhraseNor 200, "Åpne"
    sPhraseDan 200, "Åben"

    'REM:
    sPhraseEng 201, "&Tool bar"
    sPhraseIta 201, "&Tool bar"
    sPhraseFra 201, "Boîte à outils"
    sPhraseGer 201, "&Button-Leiste"
    sPhraseSpa 201, "&Caja de herramientas"
    sPhraseSwe 201, "Verktygsfält"
    sPhraseNor 201, "Verktøylinje"
    sPhraseDan 201, "Værktøjslinje"

    'REM:
    sPhraseEng 202, "New war"
    sPhraseIta 202, "Nuova partita"
    sPhraseFra 202, "Nouvelle guerre"
    sPhraseGer 202, "Neuer Krieg"
    sPhraseSpa 202, "Nueva guerra"
    sPhraseSwe 202, "Nytt krig"
    sPhraseNor 202, "Ny krig"
    sPhraseDan 202, "Ny krig"

    'REM:
    sPhraseEng 203, "Reset war"
    sPhraseIta 203, "Resetta lo scenario"
    sPhraseFra 203, "Réinitialisez la guerre"
    sPhraseGer 203, "Neuanfang des Krieges"
    sPhraseSpa 203, "Reanudar guerra"
    sPhraseSwe 203, "Börja om krig"
    sPhraseNor 203, "Start på nytt"
    sPhraseDan 203, "Genstart krig"

    'REM:
    sPhraseEng 204, "Open"
    sPhraseIta 204, "Apri scenario"
    sPhraseFra 204, "Ouvrir"
    sPhraseGer 204, "Öffnen"
    sPhraseSpa 204, "Abrir"
    sPhraseSwe 204, "Öppna"
    sPhraseNor 204, "Åpne"
    sPhraseDan 204, "Åben"

    'REM:
    sPhraseEng 205, "Save"
    sPhraseIta 205, "Salva"
    sPhraseFra 205, "Sauver"
    sPhraseGer 205, "Speichern"
    sPhraseSpa 205, "Guardar"
    sPhraseSwe 205, "Spara"
    sPhraseNor 205, "Lagre"
    sPhraseDan 205, "Gem"

    'REM:Fast war. Press + and - to adjust speed.
    sPhraseEng 206, "Fast war (press + and - to adjust speed)"
    sPhraseIta 206, "Guerra veloce"
    sPhraseFra 206, "Guerre rapide (presser + et - pour ajuster la vitesse)"
    sPhraseGer 206, "Schneller Krieg (+ und - drücken, um Geschwindigkeit anzupassen)"
    sPhraseSpa 206, "Guerra rápida (pulse + y - para ajustar la velocidad)"
    sPhraseSwe 206, "Snabbt krig (tryck på + och - för att justera hastigheten)"
    sPhraseNor 206, "Hurtig krig (trykk +/- for å justere hastigheten)"
    sPhraseDan 206, "Hurtig krig (juster farten med + og -)"

    'REM:Fast dice. Press + and - to adjust speed.
    sPhraseEng 207, "Fast dice (press + and - to adjust speed)"
    sPhraseIta 207, "Dadi veloci"
    sPhraseFra 207, "Dés rapides (presser + et - pour ajuster la vitesse)"
    sPhraseGer 207, "Schnelles Würfeln (+ und - drücken, um Geschwindigkeit anzupassen)"
    sPhraseSpa 207, "Dados rápidos (pulse + y - para ajustar la velocidad)"
    sPhraseSwe 207, "Snabba tärningsslag (tryck på + och - för att justera hastigheten)"
    sPhraseNor 207, "Hurtige terninger (trykk +/- for å justere hastigheten)"
    sPhraseDan 207, "Hurtige terninger (juster fart med + og -)"

    'REM:
    sPhraseEng 208, "Auto scroll"
    sPhraseIta 208, "Auto scroll"
    sPhraseFra 208, "Défilement automatique"
    sPhraseGer 208, "Automatisches Scrollen"
    sPhraseSpa 208, "Auto scroll"
    sPhraseSwe 208, "Automatisk scrollning"
    sPhraseNor 208, "Auto scroll"
    sPhraseDan 208, "Auto scroll"

    'REM:Replay last turn.
    sPhraseEng 209, "Start turn again"
    sPhraseIta 209, "Ricomincia il turno"
    sPhraseFra 209, "Recommencer le dernier tour"
    sPhraseGer 209, "Runde erneut starten"
    sPhraseSpa 209, "Inicia turno de nuevo"
    sPhraseSwe 209, "Börja om din tur"
    sPhraseNor 209, "Start runde på nytt"
    sPhraseDan 209, "Genstart tur"

    'REM:
    sPhraseEng 210, "Help"
    sPhraseIta 210, "Aiuto"
    sPhraseFra 210, "Aide"
    sPhraseGer 210, "Hilfe"
    sPhraseSpa 210, "Ayuda"
    sPhraseSwe 210, "Hjälp"
    sPhraseNor 210, "Hjelp"
    sPhraseDan 210, "Hjælp"

    'REM:
    sPhraseEng 211, ""
    sPhraseIta 211, ""
    sPhraseFra 211, ""
    sPhraseGer 211, ""
    sPhraseSpa 211, ""
    sPhraseSwe 211, ""
    sPhraseNor 211, ""
    sPhraseDan 211, ""

    'REM:
    sPhraseEng 212, "New version update"
    sPhraseIta 212, "Aggiornamento alla nuova versione"
    sPhraseFra 212, "Nouvelle mise à jour"
    sPhraseGer 212, "Neue Version"
    sPhraseSpa 212, "Actualización de nueva versión"
    sPhraseSwe 212, "Ny versionsuppdatering"
    sPhraseNor 212, "Oppdatering til ny versjon"
    sPhraseDan 212, "Ny version"

    'REM:
    sPhraseEng 213, ""
    sPhraseIta 213, ""
    sPhraseFra 213, ""
    sPhraseGer 213, ""
    sPhraseSpa 213, ""
    sPhraseSwe 213, ""
    sPhraseNor 213, ""
    sPhraseDan 213, ""

    'REM:
    sPhraseEng 214, "" + vbCrLf _
                 + "    > PAUSED <"
    sPhraseIta 214, "" + vbCrLf _
                 + "   > IN PAUSA <"
    sPhraseFra 214, "" + vbCrLf _
                 + "     > PAUSE <"
    sPhraseGer 214, "" + vbCrLf _
                 + "    > PAUSE <"
    sPhraseSpa 214, "" + vbCrLf _
                 + "  > Detenido <"
    sPhraseSwe 214, "" + vbCrLf _
                 + "    > PAUSAD <"
    sPhraseNor 214, "" + vbCrLf _
                 + "    > PAUSE <"
    sPhraseDan 214, "" + vbCrLf _
                 + "    > PAUSE <"

    'REM:
    sPhraseEng 215, "Could not find a browser."
    sPhraseIta 215, "Non trovo un browser."
    sPhraseFra 215, "Impossible de trouver un navigateur."
    sPhraseGer 215, "Es konnte kein Browser gefunden werden."
    sPhraseSpa 215, "No se puede encontrar un browser."
    sPhraseSwe 215, "Kunde inte hitta en web-läsare."
    sPhraseNor 215, "Kunne ikke finne en nettleser (browser) på din datamaskin."
    sPhraseDan 215, "Kan ikke finde en browser."

    'REM:
    sPhraseEng 216, "Web Page '<Var.HomePage>' not Opened"
    sPhraseIta 216, "Indirizzo Web: '<Var.HomePage>' non aperto."
    sPhraseFra 216, "La page Web ' http:/ /www.missionrisk.les com' n'est pas ouverte"
    sPhraseGer 216, "Webpage '<Var.HomePage>' nicht geöffnet"
    sPhraseSpa 216, "Página web '<Var.HomePage>' no abierta"
    sPhraseSwe 216, "Websidan '<Var.HomePage>' har inte öppnats"
    sPhraseNor 216, "Internett-siden '<Var.HomePage>' ble ikke åpnet"
    sPhraseDan 216, "Hjemmeside '<Var.HomePage>' blev ikke åbnet"

    'REM:
    sPhraseEng 217, "Web Page 'https://www.regnow.com/softsell/nph-softsell.cgi?item=1584-1' not Opened"
    sPhraseIta 217, "Pagina Web 'https://www.regnow.com/softsell/nph-softsell.cgi?item=1584-1' non aperta"
    sPhraseFra 217, "La page Web ' https:/ /www.regnow.com/softsell/nph-softsell.cgi?item=1584-1 ' n'est pas ouverte"
    sPhraseGer 217, "Webpage 'https://www.regnow.com/softsell/nph-softsell.cgi?item=1584-1' nicht geöffnet"
    sPhraseSpa 217, "Página web 'https://www.regnow.com/softsell/nph-softsell.cgi?item=1584-1' no abierta"
    sPhraseSwe 217, "Websidan 'https://www.regnow.com/softsell/nph-softsell.cgi?item=1584-1' har inte öppnats"
    sPhraseNor 217, "Internett-siden 'https://www.regnow.com/softsell/nph-softsell.cgi?item=1584-1' ble ikke åpnet"
    sPhraseDan 217, "Hjemmesiden 'https://www.regnow.com/softsell/nph-softsell.cgi?item=1584-1' blev ikke åbnet"

    'REM:A file error has occured. Perhaps X has been corrupted in some way.
    sPhraseEng 218, "A file error has occured. Perhaps "
    sPhraseIta 218, "Si è verificato un errore nei file. Forse "
    sPhraseFra 218, "Une erreur de fichier est apparue. Peut-être que le fichier "
    sPhraseGer 218, "Ein Dateifehler ist aufgetreten."
    sPhraseSpa 218, "Ocurrió un error de archivo. A lo mejor "
    sPhraseSwe 218, "Ett filfel har inträffat. Kanske "
    sPhraseNor 218, "En feil har oppstått. Kanskje "
    sPhraseDan 218, "En fil fejl er forekommet. Filen "

    'REM:
    sPhraseEng 219, " has been corrupted in some way."
    sPhraseIta 219, " e' danneggiato."
    sPhraseFra 219, " a été corrompu."
    sPhraseGer 219, " wurde korrumpiert."
    sPhraseSpa 219, " ha sido corrompido en alguna forma."
    sPhraseSwe 219, " har förstörts på något sätt."
    sPhraseNor 219, " er skadet/slettet."
    sPhraseDan 219, " virker ikke."

    'REM:
    sPhraseEng 220, "Error"
    sPhraseIta 220, "Errore"
    sPhraseFra 220, "Erreur"
    sPhraseGer 220, "Fehler"
    sPhraseSpa 220, "Error"
    sPhraseSwe 220, "Fel"
    sPhraseNor 220, "Feil"
    sPhraseDan 220, "Fejl"

    'REM:
    sPhraseEng 221, "Setup"
    sPhraseIta 221, "Setup"
    sPhraseFra 221, "Installation"
    sPhraseGer 221, "Einstellungen"
    sPhraseSpa 221, "Ajuste"
    sPhraseSwe 221, "Inställningar"
    sPhraseNor 221, "Oppsett"
    sPhraseDan 221, "Opsætning"

    'REM:
    sPhraseEng 222, ""
    sPhraseIta 222, ""
    sPhraseFra 222, ""
    sPhraseGer 222, ""
    sPhraseSpa 222, ""
    sPhraseSwe 222, ""
    sPhraseNor 222, ""
    sPhraseDan 222, ""

    'REM:Mission 7 North and South America
    sPhraseEng 223, "North and South America."
    sPhraseIta 223, "Nord e Sud America."
    sPhraseFra 223, "Amérique du Nord et du Sud."
    sPhraseGer 223, "Nord-und Südamerika."
    sPhraseSpa 223, "América del Norte y del Sur."
    sPhraseSwe 223, "Nord-och Sydamerika."
    sPhraseNor 223, "Nord-og Sør-Amerika."
    sPhraseDan 223, "Nord-og Sydamerika."

    'REM:Mission 8 North America And Australia
    sPhraseEng 224, "<Var.Phrase(423)> and <Var.Phrase(428)>."
    sPhraseIta 224, "<Var.Phrase(423)> e <Var.Phrase(428)>."
    sPhraseFra 224, "<Var.Phrase(423)> et <Var.Phrase(428)>."
    sPhraseGer 224, "<Var.Phrase(423)> und <Var.Phrase(428)>."
    sPhraseSpa 224, "<Var.Phrase(423)> y <Var.Phrase(428)>."
    sPhraseSwe 224, "<Var.Phrase(423)> och <Var.Phrase(428)>."
    sPhraseNor 224, "<Var.Phrase(423)> og <Var.Phrase(428)>."
    sPhraseDan 224, "<Var.Phrase(423)> og <Var.Phrase(428)>."

    'REM:Mission 9 South America And Europe
    sPhraseEng 225, "<Var.Phrase(424)> and <Var.Phrase(425)>."
    sPhraseIta 225, "<Var.Phrase(424)> e <Var.Phrase(425)>."
    sPhraseFra 225, "<Var.Phrase(424)> et <Var.Phrase(425)>."
    sPhraseGer 225, "<Var.Phrase(424)> und <Var.Phrase(425)>."
    sPhraseSpa 225, "<Var.Phrase(424)> y <Var.Phrase(425)>."
    sPhraseSwe 225, "<Var.Phrase(424)> och <Var.Phrase(425)>."
    sPhraseNor 225, "<Var.Phrase(424)> og <Var.Phrase(425)>."
    sPhraseDan 225, "<Var.Phrase(424)> og <Var.Phrase(425)>."

    'REM:Mission 10 Europe and Australia
    sPhraseEng 226, "<Var.Phrase(425)> and <Var.Phrase(428)>."
    sPhraseIta 226, "<Var.Phrase(425)> e <Var.Phrase(428)>."
    sPhraseFra 226, "<Var.Phrase(425)> et <Var.Phrase(428)>."
    sPhraseGer 226, "<Var.Phrase(425)> und <Var.Phrase(428)>."
    sPhraseSpa 226, "<Var.Phrase(425)> y <Var.Phrase(428)>."
    sPhraseSwe 226, "<Var.Phrase(425)> och <Var.Phrase(428)>."
    sPhraseNor 226, "<Var.Phrase(425)> og <Var.Phrase(428)>."
    sPhraseDan 226, "<Var.Phrase(425)> og <Var.Phrase(428)>."

    'REM:Mission 11 South America And Africa
    sPhraseEng 227, "<Var.Phrase(424)> and <Var.Phrase(426)>."
    sPhraseIta 227, "<Var.Phrase(424)> e <Var.Phrase(426)>."
    sPhraseFra 227, "<Var.Phrase(424)> et <Var.Phrase(426)>."
    sPhraseGer 227, "<Var.Phrase(424)> und <Var.Phrase(426)>."
    sPhraseSpa 227, "<Var.Phrase(424)> y <Var.Phrase(426)>."
    sPhraseSwe 227, "<Var.Phrase(424)> och <Var.Phrase(426)>."
    sPhraseNor 227, "<Var.Phrase(424)> og <Var.Phrase(426)>."
    sPhraseDan 227, "<Var.Phrase(424)> og <Var.Phrase(426)>."

    'REM:Mission 12 Africa and Australia
    sPhraseEng 228, "<Var.Phrase(426)> and <Var.Phrase(428)>."
    sPhraseIta 228, "<Var.Phrase(426)> e <Var.Phrase(428)>."
    sPhraseFra 228, "<Var.Phrase(426)> et <Var.Phrase(428)>."
    sPhraseGer 228, "<Var.Phrase(426)> und <Var.Phrase(428)>."
    sPhraseSpa 228, "<Var.Phrase(426)> y <Var.Phrase(428)>."
    sPhraseSwe 228, "<Var.Phrase(426)> och <Var.Phrase(428)>."
    sPhraseNor 228, "<Var.Phrase(426)> og <Var.Phrase(428)>."
    sPhraseDan 228, "<Var.Phrase(426)> og <Var.Phrase(428)>."

    'REM:Mission 13 South America And Australia
    sPhraseEng 229, "<Var.Phrase(424)> and <Var.Phrase(428)>."
    sPhraseIta 229, "<Var.Phrase(424)> e <Var.Phrase(428)>."
    sPhraseFra 229, "<Var.Phrase(424)> et <Var.Phrase(428)>."
    sPhraseGer 229, "<Var.Phrase(424)> und <Var.Phrase(428)>."
    sPhraseSpa 229, "<Var.Phrase(424)> y <Var.Phrase(428)>."
    sPhraseSwe 229, "<Var.Phrase(424)> och <Var.Phrase(428)>."
    sPhraseNor 229, "<Var.Phrase(424)> og <Var.Phrase(428)>."
    sPhraseDan 229, "<Var.Phrase(424)> og <Var.Phrase(428)>."

    'REM:by occupying and holding 24 countries - Mission14.
    sPhraseEng 230, "<Var.Mission14> countries."
    sPhraseIta 230, "<Var.Mission14> territori."
    sPhraseFra 230, "<Var.Mission14> pays."
    sPhraseGer 230, "<Var.Mission14> Länder."
    sPhraseSpa 230, "<Var.Mission14> países."
    sPhraseSwe 230, "<Var.Mission14> länder."
    sPhraseNor 230, "<Var.Mission14> land."
    sPhraseDan 230, "<Var.Mission14> lande."

    'REM:Computer 5 has connected.
    sPhraseEng 231, " has connected."
    sPhraseIta 231, " e' connesso."
    sPhraseFra 231, " s'est connecté."
    sPhraseGer 231, " hat sich verbunden."
    sPhraseSpa 231, " se conectó."
    sPhraseSwe 231, " har anslutit."
    sPhraseNor 231, " er tilkoblet."
    sPhraseDan 231, " har forbindelse."

    'REM:
    sPhraseEng 232, "An error has occured"
    sPhraseIta 232, "Si e' verificato un errore"
    sPhraseFra 232, "Une erreur est apparue"
    sPhraseGer 232, "Ein Fehler ist aufgetreten"
    sPhraseSpa 232, "Ha ocurrido un error"
    sPhraseSwe 232, "Ett fel har inträffat"
    sPhraseNor 232, "En feil har oppstått"
    sPhraseDan 232, "En fejl har fundet sted"

    'REM:Computer. Work station. Workstation.
    sPhraseEng 233, "Terminal "
    sPhraseIta 233, "Terminale "
    sPhraseFra 233, "Terminal "
    sPhraseGer 233, "Terminal "
    sPhraseSpa 233, "Terminal "
    sPhraseSwe 233, "Terminal "
    sPhraseNor 233, "Datamaskin "
    sPhraseDan 233, "Computer "

    'REM:
    sPhraseEng 234, ""
    sPhraseIta 234, ""
    sPhraseFra 234, ""
    sPhraseGer 234, ""
    sPhraseSpa 234, ""
    sPhraseSwe 234, ""
    sPhraseNor 234, ""
    sPhraseDan 234, ""

    'REM:
    sPhraseEng 235, "Host"
    sPhraseIta 235, "Host"
    sPhraseFra 235, "Hôte"
    sPhraseGer 235, "Host"
    sPhraseSpa 235, "Anfitrión"
    sPhraseSwe 235, "Värd"
    sPhraseNor 235, "Vert"
    sPhraseDan 235, "Vært"

    'REM: You are terminal 4
    sPhraseEng 236, " >> Your are "
    sPhraseIta 236, " >> Tu sei "
    sPhraseFra 236, " >> Vous êtes "
    sPhraseGer 236, " >> Sie sind "
    sPhraseSpa 236, " >> Son sus "
    sPhraseSwe 236, " >> Ditt är"
    sPhraseNor 236, " >> Du er "
    sPhraseDan 236, " >> Du er "

    'REM:The computer's Host Name is 555.
    sPhraseEng 237, "Your Local Host Name is "
    sPhraseIta 237, "Il nome Terminal Host e' "
    sPhraseFra 237, "Le nom d'hôte terminal est "
    sPhraseGer 237, "Terminal Host Name ist"
    sPhraseSpa 237, "El Nombre del Terminal Anfitrión es "
    sPhraseSwe 237, "Terminalens värdnamn är "
    sPhraseNor 237, "Datamaskinens vertsnavn er "
    sPhraseDan 237, "Computerens vært navn er "

    'REM:
    sPhraseEng 238, "Click <IP Config> for more information."
    sPhraseIta 238, "Clicca su <IP Config> per maggiori informazioni."
    sPhraseFra 238, "Clickez <IP Config> pour plus d'information."
    sPhraseGer 238, "Klicken Sie auf <IP Config> für weitere Informationen."
    sPhraseSpa 238, "Pulse <IP Config> para mas información."
    sPhraseSwe 238, "Klicka på <IP-information> för mer information."
    sPhraseNor 238, "Klikk <IP Config> for mer informasjon."
    sPhraseDan 238, "Klik <IP Config> for mere information."

    'REM:
    sPhraseEng 239, "Do you want to close all connections?"
    sPhraseIta 239, "Vuoi chiudere tutte le connessioni?"
    sPhraseFra 239, "Est-ce que vous voulez fermer toutes les connexions?"
    sPhraseGer 239, "Möchten Sie alle Verbindungen schliessen?"
    sPhraseSpa 239, "¿Desea cerrar todas las conecciones?"
    sPhraseSwe 239, "Vill du stänga alla anslutningar?"
    sPhraseNor 239, "Ønsker du å lukke alle tilkoblinger?"
    sPhraseDan 239, "Vil du lukke for alle forbindelser?"

    'REM:
    sPhraseEng 240, "Please enter the name or IP address of the host"
    sPhraseIta 240, "Per favore inserisci il nome o l'indirizzo IP dell'host"
    sPhraseFra 240, "S'il vous plaît entrez le nom ou l'adresse IP de l'hôte"
    sPhraseGer 240, "Bitte geben Sie den Namen oder die IP-Adresse des Hosts ein"
    sPhraseSpa 240, "Por favor ingrese el nombre o la dirección IP del anfitrión"
    sPhraseSwe 240, "Var vänlig ange namn eller IP-adress till värden"
    sPhraseNor 240, "Vennligst skriv inn navn eller ID-adressen til vertsmaskinen"
    sPhraseDan 240, "Skriv navn eller IP adressen på værten"

    'REM:
    sPhraseEng 241, "Missing information"
    sPhraseIta 241, "Informazioni mancanti"
    sPhraseFra 241, "Information manquante"
    sPhraseGer 241, "Fehlende Information"
    sPhraseSpa 241, "Información perdida"
    sPhraseSwe 241, "Saknar information"
    sPhraseNor 241, "Mangler informasjon"
    sPhraseDan 241, "Manglende information"

    'REM:
    sPhraseEng 242, "Disconnect"
    sPhraseIta 242, "Scollegati"
    sPhraseFra 242, "Déconnecter"
    sPhraseGer 242, "Trennen"
    sPhraseSpa 242, "Desconectar"
    sPhraseSwe 242, "Koppla från"
    sPhraseNor 242, "Koble fra"
    sPhraseDan 242, "Afbryd"

    'REM:Connecting to X.
    sPhraseEng 243, "Connecting to "
    sPhraseIta 243, "In collegamento a "
    sPhraseFra 243, "Se connecter à "
    sPhraseGer 243, "Verbinden mit "
    sPhraseSpa 243, "Conectando a "
    sPhraseSwe 243, "Ansluter till "
    sPhraseNor 243, "Kobler til "
    sPhraseDan 243, "Etablerer forbindelse til "

    'REM:
    sPhraseEng 244, "Please restart this session."
    sPhraseIta 244, "Per favore rinizia la sessione."
    sPhraseFra 244, "S'il vous plaît recommencez cette session."
    sPhraseGer 244, "Bitte starten Sie neu."
    sPhraseSpa 244, "Por favor reiniciar esta sección."
    sPhraseSwe 244, "Var vänlig start om sessionen."
    sPhraseNor 244, "Vennligst start på nytt."
    sPhraseDan 244, "Genstart denne funktion."

    'REM:
    sPhraseEng 245, "Error... Address violation"
    sPhraseIta 245, "Errore ... violazione d'indirizzo"
    sPhraseFra 245, "Erreur... violation d'adresse"
    sPhraseGer 245, "Fehler... Address violation"
    sPhraseSpa 245, "Error... Violación de dirección"
    sPhraseSwe 245, "Fel... Adressfel"
    sPhraseNor 245, "Feil... Adressekonflikt"
    sPhraseDan 245, "Fejl... Ugyldig adresse"

    'REM:
    sPhraseEng 246, "Connection closed."
    sPhraseIta 246, "Connessione chiusa."
    sPhraseFra 246, "Connexion fermée."
    sPhraseGer 246, "Verbindung geschlossen."
    sPhraseSpa 246, "Conexión cerrada."
    sPhraseSwe 246, "Anslutning stängd."
    sPhraseNor 246, "Tilkobling brutt."
    sPhraseDan 246, "Forbindelse lukket."

    'REM:
    sPhraseEng 247, "All connections closed."
    sPhraseIta 247, "Tutte le connessioni chiuse."
    sPhraseFra 247, "Toutes les connexions sont fermées."
    sPhraseGer 247, "Alle Verbindungen geschlossen"
    sPhraseSpa 247, "Todas las conexiones cerradas."
    sPhraseSwe 247, "Alla anslutningar stängda."
    sPhraseNor 247, "Alle tilkoblinger er brutt."
    sPhraseDan 247, "Alle forbindelser lukket."

    'REM:
    sPhraseEng 248, "Listening for connections."
    sPhraseIta 248, "In attesa di connessioni."
    sPhraseFra 248, "Écoutez les connexions."
    sPhraseGer 248, "Suche nach Verbindungen."
    sPhraseSpa 248, "Escuchando por la conexión."
    sPhraseSwe 248, "Lyssnar efter anslutningar."
    sPhraseNor 248, "Venter på tilkoblinger."
    sPhraseDan 248, "Lytter efter forbindelser."

    'REM:
    sPhraseEng 249, "Another aplication could be using this port"
    sPhraseIta 249, "Un'altra applicazione potrebbe star utilizzando questa porta"
    sPhraseFra 249, "Un autre aplication pourrait être en train d'utiliser ce port"
    sPhraseGer 249, "Eine andere Anwendung könnte diesen Port benutzen"
    sPhraseSpa 249, "Otra aplicación está utilizando este puerto"
    sPhraseSwe 249, "Någon annan applikation kan täknas använda den porten"
    sPhraseNor 249, "Et annet program bruker muligens denne porten."
    sPhraseDan 249, "Et andet program kunne benytte denne port"

    'REM:
    sPhraseEng 250, "Address violation"
    sPhraseIta 250, "Violazione d'indirizzo"
    sPhraseFra 250, "Violation d'adresse"
    sPhraseGer 250, "Address violation"
    sPhraseSpa 250, "Violación de dirección"
    sPhraseSwe 250, "Adressfel"
    sPhraseNor 250, "Addressekonflikt"
    sPhraseDan 250, "Ugyldig adresse"

    'REM:
    sPhraseEng 251, "Begin"
    sPhraseIta 251, "Ascolto"
    sPhraseFra 251, "Ecoutez"
    sPhraseGer 251, "Suchen"
    sPhraseSpa 251, "Escuche"
    sPhraseSwe 251, "Lyssna"
    sPhraseNor 251, "Venter"
    sPhraseDan 251, "Åben forbindelse"

    'REM:
    sPhraseEng 252, "Connect"
    sPhraseIta 252, "Collegati"
    sPhraseFra 252, "Se connecter"
    sPhraseGer 252, "Verbinden"
    sPhraseSpa 252, "Conectar"
    sPhraseSwe 252, "Anslut"
    sPhraseNor 252, "Koble til"
    sPhraseDan 252, "Tilslut"

    'REM:
    sPhraseEng 253, "Connection has been lost."
    sPhraseIta 253, "Il collegamento e' stato perso."
    sPhraseFra 253, "La connexion a été perdue."
    sPhraseGer 253, "Verbindung verloren"
    sPhraseSpa 253, "Se ha perdido la conexión."
    sPhraseSwe 253, "Anslutningen bröts."
    sPhraseNor 253, "Tilkoblingen ble avbrutt."
    sPhraseDan 253, "Forbindelsen er blevet tabt."

    'REM:X has disconnected.
    sPhraseEng 254, " has disconnected."
    sPhraseIta 254, " si e' scollegato."
    sPhraseFra 254, " s'est déconnecté."
    sPhraseGer 254, " hat die Verbindung getrennt."
    sPhraseSpa 254, " se ha desconectado."
    sPhraseSwe 254, " har kopplat från."
    sPhraseNor 254, " har koblet fra."
    sPhraseDan 254, " har afbrudt forbindelsen."

    'REM:The buffer is too full.
    sPhraseEng 255, "Buffer overflow error has occured."
    sPhraseIta 255, "Si e' verificato un overflow del buffer."
    sPhraseFra 255, "Erreur du débordement du buffer."
    sPhraseGer 255, "Puffer-Überlauf-Fehler ist aufgetreten."
    sPhraseSpa 255, "Ha ocurrido un error de desborde de memoria."
    sPhraseSwe 255, "Ett buffertfel har inträffat."
    sPhraseNor 255, "Minnebufferet er overskredet."
    sPhraseDan 255, "Buffer overflow fejl har indtruffet."

    'REM:
    sPhraseEng 256, "Connection established."
    sPhraseIta 256, "Connessione eseguita."
    sPhraseFra 256, "La connexion est établie."
    sPhraseGer 256, "Verbindung hergestellt."
    sPhraseSpa 256, "Conexión establecida."
    sPhraseSwe 256, "Anslutning upprättad."
    sPhraseNor 256, "Tilkobling er oppnådd."
    sPhraseDan 256, "Forbindelse etableret."

    'REM:No, Please try again.
    sPhraseEng 257, ""
    sPhraseIta 257, ""
    sPhraseFra 257, ""
    sPhraseGer 257, ""
    sPhraseSpa 257, ""
    sPhraseSwe 257, ""
    sPhraseNor 257, ""
    sPhraseDan 257, ""

    'REM:
    sPhraseEng 258, ""
    sPhraseIta 258, ""
    sPhraseFra 258, ""
    sPhraseGer 258, ""
    sPhraseSpa 258, ""
    sPhraseSwe 258, ""
    sPhraseNor 258, ""
    sPhraseDan 258, ""

    'REM:System information is unavailable at this time.
    sPhraseEng 259, "System Information Is Unavailable At This Time"
    sPhraseIta 259, "Le informazioni sul sistema non sono disponibili al momento"
    sPhraseFra 259, "L'information du système n'est pas disponible actuellement"
    sPhraseGer 259, "System-Information ist z.Zt. nicht verfügbar."
    sPhraseSpa 259, "Información del Sistema No Está Disponible En Este Momento"
    sPhraseSwe 259, "Systeminformation är inte tillgängligt för tillfället"
    sPhraseNor 259, "Systeminformasjon er ikke tilgjengelig for øyeblikket"
    sPhraseDan 259, "Computer information er ikke tilgængelig på nuværende tidspunkt"

    'REM: Cheat mode has been activated.
    sPhraseEng 260, "** Cheat mode activated **"
    sPhraseIta 260, "-- Modalita' bara attiva --"
    sPhraseFra 260, "** Le mode Tricher est activé **"
    sPhraseGer 260, "** Cheat-Modus aktiviert **"
    sPhraseSpa 260, "** Modo tramposo activado **"
    sPhraseSwe 260, "** Fuskläge aktiverat **"
    sPhraseNor 260, "** Jukse-modus er aktivert **"
    sPhraseDan 260, "** Snyd er aktiveret **"

    'REM:An error occured while deleting X.
    sPhraseEng 261, "An error occured while deleting "
    sPhraseIta 261, "Si e' verificato un errore durante la cancellazione "
    sPhraseFra 261, "Une erreur est apparue en effaçant "
    sPhraseGer 261, "Ein Fehler beim Löschen ist aufgetreten "
    sPhraseSpa 261, "Ha ocurrido un error durante el borrado de "
    sPhraseSwe 261, "Ett fel uppstod vid borttagning av "
    sPhraseNor 261, "En feil oppsto under sletting av "
    sPhraseDan 261, "Der opstod en fejl under sletning af "

    'REM:
    sPhraseEng 262, "It might be already opened."
    sPhraseIta 262, "Potrebbe essere gia' aperto."
    sPhraseFra 262, "Il est peut déjà ouvert."
    sPhraseGer 262, "Könnte schon geöffnet sein."
    sPhraseSpa 262, "Debería estar aún abierto."
    sPhraseSwe 262, "Den kan vara öppen."
    sPhraseNor 262, "Den er antakelig allerede åpnet."
    sPhraseDan 262, "Er måske allerede åbnet."

    'REM:X already exists.
    sPhraseEng 263, " already exists."
    sPhraseIta 263, " esiste gia'."
    sPhraseFra 263, " existe déjà."
    sPhraseGer 263, " existiert bereits."
    sPhraseSpa 263, " todavía existe."
    sPhraseSwe 263, " finns redan."
    sPhraseNor 263, " finnes allerede."
    sPhraseDan 263, " findes allerede."

    'REM:
    sPhraseEng 264, "Overwrite any way?"
    sPhraseIta 264, "Sicuro di sovrascrivere?"
    sPhraseFra 264, "Remplacer?"
    sPhraseGer 264, "Trotzdem überschreiben?"
    sPhraseSpa 264, "¿Sobre escribe de cualquier modo?"
    sPhraseSwe 264, "Skriv över i alla fall?"
    sPhraseNor 264, "Skal den overskrives?"
    sPhraseDan 264, "Overskriv?"

    'REM:
    sPhraseEng 265, "Save as..."
    sPhraseIta 265, "Salva come..."
    sPhraseFra 265, "Sauvez sous..."
    sPhraseGer 265, "Speichern als..."
    sPhraseSpa 265, "Guardar como..."
    sPhraseSwe 265, "Spara som..."
    sPhraseNor 265, "Lagre som..."
    sPhraseDan 265, "Gem som..."

    'REM:X cannot be modified or deleted.
    sPhraseEng 266, " cannot be modified or deleted."
    sPhraseIta 266, " non puo' essere modificato o cancellato."
    sPhraseFra 266, " ne peut pas être modifié ou effacé."
    sPhraseGer 266, " kann nicht modifiziert oder gelöscht werden."
    sPhraseSpa 266, " no puede ser modificado o borrado."
    sPhraseSwe 266, " kan inte ändras eller tas bort."
    sPhraseNor 266, " kan ikke endres eller slettes."
    sPhraseDan 266, " kan ikke overskrives eller slettes."

    'REM:
    sPhraseEng 267, "&Advanced..."
    sPhraseIta 267, "Avanzate..."
    sPhraseFra 267, "Avancé..."
    sPhraseGer 267, "&Fortgeschritten..."
    sPhraseSpa 267, "&Avanzado..."
    sPhraseSwe 267, "Avancerat..."
    sPhraseNor 267, "Avansert..."
    sPhraseDan 267, "Avanceret..."

    'REM:
    sPhraseEng 268, "Multiplayer..."
    sPhraseIta 268, "&Impostazione rete..."
    sPhraseFra 268, "Paramêtres du réseau..."
    sPhraseGer 268, "Netzwerk-&Einstellungen..."
    sPhraseSpa 268, "Ajuste red de traba&jo..."
    sPhraseSwe 268, "Nätverksinställningar..."
    sPhraseNor 268, "Nettverks-oppsett..."
    sPhraseDan 268, "Netværk opsætning..."

    'REM:Compose a message.
    sPhraseEng 269, "Compose message..."
    sPhraseIta 269, "Scrivi un messaggio..."
    sPhraseFra 269, "Composez un message..."
    sPhraseGer 269, "Nachricht verfassen..."
    sPhraseSpa 269, "Redactar mensaje..."
    sPhraseSwe 269, "Komponera meddelande..."
    sPhraseNor 269, "Lag beskjed..."
    sPhraseDan 269, "Skriv besked..."

    'REM:The map of continents.
    sPhraseEng 270, "&Continent map"
    sPhraseIta 270, "Mappa dei continenti"
    sPhraseFra 270, "Carte des continents"
    sPhraseGer 270, "Weltkarte"
    sPhraseSpa 270, "Mapa c&ontinental"
    sPhraseSwe 270, "Kontinentkarta"
    sPhraseNor 270, "Kart over kontinent"
    sPhraseDan 270, "Kontinent kort"

    'REM:Information
    sPhraseEng 271, "&About"
    sPhraseIta 271, "Informazioni"
    sPhraseFra 271, "A propos de"
    sPhraseGer 271, "Ü&ber"
    sPhraseSpa 271, "&Acerca de"
    sPhraseSwe 271, "Om"
    sPhraseNor 271, "Om"
    sPhraseDan 271, "Om <Var.ExeName>"

    'REM:Change the number of players.
    sPhraseEng 272, "Change the number of starting players."
    sPhraseIta 272, "Cambia il numero dei giocatori."
    sPhraseFra 272, "Changez le nombre de joueurs."
    sPhraseGer 272, "Anzahl der Spieler ändern"
    sPhraseSpa 272, "Cambie los números de jugadores iniciales."
    sPhraseSwe 272, "Ändra antalet spelare."
    sPhraseNor 272, "Øker/reduserer antall spillere."
    sPhraseDan 272, "Skift antal startende spillere."

    'REM:
    sPhraseEng 273, "Advanced options..."
    sPhraseIta 273, "Opzioni avanzate..."
    sPhraseFra 273, "Options avancées..."
    sPhraseGer 273, "Fortgeschrittene Optionen..."
    sPhraseSpa 273, "Opciones avanzadas..."
    sPhraseSwe 273, "Avancerade inställningar..."
    sPhraseNor 273, "Avanserte funksjoner..."
    sPhraseDan 273, "Avanceret opsætning..."

    'REM:DEFUNCT
    sPhraseEng 274, "Advanced war options"
    sPhraseIta 274, "Opzioni avanzate di guerra"
    sPhraseFra 274, "Options de guerre avancées"
    sPhraseGer 274, "Fortgeschrittene Kriegs-Optionen..."
    sPhraseSpa 274, "Opciones de guerra avanzada"
    sPhraseSwe 274, "Avancerade alternativ för krig"
    sPhraseNor 274, "Avanserte funksjoner"
    sPhraseDan 274, "Avanceret krigsopsætning"

    'REM:DEFUNCT
    sPhraseEng 275, "Mission Options"
    sPhraseIta 275, "Opzioni obiettivi"
    sPhraseFra 275, "Options des missions"
    sPhraseGer 275, "Missions-Optionen"
    sPhraseSpa 275, "Opciones de Misión"
    sPhraseSwe 275, "Uppdragsinställningar"
    sPhraseNor 275, "Krigføring"
    sPhraseDan 275, "Mission opsætning"

    'REM:DEFUNCT
    sPhraseEng 276, "Graphics Options"
    sPhraseIta 276, "Opzioni grafiche"
    sPhraseFra 276, "Options graphiques"
    sPhraseGer 276, "Grafik-Optionen"
    sPhraseSpa 276, "Opciones de Gráficos"
    sPhraseSwe 276, "Grafikinställningar"
    sPhraseNor 276, "Grafikk"
    sPhraseDan 276, "Grafik opsætning"

    'REM:ENGLISH CHANGED
    sPhraseEng 277, "Army wipeout missions"
    sPhraseIta 277, "Include l'obiettivo di fare piazza pulita delle armate"
    sPhraseFra 277, "Inclure les missions d'éradication d'armées"
    sPhraseGer 277, "Inclusive 'Armee-auslöschen'-Missionen"
    sPhraseSpa 277, "Incluir misiones de exterminio de ejércitos"
    sPhraseSwe 277, "Inkludera uppdrag där andra arméer skall utplånas"
    sPhraseNor 277, "Utrydd motstander."
    sPhraseDan 277, "Inkluder 'udryd hær' missioner"

    'REM:ENGLISH CHANGED
    sPhraseEng 278, "Conquer and hold missions"
    sPhraseIta 278, "Include il conseguimento ed il mantenimento dell'obiettivo"
    sPhraseFra 278, "Inclure les missions : conquérir et tenir"
    sPhraseGer 278, "Inclusive 'Erobern und halten'-Missionen"
    sPhraseSpa 278, "Incluir misiones de conquistar y retener"
    sPhraseSwe 278, "Inkludera erövra och håll-uppdrag"
    sPhraseNor 278, "Erobre og befeste"
    sPhraseDan 278, "Inkluder 'erobre og hold' missioner"

    'REM:ENGLISH CHANGED You must complete your own mission to win.
    sPhraseEng 279, "Must complete own mission"
    sPhraseIta 279, "Devi completare il tuo obiettivo per vincere"
    sPhraseFra 279, "Vous devez compléter votre propre mission pour gagner"
    sPhraseGer 279, "Eigene Mission erfüllen, um zu gewinnen"
    sPhraseSpa 279, "Debe completar las propias misiones para ganar"
    sPhraseSwe 279, "Måste avsluta sitt eget uppdrag för att vinna"
    sPhraseNor 279, "Utføre eget oppdrag for å vinne"
    sPhraseDan 279, "Skal selv afslutte mission"

    'REM:ENGLISH CHANGED You win immediately upon completing your mission.
    sPhraseEng 280, "Win immediately upon completion"
    sPhraseIta 280, "Vinci subito dopo il completamento dell'obiettivo"
    sPhraseFra 280, "Vous gagnez immédiatement apres avoir rempli votre mission"
    sPhraseGer 280, "Sofort gewinnen, nach Erfüllen der eigenen Mission"
    sPhraseSpa 280, "Ganar inmediatamente por sobre completar la misión"
    sPhraseSwe 280, "Vinn direkt efter avslutat uppdrag"
    sPhraseNor 280, "Umiddelbar seier"
    sPhraseDan 280, "Vind straks"

    'REM:ENGLISH CHANGED Change the saturation of green.
    sPhraseEng 281, "Dark green"
    sPhraseIta 281, "Cambia la luminosita' del verde"
    sPhraseFra 281, "Changement la saturation de vert"
    sPhraseGer 281, "Grüne Sättigung ändern"
    sPhraseSpa 281, "Cambiar saturación verde"
    sPhraseSwe 281, "Ändra färgmättnad på den gröna färgen"
    sPhraseNor 281, "Endre grønn farge"
    sPhraseDan 281, "Skift grøn farve"

    'REM:ENGLISH CHANGED Change the saturation of yellow.
    sPhraseEng 282, "Dark yellow"
    sPhraseIta 282, "Cambia la luminosita' del giallo"
    sPhraseFra 282, "Changement la saturation de jaune"
    sPhraseGer 282, "Gelbe Sättigung ändern"
    sPhraseSpa 282, "Cambiar saturación amarilla"
    sPhraseSwe 282, "Ändra färgmättnad på den gula färgen"
    sPhraseNor 282, "Endre gul farge"
    sPhraseDan 282, "Skift gul farve"

    'REM:ENGLISH CHANGED Change the luminosity of gray.
    sPhraseEng 283, "Dark gray"
    sPhraseIta 283, "Cambia la luminosita' del grigio"
    sPhraseFra 283, "Changement la luminosité de gris"
    sPhraseGer 283, "Graue Helligkeit ändern"
    sPhraseSpa 283, "Cambiar saturación gris"
    sPhraseSwe 283, "Ändra ljusstyrkan på den gråa färgen"
    sPhraseNor 283, "Endre grå farge"
    sPhraseDan 283, "Skift grå farve"

    'REM:ENGLISH CHANGED
    sPhraseEng 284, "Large font"
    sPhraseIta 284, "Caratteri grandi"
    sPhraseFra 284, "Grande police"
    sPhraseGer 284, "Große Schrift"
    sPhraseSpa 284, "Tamaño de fuente expandido"
    sPhraseSwe 284, "Stora teckensnitt"
    sPhraseNor 284, "Fet tekst"
    sPhraseDan 284, "Fed skrifttype"

    'REM:DEFUNCT An option to include a secret mission, where the object is to kill another player.
    sPhraseEng 285, "Include missions where you have to wipe out the last unit of another player"
    sPhraseIta 285, "Includi un obiettivo segreto di distruggere un avversario"
    sPhraseFra 285, "Inclure des missions de destruction de la dernière unité d'un autre joueur"
    sPhraseGer 285, "Inclusive Missionen, in denen die gesamte Armee eines Spielers ausgelöscht werden muss"
    sPhraseSpa 285, "Incluir misiones donde haya aniquilado la última unidad de otro jugador"
    sPhraseSwe 285, "Ta med uppdrag där du skall slå ut en annan spelares sista styrka"
    sPhraseNor 285, "Inkluderer oppdrag hvor du må utrydde en motstanders siste enhet"
    sPhraseDan 285, "Inkluder missioner hvor man skal udrydde en anden hær"

    'REM:DEFUNCT An option to include a secret mission, where the object is to capture certain continents and hold them until it is your next turn.
    sPhraseEng 286, "Include missions where you have to conquer territories and hold them until your next turn"
    sPhraseIta 286, "Includi un obiettivo segreto di catturare alcuni continenti e mantenerli fino al turno sucessivo"
    sPhraseFra 286, "Inclure des missions où vous avez à conquérir des territoires et les tenir jusqu'à votre prochain tour"
    sPhraseGer 286, "Inclusive Missionen, in denen Territorien erobert und bis zur nächsten Runde gehalten werden müssen"
    sPhraseSpa 286, "Incluir misiones donde tenga que conquistar territorios y retenerlos hasta su próximo turno"
    sPhraseSwe 286, "Ta med uppdrag där du ska erövra kontinenter och hålla dem till ditt nästa drag"
    sPhraseNor 286, "Inkluderer oppdrag hvor du må erobre kontinenter og befeste disse til det blir din tur igjen"
    sPhraseDan 286, "Inkluder missioner hvor man skal erobre kontinenter og holde dem en tur"

    'REM:DEFUNCT
    sPhraseEng 287, "You do not win when an opponent wipes out your target army for you."
    sPhraseIta 287, "Non hai vinto quando e' un avversario a spazzare le armate nemiche."
    sPhraseFra 287, "Vous ne gagnez pas quand un adversaire détruit votre armée cible à votre place."
    sPhraseGer 287, "Sie gewinnen nicht, wenn ein Gegner Ihre Ziel-Armee für Sie auslöscht"
    sPhraseSpa 287, "Usted no gana cuando un oponente aniquila un ejercito objetivo por usted"
    sPhraseSwe 287, "Du vinner inte om en motståndare utplånar den armé du skall slå ut"
    sPhraseNor 287, "Du vinner ikke dersom en motstander utrydder en motstander for deg."
    sPhraseDan 287, "Du vinder ikke, når en anden hær udrydder den hær du skal udrydde."

    'REM:DEFUNCT
    sPhraseEng 288, "After wiping out a target army, you must wait until your next turn to win when unchecked."
    sPhraseIta 288, "Dopo aver spazzato via le armate del nemico devi aspettare il prossimo turno per la vittoria"
    sPhraseFra 288, "Après avoir détruit une armée cible, vous devez attendre jusqu'à votre prochain tour pour gagner quand déselectionné."
    sPhraseGer 288, "Nachdem die Ziel-Armee ausgelöscht wurde, müssen Sie bis zur nächsten Runde warten, um zu gewinnen (falls nicht ausgewählt)"
    sPhraseSpa 288, "Después de aniquilar un ejercito objetivo, deberá esperar hasta su próximo turno para ganar hasta ser revisado"
    sPhraseSwe 288, "Efter att ha utplånat den armé du skall slå ut måste du vänta till det är din tur för att vinna om detta inte är markerat"
    sPhraseNor 288, "Etter å ha utryddet en motstander, vinner du ikke før det er blitt din tur igjen."
    sPhraseDan 288, "Man skal ikke vente en runde efter udsletning af anden hær."

    'REM:DEFUNCT
    sPhraseEng 289, "Lighten or darken the color of the Green Army."
    sPhraseIta 289, "Schiarisci o scurisci il colore delle armate verdi"
    sPhraseFra 289, "Éclairer ou foncer la couleur de l'Armée Verte."
    sPhraseGer 289, "Erhellen oder verdunklen der Farbe der grünen Armee"
    sPhraseSpa 289, "Iluminar u oscurecer el color del Ejército Verde"
    sPhraseSwe 289, "Gör den gröna armén ljusare eller mörkare"
    sPhraseNor 289, "Setter lysere/mørkere farge på den Grønne Hæren."
    sPhraseDan 289, "Gør den grønne farve lysere eller mørkere."

    'REM:DEFUNCT
    sPhraseEng 290, "Lighten or darken the color of the Yellow Army."
    sPhraseIta 290, "Schiarisci o scurisci il colore delle armate gialle"
    sPhraseFra 290, "Éclairer ou foncer la couleur de l'Armée Jaune."
    sPhraseGer 290, "Erhellen oder verdunklen der Farbe der gelben Armee"
    sPhraseSpa 290, "Iluminar u oscurecer el color del Ejército Amarillo"
    sPhraseSwe 290, "Gör den gula armén ljusare eller mörkare"
    sPhraseNor 290, "Setter lysere/mørkere farge på den Gule Hæren."
    sPhraseDan 290, "Gør den gule farve lysere eller mørkere."

    'REM:DEFUNCT
    sPhraseEng 291, "Lighten or darken the color of the Gray Army."
    sPhraseIta 291, "Schiarisci o scurisci il colore delle armate grigie"
    sPhraseFra 291, "Éclairer ou foncer la couleur de l'Armée Grise."
    sPhraseGer 291, "Erhellen oder verdunklen der Farbe der grauen Armee"
    sPhraseSpa 291, "Iluminar u oscurecer el color del Ejército Gris"
    sPhraseSwe 291, "Gör den gråa armén ljusare eller mörkare"
    sPhraseNor 291, "Setter lysere/mørkere farge på den Grå Hæren."
    sPhraseDan 291, "Gør den grå farve lysere eller mørkere."

    'REM:DEFUNCT
    sPhraseEng 292, "Change the font size used to print country scores."
    sPhraseIta 292, "Cambia le dimensioni dei caratteri per la stampa"
    sPhraseFra 292, "Changer la taille des caractères des scores des pays."
    sPhraseGer 292, "Ändern der Schriftgröße mit der die Anzahl der Armeen angezeigt wird"
    sPhraseSpa 292, "Cambiar la fuente usada para imprimir la puntuación de los países"
    sPhraseSwe 292, "Ändra storleken på teckensnittet som anger antalet arméer i länderna"
    sPhraseNor 292, "Endrer skriftstørrelsen på antall enheter på spillebrettet."
    sPhraseDan 292, "Gør skrifttypen, som viser antal arméer, fed.<Var.ExeName> network setup"

    'REM: "<Var.ExeName> network setup"
    sPhraseEng 293, "Network Administration Panel"
    sPhraseIta 293, "Impostrazioni di rete <Var.ExeName>"
    sPhraseFra 293, "Installation de <Var.ExeName> en réseau"
    sPhraseGer 293, "<Var.ExeName> Netzwerk Einstellungen"
    sPhraseSpa 293, "Configurar la red de trabajo de Misión Riesgo"
    sPhraseSwe 293, "Nätverksinställningar för <Var.ExeName>"
    sPhraseNor 293, "<Var.ExeName> nettverks-oppsett"
    sPhraseDan 293, "<Var.ExeName> netværk opsætning"

    'REM: **
    sPhraseEng 294, "Options"
    sPhraseIta 294, "Opzioni di connessione"
    sPhraseFra 294, "Options de connexion"
    sPhraseGer 294, "Verbindungs-Optionen"
    sPhraseSpa 294, "Opciones de conexión"
    sPhraseSwe 294, "Anslutningsalternativ"
    sPhraseNor 294, "Tilkoblingsvalg"
    sPhraseDan 294, "Forbindelse opsætning"

    'REM:
    sPhraseEng 295, "Connection type"
    sPhraseIta 295, "Tipo di connessione"
    sPhraseFra 295, "Types de connexion"
    sPhraseGer 295, "Verbindungs-Typ"
    sPhraseSpa 295, "Yipos de conexión"
    sPhraseSwe 295, "Anlutningstyp"
    sPhraseNor 295, "Tilkobling"
    sPhraseDan 295, "Forbindelse type"

    'REM:
    sPhraseEng 296, "Refresh rate"
    sPhraseIta 296, "Frequenza di refresh"
    sPhraseFra 296, "Vitesse de rafraîchissement"
    sPhraseGer 296, "Aktualisierungsrate"
    sPhraseSpa 296, "Tasa de refresco"
    sPhraseSwe 296, "Uppdateringsfrekvens"
    sPhraseNor 296, "Oppdatering"
    sPhraseDan 296, "Opdateringshastighed"

    'REM:
    sPhraseEng 297, "Settings"
    sPhraseIta 297, "Impostazioni"
    sPhraseFra 297, "Paramètres"
    sPhraseGer 297, "Einstellungen"
    sPhraseSpa 297, "Ajustes"
    sPhraseSwe 297, "Inställningar"
    sPhraseNor 297, "Innstillinger"
    sPhraseDan 297, "Opsætning"

    'REM:
    sPhraseEng 298, "Session history"
    sPhraseIta 298, "Storico sessione"
    sPhraseFra 298, "Historique de la session"
    sPhraseGer 298, "'Sitzungsgeschichte'"
    sPhraseSpa 298, "Historia de la sesión"
    sPhraseSwe 298, "Sessionshistoria"
    sPhraseNor 298, "Historikk for oppkoblingen"
    sPhraseDan 298, "Forbindelsens historie"

    'REM:Name or IP address of host.
    sPhraseEng 299, "Name or IP address of host"
    sPhraseIta 299, "Nome o indirizzo IP dell'host"
    sPhraseFra 299, "Nom ou adresse IP de l'hôte"
    sPhraseGer 299, "Name oder IP-Adresse des Hosts"
    sPhraseSpa 299, "Nombre o dirección IP del anfitrión"
    sPhraseSwe 299, "Värdens namn eller IP-adress"
    sPhraseNor 299, "Vertsmaskinens navn eller ID-adresse"
    sPhraseDan 299, "Navn eller IP adresse på vært"
    Call initialisePhrases300
End Sub

Private Sub initialisePhrases300()
    'REM:Main Port.
    sPhraseEng 300, "Main Port"
    sPhraseIta 300, "Main Port"
    sPhraseFra 300, "Main Port"
    sPhraseGer 300, "Main Port"
    sPhraseSpa 300, "Main Port"
    sPhraseSwe 300, "Main Port"
    sPhraseNor 300, "Main Port"
    sPhraseDan 300, "Main Port"

    'REM:Join a war. **
    sPhraseEng 301, "Join"
    sPhraseIta 301, "Unisciti ad una guerra"
    sPhraseFra 301, "Joindre une guerre"
    sPhraseGer 301, "Einem Krieg beitreten"
    sPhraseSpa 301, "Unirse a la guerra"
    sPhraseSwe 301, "Gå med i ett krig"
    sPhraseNor 301, "Delta i krig"
    sPhraseDan 301, "Gå i anden krig"

    'REM:Host a war. **
    sPhraseEng 302, "Host"
    sPhraseIta 302, "Ospita una guerra"
    sPhraseFra 302, "Organiser une guerre"
    sPhraseGer 302, "Host für einen Krieg"
    sPhraseSpa 302, "Organizar una guerra"
    sPhraseSwe 302, "Var värd för ett krig"
    sPhraseNor 302, "Arranger krig"
    sPhraseDan 302, "Start en krig"

    'REM:
    sPhraseEng 303, "TCP / IP"
    sPhraseIta 303, "TCP / IP"
    sPhraseFra 303, "TCP / IP"
    sPhraseGer 303, "TCP / IP"
    sPhraseSpa 303, "TCP / IP"
    sPhraseSwe 303, "TCP / IP"
    sPhraseNor 303, "TCP / IP"
    sPhraseDan 303, "TCP / IP"

    'REM:
    sPhraseEng 304, "Modem"
    sPhraseIta 304, "Modem"
    sPhraseFra 304, "Modem"
    sPhraseGer 304, "Modem"
    sPhraseSpa 304, "Modem"
    sPhraseSwe 304, "Modem"
    sPhraseNor 304, "Modem"
    sPhraseDan 304, "Modem"

    'REM:
    sPhraseEng 305, "Serial"
    sPhraseIta 305, "Seriale"
    sPhraseFra 305, "Série"
    sPhraseGer 305, "Seriell"
    sPhraseSpa 305, "Serial"
    sPhraseSwe 305, "Seriell"
    sPhraseNor 305, "Seriell"
    sPhraseDan 305, "Serial"

    'REM: High **
    sPhraseEng 306, "Fast"
    sPhraseIta 306, "Alto"
    sPhraseFra 306, "Haut"
    sPhraseGer 306, "Hoch"
    sPhraseSpa 306, "Alto"
    sPhraseSwe 306, "Hög"
    sPhraseNor 306, "Høy"
    sPhraseDan 306, "Høj"

    'REM:
    sPhraseEng 307, "Medium"
    sPhraseIta 307, "Medio"
    sPhraseFra 307, "Moyen"
    sPhraseGer 307, "Mittel"
    sPhraseSpa 307, "Medio"
    sPhraseSwe 307, "Mellan"
    sPhraseNor 307, "Middels"
    sPhraseDan 307, "Mellem"

    'REM: Low **
    sPhraseEng 308, "Slow"
    sPhraseIta 308, "Basso"
    sPhraseFra 308, "Bas"
    sPhraseGer 308, "Niedrig"
    sPhraseSpa 308, "Bajo"
    sPhraseSwe 308, "Låg"
    sPhraseNor 308, "Lav"
    sPhraseDan 308, "Lav"

    'REM:
    sPhraseEng 309, "&IP Info..."
    sPhraseIta 309, "&IP Info..."
    sPhraseFra 309, "IP Info..."
    sPhraseGer 309, "&IP Info..."
    sPhraseSpa 309, "&Información IP..."
    sPhraseSwe 309, "IP-information..."
    sPhraseNor 309, "IP Info..."
    sPhraseDan 309, "IP Info..."

    'REM:
    sPhraseEng 310, "&Connect"
    sPhraseIta 310, "&Collegati"
    sPhraseFra 310, "Se connecter"
    sPhraseGer 310, "&Verbinden"
    sPhraseSpa 310, "&Conectar"
    sPhraseSwe 310, "Anslut"
    sPhraseNor 310, "Koble til"
    sPhraseDan 310, "Åben forbindelse"

    'REM:
    sPhraseEng 311, "&OK"
    sPhraseIta 311, "&OK"
    sPhraseFra 311, "OK"
    sPhraseGer 311, "&OK"
    sPhraseSpa 311, "&OK"
    sPhraseSwe 311, "OK "
    sPhraseNor 311, "OK"
    sPhraseDan 311, "OK"

    'REM:A network war must have 1 host.
    sPhraseEng 312, "Network war must have 1 host"
    sPhraseIta 312, "La guerra via rete deve avere 1 host"
    sPhraseFra 312, "La guerre en réseau doit avoir 1 hôte"
    sPhraseGer 312, "Ein Netzwerk-Krieg muss einen Host haben"
    sPhraseSpa 312, "La guerra en red de trabajo debe tener un Anfitrión"
    sPhraseSwe 312, "Nätverkskrig måste ha 1 värd"
    sPhraseNor 312, "Krig i nettverk må ha 1 vert"
    sPhraseDan 312, "Netværk krig skal have en vært"

    'REM:
    sPhraseEng 313, "Modem and direct serial connections will be available in future versions"
    sPhraseIta 313, "Connessioni via seriale o via cavo saranno disponibili in versioni future."
    sPhraseFra 313, "Les connections série et par modem seront disponibles dans les futures versions"
    sPhraseGer 313, "Modem und direkte serielle Verbindungen werden in zukünftigen Versionen verfügbar sein."
    sPhraseSpa 313, "Modem y conexión serie estarán disponibles en versiones futuras"
    sPhraseSwe 313, "Modem och direkt seriell anslutning kommer att finnas i kommande versioner"
    sPhraseNor 313, "Modem og seriell tilkobling vil bli tilgjengelig i senere versjoner"
    sPhraseDan 313, ""

    'REM:
    sPhraseEng 314, "The frequency at which remote players are updated."
    sPhraseIta 314, "La frequenza di aggiornamento dei giocatori remoti"
    sPhraseFra 314, "Fréquence à laquelle les joueurs distants sont actualisés."
    sPhraseGer 314, "Die Frequenz mit der entfernte Spieler aktualisiert werden"
    sPhraseSpa 314, "La frecuencia a la cual los jugadores remotos se actualizan"
    sPhraseSwe 314, "Frekvensen varmed fjärrspelare uppdateras"
    sPhraseNor 314, "Frekvens for hvor hurtig de andre spillerne blir oppdatert."
    sPhraseDan 314, "Hvor ofte andre spillere skal opdateres."

    'REM:
    sPhraseEng 315, "Enter the name of the host terminal."
    sPhraseIta 315, "Inserisci il nome del terminale host"
    sPhraseFra 315, "Entrez le nom du terminal de l'hôte."
    sPhraseGer 315, "Geben Sie den Namen des Host-Terminals ein"
    sPhraseSpa 315, "Ingrese el nombre del terminal Anfitrión"
    sPhraseSwe 315, "Skriv in värddatorns namn"
    sPhraseNor 315, "Skriv inn vertsmaskinens navn."
    sPhraseDan 315, "Skriv navnet på værten"

    'REM:
    sPhraseEng 316, "All terminals must use the same port number."
    sPhraseIta 316, "Tutti i terminali devono usare la stessa porta"
    sPhraseFra 316, "Tous les terminaux doivent utiliser le même numéro de port."
    sPhraseGer 316, "Alle Terminals müssen dieselbe Port-Nummer benutzen"
    sPhraseSpa 316, "Todos los terminales deberán usar el mismo número de puerto"
    sPhraseSwe 316, "Alla datorer måste använda samma portnummer"
    sPhraseNor 316, "Alle datamaskiner må bruke samme port nr."
    sPhraseDan 316, "Alle computere skal bruge det samme port nummer."

    'REM:
    sPhraseEng 317, "Connect to a listening host."
    sPhraseIta 317, "Collegati ad un host"
    sPhraseFra 317, "Se connecter à un hôte."
    sPhraseGer 317, "Mit einem suchenden Host verbinden"
    sPhraseSpa 317, "Conectar a un anfitrión que escuche"
    sPhraseSwe 317, "Anslut till en lyssnande värd"
    sPhraseNor 317, "Delta i en allerede opprettet krig."
    sPhraseDan 317, "Opret forbindelse til en vært."

    'REM:
    sPhraseEng 318, "Become a host and listen for connections."
    sPhraseIta 318, "Diventa un host e accetta connessioni entranti"
    sPhraseFra 318, "Devenez un hôte et attendre les connexions."
    sPhraseGer 318, "Host werden und nach Verbindungen suchen"
    sPhraseSpa 318, "Convertirse en un anfitrión para conexiones"
    sPhraseSwe 318, "Bli värd och lyssna efter anslutningar"
    sPhraseNor 318, "Bli vertskap for en nettverkskrig."
    sPhraseDan 318, "Bliv vært og åben forbindelse"

    'REM:
    sPhraseEng 319, "Display host name and IP configuration for this terminal."
    sPhraseIta 319, "Mostra il nome host e la configurazione IP di questo terminale"
    sPhraseFra 319, "Affichez nom d'hôte et configuration de l'IP pour ce terminal."
    sPhraseGer 319, "Host-Name und IP-Konfiguration für diesen Terminal anzeigen"
    sPhraseSpa 319, "Despliegue nombre de anfitrión y configuración de IP para este terminal"
    sPhraseSwe 319, "Visa värdnamn och IP-konfiguration för denna dator"
    sPhraseNor 319, "Vis vertsnavn og IP-konfigurasjon for denne datamaskinen."
    sPhraseDan 319, "Vis vært navn og IP konfiguration for denne computer."

    'REM:
    sPhraseEng 320, "Connect with these settings"
    sPhraseIta 320, "Collegati con questi paramentri"
    sPhraseFra 320, "Se connecter avec ces paramètres"
    sPhraseGer 320, "Mit diesen Einstellungen verbinden"
    sPhraseSpa 320, "Conectar con estas configuraciones"
    sPhraseSwe 320, "Anslut med dessa inställningar"
    sPhraseNor 320, "Koble til med disse innstillinger"
    sPhraseDan 320, "Opret forbindelse med denne opsætning"

    'REM:
    sPhraseEng 321, "Hide the network setup dialog box without disconnecting."
    sPhraseIta 321, "Nascondi la finestra di connessione senza scollegarsi"
    sPhraseFra 321, "Cachez la boîte de dialogue de configuration du réseau sans déconnecter."
    sPhraseGer 321, "Die Netzwerk-Dialog-Box verbergen, ohne die Verbindung zu trennen"
    sPhraseSpa 321, "Esconda la caja de diálogo de la configuración de la red de trabajo sin desconectarse"
    sPhraseSwe 321, "Dölj nätverksanslutningsdialogrutan utan att koppla från"
    sPhraseNor 321, "Lukker denne dialogboksen uten å koble deg fra."
    sPhraseDan 321, "Skjul denne dialog boks uden at lukke forbindelsen"

    'REM:
    sPhraseEng 322, "Lists connected terminals, TCP errors, etc..."
    sPhraseIta 322, "Elenca i terminali collegati, errori TCP/IP, ecc...."
    sPhraseFra 322, "Afficher les terminaux connectés, les erreurs TCP, etc..."
    sPhraseGer 322, "Zeigt verbundene Terminals, TCP Fehler, etc. an"
    sPhraseSpa 322, "Enumere los terminales conectados, errores TCP, etc..."
    sPhraseSwe 322, "Listar anslutna datorer, TCP-fel etc..."
    sPhraseNor 322, "Viser hvilke datamaskiner som er oppkoblet, TCP-feil m.m."
    sPhraseDan 322, "Viser hvilke computere der har fået forbindelse m.m."

    'REM:
    sPhraseEng 323, "Delete the selected war."
    sPhraseIta 323, "Cancella lo scenario selezionato"
    sPhraseFra 323, "Effacer la guerre sélectionnée."
    sPhraseGer 323, "Den ausgewählten Krieg löschen"
    sPhraseSpa 323, "Borre la guerra elegida"
    sPhraseSwe 323, "Ta bort det markerade kriget"
    sPhraseNor 323, "Slett den valgte krig."
    sPhraseDan 323, "Slet den valgte krig"

    'REM:
    sPhraseEng 324, "Open this war the next time <Var.ExeName> starts."
    sPhraseIta 324, "Apri questo scenario la prossima volta che lanci <Var.ExeName>"
    sPhraseFra 324, "Ouvrir cette guerre lors du prochain démarrage de <Var.ExeName>."
    sPhraseGer 324, "Diesen Krieg beim nächsten Start öffnen"
    sPhraseSpa 324, "Inicie esta guerra la próxima vez que <Var.ExeName> comience"
    sPhraseSwe 324, "Öppna detta krig nästa gång <Var.ExeName> startas"
    sPhraseNor 324, "Åpner denne krigen neste gang <Var.ExeName> startes."
    sPhraseDan 324, "Åben denne krig næste gang <Var.ExeName> starter."

    'REM:
    sPhraseEng 325, "Lock war to prevent accidental deletion."
    sPhraseIta 325, "Blocca lo scenario per evitare cancellazioni involontarie"
    sPhraseFra 325, "Bloquer clé la guerre contre une éventuelle suppression accidentelle."
    sPhraseGer 325, "Krieg sperren, um versehentliches Löschen zu verhindern."
    sPhraseSpa 325, "Asegure la guerra para prevenir el borrado accidental"
    sPhraseSwe 325, "Lås krig för att förhindra oavsiktlig borttagning"
    sPhraseNor 325, "Lås krigen for å hindre uønsket sletting."
    sPhraseDan 325, "Lås denne krig, for at undgå utilsigtet sletning."

    'REM:
    sPhraseEng 326, "Credits"
    sPhraseIta 326, "Credits"
    sPhraseFra 326, "Crédits"
    sPhraseGer 326, "Grüße"
    sPhraseSpa 326, "Créditos"
    sPhraseSwe 326, "Tack"
    sPhraseNor 326, "Takk til..."
    sPhraseDan 326, "Lavet af"

    'REM:
    sPhraseEng 327, "Message"
    sPhraseIta 327, "Messaggio"
    sPhraseFra 327, "Message"
    sPhraseGer 327, "Nachricht"
    sPhraseSpa 327, "Mensaje"
    sPhraseSwe 327, "Meddelande"
    sPhraseNor 327, "Melding"
    sPhraseDan 327, "Besked"

    'REM:Recipient Player. Recipient Players.
    sPhraseEng 328, "Recipient Player(s)"
    sPhraseIta 328, "Giocatori destinatari"
    sPhraseFra 328, "Joueur(s) destinataire(s)."
    sPhraseGer 328, "Empfänger"
    sPhraseSpa 328, "Depocito(s) de jugadores(s) "
    sPhraseSwe 328, "Mottagande spelare"
    sPhraseNor 328, "Mottaker(e)"
    sPhraseDan 328, "Modtager(e)"

    'REM:Select recipient of private messages. Select recipients of private messages.
    sPhraseEng 329, "Select recipient(s) of private messages"
    sPhraseIta 329, "Seleziona i destinatari dei messaggi privati"
    sPhraseFra 329, "Sélectionner le(s) Destinataire(s) des messages privés."
    sPhraseGer 329, "Wählen Sie den oder die Empfänger für private Nachrichten"
    sPhraseSpa 329, "Seleccione recipiente(s) de mensajes privados"
    sPhraseSwe 329, "Välj mottagare för privata meddelanden"
    sPhraseNor 329, "Velg mottaker(e) av personlig melding"
    sPhraseDan 329, "Vælg hvilke spillere der skal modtage beskeden"

    'REM:
    sPhraseEng 330, "Public"
    sPhraseIta 330, "Pubblico"
    sPhraseFra 330, "Public"
    sPhraseGer 330, "Öffentlich"
    sPhraseSpa 330, "Publico"
    sPhraseSwe 330, "Öppet"
    sPhraseNor 330, "Åpen"
    sPhraseDan 330, "Til alle"

    'REM:Send message to all terminals.
    sPhraseEng 331, "Send message to all terminals"
    sPhraseIta 331, "Manda un messaggio a tutti i terminali"
    sPhraseFra 331, "Envoyer le message à tous les terminaux"
    sPhraseGer 331, "Nachricht an alle Terminals senden"
    sPhraseSpa 331, "Enviar mensaje a todos los terminales"
    sPhraseSwe 331, "Skicka meddelande till alla datorer"
    sPhraseNor 331, "Send melding til alle deltakere"
    sPhraseDan 331, "Skriv besked til alle spillere"

    'REM:
    sPhraseEng 332, "Private"
    sPhraseIta 332, "Privato"
    sPhraseFra 332, "Privé"
    sPhraseGer 332, "Privat"
    sPhraseSpa 332, "Privado"
    sPhraseSwe 332, "Privat"
    sPhraseNor 332, "Personlig"
    sPhraseDan 332, "Privat"

    'REM:Send message to terminals controlling selected army. Send message to terminals controlling selected armies.
    sPhraseEng 333, "Send message to terminals controlling selected army(s)"
    sPhraseIta 333, "Manda un messaggio al terminale che controlla le armate selezionate"
    sPhraseFra 333, "Envoyez le message à(aux) terminal(aux) contrôlant(s) l'armée(s) sélectionnée(s)."
    sPhraseGer 333, "Nachricht an Terminals senden, die die ausgewählte Armee kontrollieren"
    sPhraseSpa 333, "Enviar mensaje al terminal controlado por el(los) ejercito(s) seleccionado(s)"
    sPhraseSwe 333, "Skicka meddelande till datorerna som kontrollerar vald(a) armé(er)"
    sPhraseNor 333, "Send melding til valgt mottaker(e)"
    sPhraseDan 333, "Send besked til de valgte hære"

    'REM:
    sPhraseEng 334, "&Hide"
    sPhraseIta 334, "&Nascondi"
    sPhraseFra 334, "Cacher"
    sPhraseGer 334, "&Verstecken"
    sPhraseSpa 334, "&Ocultar"
    sPhraseSwe 334, "Dölj"
    sPhraseNor 334, "OK"
    sPhraseDan 334, "Skjul"

    'REM:Type your message and then press enter to send.
    sPhraseEng 335, "Type your message and press enter to send"
    sPhraseIta 335, "Scrivi il tuo messaggio e premi enter per spedire"
    sPhraseFra 335, "Tapez votre message puis entrée pour l'envoyer"
    sPhraseGer 335, "Geben Sie die Nachricht ein und drücken Sie 'enter'"
    sPhraseSpa 335, "Escriba su mensaje y presione entrar para enviar"
    sPhraseSwe 335, "Skriv ditt meddelande och tryck på enter för att skicka"
    sPhraseNor 335, "Skriv din melding og trykk ENTER for å sende"
    sPhraseDan 335, "Skriv besked og tryp 'enter' for at sende den"

    'REM:
    sPhraseEng 336, "Close"
    sPhraseIta 336, "Chiudi"
    sPhraseFra 336, "Fermer"
    sPhraseGer 336, "Schliessen"
    sPhraseSpa 336, "Cerrar"
    sPhraseSwe 336, "Stäng"
    sPhraseNor 336, "Lukk"
    sPhraseDan 336, "Luk"

    'REM:
    sPhraseEng 337, "<Var.ExeName> Editor"
    sPhraseIta 337, "<Var.ExeName> Editor"
    sPhraseFra 337, "L'Éditeur de <Var.ExeName>"
    sPhraseGer 337, "<Var.ExeName> Editor"
    sPhraseSpa 337, "Editor de <Var.ExeName>"
    sPhraseSwe 337, "<Var.ExeName>s editor"
    sPhraseNor 337, "<Var.ExeName> Editor"
    sPhraseDan 337, "<Var.ExeName> Editor"

    'REM:
    sPhraseEng 338, "Occupying army"
    sPhraseIta 338, "armate occupanti"
    sPhraseFra 338, "Armée occupante"
    sPhraseGer 338, "Besetzer-Armee"
    sPhraseSpa 338, "Ejercito ocupante"
    sPhraseSwe 338, "Ockuperande armé"
    sPhraseNor 338, "Okkuperende hær"
    sPhraseDan 338, "Indtaget af"

    'REM:
    sPhraseEng 339, "Units"
    sPhraseIta 339, "Unita'"
    sPhraseFra 339, "Unités"
    sPhraseGer 339, "Einheiten"
    sPhraseSpa 339, "Unidades"
    sPhraseSwe 339, "Enheter"
    sPhraseNor 339, "Enheter"
    sPhraseDan 339, "arméer"

    'REM:
    sPhraseEng 340, "OK"
    sPhraseIta 340, "OK"
    sPhraseFra 340, "OK"
    sPhraseGer 340, "OK"
    sPhraseSpa 340, "OK"
    sPhraseSwe 340, "OK"
    sPhraseNor 340, "OK"
    sPhraseDan 340, "OK"

    'REM:
    sPhraseEng 341, ""
    sPhraseIta 341, ""
    sPhraseFra 341, ""
    sPhraseGer 341, ""
    sPhraseSpa 341, ""
    sPhraseSwe 341, ""
    sPhraseNor 341, ""
    sPhraseDan 341, ""

    'REM:
    sPhraseEng 342, ""
    sPhraseIta 342, ""
    sPhraseFra 342, ""
    sPhraseGer 342, ""
    sPhraseSpa 342, ""
    sPhraseSwe 342, ""
    sPhraseNor 342, ""
    sPhraseDan 342, ""

    'REM:
    sPhraseEng 343, ""
    sPhraseIta 343, ""
    sPhraseFra 343, ""
    sPhraseGer 343, ""
    sPhraseSpa 343, ""
    sPhraseSwe 343, ""
    sPhraseNor 343, ""
    sPhraseDan 343, ""

    'REM:
    sPhraseEng 344, ""
    sPhraseIta 344, ""
    sPhraseFra 344, ""
    sPhraseGer 344, ""
    sPhraseSpa 344, ""
    sPhraseSwe 344, ""
    sPhraseNor 344, ""
    sPhraseDan 344, ""

    'REM:
    sPhraseEng 345, "War Statistics"
    sPhraseIta 345, "Statistiche di guerra"
    sPhraseFra 345, "Statistiques de guerre"
    sPhraseGer 345, "Kriegsstatistiken"
    sPhraseSpa 345, "Estadísticas de guerra"
    sPhraseSwe 345, "Krigsstatistik"
    sPhraseNor 345, "Statistikk for krigføringen"
    sPhraseDan 345, "Statestik"

    'REM:
    sPhraseEng 346, "Not involved"
    sPhraseIta 346, "non coinvolto"
    sPhraseFra 346, "N'implique pas"
    sPhraseGer 346, "Nicht einbezogen"
    sPhraseSpa 346, "No involucrado"
    sPhraseSwe 346, "Inte inblandad"
    sPhraseNor 346, "Ikke innvolvert"
    sPhraseDan 346, "Ikke involveret"

    'REM:End of war statistics.
    sPhraseEng 347, "&End of war stats"
    sPhraseIta 347, "&Statistiche di fine guerra"
    sPhraseFra 347, "Fin des statistiques de guerre"
    sPhraseGer 347, "&Kriegsstatistiken"
    sPhraseSpa 347, "&Fin de la estadísticas de guerra"
    sPhraseSwe 347, "Krigsslutsstatistik"
    sPhraseNor 347, "Statistikk"
    sPhraseDan 347, "Afsluttende statestik"

    'REM:
    sPhraseEng 348, "Countries conquered:"
    sPhraseIta 348, "Paesi conquistati:"
    sPhraseFra 348, "Pays conquis:"
    sPhraseGer 348, "Länder erobert:"
    sPhraseSpa 348, "Países conquistados:"
    sPhraseSwe 348, "Erövrade länder:"
    sPhraseNor 348, "Beseirede land:"
    sPhraseDan 348, "Lande erobret"

    'REM:
    sPhraseEng 349, "Countries surrendered:"
    sPhraseIta 349, "Paesi che si sono arresi:"
    sPhraseFra 349, "Pays perdus:"
    sPhraseGer 349, "Verlorene Länder:"
    sPhraseSpa 349, "Países que se rindieron:"
    sPhraseSwe 349, "Förlorade länder:"
    sPhraseNor 349, "Overgitte land:"
    sPhraseDan 349, "Lande tabt"

    'REM:
    sPhraseEng 350, "Enemy casualties:"
    sPhraseIta 350, "I danni del nemico:"
    sPhraseFra 350, "Pertes de l'ennemi:"
    sPhraseGer 350, "Feindliche Verluste:"
    sPhraseSpa 350, "Bajas del enemigo:"
    sPhraseSwe 350, "Fiendeförluster:"
    sPhraseNor 350, "Beseirede styrker:"
    sPhraseDan 350, "Arméer dræbt"

    'REM:
    sPhraseEng 351, "Your casualties:"
    sPhraseIta 351, "I tuoi danni:"
    sPhraseFra 351, "Vos pertes:"
    sPhraseGer 351, "Eigene Verluste:"
    sPhraseSpa 351, "Sus bajas:"
    sPhraseSwe 351, "Dina förluster:"
    sPhraseNor 351, "Tap av egne styrker:"
    sPhraseDan 351, "Arméer tabt"

    'REM:
    sPhraseEng 352, "IP Information"
    sPhraseIta 352, "Informazioni IP"
    sPhraseFra 352, "IP Information"
    sPhraseGer 352, "IP Information"
    sPhraseSpa 352, "Información IP"
    sPhraseSwe 352, "IP-information"
    sPhraseNor 352, "IP-informasjon"
    sPhraseDan 352, "IP information"

    'REM: Italian: If it's a command, otherwise if is a warning like I CAUGHT YOU then is HAI BARATO!
    sPhraseEng 353, "CHEAT!"
    sPhraseIta 353, "HAI BARATO!"
    sPhraseFra 353, "TRICHEUR!"
    sPhraseGer 353, "CHEAT!"
    sPhraseSpa 353, "¡TRAMPA!"
    sPhraseSwe 353, "FUSK!"
    sPhraseNor 353, "JUKS!"
    sPhraseDan 353, "SNYD!"

    'REM:
    sPhraseEng 354, "Score: "
    sPhraseIta 354, "Punteggio: "
    sPhraseFra 354, "Score: "
    sPhraseGer 354, "Punkte: "
    sPhraseSpa 354, "Puntaje: "
    sPhraseSwe 354, "Poäng: "
    sPhraseNor 354, "Poengsum: "
    sPhraseDan 354, "Point: "

    'REM:Ultra secret.
    sPhraseEng 355, "TOP SECRET"
    sPhraseIta 355, "TOP SECRET"
    sPhraseFra 355, "TOP SECRET"
    sPhraseGer 355, "TOP SECRET"
    sPhraseSpa 355, "ULTRA SECRETO"
    sPhraseSwe 355, "HEMLIGT"
    sPhraseNor 355, "HEMMELIG"
    sPhraseDan 355, "YDERST HEMMELIGT"

    'REM:
    sPhraseEng 356, "Your Mission Briefing"
    sPhraseIta 356, "Il tuo Briefing della missione"
    sPhraseFra 356, "Votre mission d'infor- mation"
    sPhraseGer 356, "Ihre Mission Briefing"
    sPhraseSpa 356, "Tu misión informa- tiva"
    sPhraseSwe 356, "Din Mission Briefing"
    sPhraseNor 356, "Din Mission Briefing"
    sPhraseDan 356, "Din mission Briefing"

    'REM:
    sPhraseEng 357, "Ignore"
    sPhraseIta 357, "Ignora"
    sPhraseFra 357, "Ignorer"
    sPhraseGer 357, "Ignorieren"
    sPhraseSpa 357, "Ignorar"
    sPhraseSwe 357, "Ignorera"
    sPhraseNor 357, "Ignorer"
    sPhraseDan 357, "Ignorer"

    'REM: Open mission briefing
    sPhraseEng 358, "Open"
    sPhraseIta 358, "Apri"
    sPhraseFra 358, "Ouvrir"
    sPhraseGer 358, "Öffnen"
    sPhraseSpa 358, "Abrir"
    sPhraseSwe 358, "Öppna"
    sPhraseNor 358, "Åpne"
    sPhraseDan 358, "Åben"

    'REM:
    sPhraseEng 359, "&View"
    sPhraseIta 359, "&Visualizza"
    sPhraseFra 359, "Voir"
    sPhraseGer 359, "A&nzeigen"
    sPhraseSpa 359, "&Ver"
    sPhraseSwe 359, "Visa"
    sPhraseNor 359, "Vis"
    sPhraseDan 359, "Vis"

    'REM:
    sPhraseEng 360, "&Mission reminder"
    sPhraseIta 360, "&Ricordami l'Obiettivo"
    sPhraseFra 360, "Rappel de mission"
    sPhraseGer 360, "Missions-&Erinnerung"
    sPhraseSpa 360, "&Recordatorio de la misión"
    sPhraseSwe 360, "Uppdragspåminnelse"
    sPhraseNor 360, "Oppdrag"
    sPhraseDan 360, "Mission reminder"

    'REM:A computer when connected to a network: example -I sit at this **terminal** to play MR- The context that it is used in: -Terminal 1 has connected to the network-. Computer.
    sPhraseEng 361, "Terminal"
    sPhraseIta 361, "Teerminale"
    sPhraseFra 361, "Terminal"
    sPhraseGer 361, "Terminal"
    sPhraseSpa 361, "Terminal"
    sPhraseSwe 361, "Dator"
    sPhraseNor 361, "Bruker"
    sPhraseDan 361, "computer"

    'REM:The Red Army. The red army.
    sPhraseEng 362, "the Red Army"
    sPhraseIta 362, "armate rosse"
    sPhraseFra 362, "l'Armée Rouge"
    sPhraseGer 362, "der roten Armee"
    sPhraseSpa 362, "el Ejército Rojo"
    sPhraseSwe 362, "den Röda Armén"
    sPhraseNor 362, "den Røde Hæren"
    sPhraseDan 362, "den røde hær"

    'REM:The Green Army. The green army.
    sPhraseEng 363, "the Green Army"
    sPhraseIta 363, "armate verdi"
    sPhraseFra 363, "L'Armée Verte"
    sPhraseGer 363, "der grünen Armee"
    sPhraseSpa 363, "el Ejército Verde"
    sPhraseSwe 363, "den Gröna Armén"
    sPhraseNor 363, "den Grønne Hæren"
    sPhraseDan 363, "den grønne hær"

    'REM:
    sPhraseEng 364, "the Blue Army"
    sPhraseIta 364, "armate blu"
    sPhraseFra 364, "l'Armée Bleue"
    sPhraseGer 364, "der blauen Armee"
    sPhraseSpa 364, "el Ejército Azul"
    sPhraseSwe 364, "den Blåa Armén"
    sPhraseNor 364, "den Blå Hæren"
    sPhraseDan 364, "den blå hær"

    'REM:
    sPhraseEng 365, "the Yellow Army"
    sPhraseIta 365, "armate gialle"
    sPhraseFra 365, "l'Armée Jaune"
    sPhraseGer 365, "der gelben Armee"
    sPhraseSpa 365, "el Ejército Amarillo"
    sPhraseSwe 365, "den Gula Armén"
    sPhraseNor 365, "den Gule Hæren"
    sPhraseDan 365, "den gule hær"

    'REM:
    sPhraseEng 366, "the Purple Army"
    sPhraseIta 366, "armate viola"
    sPhraseFra 366, "l'Armée Pourpre"
    sPhraseGer 366, "der lila Armee"
    sPhraseSpa 366, "el Ejército Purpura"
    sPhraseSwe 366, "den Lila Armén"
    sPhraseNor 366, "den Lilla Hæren"
    sPhraseDan 366, "den lilla hær"

    'REM:
    sPhraseEng 367, "the Gray Army"
    sPhraseIta 367, "armate grigie"
    sPhraseFra 367, "l'Armée Grise"
    sPhraseGer 367, "der grauen Armee"
    sPhraseSpa 367, "el Ejército Gris"
    sPhraseSwe 367, "den Gråa Armén"
    sPhraseNor 367, "den Grå Hæren"
    sPhraseDan 367, "den grønne hær"

    'REM:
    sPhraseEng 368, "Rolling Dice"
    sPhraseIta 368, "Stò tirando i dati"
    sPhraseFra 368, "Dés roulants"
    sPhraseGer 368, "Rollende Würfel"
    sPhraseSpa 368, "Dados Rodando"
    sPhraseSwe 368, "Rullande tärningar"
    sPhraseNor 368, "Kast terning"
    sPhraseDan 368, "Animerede terninger"

    'REM:Not actually used directly. IT:
    sPhraseEng 369, "Please select a language."
    sPhraseIta 369, "Selezionare una lingua."
    sPhraseFra 369, "S'il vous plaît sélectionnez une langue."
    sPhraseGer 369, "Wählen Sie bitte eine Sprache aus."
    sPhraseSpa 369, "Por favor seleccione un idioma."
    sPhraseSwe 369, "Var god välj språk."
    sPhraseNor 369, "Vennligst velg språk."
    sPhraseDan 369, "Vælg et sprog"

    'REM:
    sPhraseEng 370, ""
    sPhraseIta 370, ""
    sPhraseFra 370, ""
    sPhraseGer 370, ""
    sPhraseSpa 370, ""
    sPhraseSwe 370, ""
    sPhraseNor 370, ""
    sPhraseDan 370, ""

    'REM: French is not available yet: Français n'est pas disponible cependant.
    sPhraseEng 371, "Victory!"
    sPhraseIta 371, "Vittoria"
    sPhraseFra 371, "Victoire!"
    sPhraseGer 371, "Sieg"
    sPhraseSpa 371, "¡Victoria!"
    sPhraseSwe 371, "Seger!"
    sPhraseNor 371, "Seier!"
    sPhraseDan 371, "Sejer!"

    'REM:
    sPhraseEng 372, "Alaska"
    sPhraseIta 372, "Alaska"
    sPhraseFra 372, "Alaska"
    sPhraseGer 372, "Alaska"
    sPhraseSpa 372, "Alaska"
    sPhraseSwe 372, "Alaska"
    sPhraseNor 372, "Alaska"
    sPhraseDan 372, "Alaska"

    'REM:
    sPhraseEng 373, "Northwest Territory"
    sPhraseIta 373, "Territorio nordovest"
    sPhraseFra 373, "Territoire du nord-ouest"
    sPhraseGer 373, "Nordwestlich Territorium"
    sPhraseSpa 373, "El Territorio noroeste"
    sPhraseSwe 373, "Nordvästra territoriet"
    sPhraseNor 373, "Nordvest-territoriet"
    sPhraseDan 373, "Nordvest teritoriet"

    'REM:
    sPhraseEng 374, "Greenland"
    sPhraseIta 374, "Groenlandia"
    sPhraseFra 374, "Groenland"
    sPhraseGer 374, "Grönland"
    sPhraseSpa 374, "Groenlandia"
    sPhraseSwe 374, "Grönland"
    sPhraseNor 374, "Grønland"
    sPhraseDan 374, "Grønland"

    'REM:
    sPhraseEng 375, "British Columbia"
    sPhraseIta 375, "British Columbia"
    sPhraseFra 375, "British Columbia"
    sPhraseGer 375, "British Columbia"
    sPhraseSpa 375, "British Columbia"
    sPhraseSwe 375, "British Columbia"
    sPhraseNor 375, "British Columbia"
    sPhraseDan 375, "British Columbia"

    'REM:
    sPhraseEng 376, "Ontario"
    sPhraseIta 376, "Ontario"
    sPhraseFra 376, "Ontario"
    sPhraseGer 376, "Ontario"
    sPhraseSpa 376, "Ontario"
    sPhraseSwe 376, "Ontario"
    sPhraseNor 376, "Ontario"
    sPhraseDan 376, "Ontario"

    'REM:
    sPhraseEng 377, "Quebec"
    sPhraseIta 377, "Quebec"
    sPhraseFra 377, "Québec"
    sPhraseGer 377, "Quebec"
    sPhraseSpa 377, "Quebec"
    sPhraseSwe 377, "Quebec"
    sPhraseNor 377, "Quebec"
    sPhraseDan 377, "Quebec"

    'REM:
    sPhraseEng 378, "Western United States"
    sPhraseIta 378, "Stati Uniti occidentali"
    sPhraseFra 378, "Etats Unis de l'ouest"
    sPhraseGer 378, "West USA"
    sPhraseSpa 378, "EE.UU. occidental"
    sPhraseSwe 378, "Västra USA"
    sPhraseNor 378, "Vestlige USA"
    sPhraseDan 378, "Vestlige USA"

    'REM:
    sPhraseEng 379, "Eastern United States"
    sPhraseIta 379, "Stati Uniti orientali"
    sPhraseFra 379, "Etats Unis de l'est"
    sPhraseGer 379, "Ost USA"
    sPhraseSpa 379, "EE.UU. oriental"
    sPhraseSwe 379, "Östra USA"
    sPhraseNor 379, "Østlige USA"
    sPhraseDan 379, "Østlige USA"

    'REM:
    sPhraseEng 380, "Mexico"
    sPhraseIta 380, "Mexico"
    sPhraseFra 380, "Mexico"
    sPhraseGer 380, "Mexico"
    sPhraseSpa 380, "Mexico"
    sPhraseSwe 380, "Mexico"
    sPhraseNor 380, "Mexico"
    sPhraseDan 380, "Mexico"

    'REM:
    sPhraseEng 381, "Colombia"
    sPhraseIta 381, "Colombia"
    sPhraseFra 381, "Colombia"
    sPhraseGer 381, "Colombia"
    sPhraseSpa 381, "Colombia"
    sPhraseSwe 381, "Colombia"
    sPhraseNor 381, "Colombia"
    sPhraseDan 381, "Colombia"

    'REM:
    sPhraseEng 382, "Peru"
    sPhraseIta 382, "Perù"
    sPhraseFra 382, "Pérou"
    sPhraseGer 382, "Peru"
    sPhraseSpa 382, "Perú"
    sPhraseSwe 382, "Peru"
    sPhraseNor 382, "Peru"
    sPhraseDan 382, "Peru"

    'REM:
    sPhraseEng 383, "Brazil"
    sPhraseIta 383, "Brasile"
    sPhraseFra 383, "Brésil"
    sPhraseGer 383, "Brasilien"
    sPhraseSpa 383, "Brasil"
    sPhraseSwe 383, "Brasilien"
    sPhraseNor 383, "Brasil"
    sPhraseDan 383, "Brasilien"

    'REM:
    sPhraseEng 384, "Argentina"
    sPhraseIta 384, "Argentina"
    sPhraseFra 384, "Argentine"
    sPhraseGer 384, "Argentinien"
    sPhraseSpa 384, "Argentina"
    sPhraseSwe 384, "Argentina"
    sPhraseNor 384, "Argentina"
    sPhraseDan 384, "Argentina"

    'REM:
    sPhraseEng 385, "Iceland"
    sPhraseIta 385, "Islanda"
    sPhraseFra 385, "Islande"
    sPhraseGer 385, "Island"
    sPhraseSpa 385, "Islandia"
    sPhraseSwe 385, "Island"
    sPhraseNor 385, "Island"
    sPhraseDan 385, "Island"

    'REM:
    sPhraseEng 386, "Great Britain"
    sPhraseIta 386, "Gran Bretagna"
    sPhraseFra 386, "Grande-Bretagne"
    sPhraseGer 386, "Großbritannien"
    sPhraseSpa 386, "Gran Bretaña"
    sPhraseSwe 386, "Storbritannien"
    sPhraseNor 386, "Storbritania"
    sPhraseDan 386, "Storbritanien"

    'REM:
    sPhraseEng 387, "Scandinavia"
    sPhraseIta 387, "Scandinavia"
    sPhraseFra 387, "Scandinavie"
    sPhraseGer 387, "Skandinavien"
    sPhraseSpa 387, "Escandinavia"
    sPhraseSwe 387, "Skandinavien"
    sPhraseNor 387, "Skandinavia"
    sPhraseDan 387, "Skandinavien"

    'REM:
    sPhraseEng 388, "Germania"
    sPhraseIta 388, "Germania"
    sPhraseFra 388, "Germania"
    sPhraseGer 388, "Germania"
    sPhraseSpa 388, "Germania"
    sPhraseSwe 388, "Germania"
    sPhraseNor 388, "Germania"
    sPhraseDan 388, "Germania"

    'REM:
    sPhraseEng 389, "Spain"
    sPhraseIta 389, "Spain"
    sPhraseFra 389, "Spain"
    sPhraseGer 389, "Spain"
    sPhraseSpa 389, "Spain"
    sPhraseSwe 389, "Spain"
    sPhraseNor 389, "Spain"
    sPhraseDan 389, "Spain"

    'REM:
    sPhraseEng 390, "The Mediterranean"
    sPhraseIta 390, "The Mediterranean"
    sPhraseFra 390, "The Mediterranean"
    sPhraseGer 390, "The Mediterranean"
    sPhraseSpa 390, "The Mediterranean"
    sPhraseSwe 390, "The Mediterranean"
    sPhraseNor 390, "The Mediterranean"
    sPhraseDan 390, "The Mediterranean"

    'REM:
    sPhraseEng 391, "Prussia"
    sPhraseIta 391, "Prussia"
    sPhraseFra 391, "Prussia"
    sPhraseGer 391, "Prussia"
    sPhraseSpa 391, "Prussia"
    sPhraseSwe 391, "Prussia"
    sPhraseNor 391, "Prussia"
    sPhraseDan 391, "Prussia"

    'REM:
    sPhraseEng 392, "Algeria"
    sPhraseIta 392, "Algeria"
    sPhraseFra 392, "Algeria"
    sPhraseGer 392, "Algeria"
    sPhraseSpa 392, "Algeria"
    sPhraseSwe 392, "Algeria"
    sPhraseNor 392, "Algeria"
    sPhraseDan 392, "Algeria"

    'REM:
    sPhraseEng 393, "Egypt"
    sPhraseIta 393, "Egitto"
    sPhraseFra 393, "Egypte"
    sPhraseGer 393, "Ägypten"
    sPhraseSpa 393, "Egipto"
    sPhraseSwe 393, "Egypten"
    sPhraseNor 393, "Egypt"
    sPhraseDan 393, "Egypten"

    'REM:
    sPhraseEng 394, "Ethiopia"
    sPhraseIta 394, "Ethiopia"
    sPhraseFra 394, "Ethiopia"
    sPhraseGer 394, "Ethiopia"
    sPhraseSpa 394, "Ethiopia"
    sPhraseSwe 394, "Ethiopia"
    sPhraseNor 394, "Ethiopia"
    sPhraseDan 394, "Ethiopia"

    'REM:
    sPhraseEng 395, "Congo"
    sPhraseIta 395, "Congo"
    sPhraseFra 395, "Congo"
    sPhraseGer 395, "Kongo"
    sPhraseSpa 395, "Congo"
    sPhraseSwe 395, "Kongo"
    sPhraseNor 395, "Kongo"
    sPhraseDan 395, "Congo"

    'REM:
    sPhraseEng 396, "South Africa"
    sPhraseIta 396, "Africa meridionale"
    sPhraseFra 396, "Afrique du Sud"
    sPhraseGer 396, "Südafrika"
    sPhraseSpa 396, "Africa Sur"
    sPhraseSwe 396, "Sydafrika"
    sPhraseNor 396, "Sør-Afrika"
    sPhraseDan 396, "Sydafrika"

    'REM:
    sPhraseEng 397, "Madagascar"
    sPhraseIta 397, "Madagascar"
    sPhraseFra 397, "Madagascar"
    sPhraseGer 397, "Madagaskar"
    sPhraseSpa 397, "Madagascar"
    sPhraseSwe 397, "Madagaskar"
    sPhraseNor 397, "Madagaskar"
    sPhraseDan 397, "Madagascar"

    'REM:
    sPhraseEng 398, "Saudi Arabia"
    sPhraseIta 398, "Saudi Arabia"
    sPhraseFra 398, "Saudi Arabia"
    sPhraseGer 398, "Saudi Arabia"
    sPhraseSpa 398, "Saudi Arabia"
    sPhraseSwe 398, "Saudi Arabia"
    sPhraseNor 398, "Saudi Arabia"
    sPhraseDan 398, "Saudi Arabia"

    'REM:
    sPhraseEng 399, "Afghanistan"
    sPhraseIta 399, "Afghanistan"
    sPhraseFra 399, "Afghanistan"
    sPhraseGer 399, "Afghanistan"
    sPhraseSpa 399, "Afganistán"
    sPhraseSwe 399, "Afghanistan"
    sPhraseNor 399, "Afganistan"
    sPhraseDan 399, "Afganistan"
    Call initialisePhrases400
End Sub

Private Sub initialisePhrases400()
    'REM:
    sPhraseEng 400, "India"
    sPhraseIta 400, "India"
    sPhraseFra 400, "Inde"
    sPhraseGer 400, "Indien"
    sPhraseSpa 400, "India"
    sPhraseSwe 400, "Indien"
    sPhraseNor 400, "India"
    sPhraseDan 400, "Indien"

    'REM:
    sPhraseEng 401, "Cambodia"
    sPhraseIta 401, "Cambodia"
    sPhraseFra 401, "Cambodia"
    sPhraseGer 401, "Cambodia"
    sPhraseSpa 401, "Cambodia"
    sPhraseSwe 401, "Cambodia"
    sPhraseNor 401, "Cambodia"
    sPhraseDan 401, "Cambodia"

    'REM:
    sPhraseEng 402, "Siberia"
    sPhraseIta 402, "Siberia"
    sPhraseFra 402, "Sibérie"
    sPhraseGer 402, "Sibirien"
    sPhraseSpa 402, "Siberia"
    sPhraseSwe 402, "Sibirien"
    sPhraseNor 402, "Sibir"
    sPhraseDan 402, "Siberien"

    'REM:
    sPhraseEng 403, "Krasnoyarsk"
    sPhraseIta 403, "Krasnoyarsk"
    sPhraseFra 403, "Krasnoyarsk"
    sPhraseGer 403, "Krasnoyarsk"
    sPhraseSpa 403, "Krasnoyarsk"
    sPhraseSwe 403, "Krasnoyarsk"
    sPhraseNor 403, "Krasnoyarsk"
    sPhraseDan 403, "Krasnoyarsk"

    'REM:
    sPhraseEng 404, "China"
    sPhraseIta 404, "Cina"
    sPhraseFra 404, "Chine"
    sPhraseGer 404, "China"
    sPhraseSpa 404, "China"
    sPhraseSwe 404, "Kina"
    sPhraseNor 404, "Kina"
    sPhraseDan 404, "Kina"

    'REM:
    sPhraseEng 405, "Korea"
    sPhraseIta 405, "Korea"
    sPhraseFra 405, "Korea"
    sPhraseGer 405, "Korea"
    sPhraseSpa 405, "Korea"
    sPhraseSwe 405, "Korea"
    sPhraseNor 405, "Korea"
    sPhraseDan 405, "Korea"

    'REM:
    sPhraseEng 406, "Magadan"
    sPhraseIta 406, "Magadan"
    sPhraseFra 406, "Magadan"
    sPhraseGer 406, "Magadan"
    sPhraseSpa 406, "Magadan"
    sPhraseSwe 406, "Magadan"
    sPhraseNor 406, "Magadan"
    sPhraseDan 406, "Magadan"

    'REM:
    sPhraseEng 407, "Chukotka"
    sPhraseIta 407, "Chukotka"
    sPhraseFra 407, "Chukotka"
    sPhraseGer 407, "Chukotka"
    sPhraseSpa 407, "Chukotka"
    sPhraseSwe 407, "Chukotka"
    sPhraseNor 407, "Chukotka"
    sPhraseDan 407, "Chukotka"

    'REM:
    sPhraseEng 408, "Japan"
    sPhraseIta 408, "Giappone"
    sPhraseFra 408, "Japon"
    sPhraseGer 408, "Japan"
    sPhraseSpa 408, "Japón"
    sPhraseSwe 408, "Japan"
    sPhraseNor 408, "Japan"
    sPhraseDan 408, "Japan"

    'REM:
    sPhraseEng 409, "Kamchatka"
    sPhraseIta 409, "Kamchatka"
    sPhraseFra 409, "Kamchatka"
    sPhraseGer 409, "Kamchatka"
    sPhraseSpa 409, "Kamchatka"
    sPhraseSwe 409, "Kamtjatka"
    sPhraseNor 409, "Kamtsjakta"
    sPhraseDan 409, "Kamchatka"

    'REM:
    sPhraseEng 410, "Indonesia"
    sPhraseIta 410, "Indonesia"
    sPhraseFra 410, "Indonésie"
    sPhraseGer 410, "Indonesien"
    sPhraseSpa 410, "Indonesia"
    sPhraseSwe 410, "Indonesien"
    sPhraseNor 410, "Indonesia"
    sPhraseDan 410, "Indonesien"

    'REM:
    sPhraseEng 411, "Western Australia"
    sPhraseIta 411, "Australia occidentale"
    sPhraseFra 411, "Australie de l'ouest"
    sPhraseGer 411, "West Australien"
    sPhraseSpa 411, "Australia occidental"
    sPhraseSwe 411, "Västaustralien"
    sPhraseNor 411, "Vest-Australia"
    sPhraseDan 411, "Vestlige Australien"

    'REM:
    sPhraseEng 412, "New Guinea"
    sPhraseIta 412, "Ghinea Nuova"
    sPhraseFra 412, "Nouvelle-Guinée"
    sPhraseGer 412, "Neuguinea"
    sPhraseSpa 412, "La Nueva Guinea"
    sPhraseSwe 412, "Nya Guinea"
    sPhraseNor 412, "Ny-Guinea"
    sPhraseDan 412, "Ny Guinea"

    'REM:
    sPhraseEng 413, "Eastern Australia NZ"
    sPhraseIta 413, "Australia orientale NZ"
    sPhraseFra 413, "Australie de l'est NZ"
    sPhraseGer 413, "Ost Australien NZ"
    sPhraseSpa 413, "Australia oriental NZ"
    sPhraseSwe 413, "Östaustralien NZ"
    sPhraseNor 413, "Øst-Australia NZ"
    sPhraseDan 413, "Østlige Australien NZ"

    'REM:Continent 1 not within a sentence
    sPhraseEng 414, "North America"
    sPhraseIta 414, "Nord l'America"
    sPhraseFra 414, "Amérique du Nord"
    sPhraseGer 414, "Nordamerika"
    sPhraseSpa 414, "América del Norte"
    sPhraseSwe 414, "Nordamerika"
    sPhraseNor 414, "Nord-Amerika"
    sPhraseDan 414, "Nordamerika"

    'REM:Continent 2 not within a sentence
    sPhraseEng 415, "South America"
    sPhraseIta 415, "L'America meridionale"
    sPhraseFra 415, "Amérique du Sud"
    sPhraseGer 415, "Südamerika"
    sPhraseSpa 415, "América del Sur"
    sPhraseSwe 415, "Sydamerika"
    sPhraseNor 415, "Sør-Amerika"
    sPhraseDan 415, "Sydamerika"

    'REM:Continent 3 not within a sentence
    sPhraseEng 416, "Europe"
    sPhraseIta 416, "L'Europa"
    sPhraseFra 416, "Europe"
    sPhraseGer 416, "Europa"
    sPhraseSpa 416, "Europa"
    sPhraseSwe 416, "Europa"
    sPhraseNor 416, "Europa"
    sPhraseDan 416, "Europa"

    'REM:Continent 4 not within a sentence
    sPhraseEng 417, "Africa"
    sPhraseIta 417, "L'Africa"
    sPhraseFra 417, "Afrique"
    sPhraseGer 417, "Afrika"
    sPhraseSpa 417, "Africa"
    sPhraseSwe 417, "Afrika"
    sPhraseNor 417, "Afrika"
    sPhraseDan 417, "Afrika"

    'REM:Continent 5 not within a sentence
    sPhraseEng 418, "Asia"
    sPhraseIta 418, "L'Asia"
    sPhraseFra 418, "Asie"
    sPhraseGer 418, "Asien"
    sPhraseSpa 418, "Asia"
    sPhraseSwe 418, "Asien"
    sPhraseNor 418, "Asia"
    sPhraseDan 418, "Asien"

    'REM:Continent 6 not within a sentence
    sPhraseEng 419, "Australia"
    sPhraseIta 419, "L'Australia"
    sPhraseFra 419, "Australie"
    sPhraseGer 419, "Australien"
    sPhraseSpa 419, "Australia"
    sPhraseSwe 419, "Australien"
    sPhraseNor 419, "Australia"
    sPhraseDan 419, "Australien"

    'REM:
    sPhraseEng 420, "Multiplayer"
    sPhraseIta 420, "&Impostazione rete"
    sPhraseFra 420, "Paramêtres du réseau"
    sPhraseGer 420, "Netzwerk-&Einstellungen"
    sPhraseSpa 420, "Ajuste red de traba&jo"
    sPhraseSwe 420, "Nätverksinställningar"
    sPhraseNor 420, "Nettverks-oppsett"
    sPhraseDan 420, "Multiplayer"

    'REM:This computer IP address is 555.
    sPhraseEng 421, "Your Local IP address is "
    sPhraseIta 421, "IP: "
    sPhraseFra 421, "IP: "
    sPhraseGer 421, "IP: "
    sPhraseSpa 421, "IP: "
    sPhraseSwe 421, "IP: "
    sPhraseNor 421, "IP: "
    sPhraseDan 421, "IP: "

    'REM:
    sPhraseEng 422, "Your system pallet should be set to HIGH COLOR for 3D rendering. Do you wish to continue changing to the 3D display mode?"
    sPhraseIta 422, "Per avere gli effetti 3D, bisogna settare la pallete di colori in HIGH COLOR. Continuo cambiando in questa modalita'?"
    sPhraseFra 422, "Votre palette système devrait être mise à COULEUR (16 bits ou > ) pour l'effet 3D. Est-ce que vous souhaitez continuer à changer en  mode 3D?"
    sPhraseGer 422, "Ihre Systemfarben sollten auf HIGH COLOR (16bit) für 3D Rendering eingestellt sein. Wollen Sie das Wechseln zur 3D Anzeige fortsetzen?"
    sPhraseSpa 422, "La paleta de su sistema debe ser puesta en COLOR MÁXIMO para recreación 3D. ¿Desea continuar cambiando a modo pantalla 3D?"
    sPhraseSwe 422, "Din systempalett bör vara inställd på HIGH COLOR-mode för att visa 3D. Vill du ändå växla till 3D-läge? "
    sPhraseNor 422, "Du må velge større fargedybde enn 256 farger for din skjerm. Ønsker du fortsatt å endre visning til 3D-modus?"
    sPhraseDan 422, "Din computer skal være sat til ægte farver. Vil du stadig skifte til 3D display?"

    'REM:Continent 1 within a sentence
    sPhraseEng 423, "North America"
    sPhraseIta 423, "Nord America"
    sPhraseFra 423, "l'Amérique du Nord"
    sPhraseGer 423, "Nord-Amerika"
    sPhraseSpa 423, "Norteamérica"
    sPhraseSwe 423, "Nordamerika"
    sPhraseNor 423, "Nord-Amerika"
    sPhraseDan 423, "Nordamerika"

    'REM:Continent 2 within a sentence
    sPhraseEng 424, "South America"
    sPhraseIta 424, "Sud America"
    sPhraseFra 424, "l'Amérique du Sud"
    sPhraseGer 424, "Süd-Amerika"
    sPhraseSpa 424, "Sudamérica"
    sPhraseSwe 424, "Sydamerika"
    sPhraseNor 424, "Sør-Amerika"
    sPhraseDan 424, "Sydamerika"

    'REM:Continent 3 within a sentence
    sPhraseEng 425, "Europe"
    sPhraseIta 425, "Europa"
    sPhraseFra 425, "l'Europe"
    sPhraseGer 425, "Europa"
    sPhraseSpa 425, "Europa"
    sPhraseSwe 425, "Europa"
    sPhraseNor 425, "Europa"
    sPhraseDan 425, "Europa"

    'REM:Continent 4 within a sentence
    sPhraseEng 426, "Africa"
    sPhraseIta 426, "Africa"
    sPhraseFra 426, "l'Afrique"
    sPhraseGer 426, "Afrika"
    sPhraseSpa 426, "Afrika"
    sPhraseSwe 426, "Afrika"
    sPhraseNor 426, "Afrika"
    sPhraseDan 426, "Afrika"

    'REM:Continent 5 within a sentence
    sPhraseEng 427, "Asia"
    sPhraseIta 427, "Asia"
    sPhraseFra 427, "l'Asie"
    sPhraseGer 427, "Asien"
    sPhraseSpa 427, "Asia"
    sPhraseSwe 427, "Asien"
    sPhraseNor 427, "Asia"
    sPhraseDan 427, "Asien"

    'REM:Continent 6 within a sentence
    sPhraseEng 428, "Australia"
    sPhraseIta 428, "Australia"
    sPhraseFra 428, "l'Australie"
    sPhraseGer 428, "Australien"
    sPhraseSpa 428, "Australia"
    sPhraseSwe 428, "Australien"
    sPhraseNor 428, "Australia"
    sPhraseDan 428, "Australien"

End Sub     'Do not put code beyond this point!!!
