Function GetTickerFromISIN(isinCode As String) As String
    '====================================================================================================
    ' FUNZIONE: GetTickerFromISIN
    ' OBIETTIVO TECNICO: Isolare e restituire il Ticker symbol associato a un dato codice ISIN.
    '                    L'operazione si basa sull'interrogazione della pagina di lookup di Yahoo Finance
    '                    e sul parsing mirato della risposta HTML.
    ' PARAMETRI DI INPUT:
    '   isinCode (String) - Il codice ISIN (International Securities Identification Number) da risolvere.
    '                       Viene assunta la sua validità formale; la funzione non esegue validazione ISO 6166.
    ' VALORE DI RITORNO:
    '   String - Il Ticker symbol estratto (es. "AAPL"). In caso di fallimento nel recupero o nel parsing,
    '            restituisce una stringa descrittiva dello stato (es. "Ticker non trovato",
    '            "Errore HTTP...", "ISIN vuoto fornito", etc.).
    '====================================================================================================

    ' Dichiarazione delle variabili utilizzate per la gestione della richiesta HTTP e del parsing.
    Dim http As Object          ' Oggetto per l'esecuzione della richiesta HTTP. Verrà istanziato come MSXML2.XMLHTTP.
    Dim url As String           ' Stringa contenente l'URL completo per la query di lookup su Yahoo Finance.
    Dim responseText As String  ' Stringa destinata a contenere il corpo della risposta HTML dal server.
    Dim tickerResult As String  ' Variabile per immagazzinare il ticker estratto o i messaggi di stato/errore.
    
    tickerResult = "Ticker non trovato" ' Inizializzazione prudenziale del risultato.

    ' Validazione preliminare dell'input: un ISIN nullo o vuoto non permette di procedere.
    If Trim(isinCode) = "" Then
        GetTickerFromISIN = "ISIN vuoto fornito" ' Messaggio di errore specifico.
        Exit Function ' Interruzione anticipata della funzione.
    End If

    ' Istanziazione dell'oggetto MSXML2.XMLHTTP.
    ' Si tenta di creare la versione più recente (6.0) per massimizzare compatibilità e funzionalità.
    ' In caso di fallimento (comune su sistemi meno recenti o con configurazioni particolari),
    ' si procede con tentativi fallback su versioni precedenti (3.0, poi la generica).
    ' L'istruzione 'On Error Resume Next' permette di gestire questi tentativi in modo controllato.
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP.60.0")
    If Err.Number <> 0 Then ' Se la creazione della v6.0 fallisce...
        Err.Clear ' Pulisce l'errore per il prossimo tentativo.
        Set http = CreateObject("MSXML2.XMLHTTP.3.0") ' ...tenta con la v3.0.
    End If
    If Err.Number <> 0 Then ' Se anche la v3.0 fallisce...
        Err.Clear ' Pulisce l'errore.
        Set http = CreateObject("MSXML2.XMLHTTP") ' ...tenta con la versione base.
    End If
    On Error GoTo 0 ' Disabilita la gestione errori 'Resume Next' per ripristinare il comportamento standard.
    
    ' Verifica critica: se l'oggetto http non è stato creato, la comunicazione web è impossibile.
    If http Is Nothing Then
        GetTickerFromISIN = "Errore creazione oggetto HTTP" ' Segnala fallimento critico.
        Exit Function ' Interruzione.
    End If

    ' Composizione dell'URL per interrogare il servizio di lookup di Yahoo Finance.
    ' L'ISIN viene concatenato direttamente. Si assume che gli ISIN standard (alfanumerici)
    ' non richiedano URL encoding specifico per questo endpoint.
    url = "https://finance.yahoo.com/lookup/?s=" & isinCode

    ' Configurazione e invio della richiesta HTTP GET.
    On Error Resume Next ' Abilita la gestione errori per le operazioni di rete (Open, Send).
    http.Open "GET", url, False ' Metodo GET; URL specificato; False indica una richiesta sincrona (l'esecuzione VBA attende la risposta).
    ' Impostazione dell'header User-Agent per simulare una richiesta da un browser comune.
    ' Alcuni server web potrebbero rifiutare richieste senza uno User-Agent valido o noto.
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
    http.send ' Invio effettivo della richiesta.
    
    If Err.Number <> 0 Then ' Se si verifica un errore durante .Open o .send (es. rete irraggiungibile)
        tickerResult = "Errore invio richiesta lookup: (" & Err.Number & ") " & Err.Description
        GoTo CleanupTickerFunction ' Procedura di uscita controllata.
    End If
    On Error GoTo 0 ' Disabilita gestione errori 'Resume Next'.

    ' Verifica dello status code HTTP della risposta. 200 (OK) indica successo.
    If http.Status = 200 Then
        responseText = http.responseText ' Acquisizione del corpo HTML della risposta.
        
        ' --- Inizio della fase di Parsing dell'HTML per l'estrazione del Ticker ---
        ' Questa sezione implementa una logica di string-matching per identificare e isolare il ticker.
        ' Si basa sull'osservazione della struttura HTML tipica delle pagine di risultato di Yahoo Finance.
        
        Dim markerClasse As String      ' Stringa identificativa per le classi CSS del link target.
        Dim markerHrefStart As String   ' Stringa identificativa per l'inizio dell'attributo href del ticker.
        Dim markerHrefEnd As String     ' Stringa identificativa per la fine del ticker nell'href.
        Dim posClasse As Long           ' Posizione del marcatore di classe.
        Dim posHrefStart As Long        ' Posizione iniziale dell'href utile.
        Dim posHrefEnd As Long          ' Posizione finale del ticker nell'href.

        ' Definiamo i nostri "segnali" per la caccia al ticker nell'HTML.
        markerClasse = "loud-link fin-size-medium ellipsis" ' Combinazione di classi CSS spesso usata per il link principale del ticker.
        markerHrefStart = "href=""/quote/"                  ' Prefisso dell'attributo href per i link ai profili dei ticker.
        markerHrefEnd = "/"                                 ' Suffisso (delimitatore) del ticker nell'href.

        ' Ricerca del primo marcatore (la classe CSS). `vbTextCompare` per ignorare case.
        posClasse = InStr(1, responseText, markerClasse, vbTextCompare)

        If posClasse > 0 Then ' Marcatore di classe trovato. Si procede con l'analisi del contesto.
            ' Ricerca del marcatore href a partire dalla posizione della classe trovata.
            posHrefStart = InStr(posClasse, responseText, markerHrefStart, vbTextCompare)
            If posHrefStart > 0 Then
                ' Posizionamento sull'inizio effettivo del ticker nell'href.
                posHrefStart = posHrefStart + Len(markerHrefStart)
                ' Ricerca del delimitatore finale del ticker.
                posHrefEnd = InStr(posHrefStart, responseText, markerHrefEnd, vbTextCompare)
                If posHrefEnd > 0 Then
                    ' Estrazione del ticker.
                    tickerResult = Mid(responseText, posHrefStart, posHrefEnd - posHrefStart)
                Else
                    tickerResult = "Ticker: Formato href non riconosciuto (manca / finale)"
                End If
            Else
                ' Strategia di Fallback #1: Estrazione dall'attributo 'data-ylk', componente 'slk:'.
                ' Questo attributo è spesso usato da Yahoo per metadati interni.
                Dim markerDataYlkSlk As String, posDataYlkSlkStart As Long, posDataYlkSlkEnd As Long
                Dim posDataYlk As Long, posQuoteEnd As Long ' Variabili per il parsing di data-ylk.
                
                markerDataYlkSlk = "slk:" ' Sotto-componente di interesse.
                posDataYlk = InStr(posClasse, responseText, "data-ylk=""", vbTextCompare) ' Localizza l'attributo.
                If posDataYlk > 0 Then
                    posDataYlkSlkStart = InStr(posDataYlk, responseText, markerDataYlkSlk, vbTextCompare)
                    If posDataYlkSlkStart > 0 Then
                        posDataYlkSlkStart = posDataYlkSlkStart + Len(markerDataYlkSlk)
                        ' Il ticker in 'slk:' può essere delimitato da ';' o dalle virgolette di chiusura dell'attributo.
                        posDataYlkSlkEnd = InStr(posDataYlkSlkStart, responseText, ";", vbTextCompare)
                        posQuoteEnd = InStr(posDataYlkSlkStart, responseText, """", vbTextCompare)
                        
                        If posDataYlkSlkEnd > 0 And (posQuoteEnd = 0 Or posDataYlkSlkEnd < posQuoteEnd) Then
                            tickerResult = Mid(responseText, posDataYlkSlkStart, posDataYlkSlkEnd - posDataYlkSlkStart)
                        ElseIf posQuoteEnd > 0 Then
                            tickerResult = Mid(responseText, posDataYlkSlkStart, posQuoteEnd - posDataYlkSlkStart)
                        Else
                            tickerResult = "Ticker: Formato slk in data-ylk non riconosciuto"
                        End If
                    Else
                        tickerResult = "Ticker: Marcatore slk: non trovato in data-ylk"
                    End If
                Else
                    tickerResult = "Ticker: Marcatore href/data-ylk non trovato dopo la classe"
                End If
            End If
        Else
            ' Strategia di Fallback #2: Se la classe primaria non è trovata, si verifica se la pagina
            ' è una pagina di risultati generici "Symbols similar to" e si tenta di estrarre il primo ticker.
            If InStr(1, responseText, "Symbols similar to", vbTextCompare) > 0 Then
                posHrefStart = InStr(1, responseText, markerHrefStart, vbTextCompare) ' Usa i markerHref definiti prima.
                If posHrefStart > 0 Then
                    posHrefStart = posHrefStart + Len(markerHrefStart)
                    posHrefEnd = InStr(posHrefStart, responseText, markerHrefEnd, vbTextCompare)
                    If posHrefEnd > 0 Then
                        tickerResult = Mid(responseText, posHrefStart, posHrefEnd - posHrefStart)
                    Else
                        tickerResult = "Ticker: Formato href (fallback) non riconosciuto"
                    End If
                Else
                    tickerResult = "Ticker: Nessun link /quote/ in pagina 'Symbols similar to'"
                End If
            Else
                tickerResult = "Ticker: Struttura pagina (classe primaria) non riconosciuta"
            End If
        End If
        ' --- Fine della fase di Parsing ---
    Else
        ' La richiesta HTTP non ha avuto successo (status code diverso da 200).
        tickerResult = "Ticker: Errore HTTP " & http.Status & " - " & http.statusText
    End If

CleanupTickerFunction: ' Etichetta per il cleanup e l'uscita.
    Set http = Nothing ' Rilascio esplicito dell'oggetto HTTP per liberare risorse.
    GetTickerFromISIN = tickerResult ' Assegnazione finale del risultato alla funzione.
End Function


Function GetYahooProfileData(tickerSymbol As String, ByRef sectorOutput As String, ByRef currencyOutput As String) As Boolean
    '====================================================================================================
    ' FUNZIONE: GetYahooProfileData
    ' OBIETTIVO TECNICO: Recuperare informazioni specifiche (Settore, Valuta) dalla pagina profilo
    '                    di un titolo su Yahoo Finance, dato il suo Ticker Symbol.
    ' PARAMETRI DI INPUT:
    '   tickerSymbol (String) - Il Ticker Symbol del titolo (es. "AAPL").
    ' PARAMETRI DI OUTPUT (ByRef):
    '   sectorOutput (String) - Variabile passata per riferimento che verrà popolata con il Settore del titolo.
    '   currencyOutput (String) - Variabile passata per riferimento che verrà popolata con la Valuta del titolo (codice a 3 lettere).
    ' VALORE DI RITORNO:
    '   Boolean - True se la richiesta HTTP alla pagina profilo ha successo (status 200) e il parsing
    '             viene tentato. False in caso di errore HTTP o input non valido.
    '             La riuscita effettiva del parsing dei singoli dati è riflessa nei parametri ByRef.
    '====================================================================================================

    ' Dichiarazione variabili per la richiesta HTTP e il parsing.
    Dim http As Object              ' Oggetto per la richiesta HTTP.
    Dim profileUrl As String        ' URL della pagina profilo del ticker.
    Dim profileResponseText As String ' Corpo HTML della risposta dalla pagina profilo.
    
    ' Inizializzazione dei parametri di output ByRef e del valore di ritorno della funzione.
    sectorOutput = "Settore non trovato"
    currencyOutput = "Valuta non trovata"
    GetYahooProfileData = False ' Assumiamo fallimento fino a prova contraria (successo HTTP).

    ' Validazione preliminare dell'input: un ticker non valido o vuoto non permette di procedere.
    If Trim(tickerSymbol) = "" Or Trim(tickerSymbol) Like "*Non trovato*" Or Trim(tickerSymbol) Like "*Errore*" Or Trim(tickerSymbol) Like "*non ric.*" Then
        sectorOutput = "Ticker non valido fornito alla funzione"
        currencyOutput = "Ticker non valido fornito alla funzione"
        Exit Function ' Esce, la funzione restituirà False come da inizializzazione.
    End If

    ' Istanziazione dell'oggetto MSXML2.XMLHTTP (vedi commenti in GetTickerFromISIN per i dettagli).
    On Error Resume Next
    Set http = CreateObject("MSXML2.XMLHTTP.60.0")
    If Err.Number <> 0 Then Err.Clear : Set http = CreateObject("MSXML2.XMLHTTP.3.0")
    If Err.Number <> 0 Then Err.Clear : Set http = CreateObject("MSXML2.XMLHTTP")
    On Error GoTo 0
    
    If http Is Nothing Then
        sectorOutput = "Errore creazione oggetto HTTP"
        currencyOutput = "Errore creazione oggetto HTTP"
        Exit Function ' Esce, la funzione restituirà False.
    End If

    ' Costruzione dell'URL per la pagina profilo del ticker su Yahoo Finance.
    profileUrl = "https://finance.yahoo.com/quote/" & tickerSymbol & "/profile"

    ' Configurazione e invio della richiesta HTTP GET per la pagina profilo.
    On Error Resume Next
    http.Open "GET", profileUrl, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/123.0.0.0 Safari/537.36"
    http.send
    
    If Err.Number <> 0 Then
        sectorOutput = "Errore invio richiesta profilo: (" & Err.Number & ") " & Err.Description
        currencyOutput = "Errore invio richiesta profilo: (" & Err.Number & ") " & Err.Description
        GoTo CleanupProfileDataFunction ' Procedura di uscita controllata.
    End If
    On Error GoTo 0

    ' Analisi della risposta HTTP.
    If http.Status = 200 Then
        GetYahooProfileData = True ' Segnala successo della richiesta HTTP.
        profileResponseText = http.responseText ' Acquisizione dell'HTML.

        ' Dichiarazioni comuni per il parsing.
        Dim searchPos As Long, linkStartPos As Long, linkEndPos As Long
        Dim textStartPos As Long, textEndPos As Long
        
        ' --- Estrarre Settore ---
        ' Logica di parsing per il Settore: itera sui link <a> e analizza l'attributo href.
        ' Un link di settore ha una struttura URL specifica: /sectors/NOME_SETTORE/
        searchPos = 1 ' Inizia la ricerca dall'inizio del testo HTML.
        ' sectorOutput è già inizializzato a "Non trovato".
        Do
            linkStartPos = InStr(searchPos, profileResponseText, "<a ", vbTextCompare) ' Trova l'inizio di un tag <a>.
            If linkStartPos = 0 Then Exit Do ' Nessun altro tag <a> trovato, esce dal loop.
            
            linkEndPos = InStr(linkStartPos, profileResponseText, "</a>", vbTextCompare) ' Trova la fine del tag <a>.
            If linkEndPos = 0 Then Exit Do ' Struttura HTML anomala, esce.

            Dim linkContent_Sector As String ' Stringa contenente l'intero tag <a>.
            linkContent_Sector = Mid(profileResponseText, linkStartPos, linkEndPos - linkStartPos + Len("</a>"))
            
            Dim hrefPos_Sector As Long, hrefEndPos_Sector As Long, tempHref_Sector As String ' Variabili per l'analisi dell'href.
            hrefPos_Sector = InStr(1, linkContent_Sector, "href=""", vbTextCompare) ' Trova l'inizio dell'attributo href.
            
            If hrefPos_Sector > 0 Then
                hrefPos_Sector = hrefPos_Sector + Len("href=""") ' Avanza all'inizio del valore dell'href.
                hrefEndPos_Sector = InStr(hrefPos_Sector, linkContent_Sector, """", vbTextCompare) ' Trova la virgoletta di chiusura.
                
                If hrefEndPos_Sector > 0 Then
                    tempHref_Sector = Mid(linkContent_Sector, hrefPos_Sector, hrefEndPos_Sector - hrefPos_Sector) ' Estrae il valore dell'href.
                    
                    ' Condizione per identificare un link di settore:
                    ' 1. Inizia con "/sectors/"
                    ' 2. Finisce con "/"
                    ' 3. Ha esattamente 3 segmenti dopo la divisione per "/" (es. "", "sectors", "nome-settore", "") -> UBound = 3
                    If Left(tempHref_Sector, 9) = "/sectors/" And Right(tempHref_Sector, 1) = "/" And UBound(Split(tempHref_Sector, "/")) = 3 Then
                        ' Estrazione del testo del link (il nome del settore).
                        textStartPos = InStr(1, linkContent_Sector, ">", vbTextCompare) ' Trova la fine del tag di apertura <a>.
                        If textStartPos > 0 And textStartPos < Len(linkContent_Sector) Then
                            textStartPos = textStartPos + 1 ' Avanza all'inizio del testo.
                            textEndPos = InStr(textStartPos, linkContent_Sector, "<", vbTextCompare) ' Trova l'inizio del tag di chiusura </a>.
                            If textEndPos > 0 Then
                                sectorOutput = Trim(Mid(linkContent_Sector, textStartPos, textEndPos - textStartPos))
                                Exit Do ' Settore trovato, esce dal loop di ricerca del settore.
                            End If
                        End If
                    End If
                End If
            End If
            searchPos = linkEndPos + Len("</a>") ' Prepara la ricerca per il prossimo tag <a>.
        Loop Until sectorOutput <> "Non trovato" Or linkStartPos = 0
        
        ' --- Estrarre Valuta ---
        ' Logica di parsing per la Valuta: cerca il marcatore testuale "Currency in "
        ' e estrae i 3 caratteri successivi.
        Dim currencyMarker As String, currencyMarkerPos As Long
        currencyMarker = "Currency in " ' Marcatore testuale chiave.
        ' currencyOutput è già inizializzato a "Non trovata".
        
        currencyMarkerPos = InStr(1, profileResponseText, currencyMarker, vbTextCompare) ' Ricerca case-insensitive del marcatore.
        If currencyMarkerPos > 0 Then
            Dim currencyCodeStartPos As Long
            currencyCodeStartPos = currencyMarkerPos + Len(currencyMarker) ' Posizione iniziale del codice valuta.
            ' Verifica che ci siano abbastanza caratteri da estrarre.
            If currencyCodeStartPos + 2 <= Len(profileResponseText) Then
                currencyOutput = Mid(profileResponseText, currencyCodeStartPos, 3) ' Estrae i 3 caratteri.
            Else
                currencyOutput = "Dati valuta insuff." ' Stringa HTML troppo corta dopo il marcatore.
            End If
        Else
            ' currencyOutput rimane "Non trovata" se il marcatore non è presente.
        End If
    Else
        ' La richiesta HTTP per il profilo non ha avuto successo.
        sectorOutput = "Profilo: Errore HTTP " & http.Status
        currencyOutput = "Profilo: Errore HTTP " & http.Status
        ' GetYahooProfileData rimane False (come da inizializzazione o da fallimento HTTP).
    End If

CleanupProfileDataFunction: ' Etichetta per il cleanup.
    Set http = Nothing ' Rilascio dell'oggetto HTTP.
    ' Il valore di ritorno della funzione è già stato impostato a True se http.Status = 200, altrimenti è False.
End Function
