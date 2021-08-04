Option Strict Off

Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Net
Imports System.IO
Imports System.Configuration
Imports System.Text
Imports System
Imports System.Management
Imports System.Timers
Imports System.Threading
Imports System.Threading.ThreadPriority


Module Module2


    Public i_countFindSeed, i_countTimeFindSeeds As Int32

    Public Sub FindSeeds1()

        If Form1.CheckSeed2.Checked = True Then

            If StartMeUp.b_p_showS = True Then
                'Console.WriteLine("SeedCrawler.23 -- DateTime.Now.Day  " & DateTime.Now.Day)
            End If
            'If i_countTimeFindSeeds = 0 Then
            Call FindSeeds()
            'End If

            If i_countTimeFindSeeds = 0 And i_countFindSeed = 0 And s_p_authority = "yes" Then
                'Console.WriteLine("SeedCrawler.26 -- DateTime.Now.Day  " & DateTime.Now.Day)
                Call UpdateSubQueryToDo()
                Call FindSeeds()
            End If

            'To change in production

            If (i_countFindSeed = 1 Or i_countFindSeed > 10) And s_p_authority = "yes" Then
                If StartMeUp.b_p_showS = True Then
                    'Console.WriteLine("SeedCrawler.34 -- DateTime.Now.Day  " & DateTime.Now.Day)
                End If
                Call UpdateSubQueryToDo()
                Call FindSeeds()
                i_countFindSeed = 2
            Else
                i_countFindSeed = i_countFindSeed + 1
            End If

        End If

        'Dim NewSearch As New Search

        'Call StartThread_t_ReceiveDNSFound()

        'Dim t_Seeds As Threading.Thread = New Threading.Thread(AddressOf FindSeeds)

        't_Seeds.Name = "Thread_FindSeeds"
        't_Seeds.Priority = Normal
        't_Seeds.Start()

        'If b_p_closedTheSeedingProcess Then
        '    If t_Seeds.IsAlive Then
        '        t_Seeds.Join(1)
        '    End If
        'Application.exit()
        'Environment.ExitCode
        'End If

    End Sub

    Public Sub FindSeeds()
        Try

            If Form1.CheckSeed2.Checked = True Then

                Dim s_SearchToolName As String
                Dim b_continue As Boolean = True

                Do While b_continue

                    If GetKeywordsDatabaseName() = "stop" Then
                        Exit Do
                    End If

                    If Len(s_p_datName) > 0 Then

                        'Call Google(s_p_keywords, s_p_datName)
                        's_SearchToolName = "Google"

                        'Call Sub_CheckURLInsertBySearchTools(s_SearchToolName, s_p_datName)
                        'b_p_firstTime = True

                        'Call CheckSizeOfLog()

                        Call MSN(s_p_keywords, s_p_datName)
                        s_SearchToolName = "MSN"

                        Call Sub_CheckURLInsertBySearchTools(s_p_keywords, s_p_datName)
                        b_p_firstTime = True

                        '                    Call CheckSizeOfLog()

                        'Call HotbotInktomi(s_p_keywords, s_p_datName)
                        's_SearchToolName = "HotbotInktomi"

                        '                    Call Sub_CheckURLInsertBySearchTools(s_p_keywords, s_p_datName)
                        '                    b_p_firstTime = True

                        '                    Call CheckSizeOfLog()

                        '                    Call HotbotTeoma(s_p_keywords, s_p_datName)
                        '                    s_SearchToolName = "HotbotTeoma"

                        '                    Call Sub_CheckURLInsertBySearchTools(s_p_keywords, s_p_datName)
                        '                    b_p_firstTime = True

                        '                    Call CheckSizeOfLog()

                        '                    Call Altavista(s_p_keywords, s_p_datName)
                        '                    s_SearchToolName = "Altavista"

                        '                    Call Sub_CheckURLInsertBySearchTools(s_p_keywords, s_p_datName)
                        '                    b_p_firstTime = True

                        '                    Call CheckSizeOfLog()

                        '                    Call LooksSmart(s_p_keywords, s_p_datName)
                        '                    s_SearchToolName = "LooksSmart"

                        '                    Call Sub_CheckURLInsertBySearchTools(s_p_keywords, s_p_datName)
                        '                    b_p_firstTime = True

                        '                    Call CheckSizeOfLog()

                        Call Sub_InsertQueryFoundByKeywords(s_SearchToolName, s_p_datName)

                        Call RemoveKeywordsToLookFor(s_p_keywords, s_p_datName)

                        Call RemoveSynonymsToLookFor(s_p_keywords, s_p_datName)

                        Call SendParametersAndCleanup()

                    Else
                        '                    'Console.WriteLine(" seedcrawlers.133 --> There is no database to access? s_p_datName --> " & s_p_datName)
                        '                    System.Threading.Thread.Sleep(10000)
                        '                    Exit Do
                    End If

                Loop

            End If
        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " seedcrawlers.137 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try

    End Sub

    Public s_p_keywords As String
    Public s_p_datName As String
    Public i_p_ranking As Int32

    Function GetKeywordsDatabaseName() As String
        Try
            Dim s_keepResult As String
            Dim s_no As String = "n"
            Dim b_weHaveKeywords As Boolean = False

            Dim mySelectQuery As String = "SELECT top 1 * FROM SUBQUERY where Done ='" & s_no & "' order by Rating desc"

            Dim myConnection1 As New OleDbConnection(ConnStringDNA())

            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer
            While myReader.Read()
                s_p_datName = myReader("Query")
                s_p_keywords = myReader("SubQuery")
                i_p_ranking = myReader("Rating")
                b_weHaveKeywords = True
                Exit While
            End While

            If b_weHaveKeywords Then
                s_keepResult = "go"
            Else
                s_keepResult = "stop"
            End If

            myReader.Close()
            myConnection1.Close()

            If s_keepResult = "stop" Then
                If RetrieveFromSynonyms() = "stop" Then
                    GetKeywordsDatabaseName = "stop"
                Else
                    GetKeywordsDatabaseName = "go"
                End If
            End If

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Seedcrawlers.191 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try

    End Function

    Function RetrieveFromSynonyms()

        Try

            Dim s_no As String = "n"
            Dim s_checked As String = "checked"
            Dim b_weHaveKeywords As Boolean = False

            Dim mySelectQuery As String = "SELECT top 1 * FROM SYNONYMS where Done= '" & s_no & "' and checked = '" & s_checked & "' order by Rating desc"

            Dim myConnection1 As New OleDbConnection(ConnStringDNA())

            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer
            While myReader.Read()
                s_p_datName = myReader("Query")
                s_p_keywords = myReader("SubWithSynonyms")
                i_p_ranking = myReader("Rating")
                b_weHaveKeywords = True
                Exit While
            End While

            If b_weHaveKeywords Then
                RetrieveFromSynonyms = "go"
            Else
                RetrieveFromSynonyms = "stop"
            End If

            myReader.Close()
            myConnection1.Close()

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Seedcrawlers.235 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try
    End Function

    Sub RemoveKeywordsToLookFor(ByVal s_p_keywords, ByVal s_p_datName)

        Try
            Dim s_yes As String = "y"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myConnectionString As String = "UPDATE SUBQUERY SET Done='" & (s_yes) & "' where Query = '" & (s_p_datName) & "' and SubQuery ='" & (s_p_keywords) & "' "

            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Seedcrawlers.258 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try

    End Sub


    Sub RemoveSynonymsToLookFor(ByVal s_p_keywords, ByVal s_p_datName)
        Try
            Dim s_yes As String = "y"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myConnectionString As String = "UPDATE SYNONYMS SET Done='" & (s_yes) & "' where Query = '" & (s_p_datName) & "' and SubWithSynonyms ='" & (s_p_keywords) & "' "

            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Seedcrawlers.282 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try
    End Sub


    '    Function LookAlltheWeb(ByVal s_keywords As String, ByVal strURI As String) As String
    '        'Dim Search As New Search()
    '        Dim sSendstrURI As String = strURI
    '        Dim SGetText, s_urlString As String
    '        Dim sGetHttp As String
    '        Dim sYahoo As String = "http://alltheweb.com/search?cat=web&cs=utf-8&l=any&q="
    '        Dim sFindNext20 As String = ">Next 20"
    '        Dim iEndUrlNext20 As Integer
    '        Dim iBeginNext20 As Integer
    '        Dim bContinue As Boolean = True
    '        Dim sFindQuote As String
    '        Dim iCounter As Integer = 10
    '        Dim sFindNext20Url As String
    '        Dim iLenghtNextUrl As Integer
    '        Dim sTakeCareOfString As String
    '        Dim iFirstTimeNext20 As Integer = 0
    '        Dim sOldsFindNext20Url As String
    '        Dim sKeepstrURI As String
    '        Dim i_resetPBar As Integer = 0
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        b_p_startANew = True

    '        Try
    '            sWhatRange()
    '            CheckParametersInDocument(strURI)

    '            s_keywords = Replace(s_keywords, " ", "+")

    '            s_keywords = ReplaceAccent(s_keywords)

    '            s_urlString = "http://alltheweb.com/search?cat=web&cs=utf-8&l=any&q=" & s_keywords

    '            Do While bContinue

    '                If b_p_closedTheSeedingProcess Then
    '                    Exit Do
    '                End If

    '                If Len(sFindNext20Url) > 10 Then
    '                    If sFindNext20Url = sOldsFindNext20Url Then
    '                        Exit Do
    '                    End If
    '                    s_urlString = sFindNext20Url
    '                    sOldsFindNext20Url = sFindNext20Url
    '                End If

    '                's_urlString = "http://www.golfbizzz.com"

    '                Dim objURI As Uri = New Uri(s_urlString)

    '                Dim objWebRequest As WebRequest = WebRequest.Create(objURI)

    '                Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
    '                objWebResponse.Headers.Add("Accept-Language: fr, en")
    '                Dim objStream As Stream = objWebResponse.GetResponseStream()
    '                Dim objStreamReader As StreamReader = New StreamReader(objStream)
    '                Dim strHTML As String = objStreamReader.ReadToEnd

    '                SGetText = strHTML

    '                If Len(SGetText) > 4 Then
    '                    Call fGetUrlAddDNS(SGetText, sSendstrURI)

    '                    If i_resetPBar = 10 Then
    '                        i_resetPBar = 0
    '                    End If
    '                    i_resetPBar = i_resetPBar + 1
    '                End If

    '                If Len(SGetText) > 10 And InStr(SGetText, "No Web pages found", CompareMethod.Text) = 0 Then

    '                    sGetHttp = "http://alltheweb.com/"

    '                    If bExactPhrase Then
    '                        sFindNext20Url = "http://alltheweb.com/search?q=" & s_keywords & "&c=web&o=" & iCounter & "&l=any&cn=3&cs=utf-8&phrase=on"
    '                    Else
    '                        sFindNext20Url = "http://alltheweb.com/search?q=" & s_keywords & "&c=web&o=" & iCounter & "&l=any&cn=3&cs=utf-8"
    '                    End If

    '                    iCounter = iCounter + 10

    '                Else
    '                    'EventLog.WriteEntry(sSource, " module2.100 --> AllTheWeb.com; No Web pages found that match your query" & iCounter & " " & sFindNext20Url, EventLogEntryType.Information, 99)
    '                    Exit Do
    '                End If

    '                'EventLog.WriteEntry(sSource, " module2.100 --> AllTheWeb.com page view!" & iCounter & " " & sFindNext20Url, EventLogEntryType.Information, 99)

    '                If iCounter > 1000 Then
    '                    Exit Do
    '                End If

    '            Loop

    '        Catch ex As Exception

    '            EventLog.WriteEntry(sSource, " Seedcrawlers.386 --> " & ex.Message, EventLogEntryType.Information, 50)

    '        End Try

    '    End Function


    '    Function fGetUrlAddDNS(ByVal FileToCrawl As String, ByVal sSendstrURI As String) As String

    '        Dim iFirst As Integer
    '        Dim iLast As Integer
    '        Dim iLenght As Integer
    '        Dim sStoreResult As String
    '        Dim sAllResult As String
    '        Dim bKeepGoing As Boolean
    '        Dim sURL As String
    '        Dim iURLIn As Integer
    '        Dim sCheckStop As String
    '        Dim iURLFast As String
    '        Dim i_GoodSpot As Integer = i_p_ranking
    '        Dim i_array() As String
    '        Dim b_continue As Boolean = True
    '        Dim i_len1, i_len2 As Int32
    '        Dim s_searchEngine As String = "LookAllTheWeb"
    '        Dim s_qualityOfSE As String = "C"

    '        Try

    '            i_GoodSpot = f_iGoodSpot(s_qualityOfSE)

    '            bKeepGoing = True

    '            sURL = "http"

    '            iFirst = InStr(1, FileToCrawl, "http://w")

    '            iLast = InStr(iFirst, FileToCrawl, ";")

    '            iLenght = iLast - iFirst

    '            If iLenght > 0 Then
    '                Do While bKeepGoing
    '                    sStoreResult = Mid(FileToCrawl, iFirst, iLenght)

    '                    sStoreResult = Replace(sStoreResult, """", "")

    '                    iURLIn = InStr(1, sStoreResult, sURL)
    '                    iURLFast = InStr(1, sStoreResult, "fastsearch")

    '                    If iURLIn > 0 And (Not (iURLFast > 0)) Then

    '                        sStoreResult = Replace(sStoreResult, "&nbsp;", "")

    '                        Do While b_continue
    '                            i_len1 = Len(sStoreResult)
    '                            sStoreResult = Replace(sStoreResult, "&nbsp", "")
    '                            i_len2 = Len(sStoreResult)
    '                            If i_len1 = i_len2 Then
    '                                Exit Do
    '                            End If
    '                        Loop

    '                        i_array = Split(sStoreResult)

    '                        sStoreResult = i_array(0)

    '                        sStoreResult = Trim(RemoveHtmlTag(sStoreResult))

    '                        sStoreResult = Trim(RemoveHTMLFromDNS(sStoreResult))

    '                        sAllResult = sAllResult & Chr(10) & sStoreResult

    '                        sAllResult = Replace(sAllResult, "&nbsp", "")

    '                        sCheckStop = CheckUrlIn(sStoreResult, sSendstrURI)

    '                        If sCheckStop <> "stop" Then

    '                            'Call Search.InsertUrl(sCheckStop, sSendstrURI)
    '                            Call InsertDNSFound(sCheckStop, sSendstrURI, i_GoodSpot + 100, , , s_searchEngine)
    '                            Call f_Extract_Racine_URL(sCheckStop, sSendstrURI, i_GoodSpot, s_searchEngine)
    '                            Dim b_seedCrawlers As Boolean = True
    '                            Call SendDNSFound(sSendstrURI, b_seedCrawlers)
    '                        End If

    '                    Else
    '                        sStoreResult = "http://alltheweb.com"
    '                    End If

    '                    fGetUrlAddDNS = sAllResult

    '                    iFirst = InStr(iFirst + 5, FileToCrawl, "http://w")

    '                    If iFirst <= Len(FileToCrawl) Then
    '                        If iFirst > 0 Then
    '                            iLast = InStr(iFirst, FileToCrawl, ";")
    '                        End If
    '                    End If

    '                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
    '                        Exit Do
    '                    Else
    '                        iLenght = iLast - iFirst
    '                    End If
    '                Loop
    '            End If

    '        Catch ex As Exception

    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " Seedcrawlers.497 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try

    '    End Function


    Function Google(ByVal s_keywords As String, ByVal strURI As String) As String


        Dim sSendstrURI As String = strURI
        Dim SGetText, s_urlToSend As String
        Dim sGetHttp As String
        Dim s_Google As String = "http://www.google.com/search?q=bush+onu&hl=en&lr=&ie=UTF-8&oe=UTF-8&start=0&sa=N&filter=0"
        Dim sFindNext20 As String = ">Next 20"
        Dim iEndUrlNext20 As Integer
        Dim iBeginNext20 As Integer
        Dim bContinue As Boolean = True
        Dim sFindQuote As String
        Dim iCounter As Integer = 10
        Dim sFindNext20Url As String
        Dim iLenghtNextUrl As Integer
        Dim sTakeCareOfString As String
        Dim iFirstTimeNext20 As Integer = 0
        Dim sOldsFindNext20Url As String
        Dim sKeepstrURI, s_ChangeInteger As String
        Dim i_resetPBar As Integer = 0
        Dim b_Changed As Boolean = False
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim i_GooglePage As Integer = 1
        Dim i_addTimeGoogle As Integer = 1000
        Dim bContinueTimeGoogle As Boolean = True
        Dim i_intLoopGoogle As Integer = 1
        Dim s_timeDash = "-"
        Dim b_boolJustStarted = True


        b_p_startANew = True

        Try

            s_keywords = Replace(s_keywords, " ", "+")

            s_keywords = ReplaceAccent(s_keywords)

            s_Google = Replace(s_Google, "bush+onu", s_keywords)

            s_urlToSend = ""
            s_urlToSend = s_Google

            Do While bContinue

                Do While (bContinueTimeGoogle And b_boolJustStarted = False)

                    'Console.WriteLine(s_timeDash)
                    s_timeDash = s_timeDash + "-"
                    i_addTimeGoogle = i_addTimeGoogle
                    System.Threading.Thread.Sleep(i_addTimeGoogle)
                    i_intLoopGoogle = i_intLoopGoogle + 1
                    If i_intLoopGoogle > 20 Then
                        i_intLoopGoogle = 0
                        s_timeDash = "-"
                        Exit Do
                    End If
                Loop

                b_boolJustStarted = False

                If b_p_closedTheSeedingProcess Then
                    Exit Do
                End If

                If Len(sFindNext20Url) > 10 Then

                    i_GooglePage = i_GooglePage + 1

                    If sFindNext20Url = sOldsFindNext20Url Then
                        Exit Do
                    End If

                    s_urlToSend = sFindNext20Url
                    sOldsFindNext20Url = sFindNext20Url

                End If

                Dim objURI As Uri = New Uri(s_urlToSend)
                Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
                Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
                objWebResponse.Headers.Add("Accept-Language: fr, en")
                Dim objStream As Stream = objWebResponse.GetResponseStream()
                Dim objStreamReader As StreamReader = New StreamReader(objStream)
                Dim strHTML As String = objStreamReader.ReadToEnd

                SGetText = strHTML

                If Len(SGetText) > 4 Then
                    Call fGetUrlAddDNSGoogle(SGetText, sSendstrURI, i_GooglePage)
                End If

                If Len(SGetText) > 10 Then
                    sGetHttp = "http://google.com/"
                    If b_Changed = False Then
                        b_Changed = True
                        s_Google = Replace(s_Google, "0", "****")
                    End If
                    sFindNext20Url = Replace(s_Google, "****", iCounter)
                    iCounter = iCounter + 10
                End If

                'EventLog.WriteEntry(sSource, " module2.304 --> Google.com page view!" & iCounter & " - " & sFindNext20Url, EventLogEntryType.Information, 99)

                If iCounter > 1000 Then
                    'EventLog.WriteEntry(sSource, " Google Done page view!--> " & iCounter, EventLogEntryType.Information, 99)
                    Exit Do
                End If

                'System.Threading.Thread.Sleep(100)

            Loop

        Catch ex As Exception

            EventLog.WriteEntry(sSource, " Seedcrawlers.591 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try
    End Function


    Function fGetUrlAddDNSGoogle(ByVal FileToCrawl As String, ByVal sSendstrURI As String, Optional i_GooglePage As String = "") As String

        Dim iFirst As Integer
        Dim iLast As Integer
        Dim iLenght As Integer
        Dim sStoreResult As String
        Dim sAllResult As String
        Dim bKeepGoing As Boolean
        Dim sURL As String
        Dim iURLIn As Integer
        Dim sCheckStop As String
        Dim iURLFast As String
        Dim i_GoodSpot As Integer = i_p_ranking
        Dim i_len1, i_len2 As Int32
        Dim s_array() As String
        Dim b_continue As Boolean = True
        Dim s_searchEngine As String = "Google"
        Dim s_qualityOfSE As String = "A"

        Try

            i_GoodSpot = f_iGoodSpot(s_qualityOfSE)

            sWhatRange()
            CheckParametersInDocument(sSendstrURI)

            bKeepGoing = True

            sURL = "http"

            iFirst = InStr(1, FileToCrawl, "<cite>h")

            iLast = InStr(iFirst, FileToCrawl, "</cite>")

            iLenght = iLast - iFirst

            If iLenght > 0 Then
                Do While bKeepGoing
                    sStoreResult = Mid(FileToCrawl, iFirst + 6, iLenght - 6)
                    'sStoreResult = Replace(sStoreResult, """", "")

                    iURLIn = InStr(1, sStoreResult, sURL)
                    iURLFast = InStr(1, sStoreResult, "fastsearch")

                    If iURLIn > 0 And (Not (iURLFast > 0)) Then

                        Do While b_continue
                            i_len1 = Len(sStoreResult)
                            'sStoreResult = Replace(sStoreResult, "&nbsp", "")
                            i_len2 = Len(sStoreResult)
                            If i_len1 = i_len2 Then
                                Exit Do
                            End If
                        Loop

                        s_array = Split(sStoreResult)

                        sStoreResult = s_array(0)

                        sStoreResult = Trim(RemoveHtmlTag(sStoreResult))

                        'sStoreResult = Replace(sStoreResult, "&nbsp;", "")
                        'sStoreResult = Replace(sStoreResult, "&nbsp", "")

                        If InStr(1, sStoreResult, "prev=/search", CompareMethod.Text) > 0 Then
                            Dim i_countWhereToBegin, i_lenToKeep, i_lenStoreResult As Int32
                            i_countWhereToBegin = InStr(1, sStoreResult, "prev=/search", CompareMethod.Text)
                            i_lenStoreResult = Len(sStoreResult)
                            i_lenToKeep = i_lenStoreResult - ((i_lenStoreResult - i_countWhereToBegin) + 2)
                            sStoreResult = Mid(sStoreResult, 1, i_lenToKeep)
                        End If
                        If InStr(1, sStoreResult, ".asp%", CompareMethod.Text) > 0 Then
                            Dim i_countWhereToBegin, i_lenToKeep, i_lenStoreResult As Int32
                            i_countWhereToBegin = InStr(1, sStoreResult, ".asp%", CompareMethod.Text)
                            i_lenStoreResult = Len(sStoreResult)
                            i_lenToKeep = i_lenStoreResult - ((i_lenStoreResult - i_countWhereToBegin) - 3)
                            sStoreResult = Mid(sStoreResult, 1, i_lenToKeep)
                        End If
                        If InStr(1, sStoreResult, ".htm%", CompareMethod.Text) > 0 Then
                            Dim i_countWhereToBegin, i_lenToKeep, i_lenStoreResult As Int32
                            i_countWhereToBegin = InStr(1, sStoreResult, ".htm%", CompareMethod.Text)
                            i_lenStoreResult = Len(sStoreResult)
                            i_lenToKeep = i_lenStoreResult - ((i_lenStoreResult - i_countWhereToBegin) - 3)
                            sStoreResult = Mid(sStoreResult, 1, i_lenToKeep)
                        End If
                        If InStr(1, sStoreResult, ".cfm%", CompareMethod.Text) > 0 Then
                            Dim i_countWhereToBegin, i_lenToKeep, i_lenStoreResult As Int32
                            i_countWhereToBegin = InStr(1, sStoreResult, ".cfm%", CompareMethod.Text)
                            i_lenStoreResult = Len(sStoreResult)
                            i_lenToKeep = i_lenStoreResult - ((i_lenStoreResult - i_countWhereToBegin) - 3)
                            sStoreResult = Mid(sStoreResult, 1, i_lenToKeep)
                        End If
                        If InStr(1, sStoreResult, ".html%", CompareMethod.Text) > 0 Then
                            Dim i_countWhereToBegin, i_lenToKeep, i_lenStoreResult As Int32
                            i_countWhereToBegin = InStr(1, sStoreResult, ".html%", CompareMethod.Text)
                            i_lenStoreResult = Len(sStoreResult)
                            i_lenToKeep = i_lenStoreResult - ((i_lenStoreResult - i_countWhereToBegin) - 4)
                            sStoreResult = Mid(sStoreResult, 1, i_lenToKeep)
                        End If
                        If InStr(1, sStoreResult, ".php%", CompareMethod.Text) > 0 Then
                            Dim i_countWhereToBegin, i_lenToKeep, i_lenStoreResult As Int32
                            i_countWhereToBegin = InStr(1, sStoreResult, ".php%", CompareMethod.Text)
                            i_lenStoreResult = Len(sStoreResult)
                            i_lenToKeep = i_lenStoreResult - ((i_lenStoreResult - i_countWhereToBegin) - 3)
                            sStoreResult = Mid(sStoreResult, 1, i_lenToKeep)
                        End If

                        sAllResult = sAllResult & Chr(10) & sStoreResult

                        'sAllResult = Replace(sAllResult, "&nbsp;", "")
                        'sAllResult = Replace(sAllResult, "&nbsp", "")

                        'sCheckStop = CheckUrlIn(sStoreResult, sSendstrURI)
                        'sCheckStop = CheckUrlIn_DNS(sStoreResult, sSendstrURI)

                        If sCheckStop <> "stop" Then
                            If Not (InStr(1, sCheckStop, "q=cache") > 0) Then
                                'Call Search.InsertUrl(sCheckStop, sSendstrURI)
                                If InsertDNSFound(sStoreResult, sSendstrURI, i_GoodSpot + 100, , , s_searchEngine, i_GooglePage) = True Then
                                    'System.Threading.Thread.Sleep(2000)
                                    Call f_Extract_Racine_URL(sStoreResult, sSendstrURI, i_GoodSpot, s_searchEngine)
                                    Dim b_seedCrawlers As Boolean = True
                                    'Call SendDNSFound(sSendstrURI, b_seedCrawlers)
                                End If
                            End If
                        End If

                    Else
                        sStoreResult = "http://alltheweb.com"
                    End If

                    fGetUrlAddDNSGoogle = sAllResult

                    iFirst = InStr(iFirst + 10, FileToCrawl, "<cite>h")

                    If iFirst <= Len(FileToCrawl) Then
                        If iFirst > 0 Then
                            iLast = InStr(iFirst, FileToCrawl, "</cite>")
                        End If
                    End If

                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
                        Exit Do
                    Else
                        iLenght = iLast - iFirst
                    End If
                Loop
            End If

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Seedcrawlers.748 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try
    End Function


    Function MSN(ByVal s_p_keywords As String, ByVal s_p_datName As String) As String

        Dim i_msnPageView As Int32
        Dim sSendstrURI As String = s_p_datName
        Dim SGetText As String
        Dim sGetHttp As String
        Dim s_MSN As String = "https://www.bing.com/search?q=bush+onu&qs=n&form=QBLH&sp=-1&ghc=1&pq=bush+onu&sc=4-8&sk=&cvid=0FBC985A9FC748D19B7D05780EE6D6D8"
        Dim sFindNext15 As String = ">Next 10"
        Dim iEndUrlNext15 As Integer
        Dim iBeginNext15 As Integer
        Dim bContinue As Boolean = True
        Dim sFindQuote As String
        Dim iCounter As Int32 = 18
        Dim sFindNext15Url As String
        Dim iLenghtNextUrl As Integer
        Dim sTakeCareOfString As String
        Dim iFirstTimeNext15 As Integer = 0
        Dim sOldsFindNext15Url As String
        Dim sKeepstrURI, s_ChangeInteger, s_urlToSend As String
        Dim i_resetPBar As Integer = 0
        Dim b_Changed As Boolean = False
        Dim b_fistPass As Boolean = True
        Dim i_pageMSN As Integer = 1
        Dim s_timeDash As String = "-"
        Dim i_addTimeGoogle As Integer = 1000
        Dim i_intLoopGoogle As Integer = 0
        Dim b_boolJustStarted = True
        Dim bContinueTimeGoogle As Boolean = True

        b_p_startANew = True

        Try

            Do While (bContinueTimeGoogle And b_boolJustStarted = False)
                'Console.WriteLine(s_timeDash)
                s_timeDash = s_timeDash + "-"
                i_addTimeGoogle = i_addTimeGoogle
                System.Threading.Thread.Sleep(i_addTimeGoogle)
                i_intLoopGoogle = i_intLoopGoogle + 1
                If i_intLoopGoogle > 20 Then
                    i_intLoopGoogle = 0
                    s_timeDash = "-"
                    Exit Do
                End If
            Loop

            b_boolJustStarted = False

            s_p_keywords = Replace(s_p_keywords, " ", "+")

            s_p_keywords = ReplaceAccent(s_p_keywords)

            s_MSN = Replace(s_MSN, "bush+onu", s_p_keywords)

            s_urlToSend = ""
            s_urlToSend = s_MSN

            Dim i_keepPlaceNextGif As Int32
            Dim i_keepDebutNextURL, i_keepEndNextURL, i_lenURL As Int32
            Dim s_nextURL, s_nextSearchNumber As String
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            Dim i_countMSN As Int32 = 11

            Do While bContinue

                If b_p_closedTheSeedingProcess Then
                    Exit Do
                End If

                Dim objURI As Uri = New Uri(s_urlToSend)
                Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
                Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
                objWebResponse.Headers.Add("Accept-Language: fr, en")
                Dim objStream As Stream = objWebResponse.GetResponseStream()
                Dim objStreamReader As StreamReader = New StreamReader(objStream)
                Dim strHTML As String = objStreamReader.ReadToEnd

                SGetText = strHTML

                'i_keepPlaceNextGif = InStr(SGetText, "NEXT&nbsp;&gt;&gt;", CompareMethod.Text)
                i_msnPageView = i_msnPageView + 1

                'If i_keepPlaceNextGif = 0 Then
                'Exit Do
                'End If

                If Len(SGetText) > 4 Then

                    Call fGetUrlAddDNSMSN(SGetText, sSendstrURI, i_msnPageView)

                End If

                If b_fistPass = True Then

                    b_fistPass = False
                    s_nextURL = "https://www.bing.com/search?q=bush+onu&qs=n&sp=-1&pq=bush+onu&sc=5-8&sk=&cvid=820201E796174547B072A36E550A139F&first=11&FORM=PERE"

                Else

                    i_countMSN = i_countMSN + 10
                    s_nextURL = "https://www.bing.com/search?q=bush+onu&qs=n&sp=-1&pq=bush+onu&sc=5-8&sk=&cvid=820201E796174547B072A36E550A139F&first=11&FORM=PERE"

                    s_nextSearchNumber = "first=" + CStr(i_countMSN)
                    s_nextURL = Replace(s_nextURL, "first=11", s_nextSearchNumber)
                    s_urlToSend = s_nextURL

                End If

                s_nextURL = Replace(s_nextURL, "bush+onu", s_p_keywords)

                s_urlToSend = s_nextURL

                'i_msnPageView = i_msnPageView + 1

                System.Threading.Thread.Sleep(5000)

                If i_msnPageView > 40 Then
                    Exit Do
                End If

            Loop

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Seedcrawlers.848 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try
    End Function


    Function fGetUrlAddDNSMSN(ByVal FileToCrawl As String, ByVal sSendstrURI As String, Optional i_msnPageView As Integer = 1) As String
        Dim b_continue As Boolean = True
        Dim iFirst As Integer
        Dim iLast As Integer
        Dim iLenght As Integer
        Dim sStoreResult As String
        Dim sAllResult As String
        Dim bKeepGoing As Boolean
        Dim sURL As String
        Dim iURLIn As Integer
        Dim sCheckStop As String
        Dim iURLFast As String
        Dim s_array() As String
        Dim i_GoodSpot As Integer = i_p_ranking
        Dim i_len1, i_len2 As Int32
        Dim s_array1() As String
        Dim b_continue1 As Boolean = True
        Dim s_qualityOfSE As String = "B"
        Dim s_searchEngine As String = "Bing"

        Try
            i_GoodSpot = f_iGoodSpot(s_qualityOfSE)

            sWhatRange()
            CheckParametersInDocument(sSendstrURI)

            bKeepGoing = True

            sURL = "http"

            iFirst = InStr(1, FileToCrawl, "<a href=" & Chr(34) & "http")

            iLast = InStr(iFirst, FileToCrawl, "h=")

            iLenght = iLast - iFirst

            If iLenght > 0 Then
                Do While bKeepGoing
                    sStoreResult = Mid(FileToCrawl, iFirst, iLenght)
                    sStoreResult = Replace(sStoreResult, """", "")

                    iURLIn = InStr(1, sStoreResult, sURL)
                    iURLFast = InStr(1, sStoreResult, "fastsearch")

                    If iURLIn > 0 And (Not (iURLFast > 0)) Then

                        'sStoreResult = Replace(sStoreResult, "&nbsp;", "")
                        'sStoreResult = Replace(sStoreResult, "&nbsp", "")
                        'sStoreResult = Replace(sStoreResult, ")", "")
                        's_array = Split(sStoreResult)

                        'sStoreResult = s_array(0)

                        sStoreResult = Replace(sStoreResult, "<a href=", "")
                        '   s_urlToCount = Replace(s_urlToCount, "http://www.", "")
                        '   s_urlToCount = Replace(s_urlToCount, "https://", "")
                        '   s_urlToCount = Replace(s_urlToCount, "http://", "")

                        If InStr(sStoreResult, "<a href=", CompareMethod.Text) > 0 Then
                            'sStoreResult = Mid(sStoreResult, 1, Len(sStoreResult) - InStr(sStoreResult, "<", CompareMethod.Text))
                            'sStoreResult = Trim(sStoreResult)
                        End If

                        'Do While b_continue
                        'If InStr(sStoreResult, " ", CompareMethod.Text) > 0 Then
                        'sStoreResult = Replace(sStoreResult, " ", "")
                        'Else
                        'b_continue = False
                        'End If
                        'Loop

                        sAllResult = sAllResult & Chr(10) & sStoreResult

                        'sAllResult = Replace(sAllResult, "&nbsp;", "")
                        'sAllResult = Trim((Replace(sAllResult, "&nbsp", "")))

                        'Do While b_continue1
                        'i_len1 = Len(sStoreResult)
                        'sStoreResult = Replace(sStoreResult, "&nbsp", "")
                        'i_len2 = Len(sStoreResult)
                        'If i_len1 = i_len2 Then
                        'Exit Do
                        'End If
                        'Loop

                        s_array1 = Split(sStoreResult)

                        sStoreResult = s_array1(0)

                        If InStr(1, sStoreResult, ">", CompareMethod.Text) > 0 Then
                            'Dim i_countWhereToBegin, i_lenToKeep, i_lenStoreResult As Int32
                            'i_countWhereToBegin = InStr(1, sStoreResult, ">", CompareMethod.Text)
                            'i_lenStoreResult = Len(sStoreResult)
                            'i_lenToKeep = i_lenStoreResult - ((i_lenStoreResult - i_countWhereToBegin) - 1)
                            'sStoreResult = Mid(sStoreResult, 1, i_lenToKeep)
                        End If

                        sStoreResult = Trim(RemoveHtmlTag(sStoreResult))
                        'sCheckStop = CheckUrlIn(sStoreResult, sSendstrURI)

                        If Len(sStoreResult) > 250 Then
                            sStoreResult = Mid(sStoreResult, 1, 250)
                        End If

                        'If sCheckStop <> "stop" Then
                        If Len(sStoreResult) >= 11 Then
                            'Call Search.InsertUrl(sCheckStop, sSendstrURI)
                            If InsertDNSFound(sStoreResult, sSendstrURI, i_GoodSpot + 100, , , s_searchEngine, i_msnPageView) = True Then
                                Call f_Extract_Racine_URL(sStoreResult, sSendstrURI, i_GoodSpot, s_searchEngine)
                                Dim b_seedCrawlers As Boolean = True
                                'Call SendDNSFound(sSendstrURI, b_seedCrawlers)
                            End If
                        End If
                        'End If

                    Else
                        sStoreResult = "http://search.msn.com"
                    End If

                    fGetUrlAddDNSMSN = sAllResult

                    'iFirst = InStr(1, FileToCrawl, "<a href=" & Chr(34) & "http")

                    'iLast = InStr(iFirst, FileToCrawl, "h=")

                    iFirst = InStr(iFirst + 10, FileToCrawl, "<a href=" & Chr(34) & "http")

                    If iFirst <= Len(FileToCrawl) Then
                        If iFirst > 0 Then
                            iLast = InStr(iFirst, FileToCrawl, "h=")
                        End If
                    End If

                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
                        Exit Do
                    Else
                        iLenght = iLast - iFirst
                    End If
                Loop
            End If

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Seedcrawlers.989 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try

    End Function


    '    Function HotbotInktomi(ByVal s_p_keywords As String, ByVal s_p_datName As String) As String

    '        Dim sSendstrURI As String = s_p_datName
    '        Dim SGetText, s_urlToSend As String
    '        Dim sGetHttp As String
    '        Dim s_Google As String = "http://hotbot.com/default.asp?prov=Inktomi&query=bush+hussein+war&ps=&loc=searchbox&tab=web"
    '        Dim sFindNext20 As String = ">Next 20"
    '        Dim iEndUrlNext20 As Integer
    '        Dim iBeginNext20 As Integer
    '        Dim bContinue As Boolean = True
    '        Dim sFindQuote As String
    '        Dim iCounter As Integer = 11
    '        Dim sFindNext20Url As String
    '        Dim iLenghtNextUrl As Integer
    '        Dim sTakeCareOfString, s_oldGetText As String
    '        Dim iFirstTimeNext20 As Integer = 0
    '        Dim sOldsFindNext20Url, s_keepString As String
    '        Dim sKeepstrURI, s_ChangeInteger, i_placeOfNext, i_placeOfDoubleQuote, i_lenBetween,
    '        i_keepPlaceFound, i_keepPlacePrevious As String
    '        Dim i_resetPBar As Integer = 0
    '        Dim b_Changed As Boolean = False
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        b_p_startANew = True

    '        s_p_keywords = Replace(s_p_keywords, " ", "+")

    '        s_p_keywords = ReplaceAccent(s_p_keywords)

    '        s_Google = Replace(s_Google, "bush+hussein+war", s_p_keywords)

    '        s_urlToSend = ""

    '        s_urlToSend = s_Google

    '        Try

    '            Do While bContinue

    '                If b_p_closedTheSeedingProcess Then
    '                    Exit Do
    '                End If

    '                Dim objURI As Uri = New Uri(s_urlToSend)
    '                Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
    '                Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
    '                objWebResponse.Headers.Add("Accept-Language: fr, en")
    '                Dim objStream As Stream = objWebResponse.GetResponseStream()
    '                Dim objStreamReader As StreamReader = New StreamReader(objStream)
    '                Dim strHTML As String = objStreamReader.ReadToEnd

    '                SGetText = strHTML

    '                If Len(SGetText) > 4 Then

    '                    If iCounter >= 1000 Then
    '                        Exit Do
    '                    End If

    '                    Call fGetUrlAddDNSHotbotInktomi(SGetText, sSendstrURI)

    '                End If

    '                If Len(SGetText) > 10 Or InStr(SGetText, "Sorry, your search had no web results", CompareMethod.Text) Then

    '                    s_Google = "http://hotbot.com/default.asp?query=bush+hussein+war&first=11&page=more&ca=w&prov=Inktomi&pc=&pcx=&date="

    '                    s_Google = Replace(s_Google, "bush+hussein+war", s_p_keywords)

    '                    s_Google = Replace(s_Google, " ", "+")

    '                    iCounter = iCounter + 10

    '                    s_Google = Replace(s_Google, "11", iCounter)

    '                    s_urlToSend = s_Google

    '                    'EventLog.WriteEntry(sSource, " module2.713 HotbotInktomi next page view is --> " & iCounter & " " & s_urlToSend, EventLogEntryType.Information, 99)

    '                Else
    '                    Exit Do
    '                End If

    '                If iCounter > 1000 Then
    '                    Exit Do
    '                End If

    '            Loop

    '        Catch ex As Exception

    '            EventLog.WriteEntry(sSource, " Seedcrawlers.1088 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try
    '    End Function


    '    Function fGetUrlAddDNSHotbotInktomi(ByVal FileToCrawl As String, ByVal sSendstrURI As String) As String

    '        Dim iFirst As Integer
    '        Dim iLast, i_keepbigger As Integer
    '        Dim iLenght As Integer
    '        Dim sStoreResult As String
    '        Dim sAllResult As String
    '        Dim bKeepGoing As Boolean
    '        Dim sURL As String
    '        Dim iURLIn As Integer
    '        Dim sCheckStop As String
    '        Dim iURLFast As String
    '        Dim i_GoodSpot As Integer = i_p_ranking
    '        Dim i_len1, i_len2 As Int32
    '        Dim s_array() As String
    '        Dim b_continue As Boolean = True
    '        Dim s_qualityOfSE As String = "C"
    '        Dim s_searchEngine As String = "HotbotInktomi"

    '        Try
    '            i_GoodSpot = f_iGoodSpot(s_qualityOfSE)
    '            sWhatRange()
    '            CheckParametersInDocument(sSendstrURI)

    '            bKeepGoing = True

    '            sURL = "http"

    '            iFirst = InStr(1, FileToCrawl, "http://www.")

    '            iLast = InStr(iFirst, FileToCrawl, "<")

    '            iLenght = iLast - iFirst

    '            If iLenght > 0 Then
    '                Do While bKeepGoing
    '                    sStoreResult = Mid(FileToCrawl, iFirst, iLenght)

    '                    sStoreResult = Replace(sStoreResult, """", "")

    '                    If InStr(1, sStoreResult, ">", CompareMethod.Text) > 0 Then
    '                        Dim i_countWhereToBegin, i_lenToKeep, i_lenStoreResult As Int32
    '                        i_countWhereToBegin = InStr(1, sStoreResult, ">", CompareMethod.Text)
    '                        i_lenStoreResult = Len(sStoreResult)
    '                        i_lenToKeep = i_lenStoreResult - ((i_lenStoreResult - i_countWhereToBegin) + 1)
    '                        sStoreResult = Mid(sStoreResult, 1, i_lenToKeep)
    '                    End If

    '                    iURLIn = InStr(1, sStoreResult, sURL)

    '                    If iURLIn > 0 Then

    '                        'sStoreResult = Replace(sStoreResult, "&nbsp;", "")
    '                        'sStoreResult = Replace(sStoreResult, "&nbsp", "")

    '                        'Do While b_continue
    '                        'i_len1 = Len(sStoreResult)
    '                        'sStoreResult = Replace(sStoreResult, "&nbsp", "")
    '                        'i_len2 = Len(sStoreResult)
    '                        'If i_len1 = i_len2 Then
    '                        'Exit Do
    '                        'End If
    '                        'Loop

    '                        s_array = Split(sStoreResult)

    '                        sStoreResult = s_array(0)

    '                        sStoreResult = Trim(RemoveHtmlTag(sStoreResult))
    '                        'sAllResult = sAllResult & Chr(10) & sStoreResult

    '                        'i_keepbigger = InStr(sStoreResult, ">", CompareMethod.Text)

    '                        'If i_keepbigger > 0 Then
    '                        'sCheckStop = "stop"
    '                        'Else
    '                        sCheckStop = CheckUrlIn(sStoreResult, sSendstrURI)
    '                        'End If

    '                        If sCheckStop <> "stop" Then
    '                            'If Not (InStr(1, sCheckStop, "q=cache") > 0) Then
    '                            'Call Search.InsertUrl(sCheckStop, sSendstrURI)
    '                            Call InsertDNSFound(sCheckStop, sSendstrURI, i_GoodSpot + 100, , , s_searchEngine)
    '                            Call f_Extract_Racine_URL(sCheckStop, sSendstrURI, i_GoodSpot, s_searchEngine)
    '                            Dim b_seedCrawlers As Boolean = True
    '                            Call SendDNSFound(sSendstrURI, b_seedCrawlers)
    '                            'End If
    '                        End If

    '                    Else
    '                        'sStoreResult = "http://alltheweb.com"
    '                    End If

    '                    ''fGetUrlAddDNSHotbotInktomi = sAllResult

    '                    iFirst = InStr(iFirst + 5, FileToCrawl, "http://www")

    '                    Dim i_lenFileToCrawl As Int32 = Len(FileToCrawl)

    '                    If iFirst <= Len(FileToCrawl) Then
    '                        If iFirst > 0 Then
    '                            iLast = InStr(iFirst, FileToCrawl, "<")
    '                        End If
    '                    End If

    '                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
    '                        Exit Do
    '                    Else
    '                        iLenght = iLast - iFirst
    '                    End If
    '                Loop
    '            End If

    '        Catch ex As Exception

    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " Seedcrawlers.1210 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try

    '    End Function


    '    Function HotbotTeoma(ByVal s_p_keywords As String, ByVal s_p_datName As String) As String

    '        Dim sSendstrURI As String = s_p_datName
    '        Dim SGetText As String
    '        Dim sGetHttp As String
    '        Dim s_Google As String = "http://hotbot.com/default.asp?prov=Teoma&query=bush+hussein+war&ps=&loc=searchbox&tab=web"
    '        Dim sFindNext20 As String = ">Next 20"
    '        Dim iEndUrlNext20 As Integer
    '        Dim iBeginNext20 As Integer
    '        Dim bContinue As Boolean = True
    '        Dim sFindQuote As String
    '        Dim iCounter As Integer = 11
    '        Dim sFindNext20Url As String
    '        Dim iLenghtNextUrl As Integer
    '        Dim sTakeCareOfString, s_oldGetText As String
    '        Dim iFirstTimeNext20 As Integer = 0
    '        Dim sOldsFindNext20Url, s_keepString As String
    '        Dim sKeepstrURI, s_ChangeInteger, i_placeOfNext, i_placeOfDoubleQuote, i_lenBetween,
    '        i_keepPlaceFound, i_keepPlacePrevious As String
    '        Dim i_resetPBar As Integer = 0
    '        Dim b_Changed As Boolean = False
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        b_p_startANew = True

    '        Try
    '            s_p_keywords = Replace(s_p_keywords, " ", "+")

    '            s_p_keywords = ReplaceAccent(s_p_keywords)

    '            s_Google = Replace(s_Google, "bush+hussein+war", s_p_keywords)

    '            Do While bContinue

    '                If b_p_closedTheSeedingProcess Then
    '                    Exit Do
    '                End If

    '                Dim objURI As Uri = New Uri(s_Google)
    '                Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
    '                Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
    '                objWebResponse.Headers.Add("Accept-Language: fr, en")
    '                Dim objStream As Stream = objWebResponse.GetResponseStream()
    '                Dim objStreamReader As StreamReader = New StreamReader(objStream)
    '                Dim strHTML As String = objStreamReader.ReadToEnd

    '                SGetText = strHTML

    '                If Len(SGetText) > 4 Then

    '                    If Len(SGetText) = Len(s_oldGetText) And iCounter >= 200 Then
    '                        Exit Do
    '                    Else
    '                        s_oldGetText = SGetText
    '                    End If

    '                    Call fGetUrlAddDNSHotbotTeoma(SGetText, sSendstrURI)

    '                End If

    '                If Len(SGetText) > 10 Then

    '                    s_Google = "http://hotbot.com/default.asp?query=bush+hussein+war&first=11&page=more&ca=w&prov=Teoma&pc=&pcx=&date="

    '                    s_Google = Replace(s_Google, "bush+hussein+war", s_p_keywords)

    '                    s_Google = Replace(s_Google, " ", "+")

    '                    iCounter = iCounter + 10

    '                    s_Google = Replace(s_Google, "11", iCounter)

    '                    'EventLog.WriteEntry(sSource, " module2.925 HotbotTeoma next page view is --> " & iCounter & " " & s_Google, EventLogEntryType.Information, 99)

    '                Else
    '                    Exit Do
    '                End If

    '                If iCounter > 1000 Then
    '                    Exit Do
    '                End If

    '                'System.Threading.Thread.Sleep(100)
    '            Loop

    '        Catch ex As Exception

    '            EventLog.WriteEntry(sSource, " Seedcrawlers.1305 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try
    '    End Function


    '    Function fGetUrlAddDNSHotbotTeoma(ByVal FileToCrawl As String, ByVal sSendstrURI As String) As String

    '        Dim iFirst As Integer
    '        Dim iLast, i_keepbigger As Integer
    '        Dim iLenght As Integer
    '        Dim sStoreResult As String
    '        Dim sAllResult As String
    '        Dim bKeepGoing As Boolean
    '        Dim sURL As String
    '        Dim iURLIn As Integer
    '        Dim sCheckStop As String
    '        Dim iURLFast As String
    '        Dim i_GoodSpot As Integer = i_p_ranking
    '        Dim i_len1, i_len2 As Int32
    '        Dim s_array() As String
    '        Dim b_continue As Boolean = True
    '        Dim s_qualityOfSE As String = "C"
    '        Dim s_searchEngine As String = "HotbotTeoma"

    '        Try
    '            i_GoodSpot = f_iGoodSpot(s_qualityOfSE)
    '            sWhatRange()
    '            CheckParametersInDocument(sSendstrURI)

    '            bKeepGoing = True

    '            sURL = "http"

    '            iFirst = InStr(1, FileToCrawl, "http://www.")

    '            iLast = InStr(iFirst, FileToCrawl, "<")

    '            iLenght = iLast - iFirst

    '            If iLenght > 0 Then
    '                Do While bKeepGoing
    '                    sStoreResult = Mid(FileToCrawl, iFirst, iLenght)

    '                    sStoreResult = Replace(sStoreResult, """", "")

    '                    If InStr(1, sStoreResult, ">", CompareMethod.Text) > 0 Then
    '                        Dim i_countWhereToBegin, i_lenToKeep, i_lenStoreResult As Int32
    '                        i_countWhereToBegin = InStr(1, sStoreResult, ">", CompareMethod.Text)
    '                        i_lenStoreResult = Len(sStoreResult)
    '                        i_lenToKeep = i_lenStoreResult - ((i_lenStoreResult - i_countWhereToBegin) + 1)
    '                        sStoreResult = Mid(sStoreResult, 1, i_lenToKeep)
    '                    End If

    '                    iURLIn = InStr(1, sStoreResult, sURL)
    '                    iURLFast = InStr(1, sStoreResult, "fastsearch")

    '                    If iURLIn > 0 And (Not (iURLFast > 0)) Then

    '                        'sStoreResult = Replace(sStoreResult, "&nbsp;", "")
    '                        'sStoreResult = Replace(sStoreResult, "&nbsp", "")

    '                        'Do While b_continue
    '                        'i_len1 = Len(sStoreResult)
    '                        'sStoreResult = Replace(sStoreResult, "&nbsp", "")
    '                        'i_len2 = Len(sStoreResult)
    '                        'If i_len1 = i_len2 Then
    '                        'Exit Do
    '                        'End If
    '                        'Loop

    '                        s_array = Split(sStoreResult)

    '                        sStoreResult = s_array(0)

    '                        sStoreResult = Trim(RemoveHtmlTag(sStoreResult))

    '                        sAllResult = sAllResult & Chr(10) & sStoreResult

    '                        i_keepbigger = InStr(sStoreResult, ">", CompareMethod.Text)

    '                        'If i_keepbigger > 0 Then
    '                        'sCheckStop = "stop"
    '                        'Else
    '                        'sCheckStop = CheckUrlIn(sStoreResult, sSendstrURI)
    '                        'End If

    '                        'If sCheckStop <> "stop" Then
    '                        'If Not (InStr(1, sCheckStop, "q=cache") > 0) Then
    '                        'Call Search.InsertUrl(sCheckStop, sSendstrURI)

    '                        sStoreResult = Trim(sStoreResult)
    '                        sStoreResult = Replace(sStoreResult, Chr(39), "")
    '                        sStoreResult = Replace(sStoreResult, Chr(13), "")
    '                        sStoreResult = Replace(sStoreResult, Chr(10), "")
    '                        sStoreResult = Replace(sStoreResult, Chr(9), "")
    '                        Dim i_lenght As Integer = Len(sStoreResult)


    '                        If InStr((i_lenght - 1), sStoreResult, ";", CompareMethod.Text) > 0 Then
    '                            sStoreResult = Mid(sStoreResult, 1, i_lenght - 1)
    '                            sStoreResult = Trim(sStoreResult)
    '                        End If
    '                        If InStr((i_lenght - 1), sStoreResult, ")", CompareMethod.Text) > 0 Then
    '                            i_lenght = Len(sStoreResult)
    '                            sStoreResult = Mid(sStoreResult, 1, i_lenght - 1)
    '                        End If
    '                        If InStr((i_lenght - 1), sStoreResult, "'", CompareMethod.Text) > 0 Then
    '                            i_lenght = Len(sStoreResult)
    '                            sStoreResult = Mid(sStoreResult, 1, i_lenght - 1)
    '                        End If

    '                        If InsertDNSFound(sStoreResult, sSendstrURI, i_GoodSpot + 100, , , s_searchEngine) = True Then
    '                            Call f_Extract_Racine_URL(sStoreResult, sSendstrURI, i_GoodSpot, s_searchEngine)
    '                            Dim b_seedCrawlers As Boolean = True
    '                            Call SendDNSFound(sSendstrURI, b_seedCrawlers)
    '                        End If
    '                        'End If
    '                    End If

    '                    'Else
    '                    'sStoreResult = "http://alltheweb.com"
    '                    'End If

    '                    fGetUrlAddDNSHotbotTeoma = sAllResult

    '                    iFirst = InStr(iFirst + 5, FileToCrawl, "http://www.")

    '                    If iFirst <= Len(FileToCrawl) Then
    '                        If iFirst > 0 Then
    '                            iLast = InStr(iFirst, FileToCrawl, "<")
    '                        End If
    '                    End If

    '                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
    '                        Exit Do
    '                    Else
    '                        iLenght = iLast - iFirst
    '                    End If
    '                Loop
    '            End If

    '        Catch e_search_1312 As Exception
    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " lookforfiles.e_search_1312 --> " & e_search_1312.Message, EventLogEntryType.Information, 44)
    '        End Try

    '    End Function


    '    Function Altavista(ByVal s_p_keywords As String, ByVal s_p_datName As String) As String

    '        Dim sSendstrURI As String = s_p_datName
    '        Dim SGetText As String
    '        Dim sGetHttp As String
    '        Dim s_Google As String = "http://altavista.com/web/results?q=eagle+fly&kgs=0&kls=1&avkw=qtrp"
    '        Dim sFindNext20 As String = ">Next 20"
    '        Dim iEndUrlNext20 As Integer
    '        Dim iBeginNext20 As Integer
    '        Dim bContinue As Boolean = True
    '        Dim sFindQuote As String
    '        Dim iCounter As Integer = 10
    '        Dim sFindNext20Url As String
    '        Dim iLenghtNextUrl As Integer
    '        Dim sTakeCareOfString, s_oldGetText As String
    '        Dim iFirstTimeNext20 As Integer = 0
    '        Dim sOldsFindNext20Url, s_keepString As String
    '        Dim sKeepstrURI, s_ChangeInteger, i_placeOfNext, i_placeOfDoubleQuote, i_lenBetween,
    '        i_keepPlaceFound, i_keepPlacePrevious As String
    '        Dim i_resetPBar As Integer = 0
    '        Dim b_Changed As Boolean = False
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        b_p_startANew = True

    '        Try
    '            s_p_keywords = Replace(s_p_keywords, " ", "+")

    '            s_p_keywords = ReplaceAccent(s_p_keywords)

    '            s_Google = Replace(s_Google, "eagle+fly", s_p_keywords)

    '            Do While bContinue

    '                If b_p_closedTheSeedingProcess Then
    '                    Exit Do
    '                End If

    '                Dim objURI As Uri = New Uri(s_Google)
    '                Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
    '                Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
    '                objWebResponse.Headers.Add("Accept-Language: fr, en")
    '                Dim objStream As Stream = objWebResponse.GetResponseStream()
    '                Dim objStreamReader As StreamReader = New StreamReader(objStream)
    '                Dim strHTML As String = objStreamReader.ReadToEnd

    '                SGetText = strHTML

    '                If Len(SGetText) > 12000 Then

    '                    If Len(SGetText) = Len(s_oldGetText) And iCounter >= 1000 Then
    '                        Exit Do
    '                    Else
    '                        s_oldGetText = SGetText
    '                    End If

    '                    Call fGetUrlAddDNSAltavista(SGetText, sSendstrURI)

    '                Else
    '                    Exit Do
    '                    'EventLog.WriteEntry(sSource, " module2.1127 Altavista no records found --> ", EventLogEntryType.Information, 99)

    '                End If

    '                If Len(SGetText) > 10 Then

    '                    s_Google = "http://altavista.com/web/results?q=eagle+fly&kgs=0&kls=1&avkw=qtrp&stq=10"

    '                    s_Google = Replace(s_Google, "eagle+fly", s_p_keywords)

    '                    s_Google = Replace(s_Google, " ", "+")

    '                    iCounter = iCounter + 10

    '                    s_Google = Replace(s_Google, "10", iCounter)

    '                    'EventLog.WriteEntry(sSource, " module2.1135 Altavista next page view is --> " & iCounter & " " & s_Google, EventLogEntryType.Information, 99)

    '                Else
    '                    Exit Do
    '                End If

    '                If iCounter > 1000 Then
    '                    Exit Do
    '                End If

    '            Loop

    '        Catch ex As Exception

    '            EventLog.WriteEntry(sSource, " Seedcrawlers.1524 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try
    '    End Function


    '    Function fGetUrlAddDNSAltavista(ByVal FileToCrawl As String, ByVal sSendstrURI As String) As String

    '        Dim iFirst As Integer
    '        Dim iLast, i_keepbigger As Integer
    '        Dim iLenght As Integer
    '        Dim sStoreResult As String
    '        Dim sAllResult As String
    '        Dim bKeepGoing As Boolean
    '        Dim sURL As String
    '        Dim iURLIn As Integer
    '        Dim sCheckStop As String
    '        Dim iURLFast As String
    '        Dim i_GoodSpot As Integer = i_p_ranking
    '        Dim i_len1, i_len2 As Int32
    '        Dim s_array() As String
    '        Dim b_continue As Boolean = True
    '        Dim s_qualityOfSE As String = "B"
    '        Dim s_searchEngine As String = "Altavista"

    '        Try
    '            i_GoodSpot = f_iGoodSpot(s_qualityOfSE)
    '            sWhatRange()
    '            CheckParametersInDocument(sSendstrURI)

    '            bKeepGoing = True

    '            'MessageBox.Show(" FileToCrawl.628 " & FileToCrawl)

    '            sURL = "http"

    '            iFirst = InStr(1, FileToCrawl, "status='http://")

    '            'MessageBox.Show(InStr(1, FileToCrawl, "http://"), " 638.InStr(1, FileToCrawl, http:// ")

    '            iLast = InStr(iFirst + 9, FileToCrawl, "'")

    '            'MessageBox.Show(InStr(iFirst, FileToCrawl, ">"), "  637.InStr iFirst, FileToCrawl, >")

    '            iLenght = iLast - iFirst

    '            If iLenght > 0 Then
    '                Do While bKeepGoing
    '                    sStoreResult = Mid(FileToCrawl, iFirst + 8, iLenght - 8)

    '                    sStoreResult = Replace(sStoreResult, """", "")
    '                    'MessageBox.Show(sStoreResult, "sStoreResult 2")

    '                    iURLIn = InStr(1, sStoreResult, sURL)
    '                    iURLFast = InStr(1, sStoreResult, "fastsearch")

    '                    If iURLIn > 0 And (Not (iURLFast > 0)) Then

    '                        sStoreResult = Replace(sStoreResult, "&nbsp;", "")
    '                        sStoreResult = Replace(sStoreResult, "&nbsp", "")

    '                        Do While b_continue
    '                            i_len1 = Len(sStoreResult)
    '                            sStoreResult = Replace(sStoreResult, "&nbsp", "")
    '                            i_len2 = Len(sStoreResult)
    '                            If i_len1 = i_len2 Then
    '                                Exit Do
    '                            End If
    '                        Loop

    '                        s_array = Split(sStoreResult)

    '                        sStoreResult = s_array(0)

    '                        sAllResult = sAllResult & Chr(10) & sStoreResult

    '                        sCheckStop = CheckUrlIn(sStoreResult, sSendstrURI)

    '                        If sCheckStop <> "stop" Then
    '                            Call InsertDNSFound(sCheckStop, sSendstrURI, i_GoodSpot + 100, , , s_searchEngine)
    '                            Call f_Extract_Racine_URL(sCheckStop, sSendstrURI, i_GoodSpot, s_searchEngine)
    '                            Dim b_seedCrawlers As Boolean = True
    '                            Call SendDNSFound(sSendstrURI, b_seedCrawlers)
    '                        End If

    '                    Else

    '                        sStoreResult = "http://altavista.com"

    '                    End If

    '                    fGetUrlAddDNSAltavista = sAllResult

    '                    iFirst = InStr(iFirst + 5, FileToCrawl, "status='http://")

    '                    If iFirst <= Len(FileToCrawl) Then
    '                        If iFirst > 0 Then
    '                            iLast = InStr(iFirst + 9, FileToCrawl, "'")
    '                        End If
    '                    End If

    '                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
    '                        Exit Do
    '                    Else
    '                        iLenght = iLast - iFirst
    '                    End If
    '                Loop
    '            End If

    '        Catch ex As Exception

    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " seedcrawlers.1636 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try

    '    End Function


    '    Function LooksSmart(ByVal s_p_keywords As String, ByVal s_p_datName As String) As String

    '        Dim sSendstrURI As String = s_p_datName
    '        Dim SGetText As String
    '        Dim sGetHttp As String
    '        Dim s_Google As String = "http://www.looksmart.com/r_search?look=&key=eagle+fly"
    '        Dim sFindNext20 As String = ">Next 20"
    '        Dim iEndUrlNext20 As Integer
    '        Dim iBeginNext20 As Integer
    '        Dim bContinue As Boolean = True
    '        Dim sFindQuote As String
    '        Dim iCounter As Integer = 10
    '        Dim sFindNext20Url As String
    '        Dim iLenghtNextUrl As Integer
    '        Dim sTakeCareOfString, s_oldGetText As String
    '        Dim iFirstTimeNext20 As Integer = 0
    '        Dim sOldsFindNext20Url, s_keepString As String
    '        Dim sKeepstrURI, s_ChangeInteger, i_placeOfNext, i_placeOfDoubleQuote, i_lenBetween,
    '        i_keepPlaceFound, i_keepPlacePrevious As String
    '        Dim i_resetPBar As Integer = 0
    '        Dim b_Changed As Boolean = False
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        b_p_startANew = True

    '        Try
    '            s_p_keywords = Replace(s_p_keywords, " ", "+")

    '            s_p_keywords = ReplaceAccent(s_p_keywords)

    '            s_Google = Replace(s_Google, "eagle+fly", s_p_keywords)

    '            Do While bContinue

    '                If b_p_closedTheSeedingProcess Then
    '                    Exit Do
    '                End If

    '                Dim objURI As Uri = New Uri(s_Google)
    '                Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
    '                Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
    '                objWebResponse.Headers.Add("Accept-Language: fr, en")
    '                Dim objStream As Stream = objWebResponse.GetResponseStream()
    '                Dim objStreamReader As StreamReader = New StreamReader(objStream)
    '                Dim strHTML As String = objStreamReader.ReadToEnd

    '                SGetText = strHTML

    '                If Len(SGetText) > 4 Then

    '                    If Len(SGetText) = Len(s_oldGetText) And iCounter >= 2000 Then
    '                        Exit Do
    '                    Else
    '                        s_oldGetText = SGetText
    '                    End If

    '                    Call fGetUrlAddDNSLookSmart(SGetText, sSendstrURI)

    '                End If

    '                If Len(SGetText) > 10 Then

    '                    s_Google = "http://www.looksmart.com/r_search?l&key=eagle+fly&skip=15&se=4,0,73,0&search=us302562;local_US"

    '                    s_Google = Replace(s_Google, "eagle+fly", s_p_keywords)

    '                    iCounter = iCounter + 15

    '                    s_Google = Replace(s_Google, "15", iCounter)

    '                    'EventLog.WriteEntry(sSource, " module2.1348 LookSmart next page view is --> " & iCounter & " " & s_Google, EventLogEntryType.Information, 99)

    '                Else

    '                    Exit Do

    '                End If

    '                If iCounter > 1000 Or InStr(SGetText, "NEXT&nbsp;15", CompareMethod.Text) = 0 Then
    '                    Exit Do
    '                End If

    '                If iCounter > 1000 Then
    '                    'EventLog.WriteEntry(sSource, " LookSmart done page view!--> " & iCounter, EventLogEntryType.Information, 99)
    '                    Exit Do
    '                End If
    '                'System.Threading.Thread.Sleep(100)
    '            Loop

    '        Catch ex As Exception

    '            EventLog.WriteEntry(sSource, " Seedcrawlers.1735 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try
    '    End Function


    '    Function fGetUrlAddDNSLookSmart(ByVal FileToCrawl As String, ByVal sSendstrURI As String) As String

    '        Dim iFirst, i_keepSmaller, i_len As Integer
    '        Dim iLast, i_keepbigger As Integer
    '        Dim iLenght As Integer
    '        Dim sStoreResult As String
    '        Dim sAllResult As String
    '        Dim bKeepGoing As Boolean
    '        Dim sURL As String
    '        Dim iURLIn As Integer
    '        Dim sCheckStop As String
    '        Dim iURLFast As String
    '        Dim i_GoodSpot As Integer = i_p_ranking
    '        Dim i_len1, i_len2 As Int32
    '        Dim s_array() As String
    '        Dim b_continue As Boolean = True

    '        Dim s_qualityOfSE As String = "D"
    '        Dim s_searchEngine As String = "Looksmart"

    '        Try

    '            i_GoodSpot = f_iGoodSpot(s_qualityOfSE)

    '            sWhatRange()

    '            CheckParametersInDocument(sSendstrURI)

    '            bKeepGoing = True

    '            sURL = "http"

    '            iFirst = InStr(1, FileToCrawl, "http:")

    '            If iFirst = 0 Then
    '                Exit Function
    '            End If

    '            iLast = InStr(iFirst, FileToCrawl, ">")

    '            iLenght = iLast - iFirst


    '            If iLenght > 0 Then
    '                Do While bKeepGoing
    '                    sStoreResult = Mid(FileToCrawl, iFirst, iLenght)
    '                    sStoreResult = Replace(sStoreResult, """", "")
    '                    sStoreResult = Replace(sStoreResult, """", "")
    '                    i_keepSmaller = InStr(1, sStoreResult, "<")

    '                    If i_keepSmaller > 0 Then
    '                        sStoreResult = Trim(Mid(sStoreResult, 1, i_keepSmaller - 1))
    '                    End If

    '                    iURLIn = InStr(1, sStoreResult, sURL)
    '                    iURLFast = InStr(1, sStoreResult, "fastsearch")

    '                    If InStr(sStoreResult, "ad.doubleclick.net", CompareMethod.Text) > 0 Then
    '                        sStoreResult = "http://altavista.com"
    '                    End If

    '                    If iURLIn > 0 And (Not (iURLFast > 0)) Then

    '                        sStoreResult = Replace(sStoreResult, "&nbsp;", "")
    '                        sStoreResult = Replace(sStoreResult, "&nbsp", "")

    '                        Do While b_continue
    '                            i_len1 = Len(sStoreResult)
    '                            sStoreResult = Replace(sStoreResult, "&nbsp", "")
    '                            i_len2 = Len(sStoreResult)
    '                            If i_len1 = i_len2 Then
    '                                Exit Do
    '                            End If
    '                        Loop

    '                        s_array = Split(sStoreResult)

    '                        sStoreResult = s_array(0)

    '                        sAllResult = sAllResult & Chr(10) & sStoreResult

    '                        sCheckStop = CheckUrlIn(sStoreResult, sSendstrURI)

    '                        If sCheckStop <> "stop" Then

    '                            'Call Search.InsertUrl(sCheckStop, sSendstrURI)
    '                            Call InsertDNSFound(sCheckStop, sSendstrURI, i_GoodSpot + 101, , , s_searchEngine)
    '                            Call f_Extract_Racine_URL(sCheckStop, sSendstrURI, i_GoodSpot, s_searchEngine)
    '                            Dim b_seedCrawlers As Boolean = True
    '                            Call SendDNSFound(sSendstrURI, b_seedCrawlers)
    '                        End If

    '                    Else

    '                        sStoreResult = "http://altavista.com"

    '                    End If

    '                    fGetUrlAddDNSLookSmart = sAllResult

    '                    iFirst = InStr(iFirst + 6, FileToCrawl, "http:")

    '                    If iFirst = 0 Then
    '                        Exit Function
    '                    End If

    '                    If iFirst <= Len(FileToCrawl) Then
    '                        If iFirst > 0 Then
    '                            iLast = InStr(iFirst, FileToCrawl, ">")
    '                        End If
    '                    End If

    '                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
    '                        Exit Do
    '                    Else
    '                        iLenght = iLast - iFirst
    '                    End If
    '                Loop
    '            End If

    '        Catch ex As Exception

    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " Seedcrawlers.1865 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try

    '    End Function

    Public i_p_counterGoodSpot As Int32

    Public b_p_startANew As Boolean

    Function f_iGoodSpot(Optional ByVal s_sEngineQuality As String = "D") As Int32
        Try
            If b_p_startANew Then
                If s_sEngineQuality = "A" Then i_p_counterGoodSpot = 200
                If s_sEngineQuality = "B" Then i_p_counterGoodSpot = 150
                If s_sEngineQuality = "C" Then i_p_counterGoodSpot = 125
                If s_sEngineQuality = "D" Then i_p_counterGoodSpot = 100
                b_p_startANew = False
            End If

            f_iGoodSpot = i_p_counterGoodSpot

            i_p_counterGoodSpot = i_p_counterGoodSpot - 1

        Catch seedcrawler_1840 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " seedcrawler_1840 --> " & seedcrawler_1840.Message, EventLogEntryType.Information, 44)

        End Try
    End Function


    '    Function CheckIfStartSearchCycleAgain(ByVal s_RangeMinus, ByVal RangePlus) As Boolean

    '        Dim mySelectQuery, s, s_queryName, s_dataPath, s_valideFileName As String
    '        Dim d_dateInsert As DateTime
    '        Dim myConnection As New OleDbConnection(ConnStringDNA())
    '        Dim s_verb As String = "acknowNewMachinConfig"
    '        Dim b_continue As Boolean = True
    '        Dim s_No As String = "n"
    '        Dim s_so As String = "so"
    '        Dim s_s As String = "s"

    '        Try
    '            mySelectQuery = "SELECT * FROM QUERY"
    '            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

    '            myConnection.Open()

    '            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

    '            While myReader.Read()

    '                s_queryName = myReader("QuerySearch").ToString

    '                If Len(s_queryName) > 0 Then

    '                    s_valideFileName = s_queryName & "_DNS.mdb"
    '                    s_valideFileName = Replace(s_valideFileName, " ", "_")

    '                    If File.Exists(s_valideFileName) Then

    '                        Dim objConn As New OleDbConnection(ConnStringURLDNS(s_queryName))
    '                        mySelectQuery = "UPDATE DNSFOUND SET Done='" & (s_No) & "' where ShortURL >='" & (s_RangeMinus) & "' and  ShortURL <='" & (RangePlus) & "' "
    '                        Dim myCommand2 As New OleDbCommand(mySelectQuery)
    '                        myCommand2.Connection = objConn
    '                        objConn.Open()
    '                        myCommand2.ExecuteNonQuery()
    '                        objConn.Close()

    '                        If s_p_authority <> "yes" Then
    '                            'mySelectQuery = "UPDATE DNSFOUND SET Done='" & (s_so) & "' where ShortURL <'" & (s_RangeMinus) & "' or  ShortURL > '" & (RangePlus) & "' "
    '                            'Dim myCommand4 As New OleDbCommand(mySelectQuery)
    '                            'myCommand4.Connection = objConn
    '                            'objConn.Open()
    '                            'myCommand4.ExecuteNonQuery()
    '                            'objConn.Close()
    '                        Else
    '                            'mySelectQuery = "UPDATE DNSFOUND SET Done='" & (s_s) & "' where ShortURL <'" & (s_RangeMinus) & "' or  ShortURL > '" & (RangePlus) & "' "
    '                            'Dim myCommand4 As New OleDbCommand(mySelectQuery)
    '                            'myCommand4.Connection = objConn
    '                            'objConn.Open()
    '                            'myCommand4.ExecuteNonQuery()
    '                            'objConn.Close()
    '                        End If

    '                    End If

    '                End If

    '            End While

    '            myReader.Close()
    '            myConnection.Close()

    '            CheckIfStartSearchCycleAgain = True

    '        Catch seedcrawler_1855 As Exception

    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " seedcrawler_1855 --> " & seedcrawler_1855.Message, EventLogEntryType.Information, 44)

    '        End Try

    '    End Function

    Public b_updateDone As Boolean

    Sub UpdateSubQueryToDo()

        Try
            'If DateTime.Now.Day = 6 Or DateTime.Now.Day = 21 And b_updateDone = False Then
            If DateTime.Now.Day = 6 And b_updateDone = False Then
                Dim s_no As String = "n"
                Dim myConnection As New OleDbConnection(ConnStringDNA())
                Dim myConnectionString As String = "UPDATE SUBQUERY SET Done='" & (s_no) & "'"

                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = myConnection
                myConnection.Open()
                myCommand2.ExecuteNonQuery()
                myCommand2.Connection.Close()
                myConnection.Close()

                myConnectionString = ""
                myConnectionString = "UPDATE SYNONYMS SET Done='" & (s_no) & "'"

                Dim myCommand3 As New OleDbCommand(myConnectionString)
                myCommand3.Connection = myConnection
                myConnection.Open()
                myCommand3.ExecuteNonQuery()
                myCommand3.Connection.Close()
                myConnection.Close()
                b_updateDone = True
            Else
                b_updateDone = False
            End If

        Catch seedCrawlers_1942 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " seedCrawlers_1942 --> " & seedCrawlers_1942.Message, EventLogEntryType.Information, 44)

        End Try
    End Sub

End Module

