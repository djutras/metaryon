
Imports System
Imports System.IO
Imports System.Net
Imports System.Net.Sockets
Imports System.Text
Imports System.Data
Imports System.Data.OleDb
Imports System.Data.SqlClient
Imports System.Management
Imports System.Diagnostics
Imports System.Collections
Imports System.Configuration
Imports System.Threading
Imports System.Threading.ThreadPriority
Imports System.MarshalByRefObject
'Imports System.Windows.Forms.Form


Public Module LookForFiles

    Public i_p_Duration As Double
    Public i_p_MinKPage, i_p_MaxKPage, i_Priority As Integer
    Public i_p_Level, i_p_MaxPage, i_p_MinDensity, i_p_MaxDensity As Integer
    Public i_p_KeywordTitle, i_p_KeywordBody, i_p_keywordTag, i_p_KeywordVisible As Integer
    Public i_p_KeywordURL, i_p_MinDiskPerc, i_p_MinDiskGiga, i_p_Algorythm As Integer
    Public i_p_Google, i_p_AllTheWeb, i_p_SearchExcel, i_p_SearchWord, i_p_SearchPDF As Integer
    Public i_p_DesactiveQuery, i_p_ExactPhrase As Integer
    Public s_p_UniversalNetwork, s_p_FirstNetwork, s_p_SecondNetwork, s_p_Message, s_p_Line, s_p_FunName As String
    Public s_p_VarName, s_p_VarValue, s_p_Priority, s_p_machineName, s_p_authority As String
    Public s_p_Void As String = "Void"
    Public d_p_DateTime As DateTime
    Public s_p_doneCheckParameters As String = "no"

    Public Function f_FindUrlToScan(ByVal sToInsert As String, ByVal bContinue As Boolean) As String

        Try
          
                Call f_FindUrlToScanFU()

            Dim b_Continue As Boolean = True
            Dim b_DocumentInScratchpad As Boolean = True
            Dim b_Continue1 As Boolean = True
            Dim b_CheckPassed As Boolean = False
            Dim s_URLToLookFor As String
            Dim s_DatabaseName As String = sToInsert
            Dim s_FileToCrawl As String
            Dim s_FoundRacine As String
            Dim s_UrlDocumentCheckKeywords As String
            Dim s_Parameters As String
            Dim s_URLOfFileToCrawlFromScratchpad As String
            Dim s_FileToCrawlFromScratchpad As String
            Dim a As Integer
            Dim s_InsertUrlFound As String
            Dim i_CounterToSendDocument As Int64
            Dim b_showDebug As Boolean = False
            'b_showDebug = True

            If s_p_doneCheckParameters = "no" then
                Call CheckParametersInDocument(s_DatabaseName)
                s_p_doneCheckParameters = "yes"
            End If
            'Call AdjustThreadPriority()

            Do While b_Continue

                If StartMeUp.b_p_showS = True Then

                    Console.WriteLine(vbNewLine)
                    Console.WriteLine("-------------------------------------------------------------------------------")
                    Console.WriteLine(s_p_authority & "  --> s_p_authority.lookforfiles.121")
                    Console.WriteLine("-------------------------------------------------------------------------------")
                    Console.WriteLine(vbNewLine)

                End If

                If s_p_authority = "yes" Then

                    '    'Call DistributeAlphaRange()

                    '    'Call sendNewMachineConfiguration()

                    '    'Call SendQueryInsertOverNetwork()

                    '    'Call SendQueryUpdateOverNetwork()

                    '    'Call Check_BROADCAST_QUERY_SO()

                    '    'Call SendQueryDeleteOverNetwork()

                    Call FindSeeds1()

                    '    'Call RemoveTheSite(s_DatabaseName)

                Else

                    '    'Call PublishMachineName()  'TO CHANGE ASAP

                End If

                'Call CheckDate()

                If b_p_closedTheSeedingProcess Then
                    Exit Do
                End If

                Call CheckDuration()

                'Call SendHeartbeat()

                'Call f_CheckCycle(s_DatabaseName)

                'Call CheckInternetConnection()

                s_FoundRacine = f_FetchRacine(s_DatabaseName, s_URLToLookFor)

                

                If StartMeUp.b_p_showS = True Then
                    Console.WriteLine(vbNewLine)
                    Console.WriteLine("lookforfiles.s_FileToCrawl.95 -- s_FoundRacine " & s_FoundRacine)
                End If

                If s_FoundRacine = "stop" Then
                    System.Threading.Thread.Sleep(1000)
                    s_FileToCrawl = "stop"
                Else
                    s_FileToCrawl = f_FindAllTheHTMLFromTheURL(s_FoundRacine, s_DatabaseName)
                End If

                If s_FileToCrawl <> "stop" And Len(s_FileToCrawl) > 0 Then
                    s_FileToCrawl = ReplaceAccent(s_FileToCrawl)
                    Call fGetUrlFromFileInsertInURLFOUND(s_FileToCrawl, s_DatabaseName, s_FoundRacine)
                Else
                    RemovePointBecauseOf404(s_FoundRacine, s_DatabaseName)
                    Console.WriteLine("lookforfiles.s_FileToCrawl.109 -- is equal to stop or len os 0!   " & s_FoundRacine)
                End If

                '//---- Begin to send document found
                'Call SendDNSFound(s_DatabaseName)

                'Call ShowFunction("SendDNSFound", "153", b_showDebug)

                'Call SendURLFound(s_DatabaseName)

                'Call ShowFunction("SendURLFound", "157", b_showDebug)

                '\\---- End send document found

                Do While b_DocumentInScratchpad

                    Call ShowFunction("b_DocumentInScratchpad", "163", b_showDebug)

                    If b_p_closedTheSeedingProcess Then
                        Exit Do
                    End If

                    'Call f_CheckCycle(s_DatabaseName)

                    'Call ShowFunction("f_CheckCycle", "171", b_showDebug)

                    Call CheckSizeOfLog()

                    Call ShowFunction("CheckSizeOfLog", "176", b_showDebug)

                    s_URLOfFileToCrawlFromScratchpad = CheckDocumentInScratchpad(s_DatabaseName, s_FoundRacine)

                    Call ShowFunction("CheckDocumentInScratchpad", "179", b_showDebug)

                    If s_URLOfFileToCrawlFromScratchpad = "stop" Then
                        Exit Do
                    Else

                        s_FileToCrawlFromScratchpad = f_FindAllTheHTMLFromTheURL(s_URLOfFileToCrawlFromScratchpad, s_DatabaseName)

                        Call ShowFunction("f_FindAllTheHTMLFromTheURL", "187", b_showDebug)

                        If Len(s_FileToCrawlFromScratchpad) > i_p_lenghtMinimumFile Then

                            b_CheckPassed = False

                            b_CheckPassed = CheckForKeywords(s_FileToCrawlFromScratchpad, s_DatabaseName)

                            Call ShowFunction("CheckForKeywords", "195", b_showDebug)

                            If b_CheckPassed Then

                                Dim i_findSignature As Int64

                                i_findSignature = FindSignature(s_FileToCrawlFromScratchpad, s_DatabaseName)

                                Call ShowFunction("FindSignature", "203", b_showDebug)

                                Call InsertURLFOUND(s_URLOfFileToCrawlFromScratchpad, s_DatabaseName, s_FoundRacine, i_findSignature, s_FileToCrawlFromScratchpad)

                                Call ShowFunction("InsertURLFOUND", "207", b_showDebug)

                            End If

                            s_p_DatabaseName = ""

                        End If

                    End If

                    Call CheckIfTimeToChangeQuery()

                    Call ShowFunction("CheckIfTimeToChangeQuery", "219", b_showDebug)

                Loop

                Exit Do

            Loop

        Catch e_LookForFiles_174 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_174 --> " & e_LookForFiles_174.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public d_p_DateTimeT As DateTime

    Public i_p_SendGoodSpotOverToUrlFound As Int32
    Public i_p_SendDateHourOverToUrlFound As DateTime
    Public i_p_countTimeAfterDNSFOUNDEmpty As Long

    Public Function f_TakeRacine_DNS_TO_REMOVE(ByVal sDNSName As String, ByVal sFoundRacine As String, ByVal Verb As String) As Object

        Dim s_Verb As String = Verb
        Dim myConnString As String
        Dim sKeepString As String
        Dim sSendUri As String = sDNSName
        Dim s_NA = "n/a"
        Dim SUrl As String = "n"
        Dim i_ID As Integer
        Dim mySelectQuery As String
        Dim i_GoodSpot As Integer
        Dim s_KeepDate As String = Now()
        Dim d_dateHour As Date
        Dim d_dateRatingDone As Date
        Dim b_go As Boolean = True
        Dim s_done As String
        Dim b_dnsFoundEmpty As Boolean = True

        Dim myConnection As New OleDbConnection(ConnStringURLDNS(sSendUri))
        Dim myReader As OleDbDataReader

        Try
            If s_Verb = "s_Racine" Then
                mySelectQuery = "SELECT top 1 ID,URLSite,Done,iGoodSpot,DateHour FROM DNSFOUND where Done ='" & Trim(SUrl) & "' and URLSite <> '" & Trim(s_NA) & "' order by iGoodSpot desc,DateHour"
            Else
                If Len(sFoundRacine) > 0 Then
                    mySelectQuery = "SELECT ID,URLSite,iGoodSpot,Done,DateHour FROM DNSFOUND where URLSite = '" & sFoundRacine & "'"
                Else
                    b_go = False
                End If
            End If

            If b_go Then

                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
                Dim o_object As Object
                Dim s_RetrieveIDDateHour As Array

                myConnection.Open()

                myReader = myCommand.ExecuteReader()

                While myReader.Read()

                    b_dnsFoundEmpty = False

                    sKeepString = sKeepString & (myReader("URLSite"))
                    i_ID = myReader("ID")
                    i_GoodSpot = myReader("iGoodSpot")
                    i_p_SendGoodSpotOverToUrlFound = myReader("iGoodSpot")
                    s_done = myReader("Done")
                    d_dateHour = myReader("DateHour")
                    i_p_SendDateHourOverToUrlFound = myReader("DateHour")

                    If s_Verb = "s_Racine" Then
                        d_p_DateTimeT = myReader("DateHour")
                    End If

                    If Len(sKeepString) > 4 Then
                        If s_Verb = "s_Racine" Then
                            f_TakeRacine_DNS_TO_REMOVE = sKeepString
                        Else
                            f_TakeRacine_DNS_TO_REMOVE = i_ID
                        End If
                    Else
                        f_TakeRacine_DNS_TO_REMOVE = "stop"
                    End If
                    Exit While
                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()

            End If

            If b_dnsFoundEmpty Then

                i_p_countTimeAfterDNSFOUNDEmpty = i_p_countTimeAfterDNSFOUNDEmpty + 1
                Console.WriteLine(i_p_countTimeAfterDNSFOUNDEmpty & "  -- lookforfiles.i_p_countTimeAfterDNSFOUNDEmpty.257")

                If i_p_countTimeAfterDNSFOUNDEmpty >= 50 Then
                    i_p_countTimeAfterDNSFOUNDEmpty = 0
                    'Call Sub_SendHeartBeatUpdateDNS()
                    Call Sub_checkDNSNotDone(sSendUri)
                    Call Sub_DeleteRecordsScratchPad()
                    Call sub_restartDNSFound(sSendUri)
                End If

            End If

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.283 --> " & ex.Message, EventLogEntryType.Information, 44)
            EventLog.WriteEntry(sSource, " lookforfiles.284 sSendUri ---- mySelectQuery--> " & sSendUri & " ---- " & mySelectQuery, EventLogEntryType.Information, 44)

        End Try

        d_dateRatingDone = f_DateHourRating(d_dateHour)

        If s_Verb = "s_Racine" Then

            If Len(sKeepString) > 7 Then
                Try
                    'If s_done = "n" Then
                    Dim s_yes As String = "y"
                    Dim objConn As New OleDbConnection(ConnStringURLDNS(sSendUri))
                    Dim myConnectionString As String = "UPDATE DNSFOUND SET Done= '" & (s_yes) & "',DateHour =  '" & (d_dateRatingDone) & "'  WHERE ID =" & i_ID & ""
                    Dim myCommand2 As New OleDbCommand(myConnectionString)
                    myCommand2.Connection = objConn
                    objConn.Open()
                    myCommand2.ExecuteNonQuery()
                    objConn.Close()
                    'End If

                Catch eOx As Exception
                    'CompacAccess1(sSendUri)
                    Dim sSource As String = "AP_DENIS"
                    Dim sLog As String = "Applo"
                    EventLog.WriteEntry(sSource, "LookForFiles.303 --> " & eOx.Message, EventLogEntryType.Information, 44)
                End Try
            Else
                f_TakeRacine_DNS_TO_REMOVE = "stop"
            End If

        Else

            If i_GoodSpot < 101 And i_GoodSpot >= -200 And b_p_freshPages = True Then
                i_GoodSpot = 101
            Else

                Dim i_KeepDuration As Int32 = i_p_Duration * 60

                Dim i_MinuteSinceInsert As Long = DateDiff(DateInterval.Minute, d_p_DateTime, Now)

                If i_MinuteSinceInsert >= i_KeepDuration And b_p_freshPages = True Then
                    i_GoodSpot = i_GoodSpot + 2
                ElseIf i_MinuteSinceInsert >= (i_KeepDuration * 2) Then
                    i_GoodSpot = i_GoodSpot + 1
                ElseIf i_MinuteSinceInsert >= i_KeepDuration Then
                    i_GoodSpot = i_GoodSpot + 1
                End If

            End If

            If Len(sKeepString) > 7 Then

                Try

                    Dim objConn As New OleDbConnection(ConnStringURLDNS(sSendUri))
                    Dim myConnectionString As String = "UPDATE DNSFOUND SET iGoodSpot='" & (i_GoodSpot) & "' WHERE ID =" & i_ID & ""

                    Dim myCommand2 As New OleDbCommand(myConnectionString)
                    myCommand2.Connection = objConn
                    objConn.Open()
                    myCommand2.ExecuteNonQuery()
                    myCommand2.Connection.Close()
                    objConn.Close()

                Catch e_LookForFiles_301 As Exception
                    Dim sSource As String = "AP_DENIS"
                    Dim sLog As String = "Applo"
                    EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_301 --> " & e_LookForFiles_301.Message, EventLogEntryType.Information, 44)
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine(vbNewLine)
                    Console.WriteLine("Database might be corrupted, please call programmers! e_LookForFiles_301  " & e_LookForFiles_301.Message)
                    Console.WriteLine(vbNewLine)
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    Console.WriteLine("*****************************************************************")
                    System.Threading.Thread.Sleep(60000)
                End Try
            Else
                f_TakeRacine_DNS_TO_REMOVE = "stop"
            End If
        End If

    End Function


    Public Sub sub_restartDNSFound(ByVal s_DatabaseName As String)

        Try
            Dim i_lowerID As Integer = 0
            Dim i_upperID As Integer = 10000
            Dim b_Continue As Boolean = true
           Dim s_na As String = "n/a"

            Do While b_Continue

                Dim s_No As String = "n"
                Dim s_Yes As String = "y"
                Dim objConn As New OleDbConnection(ConnStringURLDNS(s_DatabaseName))
                Dim myConnectionString As String = "UPDATE DNSFOUND SET Done='" & (s_No) & "' where Done ='" & (s_Yes) & "' and ID > " & i_lowerID & " and ID < " & i_upperID & " and URLSite <> '" & s_na & "' "
                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = objConn
                objConn.Open()
                myCommand2.ExecuteNonQuery()
                objConn.Close()

                i_lowerID=i_lowerID+9999
                i_upperID=i_upperID+10001

                System.Threading.Thread.Sleep(2000)

                If i_upperID > 1000000 Then
                    b_Continue = False
                End if
            loop

        Catch e_lookforfiles_347 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_347 --> " & e_lookforfiles_347.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub


    Public Function f_TakeURLFromDNSFound(ByVal sToInsert As String) As String

        Dim myConnString As String
        Dim sKeepString As String
        Dim sSendUri As String = sToInsert
        Dim SUrl As String = "n"

        Try
            Dim mySelectQuery As String = "SELECT top 1 url FROM URLFOUND where  GoodSpot <> -40004 and Done ='" & Trim(SUrl) & "' order by iGoodSpot desc, DateHour"
            Dim myConnection As New OleDbConnection(ConnStringURL(sSendUri))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            While myReader.Read()
                sKeepString = sKeepString & (myReader.GetString(0))
                If Len(sKeepString) > 0 Then
                    f_TakeURLFromDNSFound = sKeepString
                Else
                    f_TakeURLFromDNSFound = "stop"
                End If

                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            ''myConnection = Nothing
            myConnection.Close()


        Catch e_lookforfiles_342 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_342 --> " & e_lookforfiles_342.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Public s_p_PassUri As String = ""


    Public Function f_FindAllTheRacines(ByVal sRetrieveHTMLPerThisURL As String) As String
        Try
            Dim s_RetrieveHTML As String
            s_p_PassUri = sRetrieveHTMLPerThisURL

            Call f_GetHtmlFromURL(sRetrieveHTMLPerThisURL)

        Catch e_LookForFiles_360 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_360 --> " & e_LookForFiles_360.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Dim i_p_GotTheHtml As Integer
    Dim i_p_NoHtml As Integer
    Dim i_p_TotalPass As Integer


    '    Private Delegate Sub Starter()

    Public s_p_sDatabaseName As String

    Public Function f_FindAllTheHTMLFromTheURL(ByVal sURLToLookFor, ByVal sDatabaseName) As String

        Dim s_URLToLookFor As String = sURLToLookFor
        Dim b_WaitForThread As Boolean = True
        Dim i_TimeCounter As Integer = 0
        Dim i_TimeToJoin As Integer
        Try
            If i_p_TotalPass < 1 Then
                i_p_GotTheHtml = 1
                i_p_NoHtml = 1
                i_p_TotalPass = 1
            End If

            Dim s_RetrieveHTML As String=""
            If Len(s_URLToLookFor) > 0 Then
                s_p_PassUri = s_URLToLookFor
                s_p_sDatabaseName = sDatabaseName

                Dim s_KeepString As String
                Dim i_KeepScore As Integer = 50
                Dim i_Myresult As Integer

                ''Dim t As Thread

                ''t = New Thread(AddressOf f_GetHtmlFromURI)
                ''t.Name = "GetHtml_Thread"
                ''t.Priority = Normal
                ''t.Start()

                'Call f_GetHtmlFromURI()

                'If s_p_authority = "yes" Then
                'i_TimeToJoin = 30000
                'Else
                i_TimeToJoin = 15000
                'End If

                ''If t.Join(i_TimeToJoin) Then
                ''    f_FindAllTheHTMLFromTheURL = s_p_RetrievedHtml
                ''    i_p_TotalPass = i_p_TotalPass + 1
                ''    i_p_GotTheHtml = i_p_GotTheHtml + 1
                ''Else
                ''    s_p_RetrievedHtml = "stop"
                ''    f_FindAllTheHTMLFromTheURL = "stop"
                ''    i_p_TotalPass = i_p_TotalPass + 1
                ''End If

                ''t.Abort()

                If Len(s_p_RetrievedHtml) > 0 Then
                    f_FindAllTheHTMLFromTheURL = s_p_RetrievedHtml
                    i_p_TotalPass = i_p_TotalPass + 1
                    i_p_GotTheHtml = i_p_GotTheHtml + 1
                Else
                    s_p_RetrievedHtml = "stop"
                    f_FindAllTheHTMLFromTheURL = "stop"
                    i_p_TotalPass = i_p_TotalPass + 1
                End If

            Else
                f_FindAllTheHTMLFromTheURL = "stop"
            End If

            Call f_FindAllTheHTMLFromTheURLFU(f_FindAllTheHTMLFromTheURL)

        Catch e_lookforfiles_438 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " No html file at that url e_lookforfiles_438 --> " & e_lookforfiles_438.Message, EventLogEntryType.Information, 444)

        End Try
    End Function


    Public Function f_GetHtmlFromURL(ByVal sUrlToFind) As String
        Try
            Dim s_UrlToFind As String = sUrlToFind
            Dim b_WhileGetHtml As Boolean = True
            Dim i_CountGetHtml As Integer = 0
            If s_UrlToFind = "stop" Then
                f_GetHtmlFromURL = "stop"
                Exit Function
            End If

            If InStr(s_UrlToFind, "http://", CompareMethod.Text) <> 1 Then
                s_UrlToFind = "http://" & Trim(s_UrlToFind)
            End If

            Dim objURI As Uri = New Uri(s_UrlToFind)
            Dim objWebRequest As WebRequest = WebRequest.Create(objURI)

            'If s_p_authority = "yes" Then
            'objWebRequest.Timeout = 60000
            'Else
            objWebRequest.Timeout = 10000
            'End If

            Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
            objWebResponse.Headers.Add("Accept-Language: fr, en")

            Dim objStream As Stream = objWebResponse.GetResponseStream()
            Dim objStreamReader As StreamReader = New StreamReader(objStream)

            Dim strHTML As String = objStreamReader.ReadToEnd

            If InStr(strHTML, "Thesaurus.com", CompareMethod.Text) Then
                f_GetHtmlFromURL = strHTML
                Exit Function
            End If

            Dim s_GotHtml As String

            Do While InStr(strHTML, "  ")
                s_GotHtml = Replace(strHTML, "  ", " ")
            Loop

            s_GotHtml = Mid(strHTML, 1, i_p_MaxKPage * 1000 * 2)

            If Len(s_GotHtml) >= (i_p_MinKPage * 1000) Then
                f_GetHtmlFromURL = "f_GetHtmlFromURL yes-" & s_GotHtml
            Else
                f_GetHtmlFromURL = "f_GetHtmlFromURL didnt work!"
            End If

        Catch e_lookforfiles_490 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_490 --> " & e_lookforfiles_490.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error at e_lookforfiles_240 is " & e_lookforfiles_240.Message & " on " & s_UrlToFind)
        End Try

    End Function


    Public Function f_CheckURLInURLToLookForToIncludeThemInOurList(ByVal s_URLToLookFor)
        '/// First find the URLs
        '/// Second check them to see if we want to include them in our list
        'Call f_CheckURLToIncludeThemInOurList

    End Function


    Public Function fGetUrlFromFileInsertInDNSFOUND(ByVal FileToCrawl As String, ByVal s_DatabaseName As String) As String

        Try
            Dim iFirst As Integer
            Dim iLast As Integer
            Dim iLenght As Integer
            Dim sStoreResult As String
            Dim sAllResult As String
            Dim bKeepGoing As Boolean
            Dim sURL As String
            Dim iURLIn As Integer
            Dim sCheckStop As String
            Dim sSendstrURI As String
            Dim iGoodSpot As Integer
            Dim i_DayToRemove As Integer = 0

            bKeepGoing = True

            sURL = s_DatabaseName
            iFirst = InStr(1, FileToCrawl, "'http")
            'MessageBox.Show(iFirst, "iFirst")
            iLast = InStr(iFirst + 4, FileToCrawl, "'")
            'MessageBox.Show(iLast, "iLast")
            iLenght = iLast - iFirst

            ' MessageBox.Show(iLenght, "iLenght")

            If iLenght > 0 Then
                Do While bKeepGoing
                    sStoreResult = Mid(FileToCrawl, iFirst + 1, iLenght)
                    sStoreResult = Trim(Replace(sStoreResult, "'", ""))
                    If Len(sStoreResult) > 0 And Len(sStoreResult) < 255 Then

                        sCheckStop = CheckUrlIn_DNS(sStoreResult, s_DatabaseName)
                        'MessageBox.Show(sCheckStop, "sCheckStop")
                        If Not (sCheckStop = "stop") Then
                            InsertDNSFound(sCheckStop, s_DatabaseName, iGoodSpot, i_DayToRemove)
                            'InsertSite(sCheckStop)
                        End If
                    End If

                    sAllResult = sAllResult & Chr(10) & sStoreResult
                    'Label3.Text = sAllResult
                    'fGetUrlAlltheWeb = sAllResult

                    iFirst = InStr(iLast + 1, FileToCrawl, "'http")
                    If iFirst <= Len(FileToCrawl) Then
                        If iFirst > 0 Then
                            iLast = InStr(iFirst + 4, FileToCrawl, "'")
                        End If
                    End If

                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
                        Exit Do
                    Else
                        iLenght = iLast - iFirst
                    End If
                Loop

            End If
        Catch e_LookForFiles_568 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_568 --> " & e_LookForFiles_568.Message, EventLogEntryType.Information, 44)
            'MsgBox("e_LookForFiles_253 " & e_LookForFiles_253.Message)
        End Try

    End Function


    Public Function CheckUrlIn_DNS(Optional ByVal sUrl As String = "", Optional ByVal sDatabaseName As String = "", Optional ByVal i_goodSpot As Int64 = 0) As String

        Try
            Dim myConnString As String
            Dim sKeepString, sUrlWithoutSlash As String
            Dim s_se As String = "se"
            Dim sSendUri As String = sDatabaseName

            sUrl = Trim(sUrl)

            If Func_CheckMaxDNSToInsert(sUrl, sDatabaseName) = "stop" And i_goodSpot <= 500 Then
                If i_goodSpot <> 0 Then
                    If F_CanMakePlace(sUrl, sDatabaseName, i_goodSpot) = "stop" Then
                        CheckUrlIn_DNS = "stop"
                        Exit Function
                    End If
                Else
                    CheckUrlIn_DNS = "stop"
                    Exit Function
                End If
            End If

            If Len(sUrl) > 10 And Len(sUrl) <= 250 _
            And InStr(sUrl, "@", CompareMethod.Text) = 0 _
            And InStr(sUrl, ".css", CompareMethod.Text) = 0 _
            And InStr(sUrl, "http", CompareMethod.Text) > 0 Then

                If InStr(sUrl, "http", CompareMethod.Text) > 1 Then
                    sUrl = Mid(sUrl, InStr(sUrl, "http", CompareMethod.Text))
                End If

                Dim i_lenght As Integer = Len(sUrl)
                If InStr(i_lenght - 1, sUrl, "/", CompareMethod.Text) > 0 Then
                    sUrlWithoutSlash = Mid(sUrl, 1, i_lenght - 1)
                End If

                Dim mySelectQuery As String = "SELECT URLSite FROM DNSFOUND where URLSite ='" & sUrl & "' or URLSIte='" & sUrlWithoutSlash & "' and sCurrentQuery='" & sDatabaseName & "'"
                Dim myConnection As New OleDbConnection(ConnStringURLDNS(sSendUri))
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

                myConnection.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    sKeepString = sKeepString & (myReader.GetString(0))
                End While

                If Len(sKeepString) > 0 Then
                    CheckUrlIn_DNS = "stop"
                End If

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If
                myConnection.Close()
            Else
                CheckUrlIn_DNS = "stop"
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.720 --> " & ex.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_lookforfiles_580) is " & e_lookforfiles_580.Message)
        End Try

    End Function


    Public Function InsertDNSFound(Optional ByVal s_URLToInsert = "", Optional ByVal sDatabaseName = "", Optional ByVal iGoodSpot = "", Optional ByVal idayToRemove = "", Optional ByVal Machine = "", Optional ByVal s_searchEngine = "no", Optional ByVal i_GooglePage = 0) As Boolean

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim myInsertQuery As String = ""
        Dim b_alreadyInserted As Boolean = False

        Try

            s_URLToInsert = Replace(s_URLToInsert, "'", "/+++++\")

            s_URLToInsert = Replace(s_URLToInsert, """", "")

            s_URLToInsert = Replace(s_URLToInsert, "Char(39)", "")

            If CheckUrlIn_DNS(s_URLToInsert, sDatabaseName, iGoodSpot) <> "Stop" Then

                If Len(Machine) = 0 Then
                    Machine = "local"
                Else
                    Dim i_len As Int16 = Len(Machine)
                End If

                Dim myConnectionString As String
                Dim myConnection As New OleDbConnection(ConnStringURLDNS(sDatabaseName))
                Dim sDate As String = Now()
                Dim sCatchString As String = sDatabaseName
                Dim s_No As String = ""
                Dim s_Level As String
                Dim i_dayToRemove As Integer = 0

                Dim s_shortURL As String

                If IsNumeric(idayToRemove) Then
                    i_dayToRemove = idayToRemove
                    sDate = DateAdd(DateInterval.Day, (-1 * i_dayToRemove), Now())
                End If

                s_URLToInsert = Trim(s_URLToInsert)

                If Len(s_URLToInsert) > 10 And Len(s_URLToInsert) <= 250 _
                And InStr(s_URLToInsert, "@", CompareMethod.Text) = 0 _
                And InStr(s_URLToInsert, ".css", CompareMethod.Text) = 0 _
                And InStr(s_URLToInsert, "http", CompareMethod.Text) > 0 Then

                    If InStr(s_URLToInsert, "http", CompareMethod.Text) > 1 Then
                        s_URLToInsert = Mid(s_URLToInsert, InStr(s_URLToInsert, "http", CompareMethod.Text))
                    End If

                    s_shortURL = Func_CleanUrl(s_URLToInsert)

                    If Len(s_shortURL) < 2 Then
                        s_shortURL = "empty"
                    End If

                    If Len(s_searchEngine) = 0 Then s_searchEngine = "no"

                    b_alreadyInserted = True

                    If Len(s_searchEngine) > 2 Then

                        Form.TextBox2.Text = s_URLToInsert.ToString

                        Form.TextBox2.Update()

                        Form.Refresh()



                        Console.WriteLine(s_URLToInsert & " - " & s_searchEngine & " -- lookforfiles.insert.842 Page --> " & i_GooglePage)

                        If CheckRangeGoHere(s_URLToInsert) = True Then
                            s_No = "n"
                            myInsertQuery = "INSERT INTO DNSFound (URLSite, DateHour, Done, iGoodSpot, ShortUrl, Machine, SearchEngine,sCurrentQuery,iGooglePage) Values('" & (s_URLToInsert) & "','" & (sDate) & "','" & (s_No) & "','" & (iGoodSpot) & "','" & (s_shortURL) & "','" & Trim(Machine) & "','" & (s_searchEngine) & "','" & (sDatabaseName) & "'," & (i_GooglePage) & ")"
                        Else
                            s_No = "s"
                            myInsertQuery = "INSERT INTO DNSFound (URLSite,DateHour,Done,iGoodSpot,ShortUrl,Machine,SearchEngine,sCurrentQuery,iGooglePage) Values('" & (s_URLToInsert) & "','" & (sDate) & "','" & (s_No) & "','" & (iGoodSpot) & "','" & (s_shortURL) & "','" & Trim(Machine) & "','" & (s_searchEngine) & "','" & (sDatabaseName) & "'," & (i_GooglePage) & ")"
                        End If

                    Else

                        If CheckRangeGoHere(s_URLToInsert) = True Then
                            s_No = "n"
                            myInsertQuery = "INSERT INTO DNSFound (URLSite,DateHour,Done,iGoodSpot,ShortUrl,Machine,sCurrentQuery,iGooglePage) Values('" & (s_URLToInsert) & "','" & (sDate) & "','" & (s_No) & "','" & (iGoodSpot) & "','" & (s_shortURL) & "','" & Trim(Machine) & "','" & (sDatabaseName) & "'," & (i_GooglePage) & ")"
                        Else
                            s_No = "s"
                            myInsertQuery = "INSERT INTO DNSFound (URLSite,DateHour,Done,iGoodSpot,ShortUrl,Machine,sCurrentQuery,iGooglePage) Values('" & (s_URLToInsert) & "','" & (sDate) & "','" & (s_No) & "','" & (iGoodSpot) & "','" & (s_shortURL) & "','" & Trim(Machine) & "','" & (sDatabaseName) & "'," & (i_GooglePage) & ")"
                        End If

                    End If

                    Dim myCommand As New OleDbCommand(myInsertQuery)
                    myCommand.Connection = myConnection
                    myConnection.Open()
                    myCommand.ExecuteNonQuery()
                    myCommand.Connection.Close()
                    myConnection.Close()

                End If
                InsertDNSFound = True
            Else
                InsertDNSFound = False
            End If

            If b_alreadyInserted = False Then
                Console.WriteLine(s_URLToInsert & " - " & s_searchEngine & " -- lookforfiles.Already_inserted.878 Page --> " & i_GooglePage)
            End If

        Catch ex As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.818 --> " & ex.Message & " -- " & myInsertQuery, EventLogEntryType.Information, 44)
            EventLog.WriteEntry(sSource, " lookforfiles.819--> " & ex.Message & " ---- s_URLToInsert -" & s_URLToInsert & " ---- sDatabaseName -" & sDatabaseName & " ---- iGoodSpot -" & iGoodSpot & " ---- idayToRemove -" & idayToRemove & " ---- Machine -" & Machine & " ---- s_searchEngine -" & s_searchEngine, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public b_p_freshPages As Boolean

    Function CheckUrlIn(Optional ByVal SUrl As String = "", Optional ByVal sSendstrURI As String = "", Optional ByVal i_findSignature As Int64 = 0, Optional ByVal i_findKeywords As Int64 = 0) As String
        Dim mySelectQuery As String
        Dim sKeepString As String
        Dim sSendUri As String = sSendstrURI
        Dim i_keywordsIn As Int64
        Dim i_signature As Int64
        b_p_freshPages = False

        Try

            '////----First stage of the motor of selection

            If Len(SUrl) > 10 And InStr(SUrl, "'", CompareMethod.Text) = 0 Then

                Dim myConnString As String

                Dim s_shortURL As String = Func_CleanUrl(SUrl)

                If IsNumeric(i_findSignature) Then
                    i_findSignature = CInt(i_findSignature)
                Else
                    i_findSignature = 0
                End If

                mySelectQuery = "SELECT url FROM URLFOUND where  GoodSpot <> -40004 and ShortUrl ='" & s_shortURL & "' and Signature = " & i_findSignature & " and Keywords = " & i_findKeywords & " AND FreshPage <> 'ar'"

                Dim myConnection As New OleDbConnection(ConnStringURL(sSendUri))
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

                myConnection.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    sKeepString = myReader("url")
                End While

                If Len(sKeepString) > 0 Then
                    CheckUrlIn = "stop"
                Else
                    If Len(SUrl) > 0 Then
                        CheckUrlIn = Trim(SUrl)
                    Else
                        CheckUrlIn = "stop"
                    End If
                End If

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()

            Else
                CheckUrlIn = "stop"
            End If

            If CheckUrlIn = "stop" Then
                Exit Function
            End If

            '////----Second stage of the selection motor

            Dim b_urlIsNotGood As Boolean = False
            Dim i_keywords As Int32
            Dim i_counter As Int32

            If Len(SUrl) > 10 Then

                mySelectQuery = "SELECT TOP 1 URL, Keywords, Signature FROM URLFOUND where  GoodSpot <> -40004 and url ='" & Trim(SUrl) & "' and sQueryUrlFound='" & sSendstrURI & "' order by IDCounter desc"

                Dim myConnection As New OleDbConnection(ConnStringURL(sSendUri))
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

                myConnection.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    If IsNumeric(myReader("Keywords")) Then
                        i_keywords = myReader("Keywords")
                    End If
                    If IsNumeric(myReader("Signature")) Then
                        i_signature = myReader("Signature")
                    End If
                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()

                Dim i_handicap As Int64 = 2
                Dim i_handicapBottom, i_handicapTop As Int64

                'i_handicap = i_handicap + i_keywords
                'i_handicapBottom = i_signature - i_handicap
                'i_handicapTop = i_signature + i_handicap

                'i_handicap = i_handicap + i_keywords
                i_handicapBottom = i_signature - i_handicap
                i_handicapTop = i_signature + i_handicap

                If CInt(i_keywords) <> CInt(i_findKeywords) Then
                    b_p_freshPages = True
                ElseIf CInt(i_keywords) = CInt(i_findKeywords) And (CInt(i_findSignature) < CInt(i_handicapBottom) Or CInt(i_findSignature) > CInt(i_handicapTop)) Then
                    b_p_freshPages = True
                Else
                    b_p_freshPages = False
                End If

            Else
                CheckUrlIn = "stop"
            End If

        Catch e_lookforfiles_723 As Exception
            CheckUrlIn = "stop"
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_723 --> " & e_lookforfiles_723.Message, EventLogEntryType.Information, 44)
            EventLog.WriteEntry(sSource, " lookforfiles.e_mySelectQuery_699 --> " & mySelectQuery, EventLogEntryType.Information, 44)
        End Try

    End Function

    Dim b_p_goAheadWithInsert As Boolean


    Dim i_p_counterToGC As Int32

    Function CheckDNAIn(ByVal SUrl As String, ByVal sSendstrURI As String) As String
        Dim mySelectQuery As String
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try

            i_p_counterToGC = i_p_counterToGC + 1

            'If i_p_counterToGC >= 500 Then
            'GC.Collect()
            'GC.WaitForPendingFinalizers()
            'i_p_counterToGC = 0
            'End If


            If Len(SUrl) > 0 And Len(sSendstrURI) > 0 Then

                b_p_goAheadWithInsert = False
                Dim i_lenSUrl As Integer = Len(SUrl)
                Dim i_endSlash As Integer = InStr(i_lenSUrl, SUrl, "/", CompareMethod.Text)

                If i_lenSUrl = i_endSlash Then
                    SUrl = Mid(SUrl, 1, (i_lenSUrl - 1))
                End If

                Dim myConnString As String
                Dim sKeepString As String
                Dim date_InsertURL As DateTime
                Dim sSendUri As String = sSendstrURI

                mySelectQuery = "SELECT url,DateHour FROM ScratchPad where url ='" & Trim(SUrl) & "'"

                Dim myConnection As New OleDbConnection(ConnStringDNA())
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

                myConnection.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    sKeepString = sKeepString & (myReader.GetString(0))
                    date_InsertURL = myReader("DateHour")
                    Exit While
                End While

                If Len(sKeepString) > 0 Then

                    Dim s_string As String = SUrl
                    CheckDNAIn = "stop"
                    '//---- Check when have been inserted
                    Dim date_ofLastCycle As DateTime
                    Dim d_KeepDuration As Double
                    Dim d_MinuteSinceInsert, d_timeSinceLastCycle, d_keepTheWholePart, d_keepFraction,
                    d_minuteSinceCycle As Double

                    d_KeepDuration = i_p_Duration * 60

                    d_MinuteSinceInsert = DateDiff(DateInterval.Minute, d_p_DateTime, Now)

                    d_timeSinceLastCycle = d_MinuteSinceInsert / d_KeepDuration

                    d_keepTheWholePart = Fix(d_timeSinceLastCycle)

                    d_keepFraction = d_timeSinceLastCycle - d_keepTheWholePart

                    d_minuteSinceCycle = d_keepFraction * d_KeepDuration

                    date_ofLastCycle = DateAdd(DateInterval.Minute, (d_minuteSinceCycle * -1), Now())

                    If date_InsertURL < date_ofLastCycle Then
                        b_p_goAheadWithInsert = True
                    End If
                    '\\---- End of check when have been inserted

                Else

                    CheckDNAIn = Trim(SUrl)

                End If

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                ''myConnection = Nothing

                myConnection.Close()

            Else

                CheckDNAIn = "stop"
                'EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_733 SUrl--> " & SUrl & "   sSendstrURI --" & sSendstrURI, EventLogEntryType.Information, 49)

            End If
        Catch e_lookforfiles_733 As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_733 --> " & e_lookforfiles_733.Message, EventLogEntryType.Information, 44)
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_733 --> " & mySelectQuery, EventLogEntryType.Information, 44)

        End Try

    End Function


    Function InsertURLFOUND(ByVal SUrlInsert As String, ByVal sSendstrURI As String, ByVal sFoundRacine As String, ByVal i_foundSignature As Int64, ByVal s_pageContent As String)

        Dim d_date As DateTime
        Dim d_hour As DateTime
        Dim d_dateHours As DateTime
        Dim i_keepHour As Integer
        Dim s_retrieveIDDateHour As Array
        Dim s_Now As String = Now()
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        'Dim s_pageContent As String
        Dim i_goodSpotDNSFOUND As Int64
        Dim i_goodSpotSentOver As Int64

        d_dateHours = d_p_DateTimeT

        Try
            SUrlInsert = Replace(SUrlInsert, "'", "/+++++\")
            'Call InsertDataInURLFOUND(sSendstrURI)
            Dim sToInsert As String = SUrlInsert
            Dim myConnectionString As String
            Dim myConnection As New OleDbConnection(ConnStringURL(sSendstrURI))
            Dim sDate As String = Now()
            Dim sNo As String
            Dim i_foundKeywords As Long = 0
            Dim s_splitSignatureKeywords As String
            Dim s_arraySignKeywords() As String
            Dim i_goodSpotFromUrlFound As Int32
            Dim b_inputThatNews As Boolean = False
            Dim b_addMoreBonus As Boolean = False
            Dim i_ID As Integer

            sNo = "n"

            s_splitSignatureKeywords = CStr(i_foundSignature)
            s_splitSignatureKeywords = Replace(s_splitSignatureKeywords, "00000", " ")
            s_arraySignKeywords = Split(s_splitSignatureKeywords)

            If UBound(s_arraySignKeywords) = 1 Then
                i_foundSignature = CLng(s_arraySignKeywords(1))
                i_foundKeywords = CLng(s_arraySignKeywords(0))
            Else
                i_foundSignature = s_arraySignKeywords(0)
                i_foundKeywords = 0
            End If

            If CheckUrlIn(sToInsert, sSendstrURI, i_foundSignature, i_foundKeywords) <> "stop" Then
                b_inputThatNews = True
            End If

            If b_inputThatNews Then
                If Len(sToInsert) > 10 And Len(sToInsert) <= 250 Then
                    Dim s_freshPage As String = "n"
                    Dim s_local As String = "loc"
                    Dim s_Verb As String = "s_ID"

                    i_ID = f_TakeRacine_DNS_ID(sSendstrURI, sFoundRacine, s_Verb)

                    If Not (IsNumeric(i_p_SendGoodSpotOverToUrlFound)) Then
                        i_p_SendGoodSpotOverToUrlFound = 0
                    Else
                        i_goodSpotSentOver = i_p_SendGoodSpotOverToUrlFound
                        i_p_SendGoodSpotOverToUrlFound = 0
                    End If

                    d_date = GetDateFromContent(s_pageContent)
                    sDate = CStr(d_date)

                    ' If s_p_authority = "yes" Then
                    s_pageContent = RemoveJSTag(s_pageContent)
                    s_pageContent = PutPipeAroundTitle(s_pageContent)
                    s_pageContent = RemoveStyle(s_pageContent)
                    s_pageContent = RemoveHtmlTag(s_pageContent)
                    s_pageContent = RemoveAccent(s_pageContent)
                    s_pageContent = keepASCIIOnly(s_pageContent)
                    s_pageContent = PutPipeAroundNewContent(s_pageContent, sToInsert, sSendstrURI)
                    'Else
                    '    s_pageContent = ""
                    'End If

                    i_goodSpotSentOver = GetGoodSpotFromURLFOUND(sToInsert, i_goodSpotSentOver, sSendstrURI)
                    i_goodSpotDNSFOUND = i_goodSpotSentOver

                    If b_p_newContentFound And b_p_freshPages = True And b_p_dateFoundToday = True Then
                        s_freshPage = "yz"
                        b_addMoreBonus = True
                    ElseIf b_p_freshPages = True And b_p_dateFoundToday = True Then
                        s_freshPage = "yy"
                        b_addMoreBonus = True
                    ElseIf b_p_freshPages = True Then
                        s_freshPage = "y"
                    Else
                        s_freshPage = "n"
                    End If

                    b_p_freshPages = False
                    b_p_dateFoundToday = False

                    Form.UPDATEURLTB3.Text = sToInsert.ToString
                    Form.FreshPageTB5.Text = s_freshPage.ToString

                    Form.UPDATEURLTB3.Update()
                    Form.FreshPageTB5.Update()

                    Form.TBScratchPad.Text = ""
                    Form.INSERTURLFOUNDTB4.Text = ""
                    Form.TB40004.Text = ""
                    'Form.UPDATEURLTB3.Text = ""
                    Form.RichTextBox1.Text = ""
                    Form.TBLenghtPage.Text = ""
                    Form.TBKeywords.Text = ""
                    Form.TBSignature.Text = ""
                    Form.FreshPageTB6.Text = ""
                    Form.TextBox40004.Text = ""
                    Form.FreshPageTB5.Text = ""
                    Form.TBScratchPad.Update()
                    Form.INSERTURLFOUNDTB4.Update()
                    Form.TB40004.Update()
                    'Form.UPDATEURLTB3.Update()
                    Form.RichTextBox1.Update()
                    Form.TBLenghtPage.Update()
                    Form.TBKeywords.Update()
                    Form.TBSignature.Update()
                    Form.FreshPageTB6.Update()
                    Form.TextBox40004.Update()
                    Form.FreshPageTB5.Update()

                    Form.Refresh()

                    Dim s_archive As String = "ar"

                    Dim b_p_notfirstTime As Boolean = False
                    Dim myConn As String = "UPDATE URLFOUND SET FreshPage= '" & (s_archive) & "' WHERE URL ='" & sToInsert & "' and sQueryUrlFound='" & sSendstrURI & "'"
                    Dim myCommand2 As New OleDbCommand(myConn)
                    myCommand2.Connection = myConnection
                    myConnection.Open()
                    myCommand2.ExecuteNonQuery()
                    myCommand2.Connection.Close()

                    Form.INSERTURLFOUNDTB4.Text = sToInsert.ToString
                    Form.FreshPageTB6.Text = s_freshPage.ToString
                    Form.RichTextBox1.Text = s_pageContent.ToString
                    Form.TBSignature.Text = i_foundSignature.ToString
                    Form.TBKeywords.Text = i_foundKeywords.ToString

                    Form.INSERTURLFOUNDTB4.Update()
                    Form.FreshPageTB6.Update()
                    Form.RichTextBox1.Update()
                    Form.TBSignature.Update()
                    Form.TBKeywords.Update()

                    Form.TBScratchPad.Text = ""
                    'Form.INSERTURLFOUNDTB4.Text = ""
                    Form.TB40004.Text = ""
                    Form.UPDATEURLTB3.Text = ""
                    'Form.RichTextBox1.Text = ""
                    Form.TBLenghtPage.Text = ""
                    'Form.TBKeywords.Text = ""
                    'Form.TBSignature.Text = ""
                    'Form.FreshPageTB6.Text = ""
                    Form.TextBox40004.Text = ""
                    Form.FreshPageTB5.Text = ""
                    Form.TBScratchPad.Update()
                    'Form.INSERTURLFOUNDTB4.Update()
                    Form.TB40004.Update()
                    Form.UPDATEURLTB3.Update()
                    'Form.RichTextBox1.Update()
                    Form.TBLenghtPage.Update()
                    'Form.TBKeywords.Update()
                    'Form.TBSignature.Update()
                    ''Form.FreshPageTB6.Update()
                    Form.TextBox40004.Update()
                    Form.FreshPageTB5.Update()

                    Form.Refresh()


                    Dim s_shortURL As String = Func_CleanUrl(sToInsert)
                    Dim s_machineURL As String = "1"
                    Dim myInsertQuery As String = "INSERT INTO URLFOUND (ID,url,DateHour,GoodSpot,Done,Machine,Signature,Keywords,ShortURL,FreshPage,URL_PAGE,sQueryUrlFound)" _
                    & "Values(" & i_ID & ",'" & Trim(sToInsert) & "','" & (sDate) & "'," & i_goodSpotSentOver & ",'" & (sNo) & "','" & (s_machineURL) & "'," & i_foundSignature & "," & i_foundKeywords & ",'" & s_shortURL & "','" & s_freshPage & "','" & s_pageContent & "','" & sSendstrURI & "')"

                    Dim myCommand As New OleDbCommand(myInsertQuery)
                    myCommand.Connection = myConnection
                    myConnection.Open()
                    myCommand.ExecuteNonQuery()
                    myCommand.Connection.Close()
                    myConnection.Close()

                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")
                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")
                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")
                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")
                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")
                    Console.WriteLine(" myInsertQuery --> " & myInsertQuery)
                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")
                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")
                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")
                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")
                    Console.WriteLine("_______*_________*____________*_____________*______________*______________")

                End If

                If Len(sToInsert) >= 11 Then
                    If i_goodSpotDNSFOUND = 0 Then
                        i_goodSpotDNSFOUND = 400
                    End If

                    If b_addMoreBonus Then
                        AddBonus(sSendstrURI, i_ID, i_goodSpotSentOver)
                    End If

                    If i_goodSpotSentOver = 0 Then
                        Call InsertDNSFound(sToInsert, sSendstrURI, i_goodSpotDNSFOUND + 100)
                        Call f_Extract_Racine_URL(sToInsert, sSendstrURI, i_goodSpotDNSFOUND)
                    End If

                End If

            End If


        Catch ex As Exception
            CompacAccessUrlFound(sSendstrURI)
            EventLog.WriteEntry(sSource, " lookforfiles.1156 --> " & ex.Message, EventLogEntryType.Information, 44)
            Console.WriteLine(vbNewLine)
            Console.WriteLine("_______*_________*____________*_____________*______________*______________")
            Console.WriteLine(" lookforfiles.1116 --> " & ex.Message)
            Console.WriteLine("_______*_________*____________*_____________*______________*______________")
        End Try

    End Function

    Sub AddBonus(ByVal sSendstrURI As String, ByVal i_ID As Int64, ByVal i_goodSpotSentOver As Int64)
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try
            Dim myConnection As New OleDbConnection(ConnStringURLDNS(sSendstrURI))
            i_goodSpotSentOver = i_goodSpotSentOver + 2
            Dim myString As String = "UPDATE DNSFOUND SET iGoodSpot =  '" & (i_goodSpotSentOver) & "'  WHERE ID =" & i_ID & ""
            Dim myCommand2 As New OleDbCommand(myString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myConnection.Close()

        Catch ex As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.1237 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    '    Function InsertURLFOUNDFromBroadcast(Optional ByVal SUrlInsert = "", Optional ByVal sSendstrURI = "",
    '    Optional ByVal ID = "", Optional ByVal smachine = "", Optional ByVal signature = "", Optional ByVal i_goodSpotFromOver = 0)
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        Dim i_ID As Integer = CInt(ID)
    '        Dim d_date As DateTime
    '        Dim d_hour As DateTime
    '        Dim d_dateHours As DateTime
    '        Dim i_keepHour, i_findSignature As Int32
    '        Dim s_retrieveIDDateHour As Array
    '        Dim s_Now As String = Now()
    '        Dim sToInsert, s_splitSignatureKeywords As String
    '        Dim s_keepArray() As String = Split(SUrlInsert, "-*-*-")
    '        Dim s_arraySignKeywords() As String
    '        Dim i_foundSignature As Long = 0
    '        Dim i_foundKeywords As Long = 0
    '        Dim s_pageContent As String
    '        Dim i_goodSpotFromUrlFound As Int32
    '        Dim b_inputThatNews As Boolean = False

    '        Try
    '            SUrlInsert = Replace(SUrlInsert, "'", "/+++++\")

    '            If Not (IsNumeric(CLng(signature))) Then
    '                signature = 0
    '            End If

    '            s_splitSignatureKeywords = CStr(signature)
    '            s_splitSignatureKeywords = Replace(s_splitSignatureKeywords, "00000", " ")
    '            s_arraySignKeywords = Split(s_splitSignatureKeywords)

    '            If UBound(s_arraySignKeywords) = 1 Then

    '                i_foundSignature = CLng(s_arraySignKeywords(1))

    '                If s_arraySignKeywords(0) = " " Or Len(s_arraySignKeywords(0)) = 0 Then
    '                Else
    '                    i_foundKeywords = CLng(s_arraySignKeywords(0))
    '                End If

    '            Else
    '                i_foundSignature = s_arraySignKeywords(0)
    '            End If

    '            sSendstrURI = Replace(sSendstrURI, "_", " ")

    '            d_dateHours = d_p_DateTimeT

    '            'If s_p_authority = "yes" Then

    '            sToInsert = s_keepArray(0)

    '            Dim s_shortURL As String = Func_CleanUrl(sToInsert)
    '            Dim myConnectionString As String
    '            Dim myConnection As New OleDbConnection(ConnStringURL(sSendstrURI))
    '            Dim sDate As String
    '            Dim sNo As String
    '            Dim s_freshPage As String

    '            i_keepHour = DateDiff(DateInterval.Hour, d_dateHours, Now)

    '            'If i_keepHour < 2 * i_p_Duration Then
    '            '    sNo = "y"
    '            'Else
    '            '    sNo = "n"
    '            'End If

    '            Dim s_verb As String = "AcknowledgeInsertDocumentFound"

    '            If CheckUrlIn(sToInsert, sSendstrURI, i_foundSignature, i_foundKeywords) <> "stop" Then
    '                b_inputThatNews = True
    '            Else
    '                'Call LastCallVerification(sToInsert, sSendstrURI, i_foundSignature, i_foundKeywords)
    '            End If

    '            If b_inputThatNews Then
    '                If Len(sToInsert) > 10 And Len(sToInsert) <= 250 Then

    '                    sToInsert = Replace(sToInsert, "/////*\\\\\", "_")

    '                    s_pageContent = f_FindAllTheHTMLFromTheURL(sToInsert, sSendstrURI)

    '                    d_date = GetDateFromContent(s_pageContent)
    '                    sDate = CStr(d_date)
    '                    s_pageContent = RemoveJSTag(s_pageContent)
    '                    s_pageContent = PutPipeAroundTitle(s_pageContent)
    '                    s_pageContent = RemoveStyle(s_pageContent)
    '                    s_pageContent = RemoveHtmlTag(s_pageContent)
    '                    s_pageContent = RemoveAccent(s_pageContent)
    '                    s_pageContent = keepASCIIOnly(s_pageContent)
    '                    s_pageContent = PutPipeAroundNewContent(s_pageContent, sToInsert, sSendstrURI)
    '                    i_goodSpotFromUrlFound = GetGoodSpotFromURLFOUND(sToInsert, i_goodSpotFromOver, sSendstrURI)

    '                    If b_p_newContentFound And b_p_freshPages = True And b_p_dateFoundToday = True Then
    '                        s_freshPage = "yz"
    '                        'b_addMoreBonus = True
    '                    ElseIf b_p_freshPages = True And b_p_dateFoundToday = True Then
    '                        s_freshPage = "yy"
    '                        'b_addMoreBonus = True
    '                    ElseIf b_p_freshPages = True Then
    '                        s_freshPage = "y"
    '                    Else
    '                        s_freshPage = "n"
    '                    End If

    '                    b_p_freshPages = False
    '                    b_p_dateFoundToday = False

    '                    Dim s_archive As String = "ar"
    '                    Dim myConn As String = "UPDATE URLFOUND SET FreshPage= '" & (s_archive) & "' WHERE URL ='" & sToInsert & "'"
    '                    Dim myCommand2 As New OleDbCommand(myConn)
    '                    myCommand2.Connection = myConnection
    '                    myConnection.Open()
    '                    myCommand2.ExecuteNonQuery()
    '                    myCommand2.Connection.Close()

    '                    Dim myInsertQuery As String = "INSERT INTO URLFOUND (ID,url,DateHour,GoodSpot,Done,Machine,Signature,Keywords,ShortURL,FreshPage, URL_PAGE)" _
    '                      & "Values(" & i_ID & ",'" & Trim(sToInsert) & "','" & (sDate) & "'," & i_goodSpotFromUrlFound & ",'" & (sNo) & "','" & (smachine) & "'," & i_foundSignature & "," & i_foundKeywords & ",'" & s_shortURL & "','" & s_freshPage & "','" & s_pageContent & "')"

    '                    Dim myCommand As New OleDbCommand(myInsertQuery)
    '                    myCommand.Connection = myConnection
    '                    myConnection.Open()
    '                    myCommand.ExecuteNonQuery()
    '                    myCommand.Connection.Close()
    '                    '''myConnection = Nothing
    '                    myConnection.Close()

    '                End If
    '            End If

    '            If s_p_authority = "yes" Then
    '                Dim s_toSendAcknowledgement As String = sToInsert & "|||||" & sSendstrURI & "|||||" & smachine
    '                s_toSendAcknowledgement = Replace(s_toSendAcknowledgement, "_", "!!!!!")
    '                Call BroadCastDNSFOUND(s_verb, s_toSendAcknowledgement)
    '                'EventLog.WriteEntry(sSource, " InsertURL from other machine lookforfiles.s_toSendAcknowledgement_980 --> " & s_toSendAcknowledgement, EventLogEntryType.Information, 778)
    '                'MsgBox("s_toSendAcknowledgement_lookforfiles_992  " & s_verb & " - " & s_toSendAcknowledgement) 
    '            End If

    '            'End If

    '        Catch e_lookforfiles_965 As Exception
    '            CompacAccessUrlFound(sSendstrURI)
    '            EventLog.WriteEntry(sSource, " InsertURL from other machine lookforfiles.e_lookforfiles_965 --> " & e_lookforfiles_965.Message, EventLogEntryType.Information, 33)
    '            Console.WriteLine(vbNewLine)
    '            Console.WriteLine("_______*_________*____________*_____________*______________*______________")
    '            Console.WriteLine(" e_lookforfiles_965 --> " & e_lookforfiles_965.Message)
    '        End Try

    '    End Function

    Public Declare Function QueryPerformanceCounter _
       Lib "kernel32" Alias "QueryPerformanceCounter" _
       (ByRef lpPerformanceCount As Long) As Integer

    Public Declare Function QueryPerformanceFrequency _
       Lib "kernel32" Alias "QueryPerformanceFrequency" _
       (ByRef lpPerformanceCount As Long) As Integer

    Function InsertDNA(ByVal SUrlToInsert As String, ByVal sDatabaseName As String,
        ByVal sLevel As String) As String
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try

            SUrlToInsert = Replace(SUrlToInsert, "'", "/+++++\")

            Dim s_ToInsert As String = SUrlToInsert
            Dim s_DatabaseName As String = sDatabaseName
            Dim myConnectionString As String
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim s_Date As String = Now()
            Dim s_No As String = "n"
            Dim s_Level As String = "n"
            Dim i_CountPage As Integer = 0

            '----> If Not (InStr(s_ToInsert, "@", CompareMethod.Text) > 0) Then
            If Len(s_ToInsert) > 10 And Len(s_ToInsert) <= 250 _
            And InStr(s_ToInsert, "@", CompareMethod.Text) = 0 _
            And InStr(s_ToInsert, ".amazon.", CompareMethod.Text) = 0 _
            And InStr(s_ToInsert, ".css", CompareMethod.Text) = 0 _
            And InStr(s_ToInsert, "http", CompareMethod.Text) > 0 Then

                If InStr(12, s_ToInsert, "html", CompareMethod.Text) > 0 Or
                   InStr(s_ToInsert, "http", CompareMethod.Text) > 0 Or
                   InStr(12, s_ToInsert, ".htm", CompareMethod.Text) > 0 Or
                   InStr(12, s_ToInsert, ".asp", CompareMethod.Text) > 0 Or
                   InStr(12, s_ToInsert, ".php", CompareMethod.Text) > 0 Or
                   InStr(12, s_ToInsert, ".jsp", CompareMethod.Text) > 0 Or
                   InStr(12, s_ToInsert, ".cfm", CompareMethod.Text) > 0 Or
                   InStr(12, s_ToInsert, ".aspx", CompareMethod.Text) > 0 Or
                   CDbl(InStr(Len(s_ToInsert) - 1, s_ToInsert, "/", CompareMethod.Text)) > 0 Then

                    If InStr(s_ToInsert, "http", CompareMethod.Text) > 1 Then
                        s_ToInsert = Mid(s_ToInsert, InStr(s_ToInsert, "http", CompareMethod.Text))
                    End If

                    If CheckDNAIn(s_ToInsert, sDatabaseName) <> "stop" Then
                        s_ToInsert = Replace(s_ToInsert, "target=_top", "")

                        '//---- Remove end slash
                        Dim i_lenSUrl As Integer = Len(s_ToInsert)
                        Dim i_endSlash As Integer = InStr(i_lenSUrl, s_ToInsert, "/", CompareMethod.Text)

                        If i_lenSUrl = i_endSlash Then
                            s_ToInsert = Mid(s_ToInsert, 1, (i_lenSUrl - 1))
                        End If
                        '\\---- end of removing

                        If Len(s_ToInsert) > 4 Then

                            Dim s_shortUrl As String = Func_CleanUrl(s_ToInsert)

                            If Len(s_shortUrl) > 0 Then

                                If Funct_PageCounter(s_ToInsert, sDatabaseName, s_No) <> "stop" Then

                                    Dim frq, counter1, counter2 As Long
                                    Dim duration As Double

                                    QueryPerformanceFrequency(frq)
                                    QueryPerformanceCounter(counter1)

                                    Dim myInsertQuery As String = "INSERT INTO SCRATCHPAD (Url,DateHour,Done,ShortUrl)" _
                                    & "Values ('" & Trim(s_ToInsert) & "','" & (s_Date) & "','" & (s_No) & "','" & (s_shortUrl) & "')"

                                    Dim myCommand As New OleDbCommand(myInsertQuery)
                                    myCommand.Connection = myConnection
                                    myConnection.Open()
                                    myCommand.ExecuteNonQuery()
                                    myCommand.Connection.Close()

                                    myConnection.Close()

                                    QueryPerformanceCounter(counter2)
                                    duration = (counter2 - counter1) / frq

                                    Console.WriteLine(duration & "  -- duration.lookforfiles.InsertDNA.1379")
                                    InsertDNA = "true"

                                    Call InsertIntoSratchPadFU(myInsertQuery)
                                Else
                                    InsertDNA = "stop"
                                End If
                            End If

                            'ElseIf Len(s_shortUrl) > 0 Then

                            '    s_No = "br"
                            '    If Funct_PageCounter(s_ToInsert, sDatabaseName, s_No) <> "stop" Then

                            '        Dim myInsertQuery As String = "INSERT INTO SCRATCHPAD (Url,DateHour,Done,ShortUrl)" _
                            '         & "Values ('" & Trim(s_ToInsert) & "','" & (s_Date) & "','" & (s_No) & "','" & (s_shortUrl) & "')"

                            '        'EventLog.WriteEntry(sSource, " lookforfiles.InsertDNA.1115 --> " & myInsertQuery, EventLogEntryType.Information, 8899)
                            '        'EventLog.WriteEntry(sSource, " InsertDNA.1115 s_ToInsert --> " & s_ToInsert & " sDatabaseName --> " & sDatabaseName, EventLogEntryType.Information, 8899)
                            '        'Console.WriteLine(" InsertDNA.1115 s_ToInsert --> " & s_ToInsert & " sDatabaseName --> " & sDatabaseName)
                            '        If Len(s_ToInsert) > 0 And Len(sDatabaseName) > 0 Then


                            '            Dim myCommand As New OleDbCommand(myInsertQuery)
                            '            myCommand.Connection = myConnection
                            '            myConnection.Open()
                            '            myCommand.ExecuteNonQuery()
                            '            myCommand.Connection.Close()
                            '            'Connection = Nothing
                            '            myConnection.Close()
                            '            InsertDNA = "true"
                            '        Else
                            '            InsertDNA = "stop"
                            '            'EventLog.WriteEntry(sSource, " InsertDNA.1115 s_ToInsert --> " & s_ToInsert & " sDatabaseName --> " & sDatabaseName, EventLogEntryType.Information, 8891)
                            '        End If
                            '    End If

                            'End If

                        End If


                Else
                        'if bigger than present cycle do the update
                        If b_p_goAheadWithInsert Then

                            'myConnectionString = "UPDATE SCRATCHPAD SET Done='" & (s_No) & "', DateHour =#" & Now() & "#  where Url = '" & (s_ToInsert) & "'"
                            myConnectionString = "UPDATE SCRATCHPAD SET Done='" & (s_No) & "', DateHour ='" & Now() & "'  where Url = '" & (s_ToInsert) & "'"
                            'MessageBox.Show(myConnectionString, " -> lookforfiles.myConnectionString.1864")
                            Dim myCommand2 As New OleDbCommand(myConnectionString)
                            myCommand2.Connection = myConnection
                            myConnection.Open()
                            myCommand2.ExecuteNonQuery()
                            myCommand2.Connection.Close()
                            '''myConnection = Nothing
                            myConnection.Close()
                            InsertDNA = "true"
                            b_p_goAheadWithInsert = False

                        End If

                    End If

                End If

            End If
            '----> End If

        Catch e_lookforfiles_944 As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_944 --> " & e_lookforfiles_944.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_lookforfiles_532) is " & e_lookforfiles_532.Message)
        End Try

    End Function


    '    Sub f_ThrottleCPU()
    '        Try
    '            Dim objectName As String = "processor"
    '            Dim counterName As String = "% Processor Time"
    '            Dim instanceName As String = "_Total"
    '            Dim i_Counter As Integer
    '            Dim s_KeepNextSample As String
    '            Dim i_KeepNextValue As Integer
    '            Dim i_KeepTimeBeenThere As Integer
    '            Dim i_KeepTimeInside As Integer

    '            Dim ccounter As PerformanceCounter
    '            ccounter = New PerformanceCounter(objectName, counterName, instanceName)

    '            Dim counter As PerformanceCounter
    '            counter = New PerformanceCounter(objectName, counterName, instanceName)


    '            Dim cs1 = New CounterSample
    '            Dim cs2 = New CounterSample
    '            'float(result)
    '            Dim pc1 = New PerformanceCounter
    '            pc1.CategoryName = "Processor"
    '            pc1.CounterName = "% Processor Time"
    '            pc1.InstanceName = "_Total"
    '            pc1.MachineName = "."
    '            cs1 = pc1.NextSample()
    '            System.Threading.Thread.Sleep(50)
    '            Dim result As Integer = CounterSample.Calculate(cs1, pc1.NextSample())
    '            Dim sresult As String = CStr(result)

    '            'MsgBox(result & " result 560")
    '            'MsgBox(counter.NextValue.ToString & " bla bla bla 35")

    '            If CInt(result) > 95 Then
    '                System.Threading.Thread.Sleep(2500)
    '                '   MsgBox("95%")
    '                i_KeepTimeInside = i_KeepTimeInside + 1
    '            End If

    '            If CInt(result) > 70 Then
    '                System.Threading.Thread.Sleep(1000)
    '                i_KeepTimeInside = i_KeepTimeInside + 1
    '            End If

    '            s_KeepNextSample = ccounter.NextSample.ToString
    '            i_KeepNextValue = ccounter.NextValue.ToString
    '        Catch e_lookforfiles_1191 As Exception
    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_1191 --> " & e_lookforfiles_1191.Message, EventLogEntryType.Information, 44)
    '            'MsgBox("The error (e_lookforfiles_532) is " & e_lookforfiles_532.Message)
    '        End Try
    '    End Sub


    '    Public Sub AdjustThreadPriority()
    '        Try
    '            Dim tThread As Thread

    '            Dim s_threadName As String = tThread.CurrentThread.Name



    '            If LCase(Trim(s_p_Priority)) = "lowest" Then
    '                tThread.CurrentThread.Priority = Lowest
    '            End If
    '            If LCase(Trim(s_p_Priority)) = "normal" Then
    '                tThread.CurrentThread.Priority = Normal
    '            End If
    '            If LCase(Trim(s_p_Priority)) = "above normal" Then
    '                tThread.CurrentThread.Priority = AboveNormal
    '            End If
    '            If LCase(Trim(s_p_Priority)) = "below normal" Then
    '                tThread.CurrentThread.Priority = BelowNormal
    '            End If
    '            If LCase(Trim(s_p_Priority)) = "highest" Then
    '                tThread.CurrentThread.Priority = Highest
    '            End If


    '        Catch e_lookforfiles_1224 As Exception
    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_1224 --> " & e_lookforfiles_1224.Message, EventLogEntryType.Information, 44)
    '            'MsgBox("The error (e_lookforfiles_532) is " & e_lookforfiles_532.Message)
    '        End Try
    '    End Sub


    '    Public Function ThreadPriority(ByVal s_p_priority) As Integer
    '        Try
    '            Dim tThread As Thread

    '            Dim s_threadName As String = tThread.CurrentThread.Name

    '            If LCase(Trim(s_p_priority)) = "lowest" Then
    '                tThread.CurrentThread.Priority = Lowest
    '            End If
    '            If LCase(Trim(s_p_priority)) = "normal" Then
    '                tThread.CurrentThread.Priority = Normal
    '            End If
    '            If LCase(Trim(s_p_priority)) = "above normal" Then
    '                tThread.CurrentThread.Priority = AboveNormal
    '            End If
    '            If LCase(Trim(s_p_priority)) = "below normal" Then
    '                tThread.CurrentThread.Priority = BelowNormal
    '            End If
    '            If LCase(Trim(s_p_priority)) = "highest" Then
    '                tThread.CurrentThread.Priority = Highest
    '            End If

    '        Catch e_lookforfiles_1262 As Exception
    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_1262 --> " & e_lookforfiles_1262.Message, EventLogEntryType.Information, 44)
    '            'MsgBox("The error (e_lookforfiles_532) is " & e_lookforfiles_532.Message)
    '        End Try
    '    End Function

    Public d_p_oNow As Date = DateTime.Now

    Sub CheckDuration()
        Try

            If DateTime.Now >= d_p_oNow.AddMinutes(5) And s_p_SeedOrScan = "Scan" _
            And Not (i_p_keepModulo >= 0 And i_p_keepModulo <= 59) Then

                d_p_oNow = DateTime.Now
                Call fFindWhatToLookFor()

            End If

        Catch e_lookforfiles_1245 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_1245 --> " & e_lookforfiles_1245.Message, EventLogEntryType.Information, 44)

        End Try
    End Sub


    Function f_TakeDocument_URL(ByVal sDatabaseName) As String

        Dim myConnString As String
        Dim sKeepString As String = "stop"
        Dim s_DatabaseName As String = sDatabaseName
        'MsgBox(sSendUri & " sSendUri")
        Dim SUrl As String = "n"
        Dim s_contenant As String = "nothing"
        Dim mySelectQuery As String
        Dim myConnection As New OleDbConnection(ConnStringURL(s_DatabaseName))
        Dim myReader As OleDbDataReader

        Try
            If Len(s_DatabaseName) > 4 Then
                mySelectQuery = "SELECT top 1 URL FROM URLFOUND where  GoodSpot <> -40004 and Done ='" & Trim(SUrl) & "'"
                'MsgBox(mySelectQuery & " mySelectQuery")

                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
                myConnection.Open()

                myReader = myCommand.ExecuteReader()

                While myReader.Read()

                    sKeepString = myReader.GetString(0)
                    If Len(sKeepString) > 7 Then
                        s_contenant = sKeepString
                        'MsgBox(" f_TakeRacine_DNS 756 " & sKeepString)
                    Else
                        s_contenant = "stop"
                        'MsgBox(" f_f_TakeRacine_DNS 759 stop ")
                    End If
                End While

                If s_contenant = "nothing" Then
                    f_TakeDocument_URL = "stop"
                    Exit Function
                Else
                    f_TakeDocument_URL = s_contenant
                End If

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                '''myConnection = Nothing
                myConnection.Close()

            Else
                f_TakeDocument_URL = "stop"
                Exit Function
            End If

        Catch e_LookForFiles_1103 As Exception

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_1103 --> " & e_LookForFiles_1103.Message, EventLogEntryType.Information, 44)
            EventLog.WriteEntry(sSource, " lookforfiles.1103   s_DatabaseName ---- mySelectQuery --> " & s_DatabaseName & " ---- " & mySelectQuery, EventLogEntryType.Information, 44)

        End Try

        Try
            SUrl = "y"
            If Len(s_DatabaseName) > 4 And Len(sKeepString) > 4 Then
                Dim objConn As New OleDbConnection(ConnStringURL(s_DatabaseName))
                Dim myConnectionString As String = "UPDATE URLFOUND SET Done= '" & (SUrl) & "', Machine = '" & CStr(Now()) & "'  WHERE URL ='" & sKeepString & "'"
                'MessageBox.Show(myConnectionString, "myConnectionString")
                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = objConn
                objConn.Open()
                myCommand2.ExecuteNonQuery()
                myCommand2.Connection.Close()
                ''objconn = Nothing
                objConn.Close()

            End If
        Catch e_LookForFiles_1122 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_1122 --> " & e_LookForFiles_1122.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_LookForFiles_974) is " & e_LookForFiles_974.Message)
        End Try


    End Function


    Function f_TakeDocument_URL_First(ByVal sDatabaseName) As String

        Dim myConnString As String
        Dim sKeepString As String = "stop"
        Dim s_DatabaseName As String = sDatabaseName
        'MsgBox(sSendUri & " sSendUri")
        Dim SUrl As String = "n"
        Dim s_keepString As String = "nothing"
        Try
            Dim mySelectQuery As String = "SELECT top 1 URL FROM URLFOUND where  GoodSpot <> -40004 and ID = 0 and Done ='" & Trim(SUrl) & "' order by DateHour"
            'MsgBox(mySelectQuery & " mySelectQuery")
            Dim myConnection As New OleDbConnection(ConnStringURL(s_DatabaseName))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                sKeepString = myReader.GetString(0)
                If Len(sKeepString) > 7 Then
                    s_keepString = sKeepString
                    'MsgBox(" f_TakeRacine_DNS 979 " & sKeepString)
                Else
                    s_keepString = "stop"
                    'MsgBox(" f_f_TakeRacine_DNS 982 stop ")
                End If
            End While

            If s_keepString = "nothing" Then
                f_TakeDocument_URL_First = "stop"
                s_keepString = "stop"
            Else
                f_TakeDocument_URL_First = s_keepString
            End If

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            ''myConnection = Nothing
            myConnection.Close()

        Catch e_LookForFiles_1175 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_1175 --> " & e_LookForFiles_1175.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_LookForFiles_992) is " & e_LookForFiles_992.Message)
        End Try

        If sKeepString <> "stop" Then
            Try
                SUrl = "y"
                Dim objConn As New OleDbConnection(ConnStringURL(s_DatabaseName))
                Dim myConnectionString As String = "UPDATE URLFOUND SET ID = 1, Done= '" & (SUrl) & "', Machine = '" & CStr(Now()) & "'  WHERE URL ='" & sKeepString & "'"
                'MessageBox.Show(myConnectionString, "myConnectionString")
                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = objConn
                objConn.Open()
                myCommand2.ExecuteNonQuery()
                'myCommand2.Connection = Nothing
                myCommand2.Connection.Close()
                ''objconn = Nothing
                objConn.Close()

            Catch e_LookForFiles_1193 As Exception
                Dim sSource As String = "AP_DENIS"
                Dim sLog As String = "Applo"
                EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_1193 --> " & e_LookForFiles_1193.Message, EventLogEntryType.Information, 44)
                'MsgBox("The error (e_LookForFiles_1008) is " & e_LookForFiles_1008.Message)
            End Try
        Else
            f_TakeDocument_URL_First = "stop"
        End If

    End Function


    Function f_FetchRacine(ByVal sDatabaseName, ByVal sURLToLookFor)

        Call f_FetchRacineFU()

        Dim b_Continue As Boolean = True
        Dim b_Cont As Boolean = True
        Dim s_FoundURL As String
        Dim s_FoundRacine As String
        Dim s_URLToLookFor As String = sURLToLookFor
        Dim s_DatabaseName As String = sDatabaseName
        Dim ID_Racine As Integer
        Dim s_Verb As String = "s_Racine"
        Dim i_GoodSpot As Integer = i_p_ranking

        Try

            'Do While b_Cont

            's_FoundURL = f_TakeDocument_URL_First(s_DatabaseName)

            'If s_FoundURL = "stop" Then
            'f_FetchRacine = "stop"
            'Exit Do
            'End If

            'Call f_Extract_Racine_URL(s_FoundURL, s_DatabaseName, i_GoodSpot)

            'Loop

            Do While b_Continue

                s_URLToLookFor = f_TakeRacine_DNS_updated(s_DatabaseName, s_FoundURL, s_Verb)

                s_URLToLookFor = Replace(s_URLToLookFor, "/+++++\", "'")

                If Len(s_URLToLookFor) > 4 Then
                    f_FetchRacine = s_URLToLookFor
                    'Update DNS to say that we've taken that root
                    Exit Function
                    b_Continue = False
                Else
                    s_FoundURL = "stop"
                End If

                If s_FoundURL = "stop" Then
                    f_FetchRacine = "stop"
                    Exit Function
                End If

            Loop
        Catch e_LookForFiles_925 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_925 --> " & e_LookForFiles_925.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Function f_Extract_Racine_URL(Optional ByVal sFoundURL = "", Optional ByVal sDatabaseName = "", Optional ByVal GoodSpot = "", Optional ByVal s_searchEngine = "") As String

        Dim iGoodSpot As Integer
        Dim s_DatabaseName As String = sDatabaseName
        Dim s_FoundURL As String = sFoundURL
        Dim i_First As Integer
        Dim i_Last As Integer
        Dim i_Start As Integer = 1
        Dim i_Lenght As Integer
        Dim s_Racine, s_keepRacine As String
        Dim b_Continue As Boolean = True
        Dim s_Array() As String
        ReDim s_Array(25)
        Dim I_Counter As Integer = 1
        Dim i_dayToRemove As Integer = 30

        iGoodSpot = GoodSpot

        Try

            i_First = InStr(10, s_FoundURL, "/")

            If i_First > 0 Then
                i_Lenght = i_First

                Do While b_Continue
                    s_Racine = Mid(s_FoundURL, 1, i_Lenght)
                    s_Array(I_Counter) = s_Racine

                    'MsgBox("s_Racine 803 ->" & s_Racine)
                    i_Start = i_First + 1
                    i_First = InStr(i_Start, s_FoundURL, "/")

                    If i_First > 0 Then
                        i_Lenght = i_First
                    Else
                        b_Continue = False
                        f_Extract_Racine_URL = "stop"
                        Exit Do
                    End If
                    I_Counter = I_Counter + 1
                Loop

            End If

            If b_Continue = False Then
                If I_Counter = 1 Then
                    If CheckUrlIn_DNS(s_Array(I_Counter), s_DatabaseName) <> "stop" Then
                        Call InsertDNSFound(s_Array(I_Counter), s_DatabaseName, iGoodSpot, i_dayToRemove, , s_searchEngine)
                    End If
                Else
                    Do While I_Counter > 0
                        s_keepRacine = s_Array(I_Counter)
                        If CheckUrlIn_DNS(s_keepRacine, s_DatabaseName) <> "stop" Then
                            Call InsertDNSFound(s_keepRacine, s_DatabaseName, iGoodSpot, i_dayToRemove, , s_searchEngine)
                        End If
                        iGoodSpot = iGoodSpot - 10
                        I_Counter = I_Counter - 1
                    Loop
                End If
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.1883 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Public Function fGetUrlFromFileInsertInURLFOUND(ByVal FileToCrawl As String, ByVal sDatabaseName As String, ByVal sURLToLookFor As String) As String

        Dim b_KeepGoing As Boolean = True
        Dim i_FirstHREF As Integer
        Dim i_LastHREF As Integer
        Dim s_FileToCrawlToReturn As String
        Dim s_URLToLookFor As String = sURLToLookFor
        Dim s_Level As String
        Dim i_CountPage As Integer = 0
        Dim i_CountPage2 As Integer = 0

        s_Level = "1"
        Try

            If InsertDNA(sURLToLookFor, sDatabaseName, s_Level) = "stop" Then
                fGetUrlFromFileInsertInURLFOUND = "stop"
                Exit Function
            End If

            Do While b_KeepGoing

                If GetFirstHREF(FileToCrawl) = 0 Or (i_CountPage * 5) >= i_p_MaxPage _
                Or i_CountPage2 >= i_p_MaxPage Then
                    Exit Do
                Else
                    i_FirstHREF = GetFirstHREF(FileToCrawl)
                End If

                i_LastHREF = EndOfFirstHREF(FileToCrawl, i_FirstHREF)

                s_FileToCrawlToReturn = RemoveHREF(FileToCrawl, i_FirstHREF, i_LastHREF, s_URLToLookFor)

                If InsertDNA(s_FileToCrawlToReturn, sDatabaseName, s_Level) = "true" Then
                    i_CountPage2 = i_CountPage2 + 1
                Else
                    fGetUrlFromFileInsertInURLFOUND = "stop"
                    Exit Function
                End If

                FileToCrawl = RemoveURLFromFileToCrawl(FileToCrawl, i_FirstHREF, i_LastHREF)

                i_CountPage = i_CountPage + 1

            Loop
        Catch e_LookForFiles_1384 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_1384 --> " & e_LookForFiles_1384.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Function GetFirstHREF(ByVal FileToCrawl) As Integer
        Try
            Dim s_HREF As String = "href"

            If Len(FileToCrawl) > 50 Then
                GetFirstHREF = InStr(FileToCrawl, s_HREF, CompareMethod.Text)
            Else
                GetFirstHREF = 0
                Exit Function
            End If
        Catch e_lookforfiles_1603 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_1603 --> " & e_lookforfiles_1603.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_lookforfiles_532) is " & e_lookforfiles_532.Message)
        End Try
    End Function


    Function EndOfFirstHREF(ByVal FileToCrawl, ByVal FirstHRef) As Integer
        Try
            Dim i_LastBigger As Integer
            Dim i_First As Integer

            If FirstHRef > 0 Then
                i_First = FirstHRef
            End If

            EndOfFirstHREF = InStr(i_First + 8, FileToCrawl, ">")
        Catch e_lookforfiles_1622 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_1622 --> " & e_lookforfiles_1622.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_lookforfiles_532) is " & e_lookforfiles_532.Message)
        End Try
    End Function


    Function RemoveHREF(ByVal FileToCrawl, ByVal iFirstHREF, ByVal iLastHREF, ByVal sURLToLookFor) As String
        Try
            Dim i_lenght As Integer
            Dim s_URLToInsertInScratchpad As String
            Dim s_Array() As String

            If iFirstHREF > 0 And iLastHREF > 0 And iLastHREF > iFirstHREF Then
                i_lenght = iLastHREF - iFirstHREF
                s_URLToInsertInScratchpad = Mid(FileToCrawl, iFirstHREF + 5, i_lenght - 5)
                s_URLToInsertInScratchpad = Replace(s_URLToInsertInScratchpad, Chr(34), " ")
                'MsgBox(s_URLToInsertInScratchpad & " -> s_URLToInsertInScratchpad.886 ")
                s_Array = Split(Trim(s_URLToInsertInScratchpad))
                s_URLToInsertInScratchpad = Trim(s_Array(0))
                'MsgBox(s_Array(0) & " -> s_s_Array(0).889 ")
                'MsgBox(s_URLToInsertInScratchpad & " -> s_URLToInsertInScratchpad.889 ")
                s_URLToInsertInScratchpad = Replace(s_URLToInsertInScratchpad, "'", "")
                s_URLToInsertInScratchpad = Replace(s_URLToInsertInScratchpad, " ", "")
            End If

            'MsgBox(s_URLToInsertInScratchpad & " -> lookforfiles.s_URLToInsertInScratchpad.744 ")
            'MsgBox(sURLToLookFor & " -> lookforfiles.sURLToLookFor.745 ")
            'MsgBox(InStr(s_URLToInsertInScratchpad, ":") & " -> lookforfiles.InStr(s_URLToInsertInScratchpad, :).746 ")

            If Len(s_URLToInsertInScratchpad) > 5 Then

                If InStr(s_URLToInsertInScratchpad, ":") > 0 Then
                    RemoveHREF = s_URLToInsertInScratchpad
                Else

                    If InStr(8, sURLToLookFor, "/", CompareMethod.Text) Then
                        'MsgBox(sURLToLookFor & " -> lookforfiles.sURLToLookFor.830 ")
                        sURLToLookFor = Mid(sURLToLookFor, 1, (InStr(8, sURLToLookFor, "/", CompareMethod.Text) - 1))
                        'MsgBox(sURLToLookFor & " -> lookforfiles.sURLToLookFor.832 ")
                    End If

                    'MsgBox(s_URLToInsertInScratchpad & " -> lookforfiles.s_URLToInsertInScratchpad.881 ")
                    's_URLToInsertInScratchpad = "//" & Trim(s_URLToInsertInScratchpad)
                    's_URLToInsertInScratchpad = Replace(Trim(s_URLToInsertInScratchpad), "////", "")
                    's_URLToInsertInScratchpad = Replace(Trim(s_URLToInsertInScratchpad), "///", "")
                    's_URLToInsertInScratchpad = Replace(Trim(s_URLToInsertInScratchpad), "//", "")
                    'MsgBox(s_URLToInsertInScratchpad & " -> lookforfiles.s_URLToInsertInScratchpad.886 ")

                    If InStr(1, s_URLToInsertInScratchpad, "/", CompareMethod.Text) = 1 Then
                        s_URLToInsertInScratchpad = Mid(s_URLToInsertInScratchpad, 2, Len(s_URLToInsertInScratchpad) - 1)
                    End If
                    'MsgBox(s_URLToInsertInScratchpad & " -> lookforfiles.s_URLToInsertInScratchpad.891 ")
                    s_URLToInsertInScratchpad = sURLToLookFor & "/" & s_URLToInsertInScratchpad
                    'MsgBox(s_URLToInsertInScratchpad & " -> lookforfiles.s_URLToInsertInScratchpad.895 ")
                    RemoveHREF = s_URLToInsertInScratchpad
                End If

            End If
        Catch e_lookforfiles_1683 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_1683 --> " & e_lookforfiles_1683.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_lookforfiles_532) is " & e_lookforfiles_532.Message)
        End Try
    End Function


    Function RemoveURLFromFileToCrawl(ByVal FileToCrawl, ByVal iFirstHREF, ByVal iLastHREF) As String
        Try
            Dim i_LenFileToCrawl As Integer

            i_LenFileToCrawl = Len(FileToCrawl)
            RemoveURLFromFileToCrawl = Mid(FileToCrawl, 1, iFirstHREF) & Mid(FileToCrawl, iLastHREF, i_LenFileToCrawl)
        Catch e_lookforfiles_1698 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_1698 --> " & e_lookforfiles_1698.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_lookforfiles_532) is " & e_lookforfiles_532.Message)
        End Try
    End Function


    Function CheckDocumentInScratchpad(ByVal s_DatabaseName, ByVal s_FoundRacine) As String

        Dim s_ConnString = "ConnStringDNA"
        Dim s_TableName = "Scratchpad"
        Dim s_FieldName = "Url"
        Dim s_UrlFromScratchpad As String
        Dim i_GoodSpot As Integer = 25

        Try
            s_UrlFromScratchpad = "stop"

            If Len(s_DatabaseName) >= 3 Then
                s_UrlFromScratchpad = GetUrlFromScratchpad(s_DatabaseName, s_ConnString, s_TableName, s_FieldName)
            Else
                Console.WriteLine("lookforfiles.1894 -- len s_DatabaseName is shorter than 2!")
                s_UrlFromScratchpad = "stop"
            End If

            If s_UrlFromScratchpad <> "stop" Then
                Call f_Extract_Racine_URL(s_UrlFromScratchpad, s_DatabaseName, i_GoodSpot)
            End If

            CheckDocumentInScratchpad = s_UrlFromScratchpad

           Call CheckDocumentInScratchpadFU(s_UrlFromScratchpad)

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.1507 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Public Function GetUrlFromScratchpad(ByVal sDatabaseName, ByVal sConnString, ByVal sTableName, ByVal sUrlFieldName) As String

        Dim s_myConnString, s_mySelectQuery As String
        Dim s_KeepString As String = "stop"
        Dim s_SendUri As String = sDatabaseName
        Dim s_ConnString As String = sConnString
        Dim s_NA = "n/a"
        Dim s_Now As DateTime = Now()
        'MsgBox(sSendUri & " sSendUri")
        Dim S_Url As String = "n"
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim myReader As OleDbDataReader
        Dim s_heartBeat As String = "Heartbeat"
        Dim i_ID As Int64

        Try

            i_p_counterToGC = i_p_counterToGC + 1

            If Trim(s_ConnString) <> "ConnStringDNA" Then
                s_ConnString = s_ConnString & "(" & s_SendUri & ")"
            Else
                s_ConnString = s_ConnString & "()"
            End If

            If Len(sUrlFieldName) > 0 And Len(sTableName) > 0 Then

                Dim frq, counter1, counter2 As Long
                Dim duration As Double


                QueryPerformanceFrequency(frq)
                QueryPerformanceCounter(counter1)

                s_mySelectQuery = "SELECT top 1 " & sUrlFieldName & ", ID FROM SCRATCHPAD where Done ='" & Trim(S_Url) & "' order by DateHour"

                Dim myCommand As New OleDbCommand(s_mySelectQuery, myConnection)
                myConnection.Open()

                myReader = myCommand.ExecuteReader()
                While myReader.Read()

                    s_KeepString = myReader.GetString(0)
                    i_ID = myReader("ID")

                    If Len(s_KeepString) > 0 Then
                        If InStr(s_KeepString, "http", CompareMethod.Text) > 0 Then
                            s_KeepString = Replace(s_KeepString, "/+++++\", "'")
                            GetUrlFromScratchpad = s_KeepString
                            Console.WriteLine("lookforfiles.GetUrlFromScratchpad.8778 --> " & sTableName & " - " & i_ID & " - " & s_KeepString)
                            'EventLog.WriteEntry(sSource, " lookforfiles.GetUrlFromScratchpad.1855 --> " & s_KeepString, EventLogEntryType.Information, 8777)
                            Exit While
                        Else
                            GetUrlFromScratchpad = "stop"
                            Exit While
                        End If
                    Else
                        GetUrlFromScratchpad = "stop"
                        Exit While
                    End If

                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()

                QueryPerformanceCounter(counter2)
                duration = (counter2 - counter1) / frq

                'EventLog.WriteEntry(sSource, " lookforfiles.InsertDNA.1975 duration--> " & duration, EventLogEntryType.Information, 8778)

                'Console.WriteLine(duration & "  -- 8778 -- lookforfiles.InsertDNA.1975")

                If s_KeepString = "stop" Then
                    GetUrlFromScratchpad = "stop"
                End If
            Else
                GetUrlFromScratchpad = "stop"
            End If

        Catch e_lookforfiles_1400 As Exception

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            EventLog.WriteEntry(sSource, " -- sUrlFieldName --> " & sUrlFieldName & " -- sTableName --> " & sTableName & " -- sConnString --> " & sConnString & " -- sDatabaseName --> " & sDatabaseName & " -- The query is  --> " & s_mySelectQuery & "     -------- The error (e_lookforfiles_1400) is " & e_lookforfiles_1400.Message, EventLogEntryType.Error, 5)

        End Try



        Try

            If s_KeepString <> "stop" Then

                Form.TBScratchPad.Text = s_KeepString.ToString
                Form.TBScratchPad.Update()

                'Form.TBScratchPad.Text = ""
                Form.INSERTURLFOUNDTB4.Text = ""
                Form.TB40004.Text = ""
                Form.UPDATEURLTB3.Text = ""
                Form.RichTextBox1.Text = ""
                Form.TBLenghtPage.Text = ""
                Form.TBKeywords.Text = ""
                Form.TBSignature.Text = ""
                Form.FreshPageTB6.Text = ""
                Form.TextBox40004.Text = ""
                Form.FreshPageTB5.Text = ""
                Form.TBScratchPad.Update()
                Form.INSERTURLFOUNDTB4.Update()
                Form.TB40004.Update()
                Form.UPDATEURLTB3.Update()
                Form.RichTextBox1.Update()
                Form.TBLenghtPage.Update()
                Form.TBKeywords.Update()
                Form.TBSignature.Update()
                Form.FreshPageTB6.Update()
                Form.TextBox40004.Update()
                Form.FreshPageTB5.Update()


                Form.Refresh()



                Dim frq, counter1, counter2 As Long
                Dim duration As Double

                QueryPerformanceFrequency(frq)
                QueryPerformanceCounter(counter1)

                S_Url = "y"
                Dim objConn As New OleDbConnection(ConnStringDNA())
                Dim myConnectionString As String = "UPDATE SCRATCHPAD SET Done= '" & (S_Url) & "' WHERE ID=" & i_ID
                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = objConn
                objConn.Open()
                myCommand2.ExecuteNonQuery()
                myCommand2.Connection.Close()
                objConn.Close()

                QueryPerformanceCounter(counter2)
                duration = (counter2 - counter1) / frq

            Else

                Form.TBScratchPad.Text = s_KeepString.ToString
                Form.TBScratchPad.Update()
                Form.Refresh()

            End If

        Catch e_LookForFiles_1740 As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.e_LookForFiles_1740 --> " & e_LookForFiles_1740.Message, EventLogEntryType.Information, 44)

        End Try


    End Function

    Sub UpdateFrom()

        Form.TBScratchPad.Text = ""
        Form.INSERTURLFOUNDTB4.Text = ""
        Form.TB40004.Text = ""
        Form.UPDATEURLTB3.Text = ""
        Form.RichTextBox1.Text = ""
        Form.TBLenghtPage.Text = ""
        Form.TBKeywords.Text = ""
        Form.TBSignature.Text = ""
        Form.FreshPageTB6.Text = ""
        Form.TextBox40004.Text = ""
        Form.FreshPageTB5.Text = ""
        Form.TBScratchPad.Update()
        Form.INSERTURLFOUNDTB4.Update()
        Form.TB40004.Update()
        Form.UPDATEURLTB3.Update()
        Form.RichTextBox1.Update()
        Form.TBLenghtPage.Update()
        Form.TBKeywords.Update()
        Form.TBSignature.Update()
        Form.FreshPageTB6.Update()
        Form.TextBox40004.Update()
        Form.FreshPageTB5.Update()


        Form.Refresh()

    End Sub


    Public ReadOnly Property ir_Duration() As Double
        Get
            ir_Duration = i_p_Duration
        End Get
    End Property


    Public ReadOnly Property ir_MinKPage() As Integer
        Get
            ir_MinKPage = i_p_MinKPage
        End Get
    End Property


    Public ReadOnly Property ir_MaxKPage() As Integer
        Get
            ir_MaxKPage = i_p_MaxKPage
        End Get
    End Property


    Public ReadOnly Property ir_Priority() As Integer
        Get
            ir_Priority = i_Priority
        End Get
    End Property


    Public ReadOnly Property ir_Level() As Integer
        Get
            ir_Level = i_p_Level
        End Get
    End Property


    Public ReadOnly Property ir_MaxPage() As Integer
        Get
            ir_MaxPage = i_p_MaxPage
        End Get
    End Property


    Public ReadOnly Property ir_MinDensity() As Integer
        Get
            ir_MinDensity = i_p_MinDensity
        End Get
    End Property


    Public ReadOnly Property ir_MaxDensity() As Integer
        Get
            ir_MaxDensity = i_p_MaxDensity
        End Get
    End Property


    Public ReadOnly Property ir_KeywordTitle() As Integer
        Get
            ir_KeywordTitle = i_p_KeywordTitle
        End Get
    End Property


    Public ReadOnly Property ir_KeywordBody() As Integer
        Get
            ir_KeywordBody = i_p_KeywordBody
        End Get
    End Property


    Public ReadOnly Property ir_keywordTag() As Integer
        Get
            ir_keywordTag = i_p_keywordTag
        End Get
    End Property


    Public ReadOnly Property ir_KeywordVisible() As Integer
        Get
            ir_KeywordVisible = i_p_KeywordVisible
        End Get
    End Property


    Public ReadOnly Property ir_KeywordURL() As Integer
        Get
            ir_KeywordURL = i_p_KeywordURL
        End Get
    End Property


    Public ReadOnly Property ir_MinDiskPerc() As Integer
        Get
            ir_MinDiskPerc = i_p_MinDiskPerc
        End Get
    End Property


    Public ReadOnly Property ir_MinDiskGiga() As Integer
        Get
            ir_MinDiskGiga = i_p_MinDiskGiga
        End Get
    End Property


    Public ReadOnly Property ir_Algorythm() As Integer
        Get
            ir_Algorythm = i_p_Algorythm
        End Get
    End Property


    Public ReadOnly Property ir_Google() As Integer
        Get
            ir_Google = i_p_Google
        End Get
    End Property


    Public ReadOnly Property ir_AllTheWeb() As Integer
        Get
            ir_AllTheWeb = i_p_AllTheWeb
        End Get
    End Property


    Public ReadOnly Property ir_SearchExcel() As Integer
        Get
            ir_SearchExcel = i_p_SearchExcel
        End Get
    End Property


    Public ReadOnly Property ir_SearchPDF() As Integer
        Get
            ir_SearchPDF = i_p_SearchPDF
        End Get
    End Property


    Public ReadOnly Property ir_DesactiveQuery() As Integer
        Get
            ir_DesactiveQuery = i_p_DesactiveQuery
        End Get
    End Property


    Public ReadOnly Property ir_ExactPhrase() As Integer
        Get
            ir_ExactPhrase = i_p_ExactPhrase
        End Get
    End Property


    Public ReadOnly Property sr_FirstNetwork() As String
        Get
            sr_FirstNetwork = s_p_FirstNetwork
        End Get
    End Property

    Public ReadOnly Property sr_Priority() As String
        Get
            sr_Priority = s_p_Priority
        End Get
    End Property


    Public ReadOnly Property sr_SecondNetwork() As String
        Get
            sr_SecondNetwork = s_p_SecondNetwork
        End Get
    End Property


    Public ReadOnly Property sr_UniversalNetwork() As String
        Get
            sr_UniversalNetwork = s_p_UniversalNetwork
        End Get
    End Property


    Public ReadOnly Property ir_SearchWord() As Integer
        Get
            ir_SearchWord = i_p_SearchWord
        End Get
    End Property


    Public ReadOnly Property s_r_machineName() As String
        Get
            s_r_machineName = System.Environment.MachineName()
        End Get
    End Property

    Public s_precedentQueryForCheckParameters As String
    Public i_p_durationForCheckParameters As Int32

    Public Function CheckParametersInDocument(ByVal sDatabaseName) As Object

        Dim sToInsert As String
        Dim myConnectionString As String
        Dim sDate As String = Now()
        Dim s_Query As String = sDatabaseName
        Dim mySelectQuery, s_Website, s As String
        Dim b_PassTrough As Boolean
        Dim b_wentSmoothly As Boolean = False

        b_PassTrough = False


        Try

            If s_precedentQueryForCheckParameters <> s_Query And Len(s_p_authority) < 2 Then

                mySelectQuery = "SELECT * FROM DNAVALUE,QUERY where s_Query ='" & s_Query & "' and s_Query = QuerySearch"

                If Len(s_Query) > 0 Then

                    s_Query = sDatabaseName

                    Dim myConnection1 As New OleDbConnection(ConnStringDNA())
                    Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

                    myConnection1.Open()

                    Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                    Dim Index As Integer

                    While myReader.Read()
                        i_p_Duration = myReader("i_Duration")
                        i_Priority = myReader("i_Level")
                        s_p_Priority = myReader("s_Priority")
                        i_p_MaxPage = myReader("i_MaxPage")
                        i_p_MinKPage = myReader("i_MinKPage")
                        i_p_MaxKPage = myReader("i_MaxKPage")
                        i_p_MinDensity = myReader("i_MinDensity")
                        i_p_MaxDensity = myReader("i_MaxDensity")
                        i_p_KeywordTitle = myReader("i_KeywordTitle")
                        i_p_KeywordBody = myReader("i_KeywordBody")
                        i_p_keywordTag = myReader("I_keywordTag")
                        i_p_KeywordVisible = myReader("i_KeywordVisible")
                        i_p_KeywordURL = myReader("i_KeywordURL")
                        i_p_MinDiskPerc = myReader("i_MinDiskPerc")
                        i_p_MinDiskGiga = myReader("I_MinDiskGiga")
                        i_p_Algorythm = myReader("i_Algorythm")
                        i_p_Google = myReader("i_Google")
                        i_p_AllTheWeb = myReader("i_AllTheWeb")
                        i_p_SearchExcel = myReader("i_SearchExcel")
                        i_p_SearchWord = myReader("i_SearchWord")
                        i_p_SearchPDF = myReader("i_SearchPDF")
                        i_p_DesactiveQuery = myReader("i_DesactiveQuery")
                        i_p_ExactPhrase = myReader("i_ExactPhrase")
                        s_p_FirstNetwork = myReader("s_FirstNetwork")
                        s_p_SecondNetwork = myReader("s_SecondNetwork")
                        s_p_UniversalNetwork = myReader("s_UniversalNetwork")
                        s_Website = myReader("s_WebSite")
                        d_p_DateTime = myReader("d_DateTime")
                        s_p_authority = myReader("Machine")
                        b_wentSmoothly = True
                        Exit While
                    End While

                    If Not (myReader.IsClosed) Then
                        myReader.Close()
                    End If

                    myConnection1.Close()

                End If

                If LCase(s_r_machineName) = LCase(s_p_authority) Then  'TO CHANGE ASAP
                    s_p_authority = "yes"
                Else
                    s_p_authority = "no"
                End If

                '*************************************************
                s_p_authority = "yes"  'To change later 2016-03-26


                s_precedentQueryForCheckParameters = s_Query
                i_p_durationForCheckParameters = i_p_Duration

            End If

        Catch e_lookforfiles_2018 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2018 --> " & e_lookforfiles_2018.Message, EventLogEntryType.Information, 44)

        End Try

    End Function


    Public Function f_DateHourRating(ByVal d_dateHour) As Date
        Try
            Dim i_KeepDuration As Int32 = i_p_Duration * 60
            If i_KeepDuration = 0 Then i_KeepDuration = 1
            Dim i_MinuteSinceInsert As Int32 = DateDiff(DateInterval.Minute, d_p_DateTime, Now)
            Dim i_keepCycle As Int32 = i_MinuteSinceInsert / i_KeepDuration
            Dim d_keep2cycleBefore As DateTime = DateAdd(DateInterval.Minute, ((i_keepCycle - 2) * i_KeepDuration), d_p_DateTime)


            If d_dateHour > d_keep2cycleBefore Then
                d_dateHour = Now()
            Else
                d_dateHour = DateAdd(DateInterval.Hour, (i_p_Duration * -1), Now())
            End If

            Dim i_keepDateDiff As Integer

            f_DateHourRating = d_dateHour
        Catch e_lookforfiles_2519 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2519 --> " & e_lookforfiles_2519.Message, EventLogEntryType.Information, 44)

        End Try
    End Function


    Function CheckRangeGoHere(ByVal URLSite) As Boolean

        Try
            'Dim s_KeepResult As String = Func_CleanUrl(URLSite)

            'URLSite = Trim(FixFirstSecondLetter(s_KeepResult))

            'If URLSite >= sr_RangeMinus And URLSite <= sr_RangePlus Then
            '    CheckRangeGoHere = True
            'Else
            '    CheckRangeGoHere = False
            'End If

            CheckRangeGoHere = True

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.2555 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Function CheckRangeForUpdateDNSFound(ByVal URLSite) As String
        Try
            Dim s_urlSite As String = URLSite
            Dim s_keepResult As String
            Dim i_lenght As Integer
            Dim i_keepDotPosition As Integer
            Dim b_continue As Boolean = True

            'MsgBox("LookForFiles.Search.s_urlSite.1660-> " & s_urlSite)

            s_keepResult = Replace(URLSite, "http://", "")
            s_keepResult = Replace(s_keepResult, "www.", "")

            s_keepResult = Mid(s_keepResult, 1, 5)

            'MsgBox("LookForFiles.Search.s_KeepResult.1660-> " & s_keepResult)

            i_keepDotPosition = InStr(s_keepResult, ".")

            If i_keepDotPosition > 0 Then
                s_keepResult = Mid(s_keepResult, 1, i_keepDotPosition)
            End If

            'MsgBox("LookForFiles.Search.s_KeepResult.1668-> " & s_keepResult)

            If Len(s_keepResult) < 5 Then

                Do While b_continue
                    If Len(s_keepResult) >= 5 Then
                        Exit Do
                    End If
                    s_keepResult = s_keepResult & "a"
                    'MsgBox("LookForFiles.Search.s_KeepResult.1681-> " & s_keepResult)
                Loop

            End If

            'MsgBox("LookForFiles.Search.s_KeepResult.1683-> " & s_keepResult)

            'MsgBox("LookForFiles.Search.sr_RangeMinus.1676 -> " & Search.sr_RangeMinus)
            'MsgBox("LookForFiles.Search.sr_RangePlus.1677 -> " & Search.sr_RangePlus)

            CheckRangeForUpdateDNSFound = s_keepResult
        Catch e_lookforfiles_2232 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2232 --> " & e_lookforfiles_2232.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_lookforfiles_532) is " & e_lookforfiles_532.Message)
        End Try
    End Function

    Dim i_p_updateFromSAtoSDebut As Integer
    Dim b_p_FirstCycleInDNSFound As Boolean = True
    Dim b_p_continueTheSending As Boolean = True
    Public i_p_counterToSendDNS As Int32 = 0

    Sub SendDNSFound(Optional ByVal sSendUri As String = "", Optional ByVal b_seedCrawlers As Boolean = False)
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim i_topToSend As Integer = 1
        Dim d_KeepDuration As Double = i_p_Duration
        Dim i_MinuteSinceInsert As Int32
        Dim s_shortURL As String

        Try

            If b_p_FirstCycleInDNSFound Then

                d_KeepDuration = d_KeepDuration * 60
                i_MinuteSinceInsert = DateDiff(DateInterval.Minute, d_p_DateTime, Now)

                If i_MinuteSinceInsert > (d_KeepDuration * 7) Then
                    b_p_FirstCycleInDNSFound = False
                End If

            End If

            If b_seedCrawlers Then
                i_topToSend = 2
            Else
                If b_p_FirstCycleInDNSFound And s_p_authority = "yes" Then
                    i_topToSend = 5
                ElseIf s_p_authority = "yes" Then
                    i_topToSend = 5
                Else
                    i_topToSend = 2
                End If
            End If

            If Len(sSendUri) > 2 Then

                Dim mySelectQuery, s_keepString As String
                Dim i_ID, i_goodSpot As Integer
                Dim SUrl As String = "s"
                Dim s_NA As String = "n/a"
                Dim b_rangeIsOk As Boolean = False

                mySelectQuery = "SELECT top " & i_topToSend & " ID,URLSite,iGoodSpot,DateHour,ShortURL FROM DNSFOUND where Done ='" & Trim(SUrl) & "' and URLSite <> '" & Trim(s_NA) & "' order by iGoodSpot desc"

                Dim myConnection As New OleDbConnection(ConnStringURLDNS(sSendUri))
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
                Dim o_object As Object
                Dim s_RetrieveIDDateHour As Array
                Dim b_checkRecordLeft As Boolean = False
                Dim b_firstLettersOK, b_shortURLWellFormed As Boolean

                myConnection.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()

                    If IsNumeric(myReader("ID")) Then
                        i_ID = myReader("ID")
                    Else
                        i_ID = 0
                    End If

                    s_keepString = myReader("URLSite")

                    s_shortURL = myReader("ShortURL")

                    If IsNumeric(myReader("iGoodSpot")) Then
                        i_goodSpot = myReader("iGoodSpot")
                    Else
                        i_goodSpot = 0
                    End If

                    If IsNumeric(i_ID) Then
                        If i_ID <> 0 Then
                            b_shortURLWellFormed = VerifyShortURL(s_shortURL)
                            b_firstLettersOK = VerifyFirstLetters(s_shortURL)
                            If b_firstLettersOK And b_shortURLWellFormed Then
                                'b_rangeIsOk = CheckIfRangeIsOkBeforeSending(s_shortURL)
                                'If b_rangeIsOk Then
                                Call UpdateDNSFOUND(sSendUri, i_ID, s_keepString)
                                Call BroadCastDNSFOUND(i_ID, s_keepString, i_goodSpot, sSendUri, s_r_machineName)

                                'End If
                            Else
                                Call UpdateDNSFoundSONow(i_ID, sSendUri)
                            End If
                        End If
                    End If
                    b_checkRecordLeft = True
                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()

                If i_p_updateFromSAtoSDebut >= 350 Then
                    i_p_updateFromSAtoSDebut = 1
                    b_p_FirstCycleInDNSFound = True
                    b_p_continueTheSending = False
                End If

                If b_checkRecordLeft = False Then

                    i_p_counterToSendDNS = i_p_counterToSendDNS + 1

                    If i_p_counterToSendDNS >= 500 Then
                        Call UpdateFromSAtoS(sSendUri)
                    End If

                End If

                i_p_updateFromSAtoSDebut = i_p_updateFromSAtoSDebut + 1

            Else

            End If

        Catch e_lookforfiles_2172 As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2172 --> " & e_lookforfiles_2172.Message, EventLogEntryType.Information, 44)

        End Try

    End Sub

    '    'Private client As UdpClient
    Private ReceivePoint As IPEndPoint
    '    Dim bContinue As Boolean = True

    Function BroadCastDNSFOUND(Optional ByVal ID = 0, Optional ByVal URLSite = "", Optional ByVal i_goodSpot = 0, Optional ByVal sSendUri = "", Optional ByVal sHostName = "", Optional ByVal sAction = "")

        '        Try
        '            Dim sSource As String = "AP_DENIS"
        '            Dim sLog As String = "Applo"
        '            Dim client As UdpClient
        '            Dim buffer = New Byte
        '            Dim ascii As Encoding = Encoding.ASCII
        '            Dim s As String
        '            Dim i As Integer = 0

        '            sSendUri = Replace(sSendUri, " ", "_")

        '            sHostName = Trim(Replace(sHostName, " ", "_"))

        '            Dim i_len As Int32 = Len(sHostName)
        '            Dim i_denis As Int32
        '            Dim s_keepString As String

        '            URLSite = Replace(URLSite, " ", "_")

        '            s = ""
        '            s = ID & "----------" & Trim(URLSite) & "----------" & Trim(i_goodSpot) & "----------" & Trim(sSendUri) & "----------" & Trim(sHostName) & "----------"

        '            ReDim buffer(s.Length + 1)
        '            buffer = ascii.GetBytes(s.ToCharArray)

        '            'Subwrite(s, "LookForFiles", "2581")

        '            client = New UdpClient
        '            client.Connect("255.255.255.255", 27950)
        '            client.Send(buffer, s.Length)

        '            If sAction = "URL_sent" Then
        '                'EventLog.WriteEntry(sSource, " -- lookforfiles.s.2769 URL found sent Over by " & s_r_machineName & " (this machine)--> " & s, EventLogEntryType.Information, 12000)
        '                Console.WriteLine(" -- lookforfiles.s.2363 URL found sent Over by " & s_r_machineName & " (this machine)--> " & s)
        '            ElseIf CStr(ID) = "QueryInsertAcknowledgement" Then
        '                'EventLog.WriteEntry(sSource, " -- lookforfiles.s.2772 Query acknowledgement found sent Over by " & s_r_machineName & " (this machine)--> " & s, EventLogEntryType.Information, 12001)
        '                Console.WriteLine(" -- lookforfiles.s.2365 Query acknowledgement found sent Over by " & s_r_machineName & " (this machine)--> " & s)
        '            ElseIf CStr(ID) = "insert" Then
        '                'EventLog.WriteEntry(sSource, " -- lookforfiles.s.2775 Query insert found sent Over by " & s_r_machineName & " (this machine)--> " & s, EventLogEntryType.Information, 12002)
        '                Console.WriteLine(" -- lookforfiles.s.2367 Query insert found sent Over by " & s_r_machineName & " (this machine)--> " & s)
        '            ElseIf CStr(ID) = "update" Then
        '                'EventLog.WriteEntry(sSource, " -- lookforfiles.s.2778 Query update found sent Over by " & s_r_machineName & " (this machine)--> " & s, EventLogEntryType.Information, 12003)
        '                'Console.WriteLine(" -- lookforfiles.s.2369 Query update found sent Over by " & s_r_machineName & " (this machine)--> " & s)
        '            ElseIf CStr(ID) = "delete" Then
        '                Console.WriteLine(" -- lookforfiles.s.2369 Query delete found sent Over by " & s_r_machineName & " (this machine)--> " & s)
        '                'EventLog.WriteEntry(sSource, " -- lookforfiles.s.2782 Query delete found sent Over by " & s_r_machineName & " (this machine)--> " & s, EventLogEntryType.Information, 12004)
        '            ElseIf CStr(ID) = "AcknowledgeInsertDocumentFound" Then
        '                'EventLog.WriteEntry(sSource, " -- lookforfiles.s.2784 AcknowledgeInsertDocumentFoundfound sent Over by " & s_r_machineName & " (this machine)--> " & s, EventLogEntryType.Information, 12005)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(" ///-->>> lookforfiles.s.2787 AcknowledgeInsertDocumentFoundfound sent Over by " & s_r_machineName & " (this machine)--> " & s)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '            ElseIf CStr(ID) = "newMachineConfig" Then
        '                'EventLog.WriteEntry(sSource, " -- lookforfiles.s.2791 newMachineConfig sent Over by " & s_r_machineName & " (this machine)--> " & s, EventLogEntryType.Information, 12005)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(" ///-->>> lookforfiles.s.2794 newMachineConfig sent Over by " & s_r_machineName & " (this machine)--> " & s)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '            ElseIf CStr(ID) = "acknowNewMachinConfig" Then
        '                'EventLog.WriteEntry(sSource, " -- lookforfiles.s.2798 acknowNewMachinConfig sent Over by " & s_r_machineName & " (this machine)--> " & s, EventLogEntryType.Information, 12005)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(" ///-->>> lookforfiles.s.2801 acknowNewMachinConfig sent Over by " & s_r_machineName & " (this machine)--> " & s)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '            ElseIf CStr(ID) = "acknowHeartbeat" Then
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(" ///-->>> lookforfiles.s.2786 acknowHeartbeat sent Over by " & s_r_machineName & " (this machine)--> " & s)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '            ElseIf CStr(ID) = "NewMachineName" Then
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(" ///-->>> lookforfiles.s.2795 Machine name sent over by " & s_r_machineName & " (this machine)--> " & s)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '            ElseIf CStr(ID) = "AcknowNewMachineName" Then
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(" ///-->>> lookforfiles.s.2799 AcknowNewMachineName sent over by " & s_r_machineName & " (this machine)--> " & s)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '            ElseIf CStr(ID) = "RemoveTheSite" Then
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(" ///-->>> lookforfiles.s.2824 Sent -->" & s_r_machineName & " Remove this site --> " & s)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '            ElseIf CStr(ID) = "AcknowRemoveTheSite" Then
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(" ///-->>> lookforfiles.s.2830 Sent -->" & s_r_machineName & " AcknowRemoveTheSite --> " & s)
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(vbNewLine)
        '            Else
        '                Console.WriteLine(vbNewLine)
        '                Console.WriteLine(" lookforfiles.s.2860 This machine -->" & s_r_machineName & " this DNS--> " & s)
        '                Console.WriteLine(vbNewLine)
        '            End If
        '            Console.WriteLine(s_p_authority & "  --> s_p_authority.lookforfiles.2962")
        '            Console.WriteLine(vbNewLine)
        '        Catch ex As Exception
        '            Dim sSource As String = "AP_DENIS"
        '            Dim sLog As String = "Applo"
        '            EventLog.WriteEntry(sSource, " lookforfiles.2845 --> " & ex.Message, EventLogEntryType.Information, 44)
        '        End Try

    End Function

    Private client As UdpClient

    Public s_p_hostName As String

    Public s_p_serverStarted As Boolean

    Sub ReceiveDNSFound()

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"

        Try
            Dim s As String
            Dim s_HostName As String

            Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName)
            s = (CType(h.AddressList.GetValue(0), IPAddress).ToString)

            s_HostName = System.Net.Dns.GetHostEntry(s).HostName()

            s_p_hostName = System.Net.Dns.GetHostEntry(s).HostName()

            If s_p_serverStarted <> True Then

                client = Nothing
                client = New UdpClient(27950)

                ReceivePoint = New IPEndPoint(New IPAddress(0), 0)
                Dim readthread As Thread = New Thread(New ThreadStart(AddressOf WaitForPackets))
                readthread.Start()
                s_p_serverStarted = True

            End If

        Catch ex As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.2870 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub


    Sub WaitForPackets()
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim s As String

        Try
            While True

                Dim data As Byte() = client.Receive(ReceivePoint)

                Dim i_broadCast As Short = 1

                s = System.Text.Encoding.Default.GetString(data)

                Call ValidEntry(s)

            End While

        Catch ex As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.2434 --> " & s & " -- " & ex.Message, EventLogEntryType.Information, 44)

        End Try
    End Sub


    Sub ValidEntry(ByVal DatagramSent)

        'Dim s_DatagramSent As String = DatagramSent
        'Dim s_SplitDatagram() As String
        'Dim i_dayToRemove As Integer = 30
        'Dim o_goodSpot As Object
        'Dim o_id As Object
        'Dim s_databaseName, s_urlToInsert, s_hostName, s_MachineName As String
        'Dim sSource As String = "AP_DENIS"
        'Dim sLog As String = "Applo"

        'Try
        '    If InStr(s_DatagramSent, "emilie", CompareMethod.Text) > 0 Then
        '        Dim sssss As String
        '        sssss = "s"
        '    End If


        '    If InStr(s_DatagramSent, "----------", CompareMethod.Text) > 0 Then
        '        s_SplitDatagram = Split(s_DatagramSent, "----------")
        '    Else
        '        s_SplitDatagram = Split(s_DatagramSent)
        '    End If

        '    o_id = s_SplitDatagram(0)

        '    s_urlToInsert = s_SplitDatagram(1)

        '    s_urlToInsert = Replace(s_urlToInsert, "_", " ")

        '    If IsNumeric(s_SplitDatagram(2)) Then
        '        o_goodSpot = s_SplitDatagram(2)
        '    Else
        '        If Len(s_SplitDatagram(2)) > 0 Then
        '            o_goodSpot = s_SplitDatagram(2)
        '            o_goodSpot = Replace(o_goodSpot, "_", " ")
        '            o_goodSpot = Replace(o_goodSpot, "///*****\\\", "_")
        '        Else
        '            o_goodSpot = 0
        '        End If
        '    End If

        '    s_databaseName = Replace(s_SplitDatagram(3), "_", " ")
        '    s_databaseName = Replace(s_databaseName, "///*****\\\", "_")

        '    s_hostName = Replace(s_SplitDatagram(4), "_", " ")

        '    Dim i_len As Int32 = Len(s_hostName)
        '    s_DatagramSent = s_DatagramSent

        '    s_urlToInsert = s_urlToInsert

        '    If UBound(s_SplitDatagram) = 5 Then
        '        s_MachineName = s_SplitDatagram(5)
        '    End If

        '    If IsNumeric(o_id) Then

        '        If Trim(s_urlToInsert) <> "done" Then

        '            If CheckRangeGoHere(s_urlToInsert) = True Then

        '                Dim s_Verb As String = "DNSFoundSendOver"
        '                Dim sSendstrURI = "DNSFoundSendOver"
        '                Dim i_ID = "DNSFoundSendOver"
        '                Dim b_done As Boolean = False
        '                Dim s_keepResultFromCheckURL As String

        '                s_keepResultFromCheckURL = CheckUrlIn_DNS(s_urlToInsert, s_databaseName)

        '                If s_keepResultFromCheckURL <> "stop" Then

        '                    InsertDNSFound(s_urlToInsert, s_databaseName, o_goodSpot, i_dayToRemove, s_hostName)

        '                    s_databaseName = Replace(Trim(s_databaseName), " ", "/$$$$$$$\")
        '                    s_hostName = Replace(Trim(s_hostName), " ", "/???????\")

        '                    i_len = Len(s_hostName)

        '                    Call BroadCastDNSFOUND(s_Verb, o_id, s_databaseName, s_hostName, s_r_machineName)
        '                    b_done = True

        '                Else

        '                    If s_keepResultFromCheckURL = "stop" Then

        '                        s_databaseName = Replace(Trim(s_databaseName), " ", "/$$$$$$$\")
        '                        s_hostName = Replace(Trim(s_hostName), " ", "/???????\")

        '                        i_len = Len(s_hostName)

        '                        Call BroadCastDNSFOUND(s_Verb, o_id, s_databaseName, s_hostName, s_r_machineName)

        '                        b_done = True

        '                    End If

        '                End If

        '                s_urlToInsert = "done"

        '            End If

        '        Else

        '            Dim s_keep, s As String
        '            Dim s_presentHostName As String

        '            Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName)
        '            s = (CType(h.AddressList.GetValue(0), IPAddress).ToString)

        '            s_presentHostName = System.Net.Dns.GetHostEntry(s).HostName()

        '            If LCase(Trim(s_hostName)) = LCase(Trim(s_presentHostName)) Then

        '                Dim s_se As String = "se"
        '                Dim myConnectionString As String
        '                Dim objConn As New OleDbConnection(ConnStringURLDNS(s_databaseName))
        '                myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_se) & "' where ID = " & (o_id)
        '                Dim myCommand2 As New OleDbCommand(myConnectionString)
        '                myCommand2.Connection = objConn
        '                objConn.Open()
        '                myCommand2.ExecuteNonQuery()
        '                'objconn = Nothing
        '                objConn.Close()
        '            End If
        '        End If

        '    Else

        '        Call InsertUpdateFromAfar(o_id, s_urlToInsert, o_goodSpot, s_databaseName, s_hostName)
        '        'EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2779_s_Verb, o_id, s_databaseName, s_hostName --> " & o_id & " _ " & o_id & " _ " & s_databaseName & " _ " & s_hostName & " _ " & Len(s_hostName), EventLogEntryType.Information, 7779)

        '    End If

        'Catch e_lookforfiles_2385 As Exception
        '    EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2385 --> " & e_lookforfiles_2385.Message, EventLogEntryType.Information, 44)
        'End Try

    End Sub

    Function UpdateDNSFOUND(Optional ByVal s_databaseName = "", Optional ByVal i_id = "", Optional ByVal s_keepString = "")
        Try

            Dim myConnString As String
            Dim sKeepString As String
            Dim date_InsertURL As DateTime
            Dim s_url As String = s_keepString
            Dim s_machineToSend As String


            s_keepString = Func_CleanUrl(s_keepString)
            s_url = FindFirstTwoLettersOfSomething(s_keepString)

            Dim mySelectQuery As String = "SELECT HostName FROM RANGE where RangeMinus <='" & Trim(s_url) & "' and RangePlus >='" & Trim(s_url) & "'"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_machineToSend = myReader("HostName")
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            '''myConnection = Nothing
            myConnection.Close()

            If IsNumeric(i_id) Then
                If i_id > 0 Then
                    Dim s_sa As String = "sa"
                    Dim myConnectionString As String
                    Dim objConn As New OleDbConnection(ConnStringURLDNS(s_databaseName))
                    myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_sa) & "', Answer= '" & (s_machineToSend) & "' where ID = " & (i_id)
                    Dim myCommand2 As New OleDbCommand(myConnectionString)
                    myCommand2.Connection = objConn
                    objConn.Open()
                    myCommand2.ExecuteNonQuery()
                    'objconn = Nothing
                    objConn.Close()
                End If
            End If

        Catch e_lookforfiles_2406 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2406 --> " & e_lookforfiles_2406.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_lookforfiles_264) is " & e_lookforfiles_264.Message)
        End Try
    End Function

    '    Function UpdateDNSFoundSO(ByVal i_id, ByVal S_DatabaseName, ByVal Machine)
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        Dim myConnectionString As String
    '        Try
    '            If IsNumeric(i_id) Then
    '                If i_id > 0 And Len(S_DatabaseName) > 0 Then
    '                    Dim s_so As String = "so"

    '                    Dim objConn As New OleDbConnection(ConnStringURLDNS(S_DatabaseName))
    '                    myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_so) & "', Answer = '" & (Machine) & "' where ID = " & (i_id)

    '                    Dim myCommand2 As New OleDbCommand(myConnectionString)
    '                    myCommand2.Connection = objConn
    '                    objConn.Open()
    '                    myCommand2.ExecuteNonQuery()
    '                    objConn.Close()
    '                End If
    '            End If

    '        Catch ex As Exception

    '            EventLog.WriteEntry(sSource, " lookforfiles.3148 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try
    '    End Function

    Function UpdateDNSFoundSONow(ByVal i_id As Int64, ByVal S_DatabaseName As String)
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim myConnectionString As String
        Try
            If IsNumeric(i_id) Then
                If i_id > 0 And Len(S_DatabaseName) > 0 Then
                    Dim s_so As String = "so"

                    Dim objConn As New OleDbConnection(ConnStringURLDNS(S_DatabaseName))
                    myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_so) & "' where ID = " & (i_id)

                    Dim myCommand2 As New OleDbCommand(myConnectionString)
                    myCommand2.Connection = objConn
                    objConn.Open()
                    myCommand2.ExecuteNonQuery()
                    objConn.Close()
                End If
            End If

        Catch ex As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.3176 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try
    End Function

    Function UpdateDNSFoundSONow(ByVal s_shortURL As String, ByVal S_DatabaseName As String)
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim myConnectionString As String
        Try

            If Len(s_shortURL) > 0 And Len(S_DatabaseName) > 0 Then
                Dim s_so As String = "so"

                Dim objConn As New OleDbConnection(ConnStringURLDNS(S_DatabaseName))
                myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_so) & "' where ShortURL = '" & (s_shortURL)

                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = objConn
                objConn.Open()
                myCommand2.ExecuteNonQuery()
                objConn.Close()
            End If

        Catch ex As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.3202 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try
    End Function

    Function UpdateFromSAtoS(ByVal s_databaseName)
        Try
            Dim s_sa As String = "sa"
            Dim s_s As String = "s"
            Dim myConnectionString As String
            Dim objConn As New OleDbConnection(ConnStringURLDNS(s_databaseName))
            myConnectionString = "UPDATE DNSFOUND SET Done='" & Trim(s_s) & "' where Done = '" & Trim(s_sa) & "'"
            'MessageBox.Show(myConnectionString & " " & s_databaseName & " -> lookforfiles.myConnectionString.1864")
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = objConn
            objConn.Open()
            myCommand2.ExecuteNonQuery()
            'objconn = Nothing
            objConn.Close()
        Catch e_lookforfiles_2449 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2449 --> " & e_lookforfiles_2449.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_lookforfiles_264) is " & e_lookforfiles_264.Message)
        End Try
    End Function

    '    Function InsertUpdateFromAfar(ByVal verb, ByVal URIToPass, ByVal queryDNA, ByVal querySchedule, ByVal MachineName)

    '        Dim s_vverb As String = verb
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        Dim myInsertQuery As String = queryDNA
    '        Dim myConnection As New OleDbConnection(ConnStringDNA())
    '        Dim s_query As String = URIToPass
    '        Dim b_AlreadyThere As Boolean = False
    '        Dim s_yes As String = "yes"
    '        'Dim b_doIt As Boolean = True
    '        Dim s_machineNameAuthority As String

    '        Try

    '            'EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2907_s_Verb, o_id, s_databaseName, s_hostName --> " & verb & " _ " & URIToPass & " _ " & queryDNA & " _ " & querySchedule & " _ " & Len(MachineName), EventLogEntryType.Information, 7780)

    '            Select Case verb

    '                Case "delete"
    '                    '//----
    '                    Try
    '                        If s_r_machineName <> MachineName Then
    '                            Dim s_array() As String
    '                            ReDim s_array(1)
    '                            s_array = Split(URIToPass, "|||")
    '                            URIToPass = s_array(0)
    '                            If UBound(s_array) = 1 Then
    '                                s_machineNameAuthority = s_array(1)
    '                            End If
    '                            Call InsertInDeleted(URIToPass)
    '                            Call Delete_Click(URIToPass, s_machineNameAuthority)
    '                            Call UpdateDeleted(URIToPass)
    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3189 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try
    '                '\\----

    '                Case "AcknowledgeDelete"
    '                    Try
    '                        If s_p_authority = "yes" Or s_r_machineName = querySchedule Then
    '                            Call DeleteInBroadcastQuery(URIToPass, queryDNA)
    '                            Call DeleteQueryIfSOAll(URIToPass)
    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3196 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "QueryInsertAcknowledgement"

    '                    'EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2920_s_Verb, o_id, s_databaseName, s_hostName --> " & verb & " _ " & URIToPass & " _ " & queryDNA & " _ " & querySchedule & " _ " & Len(MachineName), EventLogEntryType.Information, 7781)
    '                    Try
    '                        If s_r_machineName = MachineName Then

    '                            Dim b_here As Boolean = False
    '                            Dim d_now As DateTime = Now()
    '                            Dim mySelectQuery As String = "SELECT s_Query FROM BROADCAST_QUERY where s_Query ='" & Trim(URIToPass) & "' and s_Machine ='" & Trim(queryDNA) & "'"
    '                            Dim myConnection1 As New OleDbConnection(ConnStringDNA())

    '                            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
    '                            myConnection1.Open()

    '                            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
    '                            Dim Index As Integer
    '                            While myReader.Read()
    '                                'If s_query = myReader("s_Query") Then
    '                                b_here = True
    '                                Exit While
    '                                'End If
    '                            End While

    '                            If Not (myReader.IsClosed) Then
    '                                myReader.Close()
    '                            End If

    '                            myConnection1.Close()

    '                            If b_here = False Then

    '                                Dim myConnectionString As String

    '                                myInsertQuery = ""
    '                                myInsertQuery = "INSERT INTO BROADCAST_QUERY (s_Query,s_Machine,d_Date)" _
    '                                           & "Values ('" & Trim(URIToPass) & "','" & (queryDNA) & "','" & (d_now) & "')"

    '                                'EventLog.WriteEntry(sSource, " lookforfiles.Acknowledgement.2943 --> " & myInsertQuery, EventLogEntryType.Information, 3)

    '                                Dim myCommand1 As New OleDbCommand(myInsertQuery)
    '                                myCommand1.Connection = myConnection
    '                                myConnection.Open()
    '                                myCommand1.ExecuteNonQuery()
    '                                myCommand1.Connection.Close()
    '                                myConnection.Close()

    '                            Else

    '                                Dim s_so As String = "so"
    '                                Dim myConnString = "UPDATE BROADCAST_QUERY SET s_Active ='" & (s_so) & "', d_Date = '" & (d_now) & "'  where s_Query = '" & (URIToPass) & "' and s_Machine = '" & (queryDNA) & "'"

    '                                Dim myCommand2 As New OleDbCommand(myConnString)
    '                                myCommand2.Connection = myConnection
    '                                myConnection.Open()
    '                                myCommand2.ExecuteNonQuery()
    '                                myCommand2.Connection.Close()
    '                                myConnection.Close()

    '                            End If

    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3257 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "QueryUpdateAcknowledgement"

    '                    'EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2920_s_Verb, o_id, s_databaseName, s_hostName --> " & verb & " _ " & URIToPass & " _ " & queryDNA & " _ " & querySchedule & " _ " & Len(MachineName), EventLogEntryType.Information, 7781)
    '                    Try
    '                        If s_p_authority = "yes" Then

    '                            URIToPass = Replace(URIToPass, "||", " ")

    '                            Dim d_now As DateTime = Now()
    '                            Dim myConnectionString As String
    '                            Dim s_so As String = "so"
    '                            Dim myConnString = "UPDATE BROADCAST_QUERY SET s_Active ='" & (s_so) & "', d_Date = '" & (d_now) & "'  where s_Query = '" & (URIToPass) & "' and s_Machine = '" & (queryDNA) & "'"
    '                            Dim myCommand2 As New OleDbCommand(myConnString)
    '                            myCommand2.Connection = myConnection
    '                            myConnection.Open()
    '                            myCommand2.ExecuteNonQuery()
    '                            myCommand2.Connection.Close()
    '                            myConnection.Close()

    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3278 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "insert"

    '                    '//----
    '                    Try
    '                        If s_r_machineName <> MachineName Then

    '                            Dim mySelectQuery As String = "SELECT s_Query FROM DNAVALUE where s_Query ='" & s_query & "'"
    '                            Dim myConnection1 As New OleDbConnection(ConnStringDNA())

    '                            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
    '                            myConnection1.Open()

    '                            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
    '                            Dim Index As Integer
    '                            While myReader.Read()
    '                                If s_query = myReader("s_Query") Then
    '                                    b_AlreadyThere = True
    '                                    Exit While
    '                                End If
    '                            End While

    '                            If Not (myReader.IsClosed) Then
    '                                myReader.Close()
    '                            End If
    '                            myConnection1.Close()

    '                            'End If
    '                            '\\----

    '                            '//----
    '                            If b_AlreadyThere = False Then

    '                                Call CreateTable(URIToPass, MachineName)
    '                                myInsertQuery = Replace(myInsertQuery, "||", "_")
    '                                Dim myCommand1 As New OleDbCommand(myInsertQuery)
    '                                myCommand1.Connection = myConnection
    '                                myConnection.Open()
    '                                myCommand1.ExecuteNonQuery()
    '                                myCommand1.Connection.Close()
    '                                myConnection.Close()

    '                                myInsertQuery = ""
    '                                querySchedule = Replace(querySchedule, "||", "_")
    '                                myInsertQuery = querySchedule

    '                                Dim myCommand2 As New OleDbCommand(myInsertQuery)
    '                                myCommand2.Connection = myConnection
    '                                myConnection.Open()
    '                                myCommand2.ExecuteNonQuery()

    '                                myCommand2.Connection.Close()

    '                                myConnection.Close()

    '                            End If

    '                            Dim s_verb As String = "QueryInsertAcknowledgement"
    '                            Dim s_querySent As String = myInsertQuery
    '                            Dim s_machine As String = s_r_machineName

    '                            BroadCastDNSFOUND(s_verb, URIToPass, s_machine, , MachineName)

    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3342 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try
    '                '\\----

    '                Case "update"
    '                    Try
    '                        '//----
    '                        If Len(MachineName) > 0 And s_r_machineName <> MachineName Then

    '                            Dim b_c As Boolean = False
    '                            Dim s_splitQuery() As String
    '                            Dim s_arrayToQuery() As String
    '                            ReDim s_arrayToQuery(1)

    '                            s_arrayToQuery = Split(URIToPass, "|||||")
    '                            queryDNA = s_arrayToQuery(0)
    '                            querySchedule = s_arrayToQuery(1)
    '                            If b_c Then
    '                                Dim mySelectQuery As String = "SELECT s_Query FROM DNAVALUE where s_Query ='" & s_query & "'"
    '                                Dim myConnection1 As New OleDbConnection(ConnStringDNA())

    '                                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
    '                                myConnection1.Open()

    '                                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
    '                                Dim Index As Integer
    '                                While myReader.Read()
    '                                    If s_query = myReader("s_Query") Then
    '                                        b_AlreadyThere = True
    '                                        Exit While
    '                                    End If
    '                                End While

    '                                If Not (myReader.IsClosed) Then
    '                                    myReader.Close()
    '                                End If

    '                                myConnection1.Close()

    '                            End If
    '                            '\\----

    '                            '//----

    '                            b_AlreadyThere = True
    '                            If b_AlreadyThere = True Then

    '                                Dim myConnectionString As String = myInsertQuery
    '                                Dim objConn As New OleDbConnection(ConnStringDNA())

    '                                myInsertQuery = Replace(queryDNA, "|", "_")
    '                                Dim myCommand2 As New OleDbCommand(myInsertQuery)
    '                                myCommand2.Connection = objConn
    '                                objConn.Open()
    '                                myCommand2.ExecuteNonQuery()

    '                                myCommand2.Connection.Close()

    '                                objConn.Close()

    '                                myConnectionString = ""
    '                                querySchedule = Replace(querySchedule, "|", "_")
    '                                myConnectionString = querySchedule

    '                                Dim myCommand3 As New OleDbCommand(myConnectionString)
    '                                myCommand3.Connection = objConn
    '                                objConn.Open()
    '                                myCommand3.ExecuteNonQuery()

    '                                myCommand3.Connection.Close()

    '                                objConn.Close()

    '                            Else
    '                                If InStr(1, querySchedule, "|||||", CompareMethod.Text) > 0 Then
    '                                    s_splitQuery = Split(querySchedule, "|||||")
    '                                    myInsertQuery = s_splitQuery(0)
    '                                    querySchedule = s_splitQuery(1)

    '                                    Call CreateTable(URIToPass, MachineName)

    '                                    Dim myCommand1 As New OleDbCommand(myInsertQuery)
    '                                    myCommand1.Connection = myConnection
    '                                    myConnection.Open()
    '                                    myCommand1.ExecuteNonQuery()

    '                                    myCommand1.Connection.Close()

    '                                    myConnection.Close()

    '                                    myInsertQuery = ""
    '                                    myInsertQuery = querySchedule

    '                                    Dim myCommand2 As New OleDbCommand(myInsertQuery)
    '                                    myCommand2.Connection = myConnection
    '                                    myConnection.Open()
    '                                    myCommand2.ExecuteNonQuery()

    '                                    myCommand2.Connection.Close()

    '                                    myConnection.Close()
    '                                End If
    '                            End If

    '                            Dim s_machine As String = s_r_machineName
    '                            Dim s_verb As String = "QueryUpdateAcknowledgement"
    '                            Dim s_whereFirst As String = "s_Query ="
    '                            Dim i_placeOfFirst As Int32 = InStr(myInsertQuery, s_whereFirst, CompareMethod.Text)
    '                            Dim i_placeLastQuote As Int32 = InStr(i_placeOfFirst + 12, myInsertQuery, "'", CompareMethod.Text)
    '                            Dim i_lenght As Int32 = i_placeLastQuote - (i_placeOfFirst + 11)
    '                            URIToPass = Mid(myInsertQuery, i_placeOfFirst + 10, i_lenght + 1)
    '                            URIToPass = Replace(URIToPass, " ", "||")

    '                            BroadCastDNSFOUND(s_verb, URIToPass, s_machine)

    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3456 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try
    '                '\\----

    '                Case "Successfull"
    '                    Try
    '                        '//----
    '                        If s_r_machineName <> MachineName Then

    '                            Dim i_cycleTime As Double = i_p_Duration * 60
    '                            Dim d_cycleDate As DateTime
    '                            Dim s_date As String

    '                            d_cycleDate = DateAdd(DateInterval.Minute, (-2), Now())

    '                            Dim mySelectQuery As String = "SELECT s_Query FROM DNAVALUE where d_DateTime < #" & d_cycleDate & "#"

    '                            '----//
    '                            Dim myConnection1 As New OleDbConnection(ConnStringDNA())

    '                            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
    '                            myConnection1.Open()

    '                            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
    '                            '----\\

    '                            While myReader.Read()

    '                                Call Delete_Click(myReader("s_Query"))

    '                            End While

    '                            If Not (myReader.IsClosed) Then
    '                                myReader.Close()
    '                            End If
    '                            '''myConnection1 = Nothing
    '                            myConnection1.Close()

    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3493 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try
    '                '\\----

    '                Case "HeartBeat"
    '                    Try
    '                        If s_p_authority = "yes" Then
    '                            Call InsertHeartBeatScratchPad(URIToPass)
    '                            Call SendAcknowHeartBeat(URIToPass)
    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3502 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "SendAcknowHeartBeat"
    '                    Try
    '                        Dim s_receiveHeartbeat As String = "yes"
    '                        Call InsertHeartBeatScratchPad(s_receiveHeartbeat, URIToPass)
    '                        Console.WriteLine("Receive SendAcknowHeartBeat from authority!!!!!!! ")
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3507 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "acknowHeartbeat"
    '                    Try
    '                        If s_p_authority <> "yes" Then
    '                            If URIToPass = s_r_machineName Then
    '                                Call InsertAcknowHeartBeat(URIToPass)
    '                            End If
    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3514 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "DNSFoundSendOver"
    '                    Try
    '                        Dim i_ID As Int32

    '                        queryDNA = Replace(queryDNA, "/$$$$$$$\", " ")

    '                        querySchedule = Replace(querySchedule, "/???????\", " ")

    '                        If querySchedule = s_r_machineName Then
    '                            If IsNumeric(URIToPass) Then
    '                                i_ID = CInt(URIToPass)
    '                            Else
    '                                i_ID = 0
    '                            End If

    '                            If i_ID > 0 And Len(queryDNA) > 0 Then
    '                                Call UpdateDNSFoundSO(i_ID, queryDNA, MachineName)
    '                            End If

    '                            'Call Subwrite(i_ID, "lookforfiles.i_ID", "3244")
    '                            'Call Subwrite(queryDNA, "lookforfiles.querydna", "3245")
    '                            Call Subwrite(MachineName & " -- " & i_ID & " -- " & queryDNA, "DNSFoundSendOver & lookforfiles.MachineName", "3246")

    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3538 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "InsertDocumentFound"
    '                    Try
    '                        If s_p_authority = "yes" Then
    '                            Console.WriteLine(vbNewLine)
    '                            Console.WriteLine("Insert Document Found from Outside.lookforfiles.3777 --> " & URIToPass)
    '                            Console.WriteLine(vbNewLine)
    '                            Dim s_array() As String = Split(URIToPass, "|||||")

    '                            Dim ID As String = s_array(0)
    '                            Dim SUrlInsert As String = s_array(1)
    '                            SUrlInsert = Replace(SUrlInsert, "!!!!!", "_")

    '                            Dim DateHour As String = s_array(2)
    '                            Dim sMachine As String = s_array(3)
    '                            sMachine = Replace(sMachine, "!!!!!", "_")
    '                            Console.WriteLine("sMachine.lookforfiles.3777 --> " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine & " " & sMachine)
    '                            Console.WriteLine(vbNewLine)
    '                            Dim Signature As String = s_array(4)

    '                            Dim sSendstrURI As String = s_array(5)
    '                            sSendstrURI = Replace(sSendstrURI, "!!!!!", "_")

    '                            Dim i_goodSpotFromOver As String = s_array(6)

    '                            Call InsertURLFOUNDFromBroadcast(SUrlInsert, sSendstrURI, ID, sMachine, Signature, i_goodSpotFromOver)

    '                            'EventLog.WriteEntry(sSource, " -- lookforfiles.InsertDocumentFound.3455 --> ", EventLogEntryType.Information, 776)

    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3563 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "AcknowledgeInsertDocumentFound"
    '                    Try
    '                        Dim s_array() As String = Split(URIToPass, "|||||")
    '                        Dim SUrlInsert As String = s_array(0)
    '                        SUrlInsert = Replace(SUrlInsert, "!!!!!", "_")
    '                        Dim sSendstrURI As String = s_array(1)
    '                        Dim sMachine As String = s_array(2)
    '                        sMachine = Replace(sMachine, "!!!!!", "_")
    '                        Dim s_y As String = "so"

    '                        If s_r_machineName = sMachine And s_p_authority = "no" Then

    '                            Dim objConn As New OleDbConnection(ConnStringURL(sSendstrURI))
    '                            Dim myConnectionString As String = "UPDATE URLFOUND SET Done= '" & (s_y) & "' WHERE URL ='" & SUrlInsert & "'"
    '                            Dim myCommand2 As New OleDbCommand(myConnectionString)
    '                            myCommand2.Connection = objConn
    '                            objConn.Open()
    '                            myCommand2.ExecuteNonQuery()
    '                            myCommand2.Connection.Close()
    '                            objConn.Close()

    '                        End If

    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3587 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "UpdateDNSFOUND"
    '                    Try
    '                        If s_p_authority = "yes" Then
    '                            Call InsertHeartBeatScratchPad(URIToPass)
    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3594 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "newMachineConfig"
    '                    Try
    '                        Dim s_array() As String = Split(URIToPass, "|||||")
    '                        Dim s_hostName As String = s_array(0)
    '                        Dim s_IPAddress As String = s_array(1)
    '                        Dim s_RangeMinus As String = s_array(2)
    '                        Dim s_RangePlus As String = s_array(3)
    '                        Dim s_CallReceived As String = s_array(4)
    '                        Dim s_Active As String = s_array(5)

    '                        If s_p_authority <> "yes" Then
    '                            Call InsertNewMachinConfigFromBroadcast(s_hostName, s_IPAddress, s_RangeMinus, s_RangePlus, s_CallReceived, s_Active)
    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3610 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "acknowNewMachinConfig"
    '                    Try
    '                        If s_p_authority = "yes" Then

    '                            Dim s_array() As String = Split(URIToPass, "|||||")

    '                            Dim s_machineFrom As String = s_array(0)
    '                            Dim s_hostName As String = s_array(1)

    '                            Call InsertAckNewConfigInAuthority(s_machineFrom, s_hostName)

    '                            'EventLog.WriteEntry(sSource, " -- lookforfiles.newMachineConfig.3521 --> ", EventLogEntryType.Information, 776)

    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3628 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try


    '                Case "NewMachineName"
    '                    Try
    '                        If s_p_authority = "yes" Then

    '                            Dim s_machineName As String = URIToPass
    '                            Dim s_fromAbroad As String = "yes"

    '                            If sWhatRange(s_machineName, s_fromAbroad) Then
    '                                Dim s_verb As String = "AcknowNewMachineName"
    '                                BroadCastDNSFOUND(s_verb, s_machineName)
    '                            Else
    '                                Call CheckMachineNamePublished(s_machineName)
    '                            End If

    '                        End If

    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3642 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try


    '                Case "AcknowNewMachineName"
    '                    Try
    '                        If s_r_machineName = URIToPass Then
    '                            Dim s_machineName As String = URIToPass
    '                            AcknowMachineNameReceived(s_machineName)
    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3648 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try


    '                Case "RemoveTheSite"
    '                    Try
    '                        If CheckRangeGoHere(queryDNA) = True Then

    '                            Dim s_Verb As String = "AcknowRemoveTheSite"
    '                            Dim i_GoodSpotMoins1000 As Int64 = -888888
    '                            Dim s_shortURL As String = queryDNA
    '                            Dim myConnectionString As String
    '                            Dim s_databaseName As String = URIToPass
    '                            If Len(s_shortURL) > 0 And Len(s_databaseName) > 0 Then
    '                                Dim objConn1 As New OleDbConnection(ConnStringURLDNS(s_databaseName))
    '                                myConnectionString = "UPDATE DNSFOUND SET iGoodSpot=" & (i_GoodSpotMoins1000) & " where ShortURL = '" & (s_shortURL) & "'"
    '                                Dim myCommand1 As New OleDbCommand(myConnectionString)
    '                                myCommand1.Connection = objConn1
    '                                objConn1.Open()
    '                                myCommand1.ExecuteNonQuery()
    '                                objConn1.Close()

    '                                Call BroadCastDNSFOUND(s_Verb, s_databaseName, s_shortURL)
    '                            Else
    '                                s_shortURL = s_shortURL
    '                                s_databaseName = s_databaseName
    '                            End If

    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3674 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try

    '                Case "AcknowRemoveTheSite"
    '                    Try
    '                        If s_p_authority = "yes" Then

    '                            Dim s_Verb As String = "AcknowRemoveTheSite"
    '                            Dim s_so As String = "so"
    '                            Dim s_shortURL As String = queryDNA
    '                            Dim myConnectionString As String
    '                            Dim s_databaseName As String = URIToPass

    '                            If Len(s_shortURL) > 0 And Len(s_databaseName) > 0 Then

    '                                Dim objConn As New OleDbConnection(ConnStringURL(s_databaseName))
    '                                myConnectionString = "UPDATE URLFOUND SET Done ='" & (s_so) & "' where ShortURL = '" & (s_shortURL) & "'"
    '                                Dim myCommand As New OleDbCommand(myConnectionString)
    '                                myCommand.Connection = objConn
    '                                objConn.Open()
    '                                myCommand.ExecuteNonQuery()
    '                                objConn.Close()
    '                            Else
    '                                s_shortURL = s_shortURL
    '                                s_databaseName = s_databaseName
    '                            End If
    '                        End If
    '                    Catch ex As Exception
    '                        EventLog.WriteEntry(sSource, " lookforfiles.3699 --> " & ex.Message, EventLogEntryType.Information, 44)
    '                    End Try
    '            End Select

    '        Catch ex As Exception

    '            EventLog.WriteEntry(sSource, " lookforfiles.3705 --> " & ex.Message, EventLogEntryType.Information, 44)

    '        End Try

    '    End Function

    Public i_p_entryLogCounter As Int32
    Public b_p_firstTime As Boolean = True

    Public Sub CheckSizeOfLog()
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"

        Try
            If b_p_firstTime = False Then
                If i_p_entryLogCounter >= 100 Or b_p_firstTime Then
                    i_p_entryLogCounter = 0
                    EventLog.Delete(sLog)
                    EventLog.CreateEventSource(sSource, sLog)
                    b_p_firstTime = False
                End If
                i_p_entryLogCounter = i_p_entryLogCounter + 1
            End If
        Catch e_lookforfiles_2964 As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_2964 --> " & e_lookforfiles_2964.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Public b_p_notInitialize As Boolean = True
    Public d_p_startTime As DateTime

    Function CheckIfTimeToChangeQuery()

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"

        Try

            If b_p_notInitialize Then
                d_p_startTime = Now
                b_p_notInitialize = False
            End If

            Dim i_differenceMinute As Integer = DateDiff(DateInterval.Minute, d_p_startTime, Now)

            If i_differenceMinute >= 10 Then
                b_p_notInitialize = True
                Call fFindWhatToLookFor()
            End If

        Catch e_lookforfiles_3238 As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_3238 --> " & e_lookforfiles_3238.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    '    Sub ThrottleDuration(ByVal duration)

    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        Try
    '            If IsNumeric(duration) Then
    '                If duration > 1 Then
    '                    'EventLog.WriteEntry(sSource, " lookforfiles_3262 wait --> 10000", EventLogEntryType.Information, 9999)
    '                    Console.WriteLine(" lookforfiles_3262 wait --> 10000")
    '                    System.Threading.Thread.Sleep(10000)
    '                ElseIf duration > 0.3 Then
    '                    'EventLog.WriteEntry(sSource, " lookforfiles_3262 wait --> 5000", EventLogEntryType.Information, 9999)
    '                    Console.WriteLine(" lookforfiles_3262 wait --> 5000")
    '                    System.Threading.Thread.Sleep(5000)
    '                ElseIf duration > 0.25 Then
    '                    'EventLog.WriteEntry(sSource, " lookforfiles_3262 wait --> 500", EventLogEntryType.Information, 9999)
    '                    Console.WriteLine(" lookforfiles_3262 wait --> 500")
    '                    System.Threading.Thread.Sleep(500)
    '                End If
    '            Else
    '                Console.WriteLine(" lookforfiles_3262 duration --> not numeric!")
    '            End If
    '        Catch e_lookforfiles_3262 As Exception
    '            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_3262 --> " & e_lookforfiles_3262.Message, EventLogEntryType.Information, 44)
    '        End Try
    '    End Sub

    Public i_p_counterOfSendQuery As Integer

    Sub SendQueryInsertOverNetwork()

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim s_Verb As String = "insert"
        Dim s_no As String = "n"
        Dim s_sa As String = "sa"
        Dim b_putSAToWork As Boolean = False
        Dim s_queryIn As String

        Try

            If i_p_counterOfSendQuery >= 10 Or i_p_counterOfSendQuery = 0 Then

                Dim b_continue As Boolean = True
                Dim mySelectQuery As String = "SELECT s_Query FROM DNAVALUE WHERE s_Authority = '" & s_no & "'"
                Dim myConnection As New OleDbConnection(ConnStringDNA())
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
                myConnection.Open()
                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    s_queryIn = myReader("s_Query")
                    'If CheckQueryWithRange(s_queryIn) = False Then
                    '    Call Insert_In_BROADCAST_QUERY((s_queryIn))
                    '    Call SendParametersToNetwork(s_Verb, s_queryIn)
                    'Else
                    '    Call SendParametersToNetwork(s_Verb, s_queryIn)
                    'End If
                    'b_putSAToWork = True
                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()

                If b_putSAToWork Then
                    Dim myConnString = "UPDATE DNAVALUE SET s_Authority ='" & (s_sa) & "'where s_Query = '" & (s_queryIn) & "'"
                    Dim myCommand2 As New OleDbCommand(myConnString)
                    myCommand2.Connection = myConnection
                    myConnection.Open()
                    myCommand2.ExecuteNonQuery()
                    myCommand2.Connection.Close()
                    myConnection.Close()
                Else
                    Dim myConnString = "UPDATE DNAVALUE SET s_Authority ='" & (s_no) & "' where s_Authority ='" & (s_sa) & "'"
                    Dim myCommand2 As New OleDbCommand(myConnString)
                    myCommand2.Connection = myConnection
                    myConnection.Open()
                    myCommand2.ExecuteNonQuery()
                    myCommand2.Connection.Close()
                    myConnection.Close()
                End If

                i_p_counterOfSendQuery = 1
            Else
                i_p_counterOfSendQuery = i_p_counterOfSendQuery + 1
            End If

        Catch e_lookforfiles_3764 As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.e_lookforfiles_3764 --> " & e_lookforfiles_3764.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub


    '    Function CheckQueryWithRange(ByVal s_queryIn) As Boolean

    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        Try

    '            Dim s_checked As String = "Checked"
    '            Dim b_querySentOnNetwork = False
    '            Dim s_hostName As String
    '            Dim b_continue As Boolean = True
    '            Dim mySelectQuery As String = "SELECT HostName FROM RANGE where active = '" & s_checked & "' and HostName  <> '" & s_r_machineName & "'"
    '            Dim myConnection As New OleDbConnection(ConnStringDNA())
    '            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
    '            myConnection.Open()

    '            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

    '            While myReader.Read()
    '                s_hostName = myReader("HostName")
    '                b_querySentOnNetwork = VerifyHostNameQuery(s_hostName, s_queryIn)
    '                If b_querySentOnNetwork = False Then
    '                    CheckQueryWithRange = False
    '                    Exit While
    '                Else
    '                    CheckQueryWithRange = True
    '                End If
    '            End While

    '            If Not (myReader.IsClosed) Then
    '                myReader.Close()
    '            End If

    '            myConnection.Close()

    '        Catch e_lookforfiles_3428 As Exception
    '            EventLog.WriteEntry(sSource, " e_lookforfiles_3428 --> " & e_lookforfiles_3428.Message, EventLogEntryType.Information, 44)
    '        End Try

    '    End Function


    '    Function VerifyHostNameQuery(ByVal HostName, ByVal s_queryIn) As Boolean

    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        Try

    '            Dim b_sent As Boolean = False
    '            Dim s_hostName, s_query As String
    '            Dim b_continue As Boolean = True
    '            Dim mySelectQuery As String = "SELECT s_Query FROM BROADCAST_QUERY where s_Query = '" & s_queryIn & "' and s_Machine = '" & HostName & "'"
    '            Dim myConnection As New OleDbConnection(ConnStringDNA())
    '            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
    '            myConnection.Open()

    '            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

    '            While myReader.Read()
    '                s_query = myReader("s_Query")
    '                b_sent = True
    '            End While

    '            If Not (myReader.IsClosed) Then
    '                myReader.Close()
    '            End If

    '            myConnection.Close()

    '            If b_sent Then
    '                VerifyHostNameQuery = True
    '            Else
    '                VerifyHostNameQuery = False
    '            End If

    '        Catch e_lookforfiles_3783 As Exception
    '            EventLog.WriteEntry(sSource, " e_lookforfiles_3783 --> " & e_lookforfiles_3783.Message, EventLogEntryType.Information, 44)
    '        End Try

    '    End Function


    '    Sub SendParametersToNetwork(ByVal s_Verb, ByVal s_Query)

    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        Try

    '            Dim sToInsert, myConnectionString As String
    '            Dim sDate As String = Now()
    '            Dim iLevel, iMaxPage As Integer
    '            Dim iDuration As Double
    '            Dim iMinKPage, iMaxKPage, iMinDensity, iMaxDensity As Integer
    '            Dim iKeywordTitle, iKeywordBody, IkeywordTag, iKeywordVisible, iKeywordURL As Integer
    '            Dim iMinDiskPerc, IDiskMinGiga As Integer
    '            Dim iAlgorythm, iGoogle, iAllTheWeb, iExactPhrase As Integer
    '            Dim iSearchExcel, IsearchWord, iSearchPDF, iDesactiveQuery As Integer
    '            Dim s_Priority, sQuery As String
    '            Dim bGoAhead As Boolean = False
    '            Dim bInsert As Boolean = False
    '            Dim bUpdate As Boolean = False
    '            Dim bStopTheInsertion As Boolean = False
    '            Dim sWebSite As String = "n/a"
    '            Dim sNetwork As String = "n/a"
    '            Dim sSecondNetwork As String = "n/a"
    '            Dim sUniversalNetwork As String = "n/a"
    '            Dim sMachineAddress As String = "127.0.0.1"
    '            Dim sSecondaryMachineAddress = "127.0.0.1"
    '            Dim CheckAllTheTime, ComboStartDayWeek, ComboStopDayWeek, ComboStartDay1,
    '            ComboStopDay1, ComboStartDay2, ComboStopDay2, ComboStartDay3,
    '            ComboStopDay3, ComboStartDay4, ComboStopDay4, ComboStartDay5,
    '            ComboStopDay5, ComboStartDay6, ComboStopDay6, ComboStartDay7,
    '            ComboStopDay7 As String
    '            Dim d_dateInsert As DateTime

    '            If s_Verb = "update" Then
    '                bUpdate = True
    '            ElseIf s_Verb = "insert" Then
    '                bUpdate = False
    '            End If

    '            '//----Start the value to do the insert in DNAVALUE

    '            Dim mySelectQuery As String = "SELECT * FROM DNAVALUE where s_Query ='" & s_Query & "'"
    '            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
    '            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

    '            myConnection1.Open()

    '            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
    '            Dim Index As Integer

    '            While myReader.Read()
    '                iDuration = myReader("i_Duration")
    '                iLevel = myReader("i_Level")
    '                iMaxPage = myReader("i_MaxPage")
    '                s_Priority = myReader("s_Priority")
    '                iMinKPage = myReader("i_MinKPage")
    '                iMaxKPage = myReader("i_MaxKPage")
    '                iMinDensity = myReader("i_MinDensity")
    '                iMaxDensity = myReader("i_MaxDensity")
    '                iKeywordTitle = myReader("i_KeywordTitle")
    '                iKeywordBody = myReader("i_KeywordBody")
    '                IkeywordTag = myReader("I_keywordTag")
    '                iKeywordVisible = myReader("i_KeywordVisible")
    '                iKeywordURL = myReader("i_KeywordURL")
    '                iMinDiskPerc = myReader("i_MinDiskPerc")
    '                IDiskMinGiga = myReader("I_MinDiskGiga")
    '                iAlgorythm = myReader("i_Algorythm")
    '                iGoogle = myReader("i_Google")
    '                iAllTheWeb = myReader("i_AllTheWeb")
    '                iSearchExcel = myReader("i_SearchExcel")
    '                IsearchWord = myReader("i_SearchWord")
    '                iSearchPDF = myReader("i_SearchPDF")
    '                iDesactiveQuery = myReader("i_DesactiveQuery")
    '                sQuery = myReader("s_Query")
    '                iExactPhrase = myReader("i_ExactPhrase")
    '                sWebSite = myReader("s_WebSite")
    '                sNetwork = myReader("s_FirstNetwork")
    '                sSecondNetwork = myReader("s_SecondNetwork")
    '                sUniversalNetwork = myReader("s_UniversalNetwork")
    '                d_dateInsert = myReader("d_DateTime")
    '                Exit While
    '            End While

    '            If Not (myReader.IsClosed) Then
    '                myReader.Close()
    '            End If

    '            myConnection1.Close()

    '            '\\----End insert in DNAVALUE


    '            '//----Start the value to do the insert in SCHEDULE

    '            mySelectQuery = ""
    '            mySelectQuery = "SELECT * FROM SCHEDULE where s_IDQuery ='" & s_Query & "'"
    '            Dim myCommand11 As New OleDbCommand(mySelectQuery, myConnection1)

    '            myConnection1.Open()

    '            Dim myReader11 As OleDbDataReader = myCommand11.ExecuteReader()

    '            While myReader11.Read()

    '                CheckAllTheTime = myReader11("s_AllTheTime")
    '                ComboStartDayWeek = myReader11("s_WeekTimeStart")
    '                ComboStopDayWeek = myReader11("s_WeekTimeStop")
    '                ComboStartDay1 = myReader11("s_MondayTimeStart")
    '                ComboStopDay1 = myReader11("s_MondayTimeStop")
    '                ComboStartDay2 = myReader11("s_TuesdayTimeStart")
    '                ComboStopDay2 = myReader11("s_TuesdayTimeStart")
    '                ComboStartDay3 = myReader11("s_WednesdayTimeStart")
    '                ComboStopDay3 = myReader11("s_WednesdayTimeStop")
    '                ComboStartDay4 = myReader11("s_ThursdayTimeStart")
    '                ComboStopDay4 = myReader11("s_ThursdayTimeStop")
    '                ComboStartDay5 = myReader11("s_FridayTimeStart")
    '                ComboStopDay5 = myReader11("s_FridayTimeStop")
    '                ComboStartDay6 = myReader11("s_SaturdayTimeStart")
    '                ComboStopDay6 = myReader11("s_SaturdayTimeStop")
    '                ComboStartDay7 = myReader11("s_SundayTimeStart")
    '                ComboStopDay7 = myReader11("s_SundayTimeStop")

    '                Exit While

    '            End While

    '            If Not (myReader11.IsClosed) Then
    '                myReader11.Close()
    '            End If

    '            myConnection1.Close()

    '            '\\----End insert in DNAVALUE

    '            If bUpdate Then

    '                Dim s_insertDNA As String = ""
    '                s_insertDNA = "UPDATE DNAVALUE SET d_DateTime= '" & (d_dateInsert) & "',s_Priority= '" & (s_Priority) & "', s_Query= '" & (s_Query) & "',i_DesactiveQuery = " & (iDesactiveQuery) & ",i_SearchPDF = " & (iSearchPDF) & ",i_SearchWord = " & (IsearchWord) & ",i_SearchExcel = " & (iSearchExcel) & ",i_ExactPhrase = " & (iExactPhrase) & ",i_Duration =  " & (iDuration) & ",i_Level =  " & (iLevel) & ",i_MaxPage =  " & (iMaxPage) & ",i_MinKPage =  " & (iMinKPage) & ",i_MaxKPage =  " & (iMaxKPage) & ",i_MinDensity = " & (iMinDensity) & ",i_MaxDensity = " & (iMaxDensity) & ",i_KeywordTitle =" & (iKeywordTitle) & ",i_KeywordURL = " & (iKeywordURL) & ",i_KeywordBody =  " & (iKeywordBody) & ",i_KeywordTag =" & (IkeywordTag) & ",i_algorythm =  " & (iAlgorythm) & ",i_KeywordVisible = " & (iKeywordVisible) & ",i_AllTheWeb =" & (iAllTheWeb) & ",i_Google = " & (iGoogle) & ",i_MinDiskPerc =" & (iMinDiskPerc) & ",i_MinDiskGiga =  " & (IDiskMinGiga) & ",s_FirstNetwork =  '" & (sNetwork) & "',s_SecondNetwork =  '" & (sSecondNetwork) & "',s_UniversalNetwork =  '" & (sUniversalNetwork) & "',s_MachineAddress =  '" & (sMachineAddress) & "',s_SecondaryMachineAddress =  '" & (sSecondaryMachineAddress) & "',s_WebSite =  '" & (sWebSite) & "' WHERE s_Query ='" & sQuery & "'"
    '                s_insertDNA = Replace(s_insertDNA, "_", "|")
    '                s_insertDNA = Replace(s_insertDNA, " ", "_")

    '                Dim s_querySchedule As String = ""
    '                s_querySchedule = "UPDATE SCHEDULE SET s_IDQuery= '" & (sQuery) & "',i_DesactiveQuery = '" & (iDesactiveQuery) & "',s_AllTheTime = '" & (CheckAllTheTime) & "',s_WeekTimeStart = '" & (ComboStartDayWeek) & "',s_WeekTimeStop = '" & (ComboStopDayWeek) & "',s_MondayTimeStart = '" & (ComboStartDay1) & "',s_MondayTimeStop = '" & (ComboStopDay1) & "',s_TuesdayTimeStart =  '" & (ComboStartDay2) & "',s_TuesdayTimeStop =  '" & (ComboStopDay2) & "',s_WednesdayTimeStart =  '" & (ComboStartDay3) & "',s_WednesdayTimeStop =  '" & (ComboStopDay3) & "',s_ThursdayTimeStart =  '" & (ComboStartDay4) & "',s_ThursdayTimeStop = '" & (ComboStopDay4) & "',s_FridayTimeStart = '" & (ComboStartDay5) & "',s_FridayTimeStop ='" & (ComboStopDay5) & "',s_SaturdayTimeStart = '" & (ComboStartDay6) & "',s_SaturdayTimeStop =  '" & (ComboStopDay6) & "',s_SundayTimeStart ='" & (ComboStartDay7) & "',s_SundayTimeStop =  '" & (ComboStopDay7) & "' WHERE s_IDQuery ='" & sQuery & "'"
    '                s_querySchedule = Replace(s_querySchedule, "_", "|")
    '                s_querySchedule = Replace(s_querySchedule, " ", "_")

    '                sURIToPass = s_insertDNA & "|||||" & s_querySchedule

    '                s_Verb = "Sayonara!"

    '                Call BroadCastDNSFOUND(s_Verb)

    '                s_Verb = "update"

    '                Call BroadCastDNSFOUND(s_Verb, sURIToPass, , , s_r_machineName)

    '            Else

    '                Dim s_insertDNA As String =
    '                "INSERT INTO DNAVALUE (s_Query," _
    '                & "s_Priority," _
    '                & "i_DesactiveQuery," _
    '                & "i_SearchPDF," _
    '                & "i_SearchWord," _
    '                & "i_SearchExcel," _
    '                & "i_ExactPhrase," _
    '                & "i_Duration," _
    '                & "i_Level," _
    '                & "i_MaxPage," _
    '                & "i_MinKPage," _
    '                & "i_MaxKPage," _
    '                & "i_MinDensity," _
    '                & "i_MaxDensity," _
    '                & "i_KeywordTitle," _
    '                & "i_KeywordURL," _
    '                & "i_KeywordBody," _
    '                & "i_KeywordTag," _
    '                & "i_algorythm," _
    '                & "i_KeywordVisible," _
    '                & "i_AllTheWeb," _
    '                & "i_Google," _
    '                & "i_MinDiskPerc," _
    '                & "i_MinDiskGiga," _
    '                & "s_FirstNetwork," _
    '                & "s_SecondNetwork," _
    '                & "s_UniversalNetwork," _
    '                & "s_MachineAddress," _
    '                & "s_SecondaryMachineAddress," _
    '                & "d_DateTime," _
    '                & "s_WebSite)" _
    '                & "Values('" & (sQuery) & "'," _
    '                & "'" & (s_Priority) & "'," _
    '                & "" & (iDesactiveQuery) & "," _
    '                & "" & (iSearchPDF) & "," _
    '                & "" & (IsearchWord) & "," _
    '                & "" & (iSearchExcel) & "," _
    '                & "" & (iExactPhrase) & "," _
    '                & "" & (iDuration) & "," _
    '                & "" & (iLevel) & "," _
    '                & "" & (iMaxPage) & "," _
    '                & "" & (iMinKPage) & "," _
    '                & "" & (iMaxKPage) & "," _
    '                & "" & (iMinDensity) & "," _
    '                & "" & (iMaxDensity) & "," _
    '                & "" & (iKeywordTitle) & "," _
    '                & "" & (iKeywordURL) & "," _
    '                & "" & (iKeywordBody) & "," _
    '                & "" & (IkeywordTag) & "," _
    '                & "" & (iAlgorythm) & "," _
    '                & "" & (iKeywordVisible) & "," _
    '                & "" & (iAllTheWeb) & "," _
    '                & "" & (iGoogle) & "," _
    '                & "" & (iMinDiskPerc) & "," _
    '                & "" & (IDiskMinGiga) & "," _
    '                & "'" & (sNetwork) & "'," _
    '                & "'" & (sSecondNetwork) & "'," _
    '                & "'" & (sUniversalNetwork) & "'," _
    '                & "'" & (sMachineAddress) & "'," _
    '                & "'" & (sSecondaryMachineAddress) & "'," _
    '                & "'" & (d_dateInsert) & "'," _
    '                & "'" & (sWebSite) & "')"

    '                s_insertDNA = Replace(s_insertDNA, "'yes'", "'no'")
    '                s_insertDNA = Replace(s_insertDNA, "_", "||")
    '                s_insertDNA = Replace(s_insertDNA, " ", "_")

    '                Dim s_querySchedule As String = ""
    '                s_querySchedule = "INSERT INTO SCHEDULE (s_IDQuery,i_DesactiveQuery,s_AllTheTime,s_WeekTimeStart,s_WeekTimeStop,s_MondayTimeStart,s_MondayTimeStop,s_TuesdayTimeStart,s_TuesdayTimeStop,s_WednesdayTimeStart,s_WednesdayTimeStop,s_ThursdayTimeStart,s_ThursdayTimeStop,s_FridayTimeStart,s_FridayTimeStop,s_SaturdayTimeStart,s_SaturdayTimeStop,s_SundayTimeStart,s_SundayTimeStop) Values('" & (sQuery) & "','" & (iDesactiveQuery) & "','" & (CheckAllTheTime) & "','" & (ComboStartDayWeek) & "','" & (ComboStopDayWeek) & "','" & (ComboStartDay1) & "','" & (ComboStopDay1) & "','" & (ComboStartDay2) & "','" & (ComboStopDay2) & "','" & (ComboStartDay3) & "','" & (ComboStopDay3) & "','" & (ComboStartDay4) & "','" & (ComboStopDay4) & "','" & (ComboStartDay5) & "','" & (ComboStopDay5) & "','" & (ComboStartDay6) & "','" & (ComboStopDay6) & "','" & (ComboStartDay7) & "','" & (ComboStopDay7) & "')"
    '                s_querySchedule = Replace(s_querySchedule, "_", "||")
    '                s_querySchedule = Replace(s_querySchedule, " ", "_")

    '                sURIToPass = s_Query

    '                sURIToPass = Replace(sURIToPass, " ", "_")

    '                s_insertDNA = Replace(s_insertDNA, "yes", "no")

    '                s_Verb = "insert"

    '                Call BroadCastDNSFOUND(s_Verb, sURIToPass, s_insertDNA, s_querySchedule, s_r_machineName)

    '            End If

    '        Catch e_lookforfiles_3705 As Exception
    '            EventLog.WriteEntry(sSource, " e_lookforfiles_3705 --> " & e_lookforfiles_3705.Message, EventLogEntryType.Information, 44)
    '        End Try

    '    End Sub

    Public i_p_sendQueryUpdateCounter As Int32

    Sub SendQueryUpdateOverNetwork()

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim s_queryToUpdate As String = ""

        Try
            If i_p_sendQueryUpdateCounter >= 1 Or i_p_sendQueryUpdateCounter = 0 Then

                Dim b_sent As Boolean = False
                Dim s_verb As String = "update"
                Dim b_putSAToWork As Boolean = False
                Dim s_upd As String = "u"
                Dim s_no As String = "n"
                Dim s_sa As String = "sa"

                Dim b_continue As Boolean = True
                Dim mySelectQuery As String = "SELECT distinct(s_Query) FROM BROADCAST_QUERY where s_Active = '" & s_upd & "' and s_Delete ='" & s_no & "' "
                Dim myConnection As New OleDbConnection(ConnStringDNA())
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
                myConnection.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    s_queryToUpdate = myReader("s_Query")
                    'Call SendParametersToNetwork(s_verb, s_queryToUpdate)
                    b_putSAToWork = True
                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()

                If b_putSAToWork Then
                    Dim myConnString = "UPDATE BROADCAST_QUERY SET s_Active ='" & (s_sa) & "' where s_Active = '" & s_upd & "' and s_Query = '" & (s_queryToUpdate) & "'"
                    Dim myCommand2 As New OleDbCommand(myConnString)
                    myCommand2.Connection = myConnection
                    myConnection.Open()
                    myCommand2.ExecuteNonQuery()
                    myCommand2.Connection.Close()
                    myConnection.Close()
                Else
                    Dim myConnString = "UPDATE BROADCAST_QUERY SET s_Active ='" & (s_upd) & "'where s_Active = '" & s_sa & "'"
                    Dim myCommand2 As New OleDbCommand(myConnString)
                    myCommand2.Connection = myConnection
                    myConnection.Open()
                    myCommand2.ExecuteNonQuery()
                    myCommand2.Connection.Close()
                    myConnection.Close()
                End If

                i_p_sendQueryUpdateCounter = 1
            Else
                i_p_sendQueryUpdateCounter = i_p_sendQueryUpdateCounter + 1
            End If

        Catch e_lookforfiles_3780 As Exception
            EventLog.WriteEntry(sSource, " e_lookforfiles_3780 --> " & e_lookforfiles_3780.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Public i_p_counterOfSendQueryDelete As Int32

    Sub SendQueryDeleteOverNetwork()

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"

        Try
            If i_p_counterOfSendQueryDelete = 0 Or i_p_counterOfSendQueryDelete >= 50 Then
                Dim b_sent As Boolean = False
                Dim s_yes As String = "y"
                Dim s_queryToDelete As String = ""
                Dim b_continue As Boolean = True
                Dim mySelectQuery As String = "SELECT distinct(s_Query) FROM BROADCAST_QUERY where s_Delete = '" & s_yes & "'"
                Dim myConnection As New OleDbConnection(ConnStringDNA())
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
                myConnection.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    s_queryToDelete = myReader("s_Query")
                    Exit While
                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()

                If Len(s_queryToDelete) > 0 Then
                    Dim s_verb As String = "delete"
                    s_queryToDelete = s_queryToDelete & "|||" & s_r_machineName
                    'Call SendQueryToDeleteToNetwork(s_verb, s_queryToDelete)
                    Dim s_sa As String = "sa"
                    Dim myConnString = "UPDATE BROADCAST_QUERY SET s_Delete ='" & (s_sa) & "' where s_Query = '" & (s_queryToDelete) & "'"
                    Dim myCommand2 As New OleDbCommand(myConnString)
                    myCommand2.Connection = myConnection
                    myConnection.Open()
                    myCommand2.ExecuteNonQuery()
                    myCommand2.Connection.Close()
                    myConnection.Close()
                Else
                    Dim s_sa As String = "sa"
                    Dim myConnString = "UPDATE BROADCAST_QUERY SET s_Delete ='" & (s_yes) & "' where s_Delete = '" & (s_sa) & "'"
                    Dim myCommand2 As New OleDbCommand(myConnString)
                    myCommand2.Connection = myConnection
                    myConnection.Open()
                    myCommand2.ExecuteNonQuery()
                    myCommand2.Connection.Close()
                    myConnection.Close()
                End If
                i_p_counterOfSendQueryDelete = 1
            Else
                i_p_counterOfSendQueryDelete = i_p_counterOfSendQueryDelete + 1
            End If
        Catch e_lookforfiles_3780 As Exception
            EventLog.WriteEntry(sSource, " e_lookforfiles_3780 --> " & e_lookforfiles_3780.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub


    '    Sub SendQueryToDeleteToNetwork(ByVal s_verb, ByVal s_queryToDelete)

    '        Call BroadCastDNSFOUND(s_verb, s_queryToDelete, , , s_r_machineName)

    '    End Sub


    '    Sub DeleteInBroadcastQuery(ByVal s_uriToPass, ByVal s_machineName)

    '        Try

    '            Dim s_so As String = "so"
    '            Dim s_no As String = "n"

    '            Dim myConnection As New OleDbConnection(ConnStringDNA())

    '            Dim myConnString = "UPDATE BROADCAST_QUERY SET s_Active ='" & (s_no) & "', s_Delete ='" & (s_so) & "' where s_Query = '" & (s_uriToPass) & "' and s_Machine = '" & (s_machineName) & "'"

    '            Dim myCommand2 As New OleDbCommand(myConnString)
    '            myCommand2.Connection = myConnection
    '            myConnection.Open()
    '            myCommand2.ExecuteNonQuery()
    '            myCommand2.Connection.Close()
    '            myConnection.Close()

    '        Catch e_lookforfiles_3891 As Exception
    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " e_lookforfiles_3891 --> " & e_lookforfiles_3891.Message, EventLogEntryType.Information, 44)
    '        End Try

    '    End Sub


    Function keepASCIIOnly(ByVal s_string) As String
        Dim b_continue As Boolean = True
        Dim s_stringToKeep, s_characterCheck As String
        Dim i_countCharacter, i_lenght As Int32
        i_countCharacter = 0
        i_lenght = Len(s_string)
        i_countCharacter = 1

        Try
            If i_lenght > 0 Then
                Dim frq, counter1, counter2 As Long
                Dim duration As Double

                QueryPerformanceFrequency(frq)
                QueryPerformanceCounter(counter1)

                Dim queryBuilder As New StringBuilder(s_string.ToString)

                'Console.WriteLine(s_string)
                If s_string <> "stop" Then
                    queryBuilder.Replace(Chr(1), " ")
                    queryBuilder.Replace(Chr(1), " ")
                    queryBuilder.Replace(Chr(2), " ")
                    queryBuilder.Replace(Chr(3), " ")
                    queryBuilder.Replace(Chr(4), " ")
                    queryBuilder.Replace(Chr(5), " ")
                    queryBuilder.Replace(Chr(6), " ")
                    queryBuilder.Replace(Chr(7), " ")
                    queryBuilder.Replace(Chr(8), " ")
                    queryBuilder.Replace(Chr(9), " ")
                    queryBuilder.Replace(Chr(10), " ")
                    queryBuilder.Replace(Chr(11), " ")
                    queryBuilder.Replace(Chr(12), " ")
                    queryBuilder.Replace(Chr(13), " ")
                    queryBuilder.Replace(Chr(14), " ")
                    queryBuilder.Replace(Chr(15), " ")
                    queryBuilder.Replace(Chr(16), " ")
                    queryBuilder.Replace(Chr(17), " ")
                    queryBuilder.Replace(Chr(18), " ")
                    queryBuilder.Replace(Chr(19), " ")
                    queryBuilder.Replace(Chr(20), " ")
                    queryBuilder.Replace(Chr(21), " ")
                    queryBuilder.Replace(Chr(22), " ")
                    queryBuilder.Replace(Chr(23), " ")
                    queryBuilder.Replace(Chr(24), " ")
                    queryBuilder.Replace(Chr(25), " ")
                    queryBuilder.Replace(Chr(26), " ")
                    queryBuilder.Replace(Chr(27), " ")
                    queryBuilder.Replace(Chr(28), " ")
                    queryBuilder.Replace(Chr(29), " ")
                    queryBuilder.Replace(Chr(30), " ")
                    queryBuilder.Replace(Chr(31), " ")
                    queryBuilder.Replace(Chr(32), " ")
                    queryBuilder.Replace(Chr(33), " ")
                    queryBuilder.Replace(Chr(34), " ")
                    queryBuilder.Replace(Chr(35), " ")
                    queryBuilder.Replace(Chr(36), " ")
                    queryBuilder.Replace(Chr(37), " ")
                    queryBuilder.Replace(Chr(38), " ")
                    queryBuilder.Replace(Chr(39), " ")
                    queryBuilder.Replace(Chr(123), " ")
                    queryBuilder.Replace(Chr(124), " ")
                    queryBuilder.Replace(Chr(125), " ")
                    queryBuilder.Replace(Chr(126), " ")
                    queryBuilder.Replace(Chr(127), " ")
                    queryBuilder.Replace(Chr(128), " ")
                    queryBuilder.Replace(Chr(129), " ")
                    queryBuilder.Replace(Chr(130), " ")
                    queryBuilder.Replace(Chr(131), " ")
                    queryBuilder.Replace(Chr(132), " ")
                    queryBuilder.Replace(Chr(133), " ")
                    queryBuilder.Replace(Chr(134), " ")
                    queryBuilder.Replace(Chr(135), " ")
                    queryBuilder.Replace(Chr(136), " ")
                    queryBuilder.Replace(Chr(137), " ")
                    queryBuilder.Replace(Chr(138), " ")
                    queryBuilder.Replace(Chr(139), " ")
                    queryBuilder.Replace(Chr(140), " ")
                    queryBuilder.Replace(Chr(141), " ")
                    queryBuilder.Replace(Chr(142), " ")
                    queryBuilder.Replace(Chr(143), " ")
                    queryBuilder.Replace(Chr(144), " ")
                    queryBuilder.Replace(Chr(145), " ")
                    queryBuilder.Replace(Chr(146), " ")
                    queryBuilder.Replace(Chr(147), " ")
                    queryBuilder.Replace(Chr(148), " ")
                    queryBuilder.Replace(Chr(149), " ")
                    queryBuilder.Replace(Chr(150), " ")
                    queryBuilder.Replace(Chr(151), " ")
                    queryBuilder.Replace(Chr(152), " ")
                    queryBuilder.Replace(Chr(153), " ")
                    queryBuilder.Replace(Chr(154), " ")
                    queryBuilder.Replace(Chr(155), " ")
                    queryBuilder.Replace(Chr(156), " ")
                    queryBuilder.Replace(Chr(157), " ")
                    queryBuilder.Replace(Chr(158), " ")
                    queryBuilder.Replace(Chr(159), " ")
                    queryBuilder.Replace(Chr(160), " ")
                    queryBuilder.Replace(Chr(161), " ")
                    queryBuilder.Replace(Chr(162), " ")
                    queryBuilder.Replace(Chr(163), " ")
                    queryBuilder.Replace(Chr(164), " ")
                    queryBuilder.Replace(Chr(165), " ")
                    queryBuilder.Replace(Chr(166), " ")
                    queryBuilder.Replace(Chr(167), " ")
                    queryBuilder.Replace(Chr(168), " ")
                    queryBuilder.Replace(Chr(169), " ")
                    queryBuilder.Replace(Chr(170), " ")
                    queryBuilder.Replace(Chr(171), " ")
                    queryBuilder.Replace(Chr(172), " ")
                    queryBuilder.Replace(Chr(173), " ")
                    queryBuilder.Replace(Chr(174), " ")
                    queryBuilder.Replace(Chr(175), " ")
                    queryBuilder.Replace(Chr(176), " ")
                    queryBuilder.Replace(Chr(177), " ")
                    queryBuilder.Replace(Chr(178), " ")
                    queryBuilder.Replace(Chr(179), " ")
                    queryBuilder.Replace(Chr(180), " ")
                    queryBuilder.Replace(Chr(181), " ")
                    queryBuilder.Replace(Chr(182), " ")
                    queryBuilder.Replace(Chr(183), " ")
                    queryBuilder.Replace(Chr(184), " ")
                    queryBuilder.Replace(Chr(185), " ")
                    queryBuilder.Replace(Chr(186), " ")
                    queryBuilder.Replace(Chr(187), " ")
                    queryBuilder.Replace(Chr(188), " ")
                    queryBuilder.Replace(Chr(189), " ")
                    queryBuilder.Replace(Chr(190), " ")
                    queryBuilder.Replace(Chr(191), " ")
                    queryBuilder.Replace(Chr(192), " ")
                    queryBuilder.Replace(Chr(245), " ")
                    queryBuilder.Replace(Chr(246), " ")
                    queryBuilder.Replace(Chr(247), " ")
                    queryBuilder.Replace(Chr(248), " ")
                    queryBuilder.Replace(Chr(249), " ")
                    queryBuilder.Replace(Chr(250), " ")
                    queryBuilder.Replace(Chr(251), " ")
                    queryBuilder.Replace(Chr(252), " ")
                    queryBuilder.Replace(Chr(253), " ")
                    queryBuilder.Replace(Chr(254), " ")
                    queryBuilder.Replace(Chr(255), " ")
                    queryBuilder.Replace(Chr(0), " ")

                    queryBuilder.Replace("         ", " ")
                    queryBuilder.Replace("        ", " ")
                    queryBuilder.Replace("       ", " ")
                    queryBuilder.Replace("      ", " ")
                    queryBuilder.Replace("     ", " ")
                    queryBuilder.Replace("    ", " ")
                    queryBuilder.Replace("   ", " ")
                    queryBuilder.Replace("  ", " ")


                    s_stringToKeep = queryBuilder.ToString()

                    If Len(s_stringToKeep) > 0 Then
                        keepASCIIOnly = s_stringToKeep
                    Else
                        s_string = Replace(s_string, Chr(39), " ")
                        s_string = Replace(s_string, "'", " ")
                        keepASCIIOnly = s_string
                    End If

                Else
                    s_string = Replace(s_string, Chr(39), " ")
                    s_string = Replace(s_string, "'", " ")
                    keepASCIIOnly = s_string
                End If

                QueryPerformanceCounter(counter2)
                duration = (counter2 - counter1) / frq
                Console.WriteLine("---------------------------------------------------------------------------")
                Console.WriteLine(duration & "   -- the lenght is ->" & i_lenght & " --  lookforfiles.keepASCIIOnly.4299")
                Console.WriteLine("---------------------------------------------------------------------------")
            End If
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_lookforfiles_4761 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public b_p_newContentFound As Boolean
    Public b_p_firstTimeInDone As Boolean

    Function PutPipeAroundNewContent(ByVal s_pageContent As String, ByVal sToInsert As String, ByVal sSendstrURI As String) As String
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"

        Try

            'If LCase(sSendstrURI) = "golf" Then
            '    sSendstrURI = sSendstrURI
            'End If

            Dim myConnString, s_keywordsToSearchOriginal, s_cutTextWithTags As String
            Dim s_urlPage As String
            Dim sSendUri As String = sToInsert
            Dim SUrl As String = "n"
            Dim b_continue As Boolean = True
            Dim b_cuttingKeywordsDone As Boolean
            Dim i_posKeywords, i_cutText, i_posCutText, i_startSearch, i_compareInURL, i_pointer, i_pointerEnd As Int64
            Dim s_cutText, s_keywordsToSearch, s_newPageContent As String
            Dim b_exitDoArrayZero As Boolean
            Dim b_goodLenght As Boolean
            b_p_newContentFound = False

            Dim mySelectQuery As String = "SELECT TOP 1 URL_PAGE FROM URLFOUND where  GoodSpot <> -40004 and URL ='" & Trim(sToInsert) & "' order by IDCounter Desc"
            Dim myConnection As New OleDbConnection(ConnStringURL(sSendstrURI))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            While myReader.Read()
                s_urlPage = myReader("URL_PAGE")
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection.Close()

            If Len(s_urlPage) > 0 Then

                s_urlPage = Replace(s_urlPage, "...->", "")

                i_startSearch = 1

                sSendstrURI = Replace(LCase(sSendstrURI), "(", "")
                sSendstrURI = Replace(LCase(sSendstrURI), ")", "")
                sSendstrURI = Replace(LCase(sSendstrURI), " or ", " ")
                sSendstrURI = Replace(LCase(sSendstrURI), " and ", " ")
                sSendstrURI = Replace(LCase(sSendstrURI), " et ", " ")
                sSendstrURI = Replace(LCase(sSendstrURI), " ou ", " ")
                sSendstrURI = Replace(LCase(sSendstrURI), "  ", " ")
                sSendstrURI = Replace(LCase(sSendstrURI), "  ", " ")
                sSendstrURI = Replace(LCase(sSendstrURI), "  ", " ")

                s_keywordsToSearch = LCase(sSendstrURI)

                Dim s_array() As String
                s_array = Split(s_keywordsToSearch)

                Dim i_countUboundArray As Integer = UBound(s_array)

                If i_countUboundArray = 0 Then
                    b_exitDoArrayZero = True
                End If

                Try
                    i_posKeywords = InStr(i_startSearch, LCase(s_pageContent), s_keywordsToSearch, CompareMethod.Text)
                Catch ex As Exception
                    EventLog.WriteEntry(sSource, " lookforfiles.4832 --> " & ex.Message, EventLogEntryType.Information, 44)
                End Try

                Do While b_continue

                    If i_posKeywords = 0 Then
                        If b_cuttingKeywordsDone Or b_exitDoArrayZero Then
                            Exit Do
                        End If

                        s_keywordsToSearch = LCase(f_KeywordsToFind(sSendstrURI))
                        b_p_firstTimeInDone = True
                        i_posKeywords = 1
                        i_startSearch = 1
                        If Len(s_r_KeepKeywords) = 0 Then
                            b_cuttingKeywordsDone = True
                        End If
                    Else

                        Dim i_lenKeyword As Int64 = Len(s_keywordsToSearch)

                        Try
                            Dim i_keepPosKeywords As Integer = i_posKeywords - 5
                            If i_keepPosKeywords < 1 Then
                                i_keepPosKeywords = 1
                            End If
                            s_cutText = Trim(Mid(s_pageContent, i_keepPosKeywords, i_lenKeyword + 10))
                        Catch ex As Exception
                            EventLog.WriteEntry(sSource, " lookforfiles.4855 --> " & ex.Message, EventLogEntryType.Information, 44)
                        End Try

                        If b_exitDoArrayZero = False Then
                            b_goodLenght = CheckLenghtOfKeywordsWithText(s_cutText, s_keywordsToSearch)
                        End If

                        If b_goodLenght Or b_exitDoArrayZero Then

                            Try
                                If (i_posKeywords - 100) < 0 Then
                                    i_pointer = 1
                                Else
                                    i_pointer = i_posKeywords - 100
                                End If

                                i_pointerEnd = Len(s_keywordsToSearch) + 200
                                If i_pointer + i_pointerEnd > Len(s_pageContent) Then
                                    i_pointerEnd = Len(s_pageContent)
                                End If

                                s_cutTextWithTags = Mid(s_pageContent, i_pointer, i_pointerEnd)

                            Catch ex As Exception
                                EventLog.WriteEntry(sSource, " lookforfiles.4868 --> " & ex.Message, EventLogEntryType.Information, 44)
                            End Try

                            Dim i_foundMarker As Int64 = InStr(s_cutTextWithTags, "...->", CompareMethod.Binary)

                            If i_foundMarker = 0 Then

                                i_compareInURL = InStr(LCase(s_urlPage), LCase(s_cutText), CompareMethod.Text)

                                If i_compareInURL = 0 Then
                                    Dim s_cutTextToInsert As String = Replace(LCase(s_cutText), LCase(s_keywordsToSearch), "...->" & s_keywordsToSearch, CompareMethod.Text)
                                    s_pageContent = Replace(s_pageContent, s_cutText, s_cutTextToInsert)
                                    b_p_newContentFound = True
                                End If

                            End If
                        End If
                    End If

                    If Len(s_keywordsToSearch) > 0 Then
                        i_startSearch = i_posKeywords + Len(s_keywordsToSearch) + 5
                        Try
                            i_posKeywords = InStr(i_startSearch, LCase(s_pageContent), LCase(s_keywordsToSearch), CompareMethod.Text)
                        Catch ex As Exception
                            EventLog.WriteEntry(sSource, " lookforfiles.4892 --> " & ex.Message, EventLogEntryType.Information, 44)
                        End Try
                    Else
                        i_posKeywords = 0
                    End If

                Loop

            End If

            b_p_firstTimeInDone = False

            s_p_DatabaseName = ""

            PutPipeAroundNewContent = s_pageContent

        Catch ex As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.4911 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function CheckLenghtOfKeywordsWithText(ByVal s_cutText, ByVal s_keywordsToSearch) As Boolean
        Try
            Dim s_arrayKeywords() As String
            s_arrayKeywords = Split(s_keywordsToSearch)
            Dim s_arrayCutText() As String
            s_arrayCutText = Split(s_cutText)
            Dim i_uboundKeywords As Integer = UBound(s_arrayKeywords)
            Dim i_uboundCutText As Integer = UBound(s_arrayCutText)
            Dim i_countA, i_countB As Integer
            Dim i_placeKeywords, i_lenKeywords, i_lenCutText As Integer
            Dim s_keywords, s_cText As String
            Dim b_goodLenghtKeywords As Boolean
            Dim b_notGoodLenghtKeywords As Boolean
            Dim i_len As Integer

            For i_countA = 0 To i_uboundKeywords
                For i_countB = 0 To i_uboundCutText
                    i_placeKeywords = InStr(LCase(s_arrayCutText(i_countB)), LCase(s_arrayKeywords(i_countA)), CompareMethod.Text)
                    If i_placeKeywords > 0 Then

                        s_keywords = s_arrayKeywords(i_countA)
                        s_cText = s_arrayCutText(i_countB)
                        s_cText = Trim(Replace(s_cText, ".", ""))
                        s_cText = Trim(Replace(s_cText, ",", ""))
                        s_cText = Trim(Replace(s_cText, "?", ""))
                        s_cText = Trim(Replace(s_cText, "/", ""))
                        s_cText = Trim(Replace(s_cText, "\", ""))
                        s_cText = Trim(Replace(s_cText, "!", ""))
                        s_cText = Trim(Replace(s_cText, ";", ""))
                        s_cText = Trim(Replace(s_cText, ":", ""))
                        s_cText = Trim(Replace(s_cText, """", ""))
                        s_cText = Trim(Replace(s_cText, "'", ""))
                        s_cText = Trim(Replace(s_cText, "<", ""))
                        s_cText = Trim(Replace(s_cText, ">", ""))
                        s_cText = Trim(Replace(s_cText, "(", ""))
                        s_cText = Trim(Replace(s_cText, ")", ""))
                        s_cText = Trim(Replace(s_cText, "[", ""))
                        s_cText = Trim(Replace(s_cText, "]", ""))
                        s_cText = Trim(Replace(s_cText, "{", ""))
                        s_cText = Trim(Replace(s_cText, "}", ""))


                        i_lenKeywords = Len(s_keywords)
                        i_lenCutText = Len(s_cText)

                        i_len = InStr(Len(s_cText), LCase(s_cText), "s", CompareMethod.Text)

                        If i_lenKeywords = i_lenCutText Or (i_lenKeywords + 1 = i_lenCutText And i_len = i_lenCutText) Then
                            b_goodLenghtKeywords = True
                            Exit For
                        Else
                            b_notGoodLenghtKeywords = True
                        End If
                    End If
                Next

                If b_goodLenghtKeywords = False And b_notGoodLenghtKeywords Then
                    CheckLenghtOfKeywordsWithText = False
                    Exit Function
                Else
                    CheckLenghtOfKeywordsWithText = True
                End If
            Next

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.4898 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Sub SendParametersAndCleanup()

        Try
            If s_p_authority = "yes" Then

                Call DistributeAlphaRange()

                Call sendNewMachineConfiguration()

                Call SendQueryInsertOverNetwork()

                Call SendQueryUpdateOverNetwork()

                Call Check_BROADCAST_QUERY_SO()

                Call SendQueryDeleteOverNetwork()

                'Call FindSeeds1()

                'Call RemoveTheSite(s_DatabaseName)

            Else

                Call PublishMachineName()

            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.5031 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub ShowFunction(ByVal s_function, ByVal i_lineNumber, ByVal b_showDebug)

        If b_showDebug Then

            Console.WriteLine("ShowFunction --> " & i_lineNumber & " " & s_function)

        End If

    End Sub


End Module





