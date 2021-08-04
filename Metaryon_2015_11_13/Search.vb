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

Module Search_copie

    Public s_p_SeedOrScan As String

    Public Sub fFindWhatToLookFor()

        'Call fFindWhatToLookForFU()

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"

        Try

            Dim sToInsert As String
            Dim myConnectionString As String
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim sDate As String = Now()
            Dim sCheckString As String
            Dim myConnString As String
            Dim sKeepString As String
            Dim sDNSFound As String
            Dim sKeepDNSFound(20) As String
            Dim iKeepDansFoundCounter As Integer = 0
            Dim iSimpleCounter As Integer
            Dim mySelectQuery As String
            Dim sSeedOrScan As String = s_p_SeedOrScan
            Dim b_goAhead As Boolean = False
            Dim i_keepTheStringToGo, i_startWeek, i_stopWeek As Int32
            Dim s_startWeek, s_stopWeek, s_allTheTime As String
            Dim iKeepNowHour As Int32

            iKeepNowHour = DateTime.Now.Hour.ToString
            If iKeepNowHour = 0 Then iKeepNowHour = 1

            mySelectQuery = "SELECT * FROM DNAVALUE, SCHEDULE where s_Query = s_IDQuery and DNAVALUE.i_DesactiveQuery = 0 order by s_IDQuery"

            Dim myCommand1 As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand1.ExecuteReader()

            While myReader.Read()
                sToInsert = (myReader("s_Query"))
                s_allTheTime = (myReader("s_AllTheTime"))
                s_startWeek = (myReader("s_WeekTimeStart"))
                If IsNumeric(s_startWeek) Then
                    i_startWeek = CInt(s_startWeek)
                End If
                s_stopWeek = (myReader("s_WeekTimeStop"))
                If IsNumeric(s_stopWeek) Then
                    i_stopWeek = CInt(s_stopWeek)
                End If
                If s_allTheTime = "1" Then
                    b_goAhead = True
                    Exit While
                ElseIf i_startWeek < i_stopWeek Then
                    If i_startWeek <= iKeepNowHour And i_stopWeek >= iKeepNowHour Then
                        b_goAhead = True
                        Exit While
                    End If
                ElseIf i_startWeek > i_stopWeek Then
                    If (i_startWeek <= iKeepNowHour And iKeepNowHour <= 24) Or (iKeepNowHour >= 0 And iKeepNowHour < i_stopWeek) Then
                        b_goAhead = True
                        Exit While
                    End If
                End If
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            If b_goAhead And Not (sSeedOrScan = "Seed") Then
                Call StartScanning(sToInsert)
            End If

        Catch e_search_428 As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.e_search_428 --> " & e_search_428.Message, EventLogEntryType.Information, 44)

        End Try

    End Sub


    Function InsertUrl(ByVal SUrlInsert As String, ByVal sSendstrURI As String)

        Try
            Dim sToInsert As String = SUrlInsert
            Dim myConnectionString As String
            Dim myConnection As New OleDbConnection(ConnStringURL(sSendstrURI))
            Dim sDate As Date = Now()

            sDate = sDate.AddDays(-3)

            Dim sNo As String = "y"

            If Len(sToInsert) > 0 And Len(sToInsert) <= 250 Then

                Dim s_shortURL As String = Func_CleanUrl(sToInsert)

                Dim myInsertQuery As String = "INSERT INTO URLFOUND (url,DateHour,Done,ShortURL) Values('" & Trim(sToInsert) & "','" & (sDate) & "','" & (sNo) & "','" & (s_shortURL) & "')"
                Dim myCommand As New OleDbCommand(myInsertQuery)
                myCommand.Connection = myConnection
                myConnection.Open()
                myCommand.ExecuteNonQuery()
                myCommand.Connection.Close()
                '''myConnection = Nothing
                myConnection.Close()
            End If

        Catch e_search_589 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_search_589 --> " & e_search_589.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Function InsertSite(ByVal SUrlInsert As String)

        Try
            Dim sToInsert As String = SUrlInsert
            'MessageBox.Show(sToInsert, "InsertSite 1 de 2")
            Dim myConnectionString As String
            ' If the connection string is null, use a default.
            Dim myConnection As New OleDbConnection(ConnStringDNS())
            Dim sDate As String = Now()
            Dim iKeepSlash As Integer
            Dim sCheckString As String
            iKeepSlash = InStr(8, sToInsert, "/")

            If iKeepSlash > 0 Then
                sToInsert = Mid(sToInsert, 1, iKeepSlash)
            End If

            If Len(sToInsert) > 0 Then

                Dim myConnString As String
                ' If the connection string is null, use a default.
                Dim sKeepString As String
                Dim sDNSFound As String
                sToInsert = Replace(sToInsert, "'", "")
                'MessageBox.Show(sToInsert, "InsertSite 2 de 3")
                Dim mySelectQuery As String = "SELECT URL_Site FROM DNSFound where URL_Site ='" & sToInsert & "' and Query ='" & sURIToPass & "'"
                'MessageBox.Show(mySelectQuery, "mySelectQuery InsertSite 3 de 3")

                Dim myCommand1 As New OleDbCommand(mySelectQuery, myConnection)
                myConnection.Open()
                'Try
                Dim myReader As OleDbDataReader = myCommand1.ExecuteReader()

                While myReader.Read()
                    sDNSFound = sDNSFound & (myReader.GetString(0)) & Chr(10)
                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                '''myConnection = Nothing
                myConnection.Close()

                'MessageBox.Show(sDNSFound, "sDNSFound2")
                If Len(sDNSFound) > 0 Then
                    sKeepString = "stop"
                    'MessageBox.Show(sToInsert, "stop")
                Else
                    'InsertDNSFound(sToInsert)
                    'MessageBox.Show(sToInsert, "InsertDNSFound")
                End If
                'Catch
                '    sKeepString = "stop"
                'End Try
            End If
        Catch e_search_1642 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_search_1642 --> " & e_search_1642.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_LookForFiles_955) is " & e_LookForFiles_955.Message)
        End Try
    End Function


    'Function LookYahooGoogle(ByVal strURI As String) As String
    '    Try
    '        Dim SGetText As String
    '        Dim sGetHttp As String
    '        Dim sYahoo As String = "http://search.yahoo.com/bin/search?p= "
    '        Dim sFindNext20 As String = ">Next 20"
    '        Dim iEndUrlNext20 As Integer
    '        Dim iBeginNext20 As Integer
    '        Dim bContinue As Boolean = True
    '        Dim sFindQuote As String
    '        Dim iCounter As Integer = -5
    '        Dim sFindNext20Url As String
    '        Dim iLenghtNextUrl As Integer
    '        Dim sTakeCareOfString As String
    '        Dim iFirstTimeNext20 As Integer = 0
    '        Dim sOldsFindNext20Url As String

    '        sYahoo = sYahoo & strURI
    '        strURI = ""
    '        strURI = sYahoo
    '        Do While bContinue

    '            If Len(sFindNext20Url) > 10 Then
    '                If sFindNext20Url = sOldsFindNext20Url Then
    '                    Exit Do
    '                End If
    '                strURI = sFindNext20Url
    '                sOldsFindNext20Url = sFindNext20Url
    '            End If

    '            Dim objURI As Uri = New Uri(strURI)
    '            Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
    '            Dim objWebResponse As WebResponse = objWebRequest.GetResponse()

    '            objWebResponse.Headers.Add("Accept-Language: fr, en")

    '            Dim objStream As Stream = objWebResponse.GetResponseStream()
    '            Dim objStreamReader As StreamReader = New StreamReader(objStream)
    '            Dim strHTML As String = objStreamReader.ReadToEnd

    '            sTakeCareOfString = strHTML
    '            iEndUrlNext20 = InStr(1, strHTML, sFindNext20)

    '            If iEndUrlNext20 > 0 Then

    '                If iFirstTimeNext20 = 0 Then
    '                    iEndUrlNext20 = InStr(iEndUrlNext20 + 10, strHTML, sFindNext20)
    '                    iFirstTimeNext20 = 1
    '                End If

    '                Do While bContinue
    '                    sFindQuote = Mid(strHTML, (iEndUrlNext20 + iCounter), 5)
    '                    If sFindQuote = "href=" Then
    '                        iBeginNext20 = (iEndUrlNext20 + iCounter) + 5
    '                        Exit Do
    '                    End If
    '                    iCounter = iCounter - 1
    '                    If iCounter = -500 Then
    '                        Exit Do
    '                    End If
    '                Loop

    '                iLenghtNextUrl = iEndUrlNext20 - iBeginNext20
    '                If iLenghtNextUrl > 20 And iLenghtNextUrl < 400 Then
    '                    sFindNext20Url = Mid(sTakeCareOfString, (iBeginNext20 + 1), (iLenghtNextUrl - 2))
    '                    'Label7.Text = sFindNext20Url
    '                Else
    '                    'Label7.Text = iLenghtNextUrl
    '                End If
    '            End If
    '            SGetText = strHTML
    '            'Label1.Text = SGetText
    '            sGetHttp = "http://www.yahoo.com"
    '            'Label2.Text = fGetUrl(SGetText, sGetHttp)
    '            LookYahooGoogle = strHTML

    '            If Not (Len(sFindNext20Url) > 10) Then
    '                Exit Do
    '            End If
    '        Loop
    '    Catch e_search_601 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " lookforfiles.e_search_601 --> " & e_search_601.Message, EventLogEntryType.Information, 44)
    '        'MsgBox("The error (e_LookForFiles_955) is " & e_LookForFiles_955.Message)
    '    End Try
    'End Function


    Function GetWebPageAsStringShort(ByVal strURI As String) As String
        Try
            Dim objURI As Uri = New Uri(strURI)
            Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
            Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
            objWebResponse.Headers.Add("Accept-Language: fr, en")

            Dim objStream As Stream = objWebResponse.GetResponseStream()
            Dim objStreamReader As StreamReader = New StreamReader(objStream)
            Dim strHTML As String = objStreamReader.ReadToEnd
            GetWebPageAsStringShort = strHTML
        Catch e_search_756 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_756--> " & e_search_756.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_LookForFiles_955) is " & e_LookForFiles_955.Message)
        End Try
    End Function


    Public Function RemoveHTMLFromDNS(ByVal sRemove As String) As String
        Try
            Dim sKeepHtml As String
            Dim sCloseQuote As String
            Dim sOpenQuote As String
            Dim iOpenCount As Integer
            Dim iCloseCount As Integer
            Dim iLenghtRemove As Integer
            Dim iStart As Integer
            Dim bKeepGoing As Boolean
            Dim sStringToRemove As String
            Dim iCheckSpace As Integer
            Dim sKeepTraceOfString As String
            bKeepGoing = True

            sKeepHtml = sRemove

            iOpenCount = InStr(1, sKeepHtml, "<")
            iCloseCount = InStr(1, sKeepHtml, ">")

            If iOpenCount > iCloseCount And iOpenCount > 0 And iCloseCount > 0 Then
                sStringToRemove = Mid(sKeepHtml, iCloseCount, (Len(sKeepHtml) + 1) - iCloseCount)
                sKeepHtml = Replace(sKeepHtml, sStringToRemove, "")
            ElseIf iCloseCount > iOpenCount And iOpenCount > 0 And iCloseCount > 0 Then
                sStringToRemove = Mid(sKeepHtml, iOpenCount, (Len(sKeepHtml) + 1) - iOpenCount)
                sKeepHtml = Replace(sKeepHtml, sStringToRemove, "")
            ElseIf iOpenCount > 0 Then
                sStringToRemove = Mid(sKeepHtml, iOpenCount, (Len(sKeepHtml) + 1) - iOpenCount)
                sKeepHtml = Replace(sKeepHtml, sStringToRemove, "")
            ElseIf iCloseCount > 0 Then
                sStringToRemove = Mid(sKeepHtml, iCloseCount, (Len(sKeepHtml) + 1) - iCloseCount)
                sKeepHtml = Replace(sKeepHtml, sStringToRemove, "")
            End If

            sKeepHtml = Replace(sKeepHtml, Chr(9), "")
            sKeepHtml = Replace(sKeepHtml, Chr(10), "")
            sKeepHtml = Replace(sKeepHtml, Chr(13), "")
            RemoveHTMLFromDNS = sKeepHtml

        Catch e_search_377 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_377--> " & e_search_377.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Public Function RemoveHtmlTag(ByVal sRemove As String) As String
        Try
            Dim sKeepHtml As String
            Dim sCloseQuote As String
            Dim sOpenQuote As String
            Dim iOpenCount As Integer
            Dim iCloseCount As Integer
            Dim iLenghtRemove As Integer
            Dim iStart As Integer
            Dim bKeepGoing As Boolean
            Dim sStringToRemove As String
            Dim iCheckSpace As Integer
            Dim sKeepTraceOfString As String
            Dim i_ccc As Int32

            bKeepGoing = True

            sKeepHtml = sRemove

            iOpenCount = InStr(1, sKeepHtml, "<")
            iCloseCount = InStr(1, sKeepHtml, ">")
            iStart = 1

            Do While bKeepGoing
                If iOpenCount > 0 And iCloseCount > 0 Then
                    iLenghtRemove = (iCloseCount - iOpenCount)
                    sKeepTraceOfString = sKeepTraceOfString & "," & iLenghtRemove

                    If iLenghtRemove > 0 Then
                        sStringToRemove = Mid(sKeepHtml, iOpenCount, (iLenghtRemove + 1))
                        sKeepHtml = Replace(sKeepHtml, sStringToRemove, "")
                    Else
                        RemoveHtmlTag = sKeepHtml
                        Exit Do
                    End If

                    iOpenCount = InStr(1, sKeepHtml, "<")

                    If iOpenCount = 0 Then
                        Exit Do
                    End If

                    iCloseCount = InStr(iOpenCount, sKeepHtml, ">")
                    i_ccc = i_ccc + 1

                    If i_ccc > 5000 Then
                        Exit Do
                    End If
                Else
                    Exit Do
                End If
            Loop
            If Len(sKeepHtml) Then
                RemoveHtmlTag = sKeepHtml
            Else
                RemoveHtmlTag = sRemove
            End If

        Catch e_search_678 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_678--> " & e_search_678.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function PutPipeAroundTitle(ByVal s_pageContent) As String
        Try
            Dim s_Title As String = "<title>"
            Dim s_TitleEnd As String = "</title>"

            s_pageContent = Replace(s_pageContent, "< title>", s_Title)
            s_pageContent = Replace(s_pageContent, "< title >", s_Title)
            s_pageContent = Replace(s_pageContent, "<title >", s_Title)

            s_pageContent = Replace(s_pageContent, "< /title>", s_TitleEnd)
            s_pageContent = Replace(s_pageContent, "< /title >", s_TitleEnd)
            s_pageContent = Replace(s_pageContent, "</title >", s_TitleEnd)

            s_pageContent = Replace(s_pageContent, "<title>", "ttitle")
            s_pageContent = Replace(s_pageContent, "</title>", "endttitle")

            PutPipeAroundTitle = s_pageContent

        Catch e_search_445 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_445--> " & e_search_445.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Public Function RemoveJSTag(ByVal sRemove As String) As String
        Try
            Dim sKeepHtml As String
            Dim sCloseQuote As String
            Dim sOpenQuote As String
            Dim iOpenCount As Integer
            Dim iCloseCount As Integer
            Dim iLenghtRemove As Integer
            Dim iStart As Integer
            Dim bKeepGoing As Boolean
            Dim sStringToRemove As String
            Dim iCheckSpace As Integer
            Dim sKeepTraceOfString As String
            Dim i_ccc As Int32

            bKeepGoing = True

            sKeepHtml = sRemove

            iOpenCount = InStr(1, LCase(sKeepHtml), "<script", CompareMethod.Text)
            iCloseCount = InStr(1, LCase(sKeepHtml), "/script", CompareMethod.Text)
            iStart = 1

            Do While bKeepGoing
                If iOpenCount > 0 And iCloseCount > 0 Then
                    iLenghtRemove = ((iCloseCount + 8) - (iOpenCount))
                    sKeepTraceOfString = sKeepTraceOfString & "," & iLenghtRemove

                    If iLenghtRemove > 0 Then
                        sStringToRemove = Mid(sKeepHtml, iOpenCount, (iLenghtRemove + 1))
                        sKeepHtml = Replace(sKeepHtml, sStringToRemove, "")
                    Else
                        Exit Do
                    End If

                    iOpenCount = InStr(1, LCase(sKeepHtml), "<script", CompareMethod.Text)

                    If iOpenCount = 0 Then
                        Exit Do
                    End If

                    iCloseCount = InStr(iOpenCount, LCase(sKeepHtml), "/script", CompareMethod.Text)
                    i_ccc = i_ccc + 1

                    If i_ccc >= 1000 Then
                        Exit Do
                    End If
                Else
                    Exit Do
                End If

            Loop

            If IsNothing(sKeepHtml) Then
                RemoveJSTag = sRemove
            ElseIf sKeepHtml = "" Then
                RemoveJSTag = sRemove
            Else
                If Len(Trim(sKeepHtml)) > 20 Then
                    RemoveJSTag = sKeepHtml
                Else
                    RemoveJSTag = sRemove
                End If
            End If

        Catch e_search_476 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_476--> " & e_search_476.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function RemoveStyle(ByVal s_pageContent) As String
        Try
            Dim sKeepHtml As String = s_pageContent
            Dim sCloseQuote As String
            Dim sOpenQuote As String
            Dim iOpenCount As Integer
            Dim iCloseCount As Integer
            Dim iLenghtRemove As Integer
            Dim iStart As Integer
            Dim bKeepGoing As Boolean
            Dim sStringToRemove As String
            Dim iCheckSpace As Integer
            Dim sKeepTraceOfString As String
            Dim i_ccc As Int32

            bKeepGoing = True

            iOpenCount = InStr(1, LCase(sKeepHtml), "<style", CompareMethod.Text)
            iCloseCount = InStr(1, LCase(sKeepHtml), "</style>", CompareMethod.Text)
            iStart = 1

            Do While bKeepGoing
                If iOpenCount > 0 And iCloseCount > 0 Then
                    iLenghtRemove = ((iCloseCount + 8) - (iOpenCount))
                    sKeepTraceOfString = sKeepTraceOfString & "," & iLenghtRemove

                    If iLenghtRemove > 0 Then
                        sStringToRemove = Mid(sKeepHtml, iOpenCount, (iLenghtRemove + 1))
                        sKeepHtml = Replace(sKeepHtml, sStringToRemove, "")
                    Else
                        Exit Do
                    End If

                    iOpenCount = InStr(1, LCase(sKeepHtml), "<style", CompareMethod.Text)

                    If iOpenCount = 0 Then
                        Exit Do
                    End If

                    iCloseCount = InStr(iOpenCount, LCase(sKeepHtml), "</style>", CompareMethod.Text)
                    i_ccc = i_ccc + 1

                    If i_ccc >= 100 Then
                        Exit Do
                    End If
                Else
                    Exit Do
                End If

            Loop

            If IsNothing(sKeepHtml) Then
                RemoveStyle = s_pageContent
            ElseIf s_pageContent = "" Then
                RemoveStyle = s_pageContent
            Else
                If Len(Trim(sKeepHtml)) > 20 Then
                    RemoveStyle = sKeepHtml
                Else
                    RemoveStyle = sKeepHtml
                End If
            End If

        Catch e_search_476 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_476--> " & e_search_476.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Function fGetUrl(ByVal FileToCrawl As String, ByVal sGetHttp As String) As String
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
            bKeepGoing = True

            sURL = sGetHttp
            iFirst = InStr(1, FileToCrawl, "href=")
            iLast = InStr((iFirst), FileToCrawl, ">")
            iLenght = iLast - iFirst
            If iLenght > 0 Then
                Do While bKeepGoing
                    sStoreResult = Mid(FileToCrawl, iFirst + 5, iLenght - 5)
                    sStoreResult = Replace(sStoreResult, """", "")
                    iURLIn = InStr(1, sStoreResult, sURL)
                    If Not (iURLIn > 0) Then
                        sStoreResult = sURL & "/" & sStoreResult
                    End If
                    'Label3.Text = sStoreResult
                    sCheckStop = CheckUrlIn(sStoreResult, sSendstrURI)

                    If Not (sCheckStop = "stop") Then
                        InsertUrl(sCheckStop, sSendstrURI)
                    End If

                    sAllResult = sAllResult & Chr(10) & sStoreResult
                    'Label2.Text = sAllResult
                    fGetUrl = sAllResult

                    iFirst = InStr(iFirst + 5, FileToCrawl, "href=")
                    If iFirst <= Len(FileToCrawl) Then
                        If iFirst > 0 Then
                            iLast = InStr(iFirst, FileToCrawl, ">")
                        End If
                    End If

                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
                        Exit Do
                    Else
                        iLenght = iLast - iFirst
                    End If
                Loop
            End If
        Catch e_search_738 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_738--> " & e_search_738.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_LookForFiles_955) is " & e_LookForFiles_955.Message)
        End Try
    End Function


    Function fGetUrlAlltheWeb(ByVal FileToCrawl As String, ByVal sGetHttp As String) As String
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

            bKeepGoing = True

            sURL = sGetHttp
            iFirst = InStr(1, FileToCrawl, "/http")
            'MessageBox.Show(iFirst, "iFirst")
            iLast = InStr(iFirst + 4, FileToCrawl, "'")
            'MessageBox.Show(iLast, "iLast")
            iLenght = iLast - iFirst

            'MessageBox.Show(iLenght, "iLenght")

            If iLenght > 0 Then
                Do While bKeepGoing
                    sStoreResult = Mid(FileToCrawl, iFirst + 1, iLenght)
                    sStoreResult = Trim(Replace(sStoreResult, "/", ""))
                    If Len(sStoreResult) > 0 And Len(sStoreResult) < 255 Then
                        sCheckStop = CheckUrlIn(sStoreResult, sSendstrURI)
                        'MessageBox.Show(sCheckStop, "sCheckStop")
                        If Not (sCheckStop = "stop") Then
                            InsertUrl(sCheckStop, sSendstrURI)
                            InsertSite(sCheckStop)
                        End If
                    End If

                    sAllResult = sAllResult & Chr(10) & sStoreResult
                    'Label3.Text = sAllResult
                    'fGetUrlAlltheWeb = sAllResult

                    iFirst = InStr(iLast + 1, FileToCrawl, "/http")
                    If iFirst <= Len(FileToCrawl) Then
                        If iFirst > 0 Then
                            iLast = InStr(iFirst + 4, FileToCrawl, "/")
                        End If
                    End If

                    If Not (iFirst > 0 And iLast > 0 And iLast > iFirst) Then
                        Exit Do
                    Else
                        iLenght = iLast - iFirst
                    End If
                Loop

            End If
        Catch e_search_803 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_803--> " & e_search_803.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_LookForFiles_955) is " & e_LookForFiles_955.Message)
        End Try
    End Function


    Function CheckUrlIn_a_Deleter(ByVal SUrl As String, ByVal sSendstrURI As String) As String

        Try

            If Len(sSendstrURI) < 4 Then
                'MsgBox(sSendstrURI & "  --  search.sSendstrURI.1034")
            End If

            Dim myConnString As String
            Dim sKeepString As String
            Dim sSendUri As String = sSendstrURI

            If Not (InStr(1, SUrl, "'") > 0) Then
                Dim mySelectQuery As String = "SELECT url FROM URLFOUND where url ='" & Trim(SUrl) & "'"
                Dim myConnection As New OleDbConnection(ConnStringURL(sSendUri))
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

                myConnection.Open()

                'MsgBox("Search.mySelectQuery.749 " & mySelectQuery)

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    sKeepString = sKeepString & (myReader.GetString(0))
                    'MsgBox("Search.sKeepString.752 " & sKeepString)
                End While
                'MsgBox("Search.sKeepString.754 " & sKeepString)
                'MsgBox("Search.Len(sKeepString).749 " & Len(sKeepString))

                If Len(sKeepString) > 0 Then
                    'MsgBox("Search.CheckUrlIn.758 -> stop ")
                    'CheckUrlIn = "stop"
                Else
                    'MsgBox("Search.CheckUrlIn.761 " & Trim(SUrl))
                    'CheckUrlIn = Trim(SUrl)
                End If

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                '''myConnection = Nothing
                myConnection.Close()

            Else
                'CheckUrlIn = "stop"
            End If

        Catch e_search_703 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_search_703 --> " & e_search_703.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Public Sub StartTheScanningProcessFromConfiguration()

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"

        Try
            Dim b_continue As Boolean = True
            Dim b_internetConnected As Boolean
            'Call StartThread_t_ReceiveDNSFound()

            'Dim t_ReceiveDNSFound As New Thread(AddressOf ReceiveDNSFound)
            't_ReceiveDNSFound.Name = "ReceiveDNSFound"
            't_ReceiveDNSFound.Priority = sr_Priority
            't_ReceiveDNSFound.Start()

            'Call CleanupFiles()

            ' Call CheckSizeOfLog()

            b_internetConnected = CheckInternetConnection()

            If b_internetConnected Then

                s_p_SeedOrScan = "Scan"

                ' Call sWhatRange()

                '            'Console.WriteLine(sr_RangeMinus & " --- search.sr_RangeMinus.596 ")
                '            'Console.WriteLine(sr_RangePlus & " --- search.sr_RangePlus.597 ")

                'Call ExportParameters()

                While b_continue

                    If b_p_closedTheSeedingProcess Then
                        Exit While
                    End If

                    '' If CheckMachineNamePublished() Then

                    '''' If CheckMachineNamePublished() = True Then
                    '''


                    Call fFindWhatToLookFor()

                    'Call CheckSizeOfLog()

                    'Console.WriteLine("Go to sleep 3 seconds, search.826")
                    'Console.WriteLine(vbNewLine)
                    ''''    End If

                    '                If CheckAuthority() Then

                    '                    Call DistributeAlphaRange()

                    '                    Call sendNewMachineConfiguration()

                    '                    Call SendQueryInsertOverNetwork()

                    '                    Call SendQueryUpdateOverNetwork()

                    '                    Call SendQueryDeleteOverNetwork()

                    '                End If

                    '                System.Threading.Thread.Sleep(3000)

                End While

                '            'Dim t_WhatToLookFor As Thread
                '            't_WhatToLookFor = New Thread(AddressOf StartProcessWhatToLookFor)
                '            't_WhatToLookFor.Name = "StartProcessWhatToLookFor"
                '            't_WhatToLookFor.Priority = Normal
                '            't_WhatToLookFor.Start()

            End If

        Catch e_search_1058 As Exception

            EventLog.WriteEntry(sSource, " e_search_1058 --> " & e_search_1058.Message, EventLogEntryType.Information, 44)

        End Try
    End Sub


    'Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)

    'End Sub


    Sub StartProcessWhatToLookFor()
        Try
            Dim b_continue As Boolean = True

            While b_continue

                If b_p_closedTheSeedingProcess Then
                    Exit While
                End If

                Call fFindWhatToLookFor()

                Call CheckSizeOfLog()
                System.Threading.Thread.Sleep(30000)
            End While

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " search.874 --> " & ex.Message, EventLogEntryType.Information, 45)
        End Try

    End Sub


    Public Function StartThread_t_ReceiveDNSFound() As Boolean

        'If (StartThread_t_ReceiveDNSFound) Then
        'Else
        '    Dim t_ReceiveDNSFound As New Thread(AddressOf ReceiveDNSFound)
        '    t_ReceiveDNSFound.Name = "ReceiveDNSFound"
        '    t_ReceiveDNSFound.Priority = Normal
        '    t_ReceiveDNSFound.Start()
        '    StartThread_t_ReceiveDNSFound = True
        'End If

    End Function


    Public s_p_Priority As String
    Public s_p_keepStringToDoSearchOn As String
    Public s_p_keepStringWebsite As String

    Sub StartScanning(ByVal sToInsert As String)

        Call StartScanningFU()

        Dim iDuration As Double
        Dim sQuery As String = sToInsert
        Dim myConnectionString As String
        Dim sDate As String = Now()
        Dim iLevel, iMaxPage As Integer
        Dim iMinKPage, iMaxKPage, iMinDensity, iMaxDensity As Integer
        Dim iKeywordTitle, iKeywordBody, IkeywordTag, iKeywordVisible, iKeywordURL As Integer
        Dim iMinDiskPerc, IDiskMinGiga As Integer
        Dim iAlgorythm, iGoogle, iAllTheWeb, iExactPhrase As Integer
        Dim iSearchExcel, IsearchWord, iSearchPDF, iDesactiveQuery As Integer
        Dim sPriority As String
        Dim bGoAhead As Boolean = False
        Dim bInsert As Boolean = False
        Dim bUpdate As Boolean = False
        Dim bStopTheInsertion As Boolean = False
        Dim sWebSite As String = "n/a"
        Dim sNetwork As String = "n/a"
        Dim sSecondNetwork As String = "n/a"
        Dim sUniversalNetwork As String = "n/a"
        Dim sMachineAddress As String = "127.0.0.1"
        Dim sSecondaryMachineAddress = "127.0.0.1"
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim mySelectQuery As String

        Try

            If s_p_keepStringToDoSearchOn <> sToInsert Then

                If Len(sToInsert) > 0 Then

                    mySelectQuery = "SELECT * FROM DNAVALUE where s_Query ='" & sQuery & "'"
                    Dim myConnection1 As New OleDbConnection(ConnStringDNA())
                    Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

                    myConnection1.Open()

                    Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                    Dim Index As Integer

                    While myReader.Read()

                        iDuration = myReader("i_Duration")
                        iLevel = myReader("i_Level") 'done
                        iMaxPage = myReader("i_MaxPage") 'done
                        s_p_Priority = myReader("s_Priority")
                        iMinKPage = myReader("i_MinKPage") 'done
                        iMaxKPage = myReader("i_MaxKPage") 'done
                        iMinDensity = myReader("i_MinDensity") 'done
                        iMaxDensity = myReader("i_MaxDensity") 'done
                        iKeywordTitle = myReader("i_KeywordTitle") 'done
                        iKeywordBody = myReader("i_KeywordBody") 'done
                        IkeywordTag = myReader("I_keywordTag") 'done
                        iKeywordVisible = myReader("i_KeywordVisible") 'done
                        iKeywordURL = myReader("i_KeywordURL") 'done
                        iMinDiskPerc = myReader("i_MinDiskPerc") 'done
                        IDiskMinGiga = myReader("I_MinDiskGiga") 'done
                        iAlgorythm = myReader("i_Algorythm") 'done
                        iGoogle = myReader("i_Google")
                        iAllTheWeb = myReader("i_AllTheWeb")
                        iSearchExcel = myReader("i_SearchExcel") 'done
                        IsearchWord = myReader("i_SearchWord") 'done
                        iSearchPDF = myReader("i_SearchPDF") 'done
                        iDesactiveQuery = myReader("i_DesactiveQuery")
                        iExactPhrase = myReader("i_ExactPhrase") 'done
                        sNetwork = myReader("s_FirstNetwork")
                        sSecondNetwork = myReader("s_SecondNetwork")
                        sUniversalNetwork = myReader("s_UniversalNetwork")
                        sWebSite = myReader("s_WebSite") 'done

                        Exit While

                    End While

                    If Not (myReader.IsClosed) Then
                        myReader.Close()
                    End If

                    myConnection1.Close()

                Else

                    'EventLog.WriteEntry(sSource, "e_search_805 mySelectQuery--> " & mySelectQuery, EventLogEntryType.Information, 49)

                    Call StartTheScanningProcessFromConfiguration()

                End If

            End If

            s_p_keepStringToDoSearchOn = sToInsert

            'Call StartThread_t_ReceiveDNSFound()

        Catch e_search_805 As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.e_search_805 --> " & e_search_805.Message, EventLogEntryType.Information, 44)
            EventLog.WriteEntry(sSource, "e_search_805 mySelectQuery--> " & mySelectQuery, EventLogEntryType.Information, 49)

        End Try

        Try

            Call SelectFromDNSFoundFU()

            If Len(sWebSite) > 4 Then
                Dim myConnection2 As New OleDbConnection(ConnStringURLDNS(sQuery))

                Dim sDate1 As String = DateAdd(DateInterval.Minute, (-1 * i_p_Duration * 60), Now())
                Dim sNo As String
                Dim iIsGood As Integer = 100
                Dim s_MyArray() As String
                Dim i As Integer
                Dim s_na As String = "n/a"

                If InStr(1, sWebSite, ",") > 0 Then
                    sWebSite = Replace(sWebSite, " ", "")
                    s_MyArray = Split(sWebSite, ",")
                Else
                    s_MyArray = Split(sWebSite, " ")
                End If

                For i = 0 To UBound(s_MyArray)

                    Dim s_mySelectQuery As String = "SELECT TOP(1) * FROM DNSFOUND where iGoodSpot <> -4000 AND URLSite ='" & s_MyArray(i) & "' AND URLSite <>'" & s_na & "' AND sCurrentQuery = '" & sToInsert & "' order by iGoodSpot desc "

                    Dim myConnection22 As New OleDbConnection(ConnStringURLDNS(sQuery))
                    Dim myCommand As New OleDbCommand(s_mySelectQuery, myConnection22)
                    Dim i_KeepTrack As Integer = 0

                    myConnection22.Open()

                    Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                    Dim Index As Integer
                    Dim s_FoundURLSite As String = ""

                    While myReader.Read()
                        s_FoundURLSite = myReader("URLSite")
                    End While

                    If Not (myReader.IsClosed) Then
                        myReader.Close()
                    End If

                    myConnection22.Close()

                    If Not (Len(s_FoundURLSite)) > 0 Then

                        If Len(s_FoundURLSite) > 0 Then

                        Else

                            If Len(s_MyArray(i)) > 0 Then

                                Dim s_shortURL As String = Func_CleanUrl(s_MyArray(i))

                                Dim myInsertQuery2 As String = ""
                                If CheckRangeGoHere(s_MyArray(i)) = True Then
                                    sNo = "n"
                                    myInsertQuery2 = "INSERT INTO DNSFOUND (URLSite,DateHour,Done,iGoodSpot,ShortURL,sCurrentQuery) Values('" & Trim(s_MyArray(i)) & "','" & (sDate1) & "','" & (sNo) & "','" & (iIsGood) & "','" & Trim(s_shortURL) & "','" & Trim(sToInsert) & "')"
                                Else
                                    sNo = "s"
                                    myInsertQuery2 = "INSERT INTO DNSFOUND (URLSite,DateHour,Done,iGoodSpot,ShortURL,sCurrentQuery) Values('" & Trim(s_MyArray(i)) & "','" & (sDate1) & "','" & (sNo) & "','" & (iIsGood) & "','" & Trim(s_shortURL) & "','" & Trim(sToInsert) & "')"
                                End If

                                Dim objConn As New OleDbConnection(ConnStringURL(ConnStringDNA()))
                                'Dim myCommand11 As New OleDbCommand(myInsertQuery2)
                                Dim myCommand2 As New OleDbCommand(myInsertQuery2)

                                myCommand2.Connection = objConn
                                objConn.Open()
                                myCommand2.ExecuteNonQuery()
                                objConn.Close()


                                'myCommand11.Connection = myConnection2
                                'myConnection2.Open()
                                'myCommand11.ExecuteNonQuery()
                                'myConnection2.Close()

                                If i_KeepTrack = 25 Then
                                    i_KeepTrack = 0
                                End If

                                i_KeepTrack = i_KeepTrack + 1

                            End If

                        End If

                    End If

                Next

            End If

        Catch e_search_1303 As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.e_search_1303 --> " & e_search_1303.Message, EventLogEntryType.Information, 44)
        End Try

        Dim strPathToFile As String = sToInsert
        Dim strPathToFile_DNS As String = sToInsert

        Try

            'If Len(sQuery) > 0 Then
            '    strPathToFile = Replace(strPathToFile, " ", "_")
            '    strPathToFile = Replace(strPathToFile, "'", "_")
            '    strPathToFile = strPathToFile & ".mdb"
            '    strPathToFile_DNS = Replace(strPathToFile_DNS, " ", "_")
            '    strPathToFile_DNS = Replace(strPathToFile_DNS, "'", "_")
            '    strPathToFile_DNS = strPathToFile_DNS & "_DNS.mdb"
            'Else
            '    MsgBox("There is no query to look on, please check the schedule to make sure!")
            'End If

        Catch e_Search_871 As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.e_Search_871 --> " & e_Search_871.Message, EventLogEntryType.Information, 44)
        End Try

        Try

            Dim bContinue As Boolean = True
            Dim bFindAnotherUrl As Boolean = True
            Dim sUrlToScan As String = ""
            Dim sFileToScan As String = ""
            Dim bMakePlace As Boolean = False
            Dim b_CheckWebsite As Boolean = True
            Dim s_Continue As String

            Do While bContinue

                If b_p_closedTheSeedingProcess Then
                    Exit Do
                End If
                '//**********************************************************************************************
                'Call f_FileToScan_iMinDiskPerc(iMinDiskPerc, bMakePlace)
                'Call f_FileToScan_IDiskMinGiga(IDiskMinGiga, bMakePlace)

                s_Continue = f_FindUrlToScan(sToInsert, bContinue)

                s_p_FunName = "StartScanning"
                s_p_Message = ""
                s_p_VarName = "s_continue"
                s_p_VarValue = s_Continue

                If s_Continue = "stop" Then
                    bContinue = False
                End If

                If bMakePlace Then
                    'Call f_MakePlace()
                End If

                If bContinue = False Then
                    'EventLog.WriteEntry(sSource, " no error! lookforfiles.e_Search_1385 --> The crawling finished at all levels, starting an other levels?", EventLogEntryType.Information, 44)
                    Exit Do
                End If

            Loop
        Catch e_Search_915 As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.e_Search_915 --> " & e_Search_915.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub


    Public ReadOnly Property sr_RangeMinus() As String
        Get
            sr_RangeMinus = s_p_RangeMinus
        End Get
    End Property

    Public ReadOnly Property sr_RangePlus() As String
        Get
            sr_RangePlus = s_p_RangePlus
        End Get
    End Property


    Public s_p_RangeMinus As String

    Public s_p_RangePlus As String


    Public Function sWhatRange() As Boolean

        Try
            Dim s As String
            Dim s_hostName As String
            Dim s_minusRange As String
            Dim s_PlusRange As String
            Dim s_no As String = "n"
            sWhatRange = False

            'Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName)
            's = (CType(h.AddressList.GetValue(0), IPAddress).ToString)
            's_hostName = System.Net.Dns.Resolve(s).HostName()
            ''s_hostName = System.Environment.MachineName()
            ''Console.WriteLine(s_hostName & " --- search.959")

            ''Dim mySelectQuery As String = "SELECT * FROM RANGE where HostName ='" & s_hostName & "'"
            ''Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            ''Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

            ''myConnection1.Open()

            ''Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            ''Dim Index As Integer
            ''Dim s_rangeMinus As String
            ''Dim s_rangePlus As String

            ''While myReader.Read()
            ''    s_p_RangeMinus = myReader("RangeMinus")
            ''    s_p_RangePlus = myReader("RangePlus")
            ''    'If s_p_RangeMinus = 0 And s_p_RangePlus = 0 Then
            ''    'sWhatRange = False
            ''    'Else
            ''    sWhatRange = True
            ''    'End If

            ''End While

            ''If Not (myReader.IsClosed) Then
            ''    myReader.Close()
            ''End If

            '    myConnection1.Close()


            sWhatRange = True
            s_p_RangeMinus = "aa"
            s_p_RangePlus = "zz"

        Catch e_search_1327 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_1327--> " & e_search_1327.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public Function sWhatRange(ByVal s_machineName As String) As Boolean

        Try
            Dim s As String
            Dim s_hostName As String
            Dim s_minusRange As String
            Dim s_PlusRange As String
            Dim s_no As String = "n"
            sWhatRange = False

            'Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName)
            's = (CType(h.AddressList.GetValue(0), IPAddress).ToString)
            's_hostName = System.Net.Dns.Resolve(s).HostName()

            'Console.WriteLine(s_hostName & " --- search.1139")

            Dim mySelectQuery As String = "SELECT * FROM RANGE where HostName ='" & s_machineName & "' and SentOver ='" & s_no & "'"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer
            Dim s_rangeMinus As String
            Dim s_rangePlus As String

            While myReader.Read()
                sWhatRange = True
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection1.Close()

        Catch e_search_1327 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_1327--> " & e_search_1327.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public Function sWhatRange(ByVal s_machineName As String, ByVal s_fromAbroad As String) As Boolean

        Try
            Dim s As String
            Dim s_hostName As String
            Dim s_minusRange As String
            Dim s_PlusRange As String
            Dim s_no As String = "n"
            sWhatRange = False

            'Dim h As System.Net.IPHostEntry = System.Net.Dns.GetHostByName(System.Net.Dns.GetHostName)
            's = (CType(h.AddressList.GetValue(0), IPAddress).ToString)
            's_hostName = System.Net.Dns.Resolve(s).HostName()

            'Console.WriteLine(s_hostName & " --- search.1139")

            Dim mySelectQuery As String = "SELECT * FROM RANGE where HostName ='" & s_machineName & "'"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer
            Dim s_rangeMinus As String
            Dim s_rangePlus As String

            While myReader.Read()
                sWhatRange = True
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection1.Close()

        Catch e_search_1327 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_1327--> " & e_search_1327.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public d_p_cleanupFilesDate As DateTime
    Public b_p_cleanupFilesDateTaken As Boolean = False

    Sub CleanupFiles()

        Dim files() As FileInfo
        Dim filese As FileInfo
        Dim dir As DirectoryInfo = New DirectoryInfo(sAppPath)
        Dim s_keepNameFile, s_query, s_keepName As String
        Dim i_keepDot, i_keepDNS As Integer
        Dim b_AlreadyThere As Boolean = False

        Try
            If Not (b_p_cleanupFilesDateTaken) Then
                d_p_cleanupFilesDate = Now
                b_p_cleanupFilesDateTaken = True
            End If

            If DateDiff(DateInterval.Day, d_p_cleanupFilesDate, Now) > 2 Then

                b_p_cleanupFilesDateTaken = False

                files = dir.GetFiles("*.mdb")

                For Each filese In files
                    'MsgBox(filese.Name)
                    s_keepName = filese.Name
                    s_keepNameFile = filese.Name

                    i_keepDNS = InStr(s_keepNameFile, "_DNS")
                    If i_keepDNS > 0 Then
                        s_keepNameFile = Replace(s_keepNameFile, "_DNS", "")
                    End If

                    'If i_keepDNS > 0 Then
                    'Else

                    i_keepDot = InStr(s_keepNameFile, ".")
                    s_keepNameFile = Mid(s_keepNameFile, 1, ((Len(s_keepNameFile) - 1) - (Len(s_keepNameFile) - i_keepDot)))
                    s_keepNameFile = Trim(Replace(s_keepNameFile, "_", " "))

                    Dim mySelectQuery As String = "SELECT * FROM DNAVALUE where s_Query ='" & s_keepNameFile & "'"
                    'MessageBox.Show(mySelectQuery, "mySelectQuery")
                    Dim myConnection1 As New OleDbConnection(ConnStringDNA())

                    Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
                    myConnection1.Open()

                    Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                    Dim Index As Integer
                    While myReader.Read()
                        If s_keepNameFile = myReader("s_Query") Then
                            b_AlreadyThere = True
                            Exit While
                        End If

                        'MsgBox("lookforfiles.b_AlreadyThere.2054 is - " & b_AlreadyThere)

                    End While

                    If Not (myReader.IsClosed) Then
                        myReader.Close()
                    End If

                    '''myConnection1 = Nothing
                    myConnection1.Close()

                    If b_AlreadyThere = True Then
                        b_AlreadyThere = False
                    Else

                        If (File.Exists(sAppPath() & s_keepName)) Then
                            File.Delete(sAppPath() & s_keepName)
                        End If
                    End If
                    'End If
                Next

            End If

        Catch e_search_1360 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_1360 --> " & e_search_1360.Message, EventLogEntryType.Information, 44)

        End Try

        Try
            If DateDiff(DateInterval.Day, d_p_cleanupFilesDate, Now) > 2 Then
                files = dir.GetFiles("*.ldb")

                For Each filese In files
                    'MsgBox(filese.Name)
                    s_keepName = filese.Name
                    s_keepNameFile = filese.Name

                    i_keepDNS = InStr(s_keepNameFile, "_DNS")

                    If i_keepDNS > 0 Then
                        s_keepNameFile = Replace(s_keepNameFile, "_DNS", "")
                    End If

                    'If i_keepDNS > 0 Then
                    'Else

                    i_keepDot = InStr(s_keepNameFile, ".")
                    s_keepNameFile = Mid(s_keepNameFile, 1, ((Len(s_keepNameFile) - 1) - (Len(s_keepNameFile) - i_keepDot)))
                    s_keepNameFile = Trim(Replace(s_keepNameFile, "_", " "))

                    Dim mySelectQuery As String = "SELECT * FROM DNAVALUE where s_Query ='" & s_keepNameFile & "'"
                    'MessageBox.Show(mySelectQuery, "mySelectQuery")
                    Dim myConnection1 As New OleDbConnection(ConnStringDNA())

                    Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
                    myConnection1.Open()

                    Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                    Dim Index As Integer
                    While myReader.Read()
                        If s_keepNameFile = myReader("s_Query") Then
                            b_AlreadyThere = True
                            Exit While
                        End If

                        'MsgBox("lookforfiles.b_AlreadyThere.2054 is - " & b_AlreadyThere)

                    End While

                    If Not (myReader.IsClosed) Then
                        myReader.Close()
                    End If

                    '''myConnection1 = Nothing
                    myConnection1.Close()

                    If b_AlreadyThere = True Then
                        b_AlreadyThere = False
                    Else
                        If (File.Exists(sAppPath() & s_keepName)) Then
                            File.Delete(sAppPath() & s_keepName)
                        End If
                    End If

                    'End If
                Next
            End If
        Catch e_search_1418 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_1418 --> " & e_search_1418.Message, EventLogEntryType.Information, 44)

        End Try
    End Sub

    'Sub ThreadReceive()
    '    Try
    '        Dim t_ReceiveDNSFound As New Thread(AddressOf ReceiveDNSFound)
    '        t_ReceiveDNSFound.Name = "ReceiveDNSFound"

    '        't_ReceiveDNSFound.Priority = priority

    '        If LCase(Trim(s_p_Priority)) = "lowest" Then
    '            t_ReceiveDNSFound.Priority = Lowest
    '        End If
    '        If LCase(Trim(s_p_Priority)) = "normal" Then
    '            t_ReceiveDNSFound.Priority = Normal
    '        End If
    '        If LCase(Trim(s_p_Priority)) = "above normal" Then
    '            t_ReceiveDNSFound.Priority = AboveNormal
    '        End If
    '        If LCase(Trim(s_p_Priority)) = "below normal" Then
    '            t_ReceiveDNSFound.Priority = BelowNormal
    '        End If
    '        If LCase(Trim(s_p_Priority)) = "highest" Then
    '            t_ReceiveDNSFound.Priority = Highest
    '        End If

    '        t_ReceiveDNSFound.Start()
    '    Catch e_search_1490 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " e_search_1490--> " & e_search_1490.Message, EventLogEntryType.Information, 44)
    '        'MsgBox("The error (e_LookForFiles_955) is " & e_LookForFiles_955.Message)
    '    End Try
    'End Sub

    'Public Sub RichTextEvent_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)

    'End Sub

    Public i_resetPBar As Int32
    Public i_setTheCounterBack As Int32

    Sub SendURLFound(ByVal sSendstrURI As String)
        Dim myConnectionString As String
        Dim mySelectQuery As String
        Dim myConnection As New OleDbConnection(ConnStringURL(sSendstrURI))
        Dim myReader As OleDbDataReader

        Try

            If s_p_authority <> "yes" Then

                Dim ID, URL, DateHour, Machine, Signature, ShortURL, Keywords As String
                Dim s_no As String = "n"
                Dim b_record As Boolean = False
                Dim myConnString As String

                mySelectQuery = "SELECT Top 1 * FROM URLFOUND where GoodSpot <> -40004 and Done ='" & Trim(s_no) & "' order by GoodSpot desc"

                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
                Dim i_goodSpot As Int32

                myConnection.Open()

                myReader = myCommand.ExecuteReader()

                While myReader.Read()

                    ID = myReader("ID")

                    URL = myReader("URL")

                    DateHour = myReader("DateHour")

                    Machine = s_r_machineName

                    If IsNumeric(myReader("Signature")) Then
                        Signature = myReader("Signature")
                    Else
                        Signature = 0
                    End If

                    If IsNumeric(myReader("Keywords")) Then
                        Keywords = myReader("Keywords")
                    Else
                        Keywords = 0
                    End If

                    If Len(Keywords) = 0 Then
                        Keywords = 0
                    End If

                    Signature = Keywords & "00000" & Signature

                    ShortURL = myReader("ShortURL")

                    If IsNumeric(myReader("GoodSpot")) Then
                        i_goodSpot = myReader("GoodSpot")
                    Else
                        i_goodSpot = 0
                    End If

                    b_record = True

                    Exit While

                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()

                If b_record Then

                    Dim s_sa As String = "sa"
                    Dim objConn As New OleDbConnection(ConnStringURL(sSendstrURI))
                    myConnectionString = "UPDATE URLFOUND SET Done= '" & (s_sa) & "', Machine = '" & CStr(Now()) & "'  WHERE URL ='" & URL & "'"
                    Dim myCommand2 As New OleDbCommand(myConnectionString)
                    myCommand2.Connection = objConn
                    objConn.Open()
                    myCommand2.ExecuteNonQuery()
                    objConn.Close()

                    '//----
                    Dim s_verb As String = "InsertDocumentFound"
                    Dim s_packageToSend As String = ID & "|||||" & URL & "|||||" & DateHour & "|||||" & Machine & "|||||" & Signature & "|||||" & sSendstrURI & "|||||" & i_goodSpot & "|||||"

                    Call Subwrite(s_packageToSend, "search", "1308")

                    s_packageToSend = Replace(s_packageToSend, "_", "!!!!!")

                    Dim sAction As String = "URL_sent"
                    Call BroadCastDNSFOUND(s_verb, s_packageToSend, , , , sAction)
                    '\\----

                Else

                    If i_setTheCounterBack >= 50 Then

                        Dim s_n As String = "n"
                        Dim s_sa As String = "sa"
                        Dim objConn As New OleDbConnection(ConnStringURL(sSendstrURI))
                        myConnectionString = "UPDATE URLFOUND SET Done= '" & (s_n) & "', Machine = '" & CStr(Now()) & "'  WHERE Done ='" & s_sa & "'"
                        Call Subwrite(myConnectionString, "search", "1321")
                        Dim myCommand2 As New OleDbCommand(myConnectionString)
                        myCommand2.Connection = objConn
                        objConn.Open()
                        myCommand2.ExecuteNonQuery()
                        objConn.Close()
                        i_setTheCounterBack = 0

                    Else

                        i_setTheCounterBack = i_setTheCounterBack + 1

                    End If

                    'Console.WriteLine(vbNewLine)
                    'Console.WriteLine(i_setTheCounterBack & "      i_setTheCounterBack")
                    'Console.WriteLine(vbNewLine)

                End If

            End If

        Catch e_search_1329 As Exception

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_search_1329 --> " & e_search_1329.Message, EventLogEntryType.Information, 44)
            EventLog.WriteEntry(sSource, " lookforfiles.e_search_1329 myConnectionString ---- mySelectQuery ---- sSendstrURI --> " & myConnectionString & " ---- " & mySelectQuery & " ---- " & sSendstrURI, EventLogEntryType.Information, 44)

        End Try
    End Sub

    Function CheckSubQueryIn(ByVal s_originalQuery As String, ByVal s_subQueryToCheck As String) As String
        Try
            Dim s_keepNameFile As String
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim mySelectQuery As String = "SELECT * FROM SUBQUERY where Query ='" & s_originalQuery & "' and SubQuery='" & s_subQueryToCheck & "'"

            Dim myConnection1 As New OleDbConnection(ConnStringDNA())

            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer

            While myReader.Read()
                s_keepNameFile = myReader("SubQuery")
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            '''myConnection1 = Nothing
            myConnection1.Close()

            If Len(s_keepNameFile) > 0 Then
                'MsgBox("This subquery is already in, to change it, please remove it and add it again!")
                CheckSubQueryIn = "stop"
            Else
                CheckSubQueryIn = "go"
            End If

        Catch e_search_1675 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_search_1675--> " & e_search_1675.Message, EventLogEntryType.Information, 44)
            'MsgBox("The error (e_LookForFiles_955) is " & e_LookForFiles_955.Message)
        End Try
    End Function


    Public i_xxx As Int32


    Public i_startTheCounter As Int32


    Public i_xxxx As Integer


    '
    Public Function CheckInternetConnection() As Boolean
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try
            Dim b_continue As Boolean = True
            Dim s_hostName As String = System.Net.Dns.GetHostEntry("google.com").HostName
            Dim b_internetConnected As Boolean = True

            If s_hostName = "google.com" Then
                'EventLog.WriteEntry(sSource, " e_search_2606, we are connected to the internet ", EventLogEntryType.Information, 600)
                CheckInternetConnection = True
            Else

                Do While b_continue
                    s_hostName = System.Net.Dns.GetHostEntry("142.169.9.118").HostName
                    If s_hostName = "www.telussolutionsdaffaires.com" Then
                        CheckInternetConnection = True
                        Exit Do
                    End If
                    'EventLog.WriteEntry(sSource, " e_search_2614, we are not connected to the internet, and that's bad! ", EventLogEntryType.Information, 600)
                    'Console.WriteLine(" e_search_2614, we are not connected to the internet, and that's bad! ")
                    System.Threading.Thread.Sleep(30000)
                Loop

            End If
        Catch e_search_2619 As Exception
            EventLog.WriteEntry(sSource, " e_search_2619--> " & e_search_2619.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Public b_p_closedTheSeedingProcess As Boolean

    Sub Subwrite(Optional ByVal s_variable As String = "", Optional ByVal s_form As String = "", Optional ByVal s_lineNumber As String = "")
        'Console.WriteLine(s_variable & " -- " & s_form & "." & s_lineNumber)
    End Sub


    Function CheckMachineNamePublished() As Boolean

        Try

            Dim s_hostName As String = System.Environment.MachineName()
            Dim s_verb As String = "NewMachineName"

            If sWhatRange() Then
                If sWhatRange(s_hostName) Then
                    BroadCastDNSFOUND(s_verb, s_hostName)
                Else
                    CheckMachineNamePublished = True
                End If
            Else
                Dim myConnection As New OleDbConnection(ConnStringDNA())
                Dim b_continue As Boolean = True
                Dim myConnectionString, s, s_queryName As String
                Dim myInsertQuery As String = "INSERT INTO RANGE (HostName)" _
                & "Values ('" & Trim(s_hostName) & "')"
                Dim myCommand As New OleDbCommand(myInsertQuery)
                myCommand.Connection = myConnection
                myConnection.Open()
                myCommand.ExecuteNonQuery()
                myCommand.Connection.Close()
                myConnection.Close()
                BroadCastDNSFOUND(s_verb, s_hostName)
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " search.1627--> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public d_p_lastTimeDone As DateTime
    Public b_p_lastTimeDone As Boolean


    Sub PublishMachineName()

        Dim i_timeDiffMinute As Int64
        Dim i_counter As Int32
        Dim b_continue As Boolean = True

        Try

            i_timeDiffMinute = DateDiff(DateInterval.Minute, d_p_lastTimeDone, Now)

            If i_timeDiffMinute >= 60 Or b_p_lastTimeDone = False Then
                Dim s_hostName As String = System.Environment.MachineName()
                Dim s_verb As String = "NewMachineName"
                Do While b_continue
                    BroadCastDNSFOUND(s_verb, s_r_machineName)
                    System.Threading.Thread.Sleep(1000)
                    i_counter = i_counter + 1
                    If i_counter >= 2 Then
                        Exit Do
                    End If
                Loop
                d_p_lastTimeDone = Now()
            End If

            If b_p_lastTimeDone = False Then
                b_p_lastTimeDone = True
                d_p_lastTimeDone = Now()
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " search.1796--> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    'Function CheckMachineNamePublished(ByVal s_machineNameFromAbroad As String) As Boolean
    '    Try
    '        Dim s_hostName As String = s_machineNameFromAbroad

    '        If sWhatRange(s_hostName) Then
    '            Dim s_verb As String = "NewMachineName"
    '            BroadCastDNSFOUND(s_verb, s_hostName)
    '            CheckMachineNamePublished = True
    '        Else
    '            Dim s_verb As String = "AcknowNewMachineName"
    '            Dim myConnection As New OleDbConnection(ConnStringDNA())
    '            Dim b_continue As Boolean = True

    '            Dim myConnectionString, s, s_queryName As String

    '            Dim myInsertQuery As String = "INSERT INTO RANGE (HostName)" _
    '            & "Values ('" & Trim(s_hostName) & "')"
    '            Dim myCommand As New OleDbCommand(myInsertQuery)
    '            myCommand.Connection = myConnection
    '            myConnection.Open()
    '            myCommand.ExecuteNonQuery()
    '            myCommand.Connection.Close()
    '            myConnection.Close()

    '            BroadCastDNSFOUND(s_verb, s_hostName)

    '            CheckMachineNamePublished = True
    '        End If

    '    Catch e_search_1715 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " e_search_1627--> " & e_search_1715.Message, EventLogEntryType.Information, 44)
    '    End Try
    'End Function


End Module
