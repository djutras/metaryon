
Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.OleDb
Imports System.Net
Imports System
Imports System.IO
Imports System.Collections
Imports System.Configuration
Imports System.Text
Imports System.Management

Module ExportQuery

    Sub ExportParameters()

        Try
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            Dim sToInsert, myConnectionString As String
            Dim sDate As String = Now()
            Dim iLevel, iMaxPage As Integer
            Dim iDuration As Double
            Dim iMinKPage, iMaxKPage, iMinDensity, iMaxDensity As Integer
            Dim iKeywordTitle, iKeywordBody, IkeywordTag, iKeywordVisible, iKeywordURL As Integer
            Dim iMinDiskPerc, IDiskMinGiga As Integer
            Dim iAlgorythm, iGoogle, iAllTheWeb, iExactPhrase As Integer
            Dim iSearchExcel, IsearchWord, iSearchPDF As Integer
            Dim idesactiveQuery As Integer
            Dim s_Priority, sQuery As String
            Dim bGoAhead As Boolean = False
            Dim bInsert As Boolean = False
            Dim bUpdate As Boolean = False
            Dim bStopTheInsertion As Boolean = False
            Dim b_goSuccessfull As Boolean = False
            Dim sWebSite As String = "n/a"
            Dim sNetwork As String = "n/a"
            Dim sSecondNetwork As String = "n/a"
            Dim sUniversalNetwork As String = "n/a"
            Dim sMachineAddress As String = "127.0.0.1"
            Dim sSecondaryMachineAddress = "127.0.0.1"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myConnString As String
            Dim sKeepString As String
            Dim sDNSFound As String
            Dim verb As String
            Dim s_Now As String = Now()
            Dim s_verb As String = ""
            Dim s_querySchedule As String = ""
            Dim s_insertDNA As String = ""
            Dim s_queryScheduleUpdate As String = ""
            Dim s_insertDNAUpdate As String = ""
            Dim s_no As String = "no"
            Dim s_yes As String = "yes"
            Dim s_machineName, s As String
            Dim s_AllTheTime, s_WeekTimeStart, s_WeekTimeStop, s_MondayTimeStart, s_MondayTimeStop,
            s_TuesdayTimeStart, s_TuesdayTimeStop, s_WednesdayTimeStart, s_WednesdayTimeStop,
            s_ThursdayTimeStart, s_ThursdayTimeStop, s_FridayTimeStart, s_FridayTimeStop,
            s_SaturdayTimeStart, s_SaturdayTimeStop, s_SundayTimeStart, s_SundayTimeStop As String

            '//----Update query

            Dim mySelectQuery As String = "SELECT * FROM DNAVALUE, SCHEDULE where s_IDQuery = s_Query and s_Authority='" & s_yes & "'"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer

            While myReader.Read()
                b_goSuccessfull = True
                'If sQuery = myReader("s_Query") Then
                sQuery = (myReader("s_Query"))
                iDuration = myReader("i_Duration")
                iLevel = myReader("i_Level")
                iMaxPage = myReader("i_MaxPage")
                s_Priority = myReader("s_Priority")
                iMinKPage = myReader("i_MinKPage")
                iMaxKPage = myReader("i_MaxKPage")
                iMinDensity = myReader("i_MinDensity")
                iMaxDensity = myReader("i_MaxDensity")
                iKeywordTitle = myReader("i_KeywordTitle")
                iKeywordBody = myReader("i_KeywordBody")
                IkeywordTag = myReader("I_keywordTag")
                iKeywordVisible = myReader("i_KeywordVisible")
                iKeywordURL = myReader("i_KeywordURL")
                iMinDiskPerc = myReader("i_MinDiskPerc")
                IDiskMinGiga = myReader("I_MinDiskGiga")
                iAlgorythm = myReader("i_Algorythm")
                iGoogle = myReader("i_Google")
                iAllTheWeb = myReader("i_AllTheWeb")
                iSearchExcel = myReader("i_SearchExcel")
                IsearchWord = myReader("i_SearchWord")
                iSearchPDF = myReader("i_SearchPDF")
                idesactiveQuery = myReader("DNAVALUE.i_DesactiveQuery")
                iExactPhrase = myReader("i_ExactPhrase")
                sUniversalNetwork = myReader("s_UniversalNetwork")
                sNetwork = myReader("s_SecondNetwork")
                sSecondNetwork = myReader("s_FirstNetwork")
                sWebSite = myReader("s_WebSite")

                s_AllTheTime = myReader("s_AllTheTime")
                s_WeekTimeStart = myReader("s_WeekTimeStart")
                s_WeekTimeStop = myReader("s_WeekTimeStop")
                s_MondayTimeStart = myReader("s_MondayTimeStart")
                s_MondayTimeStop = myReader("s_MondayTimeStop")
                s_TuesdayTimeStart = myReader("s_TuesdayTimeStart")
                s_TuesdayTimeStop = myReader("s_TuesdayTimeStop")
                s_WednesdayTimeStart = myReader("s_WednesdayTimeStart")
                s_WednesdayTimeStop = myReader("s_WednesdayTimeStop")
                s_ThursdayTimeStart = myReader("s_ThursdayTimeStart")
                s_ThursdayTimeStop = myReader("s_ThursdayTimeStop")
                s_FridayTimeStart = myReader("s_FridayTimeStart")
                s_FridayTimeStop = myReader("s_FridayTimeStop")
                s_SaturdayTimeStart = myReader("s_SaturdayTimeStart")
                s_SaturdayTimeStop = myReader("s_SaturdayTimeStop")
                s_SundayTimeStart = myReader("s_SundayTimeStart")
                s_SundayTimeStop = myReader("s_SundayTimeStop")


                ' End If
                'MsgBox("myReader is - " & myReader("s_Query"))

                '//----Update query

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                ''myConnection1 = Nothing
                myConnection1.Close()

                myConnectionString = "UPDATE DNAVALUE SET d_DateTime='" & (s_Now) & "',s_Priority= '" & (s_Priority) & "', s_Query= '" & (sQuery) & "',i_DesactiveQuery = " & (idesactiveQuery) & ",i_SearchPDF = " & (iSearchPDF) & ",i_SearchWord = " & (IsearchWord) & ",i_SearchExcel = " & (iSearchExcel) & ",i_ExactPhrase = " & (iExactPhrase) & ",i_Duration =  " & (iDuration) & ",i_Level =  " & (iLevel) & ",i_MaxPage =  " & (iMaxPage) & ",i_MinKPage =  " & (iMinKPage) & ",i_MaxKPage =  " & (iMaxKPage) & ",i_MinDensity = " & (iMinDensity) & ",i_MaxDensity = " & (iMaxDensity) & ",i_KeywordTitle =" & (iKeywordTitle) & ",i_KeywordURL = " & (iKeywordURL) & ",i_KeywordBody =  " & (iKeywordBody) & ",i_KeywordTag =" & (IkeywordTag) & ",i_algorythm =  " & (iAlgorythm) & ",i_KeywordVisible = " & (iKeywordVisible) & ",i_AllTheWeb =" & (iAllTheWeb) & ",i_Google = " & (iGoogle) & ",i_MinDiskPerc =" & (iMinDiskPerc) & ",i_MinDiskGiga =  " & (IDiskMinGiga) & ",s_FirstNetwork =  '" & (sNetwork) & "',s_SecondNetwork =  '" & (sSecondNetwork) & "',s_UniversalNetwork =  '" & (sUniversalNetwork) & "',s_MachineAddress =  '" & (sMachineAddress) & "',s_SecondaryMachineAddress =  '" & (sSecondaryMachineAddress) & "',s_WebSite =  '" & (sWebSite) & "' WHERE s_Query ='" & sQuery & "'"
                s_insertDNAUpdate = myConnectionString
                s_insertDNAUpdate = Replace(s_insertDNAUpdate, "_", "///*****\\\")
                s_insertDNAUpdate = Replace(s_insertDNAUpdate, " ", "_")


                'EventLog.WriteEntry(sSource, myConnectionString & " ExportQuery.myConnectionString.130")
                'MessageBox.Show(myConnectionString, " ExportQuery.myConnectionString.130")

                myConnectionString = ""
                myConnectionString = "UPDATE SCHEDULE SET s_IDQuery= '" & (sQuery) & "',i_DesactiveQuery = '" & (idesactiveQuery) & "',s_AllTheTime = '" & (s_AllTheTime) & "',s_WeekTimeStart = '" & (s_WeekTimeStart) & "',s_WeekTimeStop = '" & (s_WeekTimeStop) & "',s_MondayTimeStart = '" & (s_MondayTimeStart) & "',s_MondayTimeStop = '" & (s_MondayTimeStop) & "',s_TuesdayTimeStart =  '" & (s_TuesdayTimeStart) & "',s_TuesdayTimeStop =  '" & (s_TuesdayTimeStop) & "',s_WednesdayTimeStart =  '" & (s_WednesdayTimeStart) & "',s_WednesdayTimeStop =  '" & (s_WednesdayTimeStop) & "',s_ThursdayTimeStart =  '" & (s_ThursdayTimeStart) & "',s_ThursdayTimeStop = '" & (s_ThursdayTimeStop) & "',s_FridayTimeStart = '" & (s_FridayTimeStart) & "',s_FridayTimeStop ='" & (s_FridayTimeStop) & "',s_SaturdayTimeStart = '" & (s_SaturdayTimeStart) & "',s_SaturdayTimeStop =  '" & (s_SaturdayTimeStop) & "',s_SundayTimeStart ='" & (s_SundayTimeStart) & "',s_SundayTimeStop =  '" & (s_SundayTimeStop) & "' WHERE s_IDQuery ='" & sQuery & "'"
                'MessageBox.Show(myConnectionString, " exportquery.myConnectionString.131")

                s_verb = "update"

                s_queryScheduleUpdate = myConnectionString
                s_queryScheduleUpdate = Replace(s_queryScheduleUpdate, "_", "///*****\\\")
                s_queryScheduleUpdate = Replace(s_queryScheduleUpdate, " ", "_")

                'sURIToPass = sQuery

                'sURIToPass = Replace(sURIToPass, " ", "_")

                s_queryScheduleUpdate = s_insertDNAUpdate & "|||||" & s_queryScheduleUpdate

                '\\----

                '//----Insert Query
                Dim myInsertQuery As String =
                "INSERT INTO DNAVALUE (s_Query," _
                & "s_Priority," _
                & "i_DesactiveQuery," _
                & "i_SearchPDF," _
                & "i_SearchWord," _
                & "i_SearchExcel," _
                & "i_ExactPhrase," _
                & "i_Duration," _
                & "i_Level," _
                & "i_MaxPage," _
                & "i_MinKPage," _
                & "i_MaxKPage," _
                & "i_MinDensity," _
                & "i_MaxDensity," _
                & "i_KeywordTitle," _
                & "i_KeywordURL," _
                & "i_KeywordBody," _
                & "i_KeywordTag," _
                & "i_algorythm," _
                & "i_KeywordVisible," _
                & "i_AllTheWeb," _
                & "i_Google," _
                & "i_MinDiskPerc," _
                & "i_MinDiskGiga," _
                & "s_FirstNetwork," _
                & "s_SecondNetwork," _
                & "s_UniversalNetwork," _
                & "s_MachineAddress," _
                & "s_SecondaryMachineAddress," _
                & "d_DateTime," _
                & "s_Authority," _
                & "s_WebSite)" _
                & "Values('" & (sQuery) & "'," _
                & "'" & (s_Priority) & "'," _
                & "" & (idesactiveQuery) & "," _
                & "" & (iSearchPDF) & "," _
                & "" & (IsearchWord) & "," _
                & "" & (iSearchExcel) & "," _
                & "" & (iExactPhrase) & "," _
                & "" & (iDuration) & "," _
                & "" & (iLevel) & "," _
                & "" & (iMaxPage) & "," _
                & "" & (iMinKPage) & "," _
                & "" & (iMaxKPage) & "," _
                & "" & (iMinDensity) & "," _
                & "" & (iMaxDensity) & "," _
                & "" & (iKeywordTitle) & "," _
                & "" & (iKeywordURL) & "," _
                & "" & (iKeywordBody) & "," _
                & "" & (IkeywordTag) & "," _
                & "" & (iAlgorythm) & "," _
                & "" & (iKeywordVisible) & "," _
                & "" & (iAllTheWeb) & "," _
                & "" & (iGoogle) & "," _
                & "" & (iMinDiskPerc) & "," _
                & "" & (IDiskMinGiga) & "," _
                & "'" & (sNetwork) & "'," _
                & "'" & (sSecondNetwork) & "'," _
                & "'" & (sUniversalNetwork) & "'," _
                & "'" & (sMachineAddress) & "'," _
                & "'" & (sSecondaryMachineAddress) & "'," _
                & "'" & (s_Now) & "'," _
                & "'" & (s_no) & "'," _
                & "'" & (sWebSite) & "')"

                s_insertDNA = myInsertQuery
                s_insertDNA = Replace(s_insertDNA, "_", "///*****\\\")
                s_insertDNA = Replace(s_insertDNA, " ", "_")

                'MsgBox("exportquery.myInsertQuery.215 - " & myInsertQuery)

                myInsertQuery = ""
                myInsertQuery = "INSERT INTO SCHEDULE (s_IDQuery,i_DesactiveQuery,s_AllTheTime,s_WeekTimeStart,s_WeekTimeStop,s_MondayTimeStart,s_MondayTimeStop,s_TuesdayTimeStart,s_TuesdayTimeStop,s_WednesdayTimeStart,s_WednesdayTimeStop,s_ThursdayTimeStart,s_ThursdayTimeStop,s_FridayTimeStart,s_FridayTimeStop,s_SaturdayTimeStart,s_SaturdayTimeStop,s_SundayTimeStart,s_SundayTimeStop) Values('" & (sQuery) & "','" & (idesactiveQuery) & "','" & (s_AllTheTime) & "','" & (s_WeekTimeStart) & "','" & (s_WeekTimeStop) & "','" & (s_MondayTimeStart) & "','" & (s_MondayTimeStop) & "','" & (s_TuesdayTimeStart) & "','" & (s_TuesdayTimeStop) & "','" & (s_WednesdayTimeStart) & "','" & (s_WednesdayTimeStop) & "','" & (s_ThursdayTimeStart) & "','" & (s_ThursdayTimeStop) & "','" & (s_FridayTimeStart) & "','" & (s_FridayTimeStop) & "','" & (s_SaturdayTimeStart) & "','" & (s_SaturdayTimeStop) & "','" & (s_SundayTimeStart) & "','" & (s_SundayTimeStop) & "')"

                s_querySchedule = myInsertQuery
                s_querySchedule = Replace(s_querySchedule, "_", "///*****\\\")
                s_querySchedule = Replace(s_querySchedule, " ", "_")

                s_querySchedule = s_insertDNA & "|||||" & s_querySchedule

                '\\----

                'sURIToPass = Replace(sURIToPass, " ", "_")

                'MsgBox("ExportQuery.BroadCastDNSFOUND.235 - " & s_verb & " - " & sURIToPass & " - " & s_queryScheduleUpdate & " - " & s_querySchedule & " - " & s_r_machineName)

                Call BroadCastDNSFOUND(s_verb, sURIToPass, s_queryScheduleUpdate, s_querySchedule, s_r_machineName)

                System.Threading.Thread.Sleep(250)

            End While

            '//----
            If b_goSuccessfull Then
                s_verb = "Successfull"
                sURIToPass = "Successfull"
                s_queryScheduleUpdate = ""
                s_querySchedule = ""
                'MsgBox("ExportQuery.BroadCastDNSFOUND.249 - " & s_verb & " - " & sURIToPass & " - " & s_queryScheduleUpdate & " - " & s_querySchedule & " - " & s_r_machineName)
                Call BroadCastDNSFOUND(s_verb, sURIToPass, s_queryScheduleUpdate, s_querySchedule, s_r_machineName)
            End If
            '\\----


            '//----
            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection1.Close()

            '\\----
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " exportQuery.278 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Public i_p_KeepHeartBeat As Integer
    Public d_p_dateHeartBeat As DateTime

    Public d_p_timeOfPrecedentHeartbeat As DateTime

    Sub SendHeartbeat()
        Try
            If s_p_authority <> "yes" Then
                Dim iKeepMinute As Integer
                Dim s_verb As String = "HeartBeat"
                If Now() >= DateAdd(DateInterval.Minute, 15, d_p_timeOfPrecedentHeartbeat) Then
                    Console.WriteLine(vbNewLine)
                    Console.WriteLine("**************************************************************************")
                    Console.WriteLine(s_verb & s_r_machineName & " -- this machine sent a heartbeat!")
                    Console.WriteLine("**************************************************************************")
                    Console.WriteLine(vbNewLine)
                    ' Call BroadCastDNSFOUND(s_verb, s_r_machineName)
                    d_p_timeOfPrecedentHeartbeat = Now()
                End If
            End If
        Catch e_exportQuery_314 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_exportQuery_314 --> " & e_exportQuery_314.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    'Sub SendAcknowHeartBeat(ByVal URIToPass)
    '    Try

    '        If s_p_authority = "yes" Then

    '            Dim iKeepMinute As Integer
    '            Dim s_verb As String = "SendAcknowHeartBeat"

    '            Console.WriteLine(vbNewLine)
    '            Console.WriteLine("***************************************************************************")
    '            Console.WriteLine("***************************************************************************")
    '            Console.WriteLine(s_verb & " from " & s_r_machineName & " to " & URIToPass & " " & Now())
    '            Console.WriteLine("***************************************************************************")
    '            Console.WriteLine("***************************************************************************")
    '            Console.WriteLine(vbNewLine)

    '            Call BroadCastDNSFOUND(s_verb, URIToPass)

    '        End If

    '    Catch eOX As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, "exportQuery.339 --> " & eOX.Message, EventLogEntryType.Information, 44)
    '    End Try

    'End Sub

    'Sub SendHeartbeat(ByVal i_totalCount As Int32, ByVal i_totalDNSCount As Int32, ByVal i_timeDNSFOUNDEmpty As Int32, ByVal i_pageView As Int64, ByVal s_databaseName As String)

    '    Try

    '        Dim iKeepMinute As Integer
    '        Dim s_verb As String = "HeartBeat"
    '        Dim s_sendTotalCount As String
    '        s_databaseName = Replace(s_databaseName, " ", "_")

    '        s_sendTotalCount = s_r_machineName & "||||||||||" & CStr(i_totalCount) & "||||||||||" & CStr(i_totalDNSCount) & "||||||||||" & CStr(i_timeDNSFOUNDEmpty) & "||||||||||" & CStr(i_pageView) & "||||||||||" & CStr(s_databaseName)

    '        iKeepMinute = DateDiff(DateInterval.Minute, d_p_dateHeartBeat, Now)

    '        Console.WriteLine(vbNewLine)
    '        Console.WriteLine("***********************************************************************************")
    '        Console.WriteLine(s_verb & s_r_machineName & " -- this machine sent a heartbeat!")
    '        Console.WriteLine("***********************************************************************************")
    '        Console.WriteLine(vbNewLine)

    '        Call BroadCastDNSFOUND(s_verb, s_sendTotalCount)

    '    Catch e_exportQuery_357 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " lookforfiles.e_exportQuery_357 --> " & e_exportQuery_357.Message, EventLogEntryType.Information, 44)
    '    End Try

    'End Sub


    'Sub InsertHeartBeatScratchPad(ByVal URIToPass)
    '    Try
    '        Dim myConnection As New OleDbConnection(ConnStringDNA())
    '        Dim s_ToInsert As String = URIToPass
    '        Dim s_date As String = Now()
    '        Dim s_yes As String = "y"
    '        Dim s_splitUriToPass() As String
    '        ReDim s_splitUriToPass(3)
    '        Dim d_date As DateTime = Now()
    '        Dim s_heartBeat As String = "Heartbeat"

    '        s_splitUriToPass = Split(URIToPass, "||||||||||")

    '        If s_p_authority = "yes" Then
    '            If UBound(s_splitUriToPass) >= 1 Then
    '                Dim s_machine As String = s_splitUriToPass(0)
    '                Dim i_count As Int32 = s_splitUriToPass(1)
    '                Dim i_totalDNSCount As Int32 = s_splitUriToPass(2)
    '                Dim i_DNSFOUNDEmpty As Int32 = s_splitUriToPass(3)
    '                Dim i_pageView As Int32 = s_splitUriToPass(4)
    '                Dim s_query As String = s_splitUriToPass(5)
    '                s_query = Replace(s_query, "_", " ")
    '                InsertIntoLoad(s_machine, i_count, i_totalDNSCount, i_DNSFOUNDEmpty, i_pageView, s_query)
    '            Else
    '                URIToPass = "HeartBeat_from_" & URIToPass
    '                Dim myInsertQuery As String = "INSERT INTO SCRATCHPAD (Url,DateHour,Done,ShortUrl)" _
    '                            & "Values ('" & Trim(s_ToInsert) & "','" & (s_date) & "','" & (s_yes) & "','" & (s_heartBeat) & "')"
    '                Dim myCommand As New OleDbCommand(myInsertQuery)
    '                myCommand.Connection = myConnection
    '                myConnection.Open()
    '                myCommand.ExecuteNonQuery()
    '                myCommand.Connection.Close()
    '                myConnection.Close()
    '            End If
    '        End If

    '    Catch e_exportQuery_326 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " lookforfiles.e_exportQuery_326 --> " & e_exportQuery_326.Message, EventLogEntryType.Information, 44)
    '    End Try
    'End Sub

    'Sub InsertHeartBeatScratchPad(ByVal acknowledgeHeartbeatFromAuthority, ByVal URIToPass)

    '    Try
    '        Dim myConnection As New OleDbConnection(ConnStringDNA())
    '        Dim s_ToInsert As String = "Authority"
    '        Dim s_date As String = Now()
    '        Dim s_yes As String = "y"
    '        Dim s_splitUriToPass() As String
    '        ReDim s_splitUriToPass(3)
    '        Dim d_date As DateTime = Now()

    '        If URIToPass = s_r_machineName And s_p_authority <> "yes" Then

    '            Dim myInsertQuery As String = "INSERT INTO SCRATCHPAD (Url,DateHour,Done)" _
    '                        & "Values ('" & Trim(s_ToInsert) & "','" & (s_date) & "','" & (s_yes) & "')"
    '            Dim myCommand As New OleDbCommand(myInsertQuery)
    '            myCommand.Connection = myConnection
    '            myConnection.Open()
    '            myCommand.ExecuteNonQuery()
    '            myCommand.Connection.Close()
    '            myConnection.Close()

    '        End If

    '    Catch eOX As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, "exportQuery.437 --> " & eOX.Message, EventLogEntryType.Information, 44)
    '    End Try
    'End Sub

    'Sub UpdateDNSFOUNDReport(ByVal URIToPass)

    '    Try

    '        Dim myConnection As New OleDbConnection(ConnStringDNA())
    '        Dim s_ToInsert As String = URIToPass
    '        Dim s_date As String = Now()
    '        Dim s_No As String = "n"
    '        Dim i_count As Int32
    '        Dim myInsertQuery As String

    '        If s_p_authority = "yes" Then

    '            Dim mySelectQuery As String = "SELECT Count(*) as CountDNS FROM REP_DNSFOUND"

    '            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
    '            myConnection.Open()
    '            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

    '            While myReader.Read()
    '                i_count = (myReader("CountDNS"))
    '            End While

    '            If Not (myReader.IsClosed) Then
    '                myReader.Close()
    '            End If
    '            myConnection.Close()

    '            If i_count >= 2000 Then
    '                myInsertQuery = "Delete from SCRATCHPAD"
    '                Call ExecuteNonQueryDNA(myInsertQuery)
    '            End If

    '            myInsertQuery = ""
    '            myInsertQuery = "INSERT INTO REP_DNS (REP_MESS,REP_TIME)" _
    '            & "Values ('" & Trim(URIToPass) & "','" & (s_date) & "')"
    '            Call ExecuteNonQueryDNA(myInsertQuery)

    '        End If

    '    Catch e_exportQuery_408 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " lookforfiles.e_exportQuery_408 --> " & e_exportQuery_408.Message, EventLogEntryType.Information, 44)
    '    End Try
    'End Sub


    'Sub ExecuteNonQueryDNA(ByVal s_myInsertQuery)

    '    Try

    '        Dim myConnection As New OleDbConnection(ConnStringDNA())

    '        Dim myCommand1 As New OleDbCommand(s_myInsertQuery)
    '        myCommand1.Connection = myConnection
    '        myConnection.Open()
    '        myCommand1.ExecuteNonQuery()
    '        myCommand1.Connection.Close()
    '        myConnection.Close()

    '    Catch e_exportQuery_408 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " lookforfiles.e_exportQuery_408 --> " & e_exportQuery_408.Message, EventLogEntryType.Information, 44)
    '    End Try

    'End Sub

    'Function CheckAuthority() As Boolean
    '    Try
    '        Dim s_keepHostName As String
    '        Dim b_nameIsIn As Boolean = False

    '        Dim mySelectQuery, s_Website, s As String

    '        mySelectQuery = "SELECT Machine FROM Query where Machine ='" & System.Environment.MachineName() & "'"

    '        Dim myConnection1 As New OleDbConnection(ConnStringDNA())
    '        Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

    '        myConnection1.Open()

    '        Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
    '        Dim Index As Integer

    '        While myReader.Read()
    '            s_keepHostName = myReader("Machine").ToString
    '            Exit While
    '        End While
    '        If Len(s_keepHostName) > 0 Then
    '            CheckAuthority = True
    '        Else
    '            CheckAuthority = False
    '        End If
    '        myReader.Close()
    '        myConnection1.Close()

    '    Catch ex As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " async.540 --> " & ex.Message, EventLogEntryType.Information, 44)
    '    End Try


    'End Function




End Module

