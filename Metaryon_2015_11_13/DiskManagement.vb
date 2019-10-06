
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

Public Module DiskManagement

    Public s_p_Cycle As String
    Public i_p_KeepHour As Integer
    Public d_p_KeepMinute As Double
    Public d_p_KeepOldMinute As Double
    Dim i_p_updateFromSAtoS As Integer


    Public i_p_keepModulo As Int32
    Public i_p_scrathpad As Int32
    Public s_p_precedentDatabaseName As String

    Public Function f_CheckCycle(ByVal s_DatabaseName)
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"

        Try

            Dim i_MinuteSinceInsert As Integer
            Dim i_HourSinceInsert As Double
            Dim i_KeepModulo As Integer
            Dim myConnectionString As String
            Dim s_No As String = "n"
            Dim s_Yes As String = "y"
            Dim d_KeepDuration As Double = i_p_Duration
            Dim d_Now As DateTime = Now()
            Dim i_KeepEntier As Integer
            Dim b_con As Boolean = True
            Dim i As Integer = 0
            Dim i_sixty As Integer = 60
            Dim d_Mod As Double
            Dim i_modUpdate As Double
            Dim b_beginCycle As Boolean

            If Len(s_p_precedentDatabaseName) > 2 Then
                If s_p_precedentDatabaseName <> s_DatabaseName Then
                    Call Sub_checkDNSNotDone(s_p_precedentDatabaseName)
                End If
            End If

            s_p_precedentDatabaseName = s_DatabaseName

            If s_p_Cycle <> "go" And s_p_Cycle <> "stop" Then
                s_p_Cycle = "go"
            End If

            If s_p_Cycle = "go" Then

                d_KeepDuration = d_KeepDuration * 60

                i_MinuteSinceInsert = DateDiff(DateInterval.Minute, d_p_DateTime, Now)

                d_p_KeepMinute = i_MinuteSinceInsert

                i_KeepModulo = i_MinuteSinceInsert Mod d_KeepDuration
                i_p_keepModulo = i_MinuteSinceInsert Mod d_KeepDuration

                d_p_KeepOldMinute = d_p_KeepMinute

                If i_p_scrathpad > 10 Or b_beginCycle Or i_p_scrathpad = 0 Then
                    b_beginCycle = CheckOldestEntryInScratchpad()
                    i_p_scrathpad = 1
                End If

                i_p_scrathpad = i_p_scrathpad + 1

                If (i_KeepModulo >= 0 And i_KeepModulo <= 59) Or b_beginCycle Then

                    Dim objConn As New OleDbConnection(ConnStringURLDNS(s_DatabaseName))
                    myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_No) & "' where Done ='" & (s_Yes) & "' "
                    Dim myCommand2 As New OleDbCommand(myConnectionString)
                    myCommand2.Connection = objConn
                    objConn.Open()
                    myCommand2.ExecuteNonQuery()
                    objConn.Close()

                    s_p_Cycle = "stop"

                    Call Sub_DeleteRecordsScratchPad()

                    'Call CompacAccess()

                    'Call UpdateFromSAtoS(s_DatabaseName)

                    Call Sub_DeleteRecordsPageCounter(s_DatabaseName)

                    If s_p_authority = "yes" Then

                        Call DeleteArchive(s_DatabaseName)
                        'Call CompacAccessUrlFound(s_DatabaseName)

                        Call ExportParameters()
                        Call DeleteFromDATADNSOlderInsert()
                        Call LoadBalanceFromDNSFOUND()
                        Call DistributeLoadFromDNSFOUND()
                        Call UpdateDNSFOUNDWithRangeEndOfCycle(s_DatabaseName)
                        Call RemoveMachineFromRAN_Broadcast()

                    End If

                End If

            Else

                i_MinuteSinceInsert = DateDiff(DateInterval.Minute, d_p_DateTime, Now)

                If (i_MinuteSinceInsert) > (d_p_KeepMinute + 60) Then

                    'EventLog.WriteEntry(sSource, " Cycle Go -- d_KeepDuration(" & d_KeepDuration & ") i_MinuteSinceInsert(" & i_MinuteSinceInsert & ") i_KeepModulo(" & i_KeepModulo & ")", EventLogEntryType.Information, 500)

                    s_p_Cycle = "go"

                End If

            End If

            'EventLog.WriteEntry(sSource, " en sortie - s_p_Cycle(" & s_p_Cycle & ")", EventLogEntryType.Information, 500)

        Catch ex As Exception
            EventLog.WriteEntry(sSource, " DiskManagement.127 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public Function f_DNSFOUNDStartANew(ByVal s_DatabaseName)

        Dim myConnectionString As String
        Dim s_no As String = "n"
        Dim s_yes As String = "y"
        Dim d_now As DateTime = Now()

        Try

            Dim objConn As New OleDbConnection(ConnStringURLDNS(s_DatabaseName))

            myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_no) & "' where Done = '" & (s_yes) & "'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = objConn
            objConn.Open()
            myCommand2.ExecuteNonQuery()
            objConn.Close()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.155 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try

    End Function

    Public Function f_FileToScan_iMinDiskPerc(ByVal iMinDiskPerc As Integer, ByVal bMakePlace As Boolean)
        Try
            Dim iPercFreeSpace As Integer
            Dim iMinFreeSpace As Integer
            Dim dSpecFreeSpace As Double = iMinDiskPerc
            'Dim path As ManagementPath
            'Dim o As ManagementObject
            Dim sTotalSpace As String
            Dim sLeftSpace As String
            Dim sPercSpace As String

            Dim iTotalSpace As Int64
            Dim iLeftSpace As Int64
            Dim dPercFreeSpace As Double
            Dim sPercFreeSpace As String

            'path = New ManagementPath
            'path.Server = System.Environment.MachineName()
            'path.NamespacePath = "root\CIMV2"
            'path.RelativePath = "Win32_LogicalDisk='c:'"
            'o = New ManagementObject(path)

            'sTotalSpace = o.Item("Size").ToString
            iTotalSpace = Convert.ToInt64(Val(sTotalSpace.ToString))

            'sLeftSpace = o.Item("freespace").ToString
            iLeftSpace = Convert.ToInt64(Val(sLeftSpace.ToString))
            'MessageBox.Show(iLeftSpace, "iLeftSpace")

            dPercFreeSpace = iLeftSpace / iTotalSpace
            dPercFreeSpace = FormatCurrency(dPercFreeSpace, 2)
            dSpecFreeSpace = dSpecFreeSpace / 100
            dSpecFreeSpace = FormatCurrency(dSpecFreeSpace, 2)

            If dSpecFreeSpace < dPercFreeSpace Then
                'MessageBox.Show(dSpecFreeSpace & " vs " & dPercFreeSpace, "Ok enough disk space in percentage!")
                bMakePlace = False
            Else
                'MessageBox.Show(dSpecFreeSpace & " vs " & dPercFreeSpace, "Not enough disk space in percentage!")
                bMakePlace = True
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.207 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try

    End Function

    Public Function f_FileToScan_IDiskMinGiga(ByVal IDiskMinGiga, ByVal bMakePlace)
        Try
            Dim iPercFreeSpace As Integer
            Dim iMinFreeSpace As Integer
            Dim dSpecFreeSpace As Double = IDiskMinGiga
            ' Dim path As ManagementPath
            ' Dim o As ManagementObject
            Dim sTotalSpace As String
            Dim sLeftSpace As String
            Dim sPercSpace As String

            Dim iTotalSpace As Int64
            Dim iLeftSpace As Int64
            Dim dPercFreeSpace As Double
            Dim sPercFreeSpace As String

            'path = New ManagementPath
            'path.Server = System.Environment.MachineName()
            'path.NamespacePath = "root\CIMV2"
            'path.RelativePath = "Win32_LogicalDisk='c:'"
            'o = New ManagementObject(path)

            'Console.WriteLine(o.Item("DriveType"))
            'MessageBox.Show(o.Item("DriveType").ToString, "DriveType")

            'Console.WriteLine(o.Item("Size"))
            'sTotalSpace = o.Item("Size").ToString
            iTotalSpace = Convert.ToInt64(Val(sTotalSpace.ToString))
            'MessageBox.Show(iTotalSpace, "Size")

            'Console.WriteLine(o.Item("freespace"))
            'sLeftSpace = o.Item("freespace").ToString
            iLeftSpace = Convert.ToInt64(Val(sLeftSpace.ToString))
            'MessageBox.Show(iLeftSpace, "iLeftSpace")

            dPercFreeSpace = iLeftSpace
            dPercFreeSpace = FormatCurrency(dPercFreeSpace, 2)
            dSpecFreeSpace = dSpecFreeSpace
            dSpecFreeSpace = FormatCurrency(dSpecFreeSpace, 2)

            If dSpecFreeSpace < dPercFreeSpace Then
                bMakePlace = False
            Else
                bMakePlace = True
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.261 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Sub Sub_checkDNSNotDone(ByVal s_databaseName)

        Dim s_no As String = "n"
        Dim s_yes As String = "y"
        Dim i_ID As Integer
        Dim mySelectQuery As String
        Dim i_totalCount, i_totalDNSCount As Int32
        Dim s_KeepDate As String = Now()
        Dim d_dateHour As Date
        Dim d_dateRatingDone As Date
        Dim b_go As Boolean = True
        Dim s_done As String
        Dim b_isThere As Boolean = False
        Dim i_pageView As Int64

        Try

            mySelectQuery = "SELECT Count(DONE) AS TOTALCOUNT FROM DNSFOUND where DONE = '" & s_no & "'"

            Dim myConnection As New OleDbConnection(ConnStringURLDNS(s_databaseName))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            Dim o_object As Object
            Dim s_RetrieveIDDateHour As Array

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            While myReader.Read()
                i_totalCount = myReader("TOTALCOUNT")
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            mySelectQuery = ""
            mySelectQuery = "SELECT Count(ID) AS TOTALIDCOUNT FROM DNSFOUND WHERE DONE = '" & s_no & "' OR DONE = '" & s_yes & "'"

            Dim myCommand1 As New OleDbCommand(mySelectQuery, myConnection)
            Dim o_object1 As Object
            Dim s_RetrieveIDDateHour1 As Array

            myConnection.Open()

            Dim myReader1 As OleDbDataReader = myCommand1.ExecuteReader()
            While myReader1.Read()
                i_totalDNSCount = myReader1("TOTALIDCOUNT")
                Exit While
            End While

            If Not (myReader1.IsClosed) Then
                myReader1.Close()
            End If

            myConnection.Close()

            i_pageView = HowManyPageViews()

            'If s_p_authority = "yes" Then
            '    InsertIntoLoad(s_r_machineName, i_totalCount, i_totalDNSCount, i_p_countTimeAfterDNSFOUNDEmpty, i_pageView, s_databaseName)
            'Else
            '    InsertIntoLoad(s_r_machineName, i_totalCount, i_totalDNSCount, i_p_countTimeAfterDNSFOUNDEmpty, i_pageView, s_databaseName)
            '    SendHeartbeat(i_totalCount, i_totalDNSCount, i_p_countTimeAfterDNSFOUNDEmpty, i_pageView, s_databaseName)
            'End If

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.296 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try

    End Sub

    Sub InsertAcknowHeartBeat(ByVal s_machineName)
        Try

            Dim s_so As String = "so"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myConnString1 = "UPDATE LOADBALANCING SET SENTTOAUTH ='" & (s_so) & "' where MACHINENAME = '" & (s_machineName) & "'"
            Dim myCommand2 As New OleDbCommand(myConnString1)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch DiskManagement_328 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.DiskManagement_328 --> " & DiskManagement_328.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub InsertIntoLoad(Optional ByVal s_machineName = "", Optional ByVal i_totalCount = 0, Optional ByVal i_totalDNSCount = 0, Optional ByVal i_DNSFOUNDEmpty = 0, Optional ByVal i_pageView = 0, Optional ByVal s_databaseName = "")
        Try

            Dim myConnString, mySelectQuery As String
            Dim sKeepString As String
            Dim date_InsertURL As DateTime = Now()
            Dim i_keepCumulative As Int32
            Dim b_alreadyIn As Boolean = False

            If Len(s_machineName) = 0 Then
                s_machineName = s_r_machineName
            End If

            mySelectQuery = "SELECT SUM(TOTALCOUNT) as CUMULATIVECOUNT FROM LOADBALANCING WHERE MACHINENAME ='" & (s_machineName) & "'"

            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                If IsNumeric(myReader("CUMULATIVECOUNT")) Then
                    i_keepCumulative = myReader("CUMULATIVECOUNT")
                    i_keepCumulative = i_keepCumulative + i_totalCount
                Else
                    i_keepCumulative = i_totalCount
                End If
                b_alreadyIn = True
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            b_alreadyIn = False

            If b_alreadyIn Then
                Dim s_so As String = "so"
                i_keepCumulative = i_keepCumulative + i_totalCount
                Dim myConnString1 = "UPDATE LOADBALANCING SET QUERY =" & (s_databaseName) & ",PAGEVIEW =" & (i_pageView) & ",TOTALCOUNT =" & (i_totalCount) & ", CUMULATIVECOUNT = " & (i_keepCumulative) & ", DATEOFENTRY = #" & (date_InsertURL) & "#  where MACHINENAME = '" & (s_machineName) & "'"
                Dim myCommand2 As New OleDbCommand(myConnString1)
                myCommand2.Connection = myConnection
                myConnection.Open()
                myCommand2.ExecuteNonQuery()
                myCommand2.Connection.Close()
                myConnection.Close()
            Else
                Dim myConnection1 As New OleDbConnection(ConnStringDNA())
                Dim s_date As String = Now()
                Dim s_splitUriToPass() As String
                ReDim s_splitUriToPass(1)
                Dim d_date As DateTime = Now()
                Dim myInsertQuery As String = "INSERT INTO LOADBALANCING (MACHINENAME, TOTALCOUNT, CUMULATIVECOUNT, TOTALDNSCOUNT, ITERATIONEMPTY, DATEOFENTRY,PAGEVIEW, QUERY)" _
                            & "Values ('" & Trim(s_machineName) & "'," & i_totalCount & "," & i_keepCumulative & "," & i_totalDNSCount & "," & i_DNSFOUNDEmpty & ",#" & (date_InsertURL) & "#," & i_pageView & ",'" & Trim(s_databaseName) & "')"
                Dim myCommand1 As New OleDbCommand(myInsertQuery)
                myCommand1.Connection = myConnection1
                myConnection1.Open()
                myCommand1.ExecuteNonQuery()
                myCommand1.Connection.Close()
                myConnection1.Close()
            End If

            If s_p_authority = "yes" Then
                Dim s_verb As String = "acknowHeartbeat"
                BroadCastDNSFOUND(s_verb, s_machineName)
            End If

        Catch DiskManagement_401 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.DiskManagement_401 --> " & DiskManagement_401.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Sub LoadBalanceFromDNSFOUND()
        Try
            Dim myConnString As String
            Dim sKeepString, s_rangeMinus, s_rangePlus, s_HostName As String
            Dim i_CounthostName As Int32
            Dim date_InsertURL As DateTime
            Dim s_machineToSend As String
            Dim s_queryBefore As String

            Dim mySelectQuery As String = "SELECT RangeMinus,RangePlus,HostName FROM RANGE order by HostName"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_HostName = myReader("HostName")
                s_rangeMinus = myReader("RangeMinus")
                s_rangePlus = myReader("RangePlus")
                Call CountDNSFOUNDByRange(s_rangeMinus, s_rangePlus, s_HostName)
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.443 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Sub CountDNSFOUNDByRange(ByVal s_rangeMinus, ByVal s_rangePlus, ByVal s_HostName)
        Dim s_query, s_queryToKeep As String
        Dim b_nextTurnStop As Boolean
        Dim b_enoughDNS As Boolean
        Dim i_keepSum, i_countDNS, i_countDNSToKeep As Int64

        Try
            Dim mySelectQuery As String = "SELECT s_Query FROM DNAVALUE"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())

            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer

            While myReader.Read()
                s_query = myReader("s_query")
                i_countDNS = CountDNSInDNSFOUND(s_query)

                If i_countDNS > 20000 Then
                    If i_countDNS > i_countDNSToKeep Then
                        i_countDNSToKeep = i_countDNS
                        s_queryToKeep = s_query
                    End If
                End If

            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection1.Close()

            Call InsertCountOfDNSInData(s_rangeMinus, s_rangePlus, s_queryToKeep, s_HostName)

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.454 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub InsertCountOfDNSInData(ByVal s_rangeMinus, ByVal s_rangePlus, ByVal s_query, ByVal s_HostName)

        Dim s_no As String = "n"
        Dim s_yes As String = "y"
        Dim i_ID As Integer
        Dim mySelectQuery As String
        Dim i_totalCount, i_totalDNSCount As Int32
        Dim s_KeepDate As String = Now()
        Dim d_dateHour As Date
        Dim d_dateRatingDone As Date
        Dim b_go As Boolean = True
        Dim s_done As String
        Dim b_isThere As Boolean = False

        Try

            mySelectQuery = "SELECT Count(ShortURL) AS TOTALShortURL FROM DNSFOUND where ShortURL >= '" & s_rangeMinus & "' AND ShortURL < '" & s_rangePlus & "'"

            Dim myConnection As New OleDbConnection(ConnStringURLDNS(s_query))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            Dim o_object As Object
            Dim s_RetrieveIDDateHour As Array

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            While myReader.Read()
                i_totalCount = myReader("TOTALShortURL")
                Call InsertCountDNSInDataDNS(s_rangeMinus, s_rangePlus, i_totalCount, s_query, s_HostName)
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.485 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub InsertCountDNSInDataDNS(ByVal s_rangeMinus, ByVal s_rangePlus, ByVal i_totalCountInRange, ByVal s_query, ByVal s_HostName)

        Dim mySelectQuery As String
        Dim b_go As Boolean = True
        Dim s_done As String
        Dim i_totalCountOfDNS, i_CounthostName As Int32
        Dim i_percentageVised, i_percentageObtain As Int32

        i_CounthostName = findCountHostname()

        If IsNumeric(i_CounthostName) Then
            If i_CounthostName > 0 Then
                i_percentageVised = CInt(Fix(100 / i_CounthostName))
            End If
        End If

        Try
            mySelectQuery = "SELECT Count(ShortURL) AS TOTALShortURL FROM DNSFOUND"

            Dim myConnection As New OleDbConnection(ConnStringURLDNS(s_query))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            Dim o_object As Object
            Dim s_RetrieveIDDateHour As Array

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            While myReader.Read()
                i_totalCountOfDNS = myReader("TOTALShortURL")
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            If IsNumeric(i_totalCountInRange) And IsNumeric(i_totalCountOfDNS) Then
                i_percentageObtain = CInt(Fix((i_totalCountInRange / i_totalCountOfDNS) * 100))
            End If

            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim s_date As String = Now()
            Dim s_splitUriToPass() As String
            ReDim s_splitUriToPass(1)
            Dim d_date As DateTime = Now()
            Dim myInsertQuery As String = "INSERT INTO DATADNS (Machine, RangeMinus, RangePlus, TotalDNSInRange, TotalDNS, PercentageVised, PercentageObtain,DateOfInsert)" _
                        & "Values ('" & Trim(s_HostName) & "','" & Trim(s_rangeMinus) & "','" & s_rangePlus & "'," & i_totalCountInRange & "," & i_totalCountOfDNS & "," & i_percentageVised & "," & i_percentageObtain & ",#" & (d_date) & "#)"
            Dim myCommand1 As New OleDbCommand(myInsertQuery)
            myCommand1.Connection = myConnection1
            myConnection1.Open()
            myCommand1.ExecuteNonQuery()
            myCommand1.Connection.Close()
            myConnection1.Close()

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.531 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub


    Function findCountHostname() As Integer
        Try
            Dim myConnString As String
            Dim sKeepString, s_rangeMinus, s_rangePlus As String
            Dim i_CounthostName As Int32
            Dim date_InsertURL As DateTime
            Dim s_machineToSend As String
            Dim s_queryBefore As String

            Dim mySelectQuery As String = "SELECT count(HostName) as TotalHostName FROM RANGE"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()

                i_CounthostName = myReader("TotalHostName")
                findCountHostname = i_CounthostName
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.617 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Function CheckOldestEntryInScratchpad() As Boolean
        Try
            Dim d_KeepDuration As Double = i_p_Duration
            Dim b_c As Boolean = False
            Dim s_splitQuery() As String
            Dim s_arrayToQuery() As String
            ReDim s_arrayToQuery(1)
            d_KeepDuration = d_KeepDuration * 60
            Dim d_dateLastCycle As DateTime = DateAdd(DateInterval.Minute, -d_KeepDuration, Now())


            Dim mySelectQuery As String = "SELECT URL FROM SCRATCHPAD where DateHour < #" & d_dateLastCycle & "#"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer

            While myReader.Read()
                CheckOldestEntryInScratchpad = True
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection1.Close()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.660 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public d_p_LastUpdateOfLoadFromDNS As DateTime
    Public d_p_dateFutureDay As DateTime
    Public b_p_beginWorkInDistributeLoad As Boolean = True

    Sub DistributeLoadFromDNSFOUND()
        Try
            Dim d_KeepDuration As Double = i_p_Duration
            Dim b_c As Boolean = False
            Dim s_splitQuery() As String
            Dim s_arrayToQuery() As String
            ReDim s_arrayToQuery(1)
            d_KeepDuration = d_KeepDuration * 60
            Dim d_dateLastHour As DateTime = DateAdd(DateInterval.Minute, -120, Now())
            Dim s_machine, s_machineBefore As String
            Dim i_percVised, i_percVisedLocal, i_percVisedOther, i_percObtain, i_totalPercObtain, i_resultPercObtain, i_minRange, i_maxRange As Int32
            Dim b_alert As Boolean
            Dim i_machineNumber As Int64 = FindMachineNumber()
            Dim i_id As Int64

            If d_p_LastUpdateOfLoadFromDNS >= d_p_dateFutureDay Or b_p_beginWorkInDistributeLoad Then

                b_p_beginWorkInDistributeLoad = False
                d_p_dateFutureDay = DateAdd(DateInterval.Day, 1, Now())
                d_p_LastUpdateOfLoadFromDNS = Now()

                Dim mySelectQuery As String = "SELECT TOP " & i_machineNumber & " Machine,PercentageObtain,PercentageVised, ID FROM DATADNS ORDER BY ID DESC"
                Dim myConnection1 As New OleDbConnection(ConnStringDNA())
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
                myConnection1.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                Dim i_index, i_indexToKeep As Integer

                While myReader.Read()

                    i_id = myReader("ID")
                    s_machine = myReader("Machine")
                    i_percVised = myReader("PercentageVised")
                    i_percObtain = myReader("PercentageObtain")

                    If s_r_machineName = s_machine Then
                        i_percVisedLocal = i_percVised / 2
                        i_minRange = i_percVisedLocal - (i_percVised / 2)
                        If i_minRange < 0 Then i_minRange = 0
                        i_maxRange = i_percVisedLocal + (i_percVised / 2)
                    Else
                        i_percVisedOther = i_percVised
                        i_percVisedOther = (i_percVisedOther + (i_percVisedOther * ((i_percVisedOther / 100)) / 2))
                        i_minRange = i_percVisedOther - (i_percVised / 2)
                        If i_minRange < 0 Then i_minRange = 0
                        i_maxRange = i_percVisedOther + (i_percVised / 2)
                    End If

                    If i_percObtain < i_minRange Or i_percObtain > i_maxRange Then
                        b_alert = True
                        Exit While
                    End If

                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection1.Close()

                If b_alert Then
                    DoTheNewRange(i_percVised)
                End If

            End If


        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.764 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub


    Sub DoTheNewRange(ByVal i_percVised)

        Dim i_ID As Integer
        Dim mySelectQuery As String
        Dim i_totalCount, i_totalDNSCount As Int64
        Dim s_KeepDate As String = Now()
        Dim d_dateHour As Date
        Dim d_dateRatingDone As Date
        Dim b_go As Boolean = True
        Dim s_done, s_refreshing As String
        Dim b_isThere As Boolean = False
        Dim s_query, s_minRange, s_maxRange As String
        Dim i_rangeNumberOfDNS, i_rangeNumberOfDNSToAskFor As Int64
        Dim b_continue As Boolean = True
        Dim s_shortURL, s_shortURLBefore, s_hostMachine, s_firstLetter, s_secondLetter As String
        Dim i_counter As Int64 = 0
        Dim i_minRange, i_maxRange, i_percVisedLocal, i_percVisedOther As Int64
        Dim b_transaction As Boolean
        Dim i_machineNumber, i_lenShortURL, i_counterOfDNS, i_fLetter, i_sLetter As Int64

        Try

            'Begin transaction put an old insert in scratchpad
            Call PutOldInsertInScratchpad()

            i_machineNumber = FindMachineNumber()

            s_query = Find_s_Query()

            If s_query <> "stop" Then

                i_totalDNSCount = i_p_countDNSFoundFroms_s_Query

                While b_continue

                    s_hostMachine = GetNameOfHostMachine(s_hostMachine)

                    If s_r_machineName = s_hostMachine Then
                        i_percVisedLocal = i_percVised - (i_percVised / 2)
                    Else
                        i_percVisedLocal = (i_percVised + (i_percVised * ((i_percVised / 100) / 2)))
                    End If

                    i_rangeNumberOfDNS = (i_totalDNSCount * (i_percVisedLocal / 100)) + 300

                    i_rangeNumberOfDNSToAskFor = i_rangeNumberOfDNSToAskFor + i_rangeNumberOfDNS

                    If i_rangeNumberOfDNSToAskFor >= i_totalDNSCount Then
                        i_rangeNumberOfDNSToAskFor = i_totalDNSCount
                    End If

                    mySelectQuery = "SELECT Top " & i_rangeNumberOfDNSToAskFor & " ShortURL FROM DNSFOUND order by ShortURL"

                    Dim myConnection As New OleDbConnection(ConnStringURLDNS(s_query))
                    Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
                    Dim o_object As Object
                    Dim s_RetrieveIDDateHour As Array

                    myConnection.Open()

                    Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                    i_counterOfDNS = 0

                    While myReader.Read()
                        If i_counterOfDNS > i_rangeNumberOfDNSToAskFor - 300 Then
                            s_shortURL = myReader("ShortURL")
                            i_lenShortURL = Len(s_shortURL)

                            If i_lenShortURL = 1 Then
                                s_shortURL = Mid(s_shortURL, 1, 1)
                                If s_shortURL = s_shortURLBefore Then
                                    s_shortURL = FindNextMinRangeMoreTwo(s_shortURL)
                                End If
                                s_shortURLBefore = s_shortURL
                                If IsNumeric(GiveIntFromLetter(s_firstLetter)) Then
                                    Exit While
                                End If
                            Else
                                s_shortURL = Mid(s_shortURL, 1, 2)
                                If s_shortURL = s_shortURLBefore Then
                                    s_shortURL = FindNextMinRangeMoreTwo(s_shortURL)
                                End If
                                s_shortURLBefore = s_shortURL
                                s_firstLetter = Trim(Mid(s_shortURL, 1, 1))
                                s_secondLetter = Trim(Mid(s_shortURL, 2, 1))
                                If IsNumeric(s_firstLetter) Then
                                    s_firstLetter = GiveLetterFromInt(s_firstLetter)
                                    s_shortURL = s_firstLetter & s_secondLetter
                                End If
                                If IsNumeric(s_secondLetter) Then
                                    s_secondLetter = GiveLetterFromInt(s_secondLetter)
                                    s_shortURL = s_firstLetter & s_secondLetter
                                End If
                                i_fLetter = GiveIntFromLetter(s_firstLetter)
                                i_sLetter = GiveIntFromLetter(s_secondLetter)
                                If IsNumeric(i_fLetter) And IsNumeric(i_sLetter) Then
                                    If i_fLetter <> 0 And i_sLetter <> 0 Then
                                        Exit While
                                    End If
                                End If
                            End If
                        End If

                        i_counterOfDNS = i_counterOfDNS + 1

                    End While

                    If Not (myReader.IsClosed) Then
                        myReader.Close()
                    End If

                    myConnection.Close()

                    s_firstLetter = Trim(Mid(s_shortURL, 1, 1))
                    s_secondLetter = Trim(Mid(s_shortURL, 2, 1))
                    i_fLetter = GiveIntFromLetter(s_firstLetter)
                    i_sLetter = GiveIntFromLetter(s_secondLetter)
                    If IsNumeric(i_fLetter) And IsNumeric(i_sLetter) Then
                        If i_sLetter = 0 Then
                            s_shortURL = i_fLetter & "a"
                        End If
                    End If

                    i_counter = i_counter + 1

                    If i_counter = 1 Then
                        s_minRange = "0"
                        s_maxRange = s_shortURL
                        s_refreshing = "False"
                    ElseIf i_counter = i_machineNumber Then
                        s_minRange = FindNextMinRange(s_maxRange)
                        s_maxRange = "zz"
                        s_refreshing = "True"
                    Else
                        s_minRange = FindNextMinRange(s_maxRange)
                        s_maxRange = s_shortURL
                        s_refreshing = "False"
                    End If

                    Call InsertAlphaRange(s_minRange, s_maxRange, s_hostMachine)

                    If i_rangeNumberOfDNSToAskFor >= i_totalDNSCount _
                    Or i_counter >= i_machineNumber Then
                        Exit While
                    End If

                End While

            End If

            'Close transaction here, everything went fine
            ' Call Sub_DeleteRecordsScratchPad()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.831--> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub


    Function FindTotalDNSFOUND(ByVal i_percVised) As Int64

        Dim s_query, mySelectQuery As String
        Dim b_nextTurnStop As Boolean
        Dim i_totalCountDNSFound As Int64

        Try

            s_query = Find_s_Query()
            If s_query <> "stop" Then
                FindTotalDNSFOUND = i_p_countDNSFoundFroms_s_Query
            Else
                FindTotalDNSFOUND = 0
            End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.862 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Dim i_p_countDNSFoundFroms_s_Query As Int64

    Function Find_s_Query() As String
        Dim s_query, s_queryToKeep As String
        Dim b_nextTurnStop As Boolean
        Dim i_totalCountDNSFound As Int64
        Dim b_goodSample As Boolean
        Dim i_countDNS, i_countDNSToKeep As Int64

        Try
            Dim mySelectQuery As String = "SELECT s_Query FROM DNAVALUE"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())

            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer

            While myReader.Read()
                s_query = myReader("s_query")
                i_countDNS = CountDNSInDNSFOUND(s_query)
                If i_countDNS > 20000 Then
                    If i_countDNS > i_countDNSToKeep Then
                        i_countDNSToKeep = i_countDNS
                        i_p_countDNSFoundFroms_s_Query = i_countDNS
                        s_queryToKeep = s_query
                        b_goodSample = True
                    End If
                End If
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection1.Close()

            If b_goodSample Then
                Find_s_Query = s_queryToKeep
            Else
                Find_s_Query = "stop"
            End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.924 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Function CountDNSInDNSFOUND(ByVal s_query) As Int64
        Dim mySelectQuery As String
        Dim b_nextTurnStop As Boolean
        Dim i_totalCountDNSFound As Int64

        Try

            mySelectQuery = "SELECT Count(ShortURL) AS TOTALCOUNTShortURL FROM DNSFOUND"

            Dim myConnection As New OleDbConnection(ConnStringURLDNS(s_query))
            Dim myCommand1 As New OleDbCommand(mySelectQuery, myConnection)
            Dim o_object As Object
            Dim s_RetrieveIDDateHour As Array

            myConnection.Open()

            Dim myReader1 As OleDbDataReader = myCommand1.ExecuteReader()
            While myReader1.Read()
                i_totalCountDNSFound = myReader1("TOTALCOUNTShortURL")
            End While

            If Not (myReader1.IsClosed) Then
                myReader1.Close()
            End If

            myConnection.Close()

            If i_totalCountDNSFound > 20000 Then
                CountDNSInDNSFOUND = i_totalCountDNSFound
            Else
                CountDNSInDNSFOUND = 0
            End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.977 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Function GetNameOfHostMachine(ByVal s_hostMachine) As String

        Try
            Dim myConnString As String
            Dim sKeepString, s_rangeMinus, s_rangePlus, s_HostName As String
            Dim i_CounthostName As Int32
            Dim date_InsertURL As DateTime
            Dim s_machineToSend As String
            Dim s_queryBefore As String
            Dim b_readyForNextTime As Boolean = False

            Dim mySelectQuery As String = "SELECT HostName FROM RANGE order by HostName"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_HostName = myReader("HostName")
                If Len(Trim(s_hostMachine)) = 0 Then
                    Exit While
                ElseIf s_HostName = s_hostMachine Then
                    b_readyForNextTime = True
                ElseIf b_readyForNextTime Then
                    Exit While
                End If

            End While

            GetNameOfHostMachine = s_HostName

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.974 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Function PutOldInsertInScratchpad() As Boolean

        Dim s_ToInsert As String = "Transaction_start_from_diskmanagement.996"
        Dim s_shortUrl As String = ""
        Dim s_no As String = "n"
        Dim myConnectionString As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_Date As String = DateAdd(DateInterval.Day, -100, Now())

        Try
            Dim myInsertQuery As String = "INSERT INTO SCRATCHPAD (Url,DateHour,Done,ShortUrl)" _
                                            & "Values ('" & Trim(s_ToInsert) & "','" & (s_Date) & "','" & (s_no) & "','" & (s_shortUrl) & "')"

            Dim myCommand As New OleDbCommand(myInsertQuery)
            myCommand.Connection = myConnection
            myConnection.Open()
            myCommand.ExecuteNonQuery()
            myCommand.Connection.Close()

            myConnection.Close()
        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.974 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Function InsertNewAlphaRange(ByVal s_hostMachine, ByVal s_minRange, ByVal s_maxRange) As Boolean

    End Function

    Function FindNextMinRange(ByVal s_maxRange) As String
        Try
            Dim s_findFirstLetter As String
            Dim s_findSecondLetter As String
            Dim i_sum As Int64
            Dim i_len As Int64 = Len(s_maxRange)

            If i_len = 2 Then
                s_findFirstLetter = Mid(s_maxRange, 1, 1)
                s_findSecondLetter = Mid(s_maxRange, 2, 1)
                i_sum = (GiveIntFromLetter(s_findFirstLetter) * 26) + GiveIntFromLetter(s_findSecondLetter)
            Else
                s_findFirstLetter = Mid(s_maxRange, 1, 1)
                i_sum = GiveIntFromLetter(s_findFirstLetter)
            End If

            If i_sum >= 676 Then
                FindNextMinRange = "zz"
            Else
                FindNextMinRange = ConvertSumToString(i_sum + 1)
            End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.1042 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Function FindNextMinRangeMoreTwo(ByVal s_maxRange) As String
        Try
            Dim s_findFirstLetter As String
            Dim s_findSecondLetter As String
            Dim i_sum As Int64
            Dim i_len As Int64 = Len(s_maxRange)

            If i_len = 2 Then
                s_findFirstLetter = Mid(s_maxRange, 1, 1)
                s_findSecondLetter = Mid(s_maxRange, 2, 1)
                i_sum = (GiveIntFromLetter(s_findFirstLetter) * 26) + GiveIntFromLetter(s_findSecondLetter)
            Else
                s_findFirstLetter = Mid(s_maxRange, 1, 1)
                i_sum = GiveIntFromLetter(s_findFirstLetter)
            End If

            If i_sum >= 676 Then
                FindNextMinRangeMoreTwo = "zz"
            Else
                FindNextMinRangeMoreTwo = ConvertSumToString(i_sum + 2)
            End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.1042 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function FindMachineNumber() As Int64
        Try
            Dim myConnString As String
            Dim sKeepString, s_rangeMinus, s_rangePlus, s_HostName As String
            Dim i_CounthostName As Int32
            Dim date_InsertURL As DateTime
            Dim s_machineToSend As String
            Dim s_queryBefore As String

            Dim mySelectQuery As String = "SELECT count(HostName) as TotalHostName FROM RANGE"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                FindMachineNumber = myReader("TotalHostName")
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.1076 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Sub DeleteFromDATADNSOlderInsert()
        Try
            Dim d_dateToday As DateTime = DateAdd(DateInterval.Day, -10, Now())
            Dim myConnection As New OleDbConnection(ConnStringDNA())

            Dim myInsertQuery As String = "Delete from DATADNS where DateOfInsert <= #" & d_dateToday & "#"

            Dim myCommand1 As New OleDbCommand(myInsertQuery)
            myCommand1.Connection = myConnection
            myConnection.Open()
            myCommand1.ExecuteNonQuery()
            myCommand1.Connection.Close()
            myConnection.Close()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.635 --> " & ex.Message, EventLogEntryType.Information, 45)
        End Try

    End Sub

    Function CheckIfRangeIsOkBeforeSending(ByVal ShortUrl) As Boolean
        Try
            Dim myConnString As String
            Dim sKeepString, s_rangeMinus, s_rangePlus, s_HostName, s_sentOver As String
            Dim i_CounthostName As Int32
            Dim date_InsertURL As DateTime
            Dim s_machineToSend As String
            Dim s_queryBefore As String

            'ShortUrl = FindFirstTwoLettersOfSomething(ShortUrl)

            Dim mySelectQuery As String = "SELECT count(DONE) as COUNTTOTAL FROM RAN_BROADCAST"

            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_sentOver = myReader("SentOver")

                If s_sentOver = "so" Then
                    CheckIfRangeIsOkBeforeSending = True
                    Exit While
                End If
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.1219 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Sub RemoveSOFromRange()
        Dim myConnectionString As String
        Dim s_no As String = "n"
        Dim s_yes As String = "y"
        Dim d_now As DateTime = Now()

        Try
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            myConnectionString = "UPDATE RANGE SET SentOver='" & (s_no) & "' where HostName <> '" & s_r_machineName & "'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myConnection.Close()
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.1258 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Function FindFirstTwoLettersOfSomething(ByVal s_stringToEvaluate) As String
        Dim s_string As String
        Try
            Dim i_dot As Integer
            Dim i_http As Integer
            Dim i_www As Integer

            If Len(s_stringToEvaluate) > 1 Then
                'i_dot = InStr(s_stringToEvaluate, ".", CompareMethod.Text)
                'i_http = InStr(s_stringToEvaluate, "http", CompareMethod.Text)
                'i_www = InStr(s_stringToEvaluate, "www", CompareMethod.Text)

                'If (i_http > 0 Or i_www > 0) And i_dot > 0 Then
                'i_dot = InStr(s_stringToEvaluate, ".", CompareMethod.Text)
                's_string = Mid(s_stringToEvaluate, i_dot + 1, 2)
                'Else
                s_string = Mid(s_stringToEvaluate, 1, 2)
                'End If

            ElseIf Len(s_stringToEvaluate) > 0 Then
                s_string = Mid(s_stringToEvaluate, 1, 1)
            Else
                s_string = s_stringToEvaluate
            End If
            FindFirstTwoLettersOfSomething = s_string
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.1331 --> " & ex.Message, EventLogEntryType.Information, 44)

        End Try

    End Function

    Public i_p_countRemoveTheSite As Int64

    Sub RemoveTheSite(ByVal s_DatabaseName)
        Try
            If i_p_countRemoveTheSite >= 25 Or i_p_countRemoveTheSite = 0 Then
                i_p_countRemoveTheSite = 1
                Dim myConnString As String
                Dim sKeepString As String
                Dim s_se As String = "se"
                Dim i_countURL As Int32 = 0
                Dim sSource As String = "AP_DENIS"
                Dim sLog As String = "Applo"
                Dim i_goodSpot As Int64 = -888888
                Dim s_verb As String = "RemoveTheSite"
                Dim s_so As String = "so"

                Dim s_shortURL As String

                Dim mySelectQuery As String = "SELECT distinct ShortURL FROM URLFound where GoodSpot = " & i_goodSpot & " AND Done <> '" & s_so & "'order by ShortURL desc"
                Dim myConnection As New OleDbConnection(ConnStringURL(s_DatabaseName))
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

                myConnection.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    s_shortURL = (myReader("ShortURL"))
                    Call BroadCastDNSFOUND(s_verb, s_DatabaseName, s_shortURL)
                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()
            Else
                i_p_countRemoveTheSite = i_p_countRemoveTheSite + 1
            End If

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.1308 --> " & ex.Message, EventLogEntryType.Information, 45)

        End Try

    End Sub

    Function HowManyPageViews() As Int64
        Try
            Dim d_KeepDuration As Double = i_p_Duration
            Dim s_yes As String = "y"
            Dim b_c As Boolean = False
            Dim s_splitQuery() As String
            Dim s_arrayToQuery() As String
            ReDim s_arrayToQuery(1)
            Dim s_heartBeat As String = ""

            Dim mySelectQuery As String = "SELECT count(URL) as TOTALCOUNT FROM SCRATCHPAD where ShortUrl <> '" & s_heartBeat & "'"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer

            While myReader.Read()
                HowManyPageViews = myReader("TOTALCOUNT")
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection1.Close()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.1356 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Sub DeleteArchive(ByVal s_databaseString)
        Try
            Dim objConn As New OleDbConnection(ConnStringURL(s_databaseString))
            Dim myConnectionString As String = "DELETE FROM URLFOUND WHERE Freshpage ='ar'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = objConn
            objConn.Open()
            myCommand2.ExecuteNonQuery()
            objConn.Close()
            Console.WriteLine(vbNewLine)
            Console.WriteLine("DeleteArchive.1445 successfull!")
            Console.WriteLine(vbNewLine)
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " DiskManagement.1356 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub CompacURLFOUND(ByVal s_DatabaseName)


    End Sub

End Module

