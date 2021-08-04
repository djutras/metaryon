Imports System
Imports System.Net
Imports System.Threading
Imports System.Text
Imports System.IO
Imports System.Object
Imports System.MarshalByRefObject
Imports System.Runtime.Remoting.Lifetime
Imports System.Data
Imports System.Data.OleDb

Public Module Async
    Public Class LSClass
        Inherits MarshalByRefObject

        Public Overrides Function InitializeLifetimeService() As Object
            Dim lease As ILease = CType(MyBase.InitializeLifetimeService(), ILease)

            'MsgBox(" lease.CurrentState.13 " & lease.CurrentState.ToString)
            'MsgBox(" LeaseState.Initial.13 " & LeaseState.Initial.ToString)

            If lease.CurrentState = LeaseState.Initial Then
                lease.InitialLeaseTime = TimeSpan.FromSeconds(5)
                lease.SponsorshipTimeout = TimeSpan.FromSeconds(0)
                lease.RenewOnCallTime = TimeSpan.FromSeconds(0)
            End If
            Return lease
        End Function
    End Class


    Public s_p_RetrievedHtml As String


    Public Sub f_GetHtmlFromURI()

        s_p_RetrievedHtml = ""
        Dim s_UrlToFind As String = s_p_PassUri
        Dim s_databaseName As String = s_p_sDatabaseName
        Dim i_TimeToJoin As Int32
        Dim b_WhileGetHtml As Boolean = True
        Dim i_CountGetHtml As Integer = 0
        Dim i_lenghtOfThePage As Int32
        Dim b_block_404 As Boolean = False
        Dim s_string As String

        If s_UrlToFind = "stop" Then
            s_p_RetrievedHtml = "stop"
            Exit Sub
        End If

        Dim i_integer As Integer = InStr(s_UrlToFind, "https://", CompareMethod.Text)

        If InStr(s_UrlToFind, "http://", CompareMethod.Text) <> 1 And InStr(s_UrlToFind, "https://", CompareMethod.Text) <> 1 Then
            s_UrlToFind = "http://" & Trim(s_UrlToFind)
        End If

        i_TimeToJoin = 5000

        Try
            Dim objURI As Uri = New Uri(s_UrlToFind)
            Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
            Call objWebRequest.InitializeLifetimeService()
            'objWebRequest.Timeout = i_TimeToJoin

            Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
            objWebResponse.Headers.Add("Accept-Language: fr, en")

            Call objWebResponse.InitializeLifetimeService()

            objWebResponse.InitializeLifetimeService()

            Dim objStream As Stream = objWebResponse.GetResponseStream()
            Call objStream.InitializeLifetimeService()

            Dim objStreamReader As StreamReader = New StreamReader(objStream)
            Call objStreamReader.InitializeLifetimeService()

            Dim strHTML As String = objStreamReader.ReadToEnd
            Dim s_GotHtml As String

            s_GotHtml = Mid(strHTML, 1, i_p_MaxKPage * 1000)
            i_lenghtOfThePage = Len(s_GotHtml)

            Call f_GetHtmlFromURIFU(s_GotHtml)

            If (InStr(s_GotHtml, "404 not found", CompareMethod.Text) > 0) Then
                s_GotHtml = s_GotHtml
                b_block_404 = True
            End If

            If (InStr(s_GotHtml, LCase("Sorry, we can’t find"), CompareMethod.Text) > 0) Then
                s_GotHtml = s_GotHtml
                b_block_404 = True
            End If

            If (InStr(s_GotHtml, "404-not-found", CompareMethod.Text) > 0) Then
                s_GotHtml = s_GotHtml
                b_block_404 = True
            End If

            If (InStr(s_GotHtml, LCase("You don't have permission"), CompareMethod.Text) > 0) Then
                s_GotHtml = s_GotHtml
                b_block_404 = True
            End If

            If (InStr(s_GotHtml, LCase("Page not found"), CompareMethod.Text) > 0) Then
                s_GotHtml = s_GotHtml
                b_block_404 = True
            End If

            If (InStr(s_GotHtml, LCase("404 Page"), CompareMethod.Text) > 0) Then
                s_GotHtml = s_GotHtml
                b_block_404 = True
            End If

            If (InStr(s_GotHtml, LCase("PAGE CANNOT BE FOUND"), CompareMethod.Text) > 0) Then
                s_GotHtml = s_GotHtml
                b_block_404 = True
            End If

            If (i_lenghtOfThePage <= i_p_lenghtMinimumFile) Then
                s_GotHtml = s_GotHtml
                b_block_404 = True
            End If

            Form1.TBLenghtPage.Text = i_lenghtOfThePage.ToString
            Form1.TBLenghtPage.Update()

            Form1.Refresh()

            'Console.WriteLine("Async.81 -- i_lenghtOfThePage -- " & i_lenghtOfThePage)

            If i_lenghtOfThePage >= (i_p_MinKPage * 1000) Then
                s_p_RetrievedHtml = s_GotHtml
            Else
                s_p_RetrievedHtml = "stop"
            End If

        Catch e_async_80 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"

            'Call RemovePointBecauseOf404(s_UrlToFind, s_databaseName)
            s_string = e_async_80.ToString()

            If InStr(s_string, "404", CompareMethod.Text) = 0 _
            And InStr(s_string, "The remote server returned an error", CompareMethod.Text) = 0 _
            And InStr(s_string, "Invalid URI", CompareMethod.Text) = 0 _
            And InStr(s_string, "400", CompareMethod.Text) = 0 _
            And InStr(s_string, "403", CompareMethod.Text) = 0 _
            And InStr(s_string, "The underlying connection was closed:", CompareMethod.Text) = 0 _
            And InStr(s_string, "Thread was being aborted.", CompareMethod.Text) = 0 _
            And InStr(s_string, "500", CompareMethod.Text) = 0 Then

                EventLog.WriteEntry(sSource, " e_async_80 --> " & e_async_80.Message, EventLogEntryType.Information, 44)

            End If

        Finally

            If (b_block_404 Or Len(s_string) > 0) Then

                Block404ThatURLInDNSFOUND(s_UrlToFind, s_p_sDatabaseName, i_lenghtOfThePage, s_string)
                Block404ThatURLInURLFOUND(s_UrlToFind, s_p_sDatabaseName, i_lenghtOfThePage, s_string)

            End If

        End Try

    End Sub


    Sub Block404ThatURLInDNSFOUND(ByVal s_URLSite, ByVal sDatabaseName, ByVal i_lenghtOfThePage, ByVal s_error)
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try
            'Dim s_shortUrl As String = Func_CleanUrl(s_urlToCount)

            If Len(s_URLSite) > 0 And Len(sDatabaseName) > 0 Then

                Form1.TB40004.Text = s_URLSite.ToString
                Form1.TextBox40004.Text = "-40004"
                Form1.TBLenghtPage.Text = i_lenghtOfThePage.ToString

                Form1.TB40004.Update()
                Form1.TextBox40004.Update()
                Form1.TBLenghtPage.Update()
                Form1.Refresh()

                Dim s_Yes As String = "y"

                Dim objConn As New OleDbConnection(ConnStringURLDNS(sDatabaseName))

                Dim myConnectionString As String =
                "UPDATE DNSFOUND SET Done='" & (s_Yes) & "', iGoodSpot=-40004, i_lenghtPage=" & i_lenghtOfThePage & ",s_error='" & s_error & "' where URLSite ='" & (s_URLSite) & "' and sCurrentQuery='" & (sDatabaseName) & "'"

                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = objConn
                objConn.Open()
                myCommand2.ExecuteNonQuery()
                'objconn = Nothing
                objConn.Close()

            Else

                'EventLog.WriteEntry(sSource, " lookforfiles.e_module1_600 s_urlToCount--> " & s_urlToCount & "  sDatabaseName--> " & sDatabaseName & "  s_shortUrl--> " & s_shortUrl, EventLogEntryType.Information, 49)

            End If

            'Console.WriteLine("Insert 404 Block404ThatURLInDNSFOUND ---> " & s_URLSite)

        Catch e_module1_600 As Exception


            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_600 --> " & e_module1_600.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)

        End Try

    End Sub

    Sub Block404ThatURLInURLFOUND(ByVal s_URLSite, ByVal sDatabaseName, ByVal i_lenghtOfThePage, ByVal s_string)
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try
            'Dim s_shortUrl As String = Func_CleanUrl(s_urlToCount)

            If Len(s_URLSite) > 0 And Len(sDatabaseName) > 0 Then

                Form1.TB40004.Text = s_URLSite.ToString
                Form1.TextBox40004.Text = "-40004"

                Form1.TB40004.Update()
                Form1.TextBox40004.Update()
                Form1.Refresh()


                Dim s_Yes As String = "y"

                Dim objConn As New OleDbConnection(ConnStringURLDNS(sDatabaseName))

                Dim myConnectionString As String =
                "UPDATE URLFOUND SET Done='" & (s_Yes) & "', GoodSpot=-40004, i_lenghtPage=" & i_lenghtOfThePage & ", s_error='" & s_string & "' where URL ='" & (s_URLSite) & "' and sQueryUrlFound='" & (sDatabaseName) & "'"

                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = objConn
                objConn.Open()
                myCommand2.ExecuteNonQuery()
                'objconn = Nothing
                objConn.Close()

                'Console.WriteLine("Insert 404 Block404ThatURLInURLFOUND ---> " & s_URLSite)

            Else

                'EventLog.WriteEntry(sSource, " lookforfiles.e_module1_600 s_urlToCount--> " & s_urlToCount & "  sDatabaseName--> " & sDatabaseName & "  s_shortUrl--> " & s_shortUrl, EventLogEntryType.Information, 49)

            End If
        Catch e_module1_600 As Exception


            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_600 --> " & e_module1_600.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)

        End Try

    End Sub

    Public Class RequestState

        Public RequestData As New StringBuilder("")
        Public BufferRead(1024) As Byte
        Public Request As HttpWebRequest
        Public ResponseStream As Stream
        ' Create Decoder for appropriate encoding type
        Public StreamDecode As Decoder = Encoding.UTF8.GetDecoder()

        Public Sub New()
            Request = Nothing
            ResponseStream = Nothing
        End Sub
    End Class

    ' ClientGetAsync issues the async request
    'Public Class ClientGetAsync
    Public allDone As New ManualResetEvent(False)
    Const BUFFER_SIZE As Integer = 1024

    Public Function Main(ByVal sURI)
        'Dim Args As String() = sURI
        'MsgBox("async.sURI.272 - >" & sURI)
        ' Get the URI from the command line
        Dim HttpSite As Uri = New Uri(sURI)

        ' Create the request object
        Dim wreq As HttpWebRequest = CType(WebRequest.Create(HttpSite), HttpWebRequest)

        ' Create the state object
        Dim rs As RequestState = New RequestState()

        ' Add the request into the state so it can be passed around
        rs.Request = wreq

        ' Issue the async request
        Dim r As IAsyncResult = CType(wreq.BeginGetResponse(New AsyncCallback(AddressOf RespCallback), rs), IAsyncResult)

        ' Set the ManualResetEvent to Wait so that the app
        ' doesn't exit until after the callback is called
        allDone.WaitOne()
    End Function


    Sub RespCallback(ByVal ar As IAsyncResult)
        ' Get the RequestState object from the async result
        Dim rs As RequestState = CType(ar.AsyncState, RequestState)

        ' Get the HttpWebRequest from RequestState
        Dim req As HttpWebRequest = rs.Request

        ' Calling EndGetResponse produces the HttpWebResponse object
        ' which came from the request issued above
        Dim resp As HttpWebResponse = CType(req.EndGetResponse(ar), HttpWebResponse)

        ' Now that we have the response, it is time to start reading
        ' data from the response stream
        Dim ResponseStream As Stream = resp.GetResponseStream()

        ' The read is also done using async so we'll want
        ' to store the stream in RequestState
        rs.ResponseStream = ResponseStream

        ' Note that rs.BufferRead is passed in to BeginRead.  This is
        ' where the data will be read into.
        'Dim iarRead As IAsyncResult = ResponseStream.BeginRead(rs.BufferRead, 0, BUFFER_SIZE, New AsyncCallback(AddressOf ReadCallBack), rs)
        Dim iarRead As IAsyncResult = ResponseStream.BeginRead(rs.BufferRead, 0, BUFFER_SIZE, New AsyncCallback(AddressOf RespCallback), rs)

    End Sub

    Public s_p_GetHtml As String

    Public Sub ReadCallBack(ByVal asyncResult As IAsyncResult)
        ' Get the RequestState object from the asyncresult
        Dim rs As RequestState = CType(asyncResult.AsyncState, RequestState)

        ' Pull out the ResponseStream that was set in RespCallback
        Dim responseStream As Stream = rs.ResponseStream

        ' At this point rs.BufferRead should have some data in it.
        ' Read will tell us if there is any data there
        Dim read As Integer = responseStream.EndRead(asyncResult)
        If read > 0 Then
            ' Prepare Char array buffer for converting to Unicode
            Dim charBuffer(1024) As Char

            ' Convert byte stream to Char array and then String
            ' len shows how many characters are converted to Unicode
            Dim len As Integer = rs.StreamDecode.GetChars(rs.BufferRead, 0, read, charBuffer, 0)
            Dim str As String = New String(charBuffer, 0, len)

            ' Append the recently read data to the RequestData stringbuilder object
            ' contained in RequestState
            rs.RequestData.Append(str)

            ' Now fire off another async call to read some more data
            ' Note that this will continue to get called until
            ' responseStream.EndRead returns -1
            Dim ar As IAsyncResult = responseStream.BeginRead(rs.BufferRead, 0, BUFFER_SIZE, New AsyncCallback(AddressOf RespCallback), rs)
        Else
            If rs.RequestData.Length > 1 Then
                ' All of the data has been read, so display it to the console
                Dim strContent As String
                s_p_GetHtml = rs.RequestData.ToString()

            End If

            ' Close down the response stream
            responseStream.Close()

            ' Set the ManualResetEvent so the main thread can exit
            allDone.Set()
        End If


    End Sub
    'End Class


    Sub RemovePointBecauseOf404(ByVal s_urlToFind, ByVal s_databaseName)

        Dim SUrl As String = "n"
        Dim i_ID As Integer
        Dim mySelectQuery As String
        Dim i_GoodSpot As Int32 = -2000
        Dim s_KeepDate As String = Now()
        Dim d_dateHour As Date
        Dim d_dateRatingDone As Date
        Dim b_go As Boolean = True
        Dim s_done As String
        Dim b_isThere As Boolean = False

        Try

            mySelectQuery = "SELECT iGoodSpot FROM DNSFOUND where iGoodSpot<>-40004 and URLSite = '" & s_urlToFind & "' and sCurrentQuery = '" & s_databaseName & "'"

            Dim myConnection As New OleDbConnection(ConnStringURLDNS(s_databaseName))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            Dim o_object As Object
            Dim s_RetrieveIDDateHour As Array

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            While myReader.Read()

                i_GoodSpot = myReader("iGoodSpot")
                b_isThere = True
                Exit While

            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            'myConnection = Nothing
            myConnection.Close()

            'If i_GoodSpot > -1000 Then

            If b_isThere Then

                i_GoodSpot = i_GoodSpot - 2
                Dim myConnectionString As String
                Dim objConn As New OleDbConnection(ConnStringURLDNS(s_databaseName))
                myConnectionString = "UPDATE DNSFOUND SET iGoodSpot=" & (i_GoodSpot) & " where URLSite = '" & (s_urlToFind) & "' and sCurrentQuery='" & s_databaseName & "'"
                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = objConn
                objConn.Open()
                myCommand2.ExecuteNonQuery()
                'objconn = Nothing
                objConn.Close()
                'End If

                Call RemovePointBecauseOf404FU(myConnectionString)

            End If

        Catch e_async_272 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_272 --> " & e_async_272.Message, EventLogEntryType.Information, 44)

        End Try
    End Sub

    Public i_p_countsaInRange As Int32
    Public i_p_countSendNewMachineConfig As Int32
    Public i_p_countRecursive As Int64

    Sub sendNewMachineConfiguration()

        Dim s_Noo As String = "no"
        Dim s_no As String = "n"
        Dim s_KeepDate As String = Now()
        Dim d_dateHour As Date
        Dim s_done As String
        Dim s_hostName, s_RangeMinus, s_RangePlus, s_CallReceived, s_Active As String

        Try

            If i_p_countSendNewMachineConfig >= 1 Or i_p_countSendNewMachineConfig = 0 Then

                Dim s As String
                Dim s_so As String = "so"
                Dim s_verb As String = "newMachineConfig"
                Dim s_IPAddress As String = "nothing"
                Dim mySelectQuery As String
                Dim i_countHostName As Int32

                'Console.WriteLine(s_hostName & " --- search.959")

                mySelectQuery = "SELECT count(HOSTNAME) as CountHOSTNAME FROM RANGE where (Authority ='" & s_no & "' or Authority ='" & s_Noo & "')"

                Dim myConnection As New OleDbConnection(ConnStringDNA())
                Dim myCommand1 As New OleDbCommand(mySelectQuery, myConnection)
                myConnection.Open()
                Dim myReader1 As OleDbDataReader = myCommand1.ExecuteReader()
                While myReader1.Read()
                    i_countHostName = myReader1("CountHOSTNAME")
                End While
                If Not (myReader1.IsClosed) Then
                    myReader1.Close()
                End If
                myConnection.Close()

                If i_countHostName > 0 Then
                    mySelectQuery = ""
                    mySelectQuery = "SELECT * FROM RANGE"
                    Dim myConnection1 As New OleDbConnection(ConnStringDNA())
                    Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

                    myConnection1.Open()

                    Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                    Dim Index As Integer

                    While myReader.Read()

                        s_hostName = myReader("HostName")
                        If Len(myReader("IPAddress")) > 1 Then
                            s_IPAddress = myReader("IPAddress")
                        Else
                            s_IPAddress = "nothing"
                        End If
                        s_RangeMinus = myReader("RangeMinus")
                        s_RangePlus = myReader("RangePlus")
                        s_CallReceived = myReader("CallReceived")
                        s_Active = myReader("Active")

                        'If i_p_countRecursive = 0 Then
                        Call InsertInRAN_BROADCAST(s_hostName)
                        'Else
                        '   Dim se As String
                        '   s = se
                        'End If

                        s = s_hostName &
                        "|||||" & s_IPAddress &
                        "|||||" & s_RangeMinus &
                        "|||||" & s_RangePlus &
                        "|||||" & s_CallReceived &
                        "|||||" & s_Active

                        If s_hostName = s_r_machineName Then
                            BroadCastDNSFOUND(s_verb, s)
                            Call UpdateDNSFOUNDWithNewRange(s_RangeMinus, s_RangePlus)
                        Else
                            BroadCastDNSFOUND(s_verb, s)
                        End If

                    End While

                    If Not (myReader.IsClosed) Then
                        myReader.Close()
                    End If

                    myConnection1.Close()
                Else
                    i_p_countRecursive = 0
                    b_p_doTheBigJobToSendNewConfiguration = False
                    Exit Sub
                End If

                i_p_countSendNewMachineConfig = 1

                'If i_p_countRecursive > 10 Then
                '    i_p_countRecursive = 0
                '    Exit Sub
                'Else
                '    i_p_countRecursive = i_p_countRecursive + 1
                '    System.Threading.Thread.Sleep(1000)
                '    Call sendNewMachineConfiguration()
                'End If

            Else
                i_p_countSendNewMachineConfig = i_p_countSendNewMachineConfig + 1
            End If

            If Len(s_hostName) > 0 Then
                Call InsertsaInRange(s_hostName)
                Call InsertCallSend(s_hostName, s_CallReceived)
            Else
                If i_p_countsaInRange >= 50 Then
                    Call InsertnoInRange(s_hostName)
                    i_p_countsaInRange = 0
                End If
                i_p_countsaInRange = i_p_countsaInRange + 1
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " async.397 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Sub InsertCallSend(ByVal s_hostName, ByVal s_CallReceived)
        Dim myConnectionString, s As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())

        Try
            If IsNumeric(s_CallReceived) Then
                s_CallReceived = CInt(s_CallReceived)
                s_CallReceived = s_CallReceived + 1
            Else
                s_CallReceived = 0
            End If

            myConnectionString = "UPDATE RANGE SET CallReceived=" & (s_CallReceived) & " where HostName = '" & Trim(s_hostName) & "'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch e_Async_378 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_AddMachine_385 --> " & e_Async_378.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Sub InsertNewMachinConfigFromBroadcast(ByVal s_hostName, ByVal s_IPAddress, ByVal s_RangeMinus, ByVal s_RangePlus, ByVal s_CallReceived, ByVal s_Active)

        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_verb As String = "acknowNewMachinConfig"
        Dim b_continue As Boolean = True

        Try
            Dim myConnectionString, s, s_queryName As String

            If s_p_authority <> "yes" Then
                If CheckHostNameIn(s_hostName) <> "stop" Then
                    If Len(s_hostName) > 0 Then

                        Dim myInsertQuery As String = "INSERT INTO RANGE (HostName,IPAddress,RangeMinus,RangePlus,Active)" _
                         & "Values ('" & Trim(s_hostName) & "','" & (s_IPAddress) & "','" & (s_RangeMinus) & "','" & (s_RangePlus) & "','" & (s_Active) & "')"
                        Dim myCommand As New OleDbCommand(myInsertQuery)
                        myCommand.Connection = myConnection
                        myConnection.Open()
                        myCommand.ExecuteNonQuery()
                        myCommand.Connection.Close()
                        myConnection.Close()

                    End If
                Else

                    myConnectionString = "UPDATE RANGE SET HostName='" & (s_hostName) & "', IPAddress='" & (s_IPAddress) & "', RangeMinus='" & (s_RangeMinus) & "', RangePlus='" & (s_RangePlus) & "', CallReceived='" & (s_CallReceived) & "', Active='" & (s_Active) & "'  where HostName = '" & Trim(s_hostName) & "'"
                    Dim myCommand2 As New OleDbCommand(myConnectionString)
                    myCommand2.Connection = myConnection
                    myConnection.Open()
                    myCommand2.ExecuteNonQuery()
                    myCommand2.Connection.Close()
                    myConnection.Close()
                End If
                If s_r_machineName = s_hostName Then
                    'If RefreshDNSFOUND(s_hostName, s_RangeMinus, s_RangePlus) Then
                    Call UpdateDNSFOUNDWithNewRange(s_RangeMinus, s_RangePlus)
                    'End If
                End If
            End If

            'If s_r_machineName = s_hostName Then
            'Call CheckIfStartSearchCycleAgain(s_RangeMinus, s_RangePlus)
            'End If

            If s_p_authority <> "yes" Then
                s = s_r_machineName & "|||||" & s_hostName
                Call BroadCastDNSFOUND(s_verb, s)
            End If

        Catch e_Async_415 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_AddMachine_385 --> " & e_Async_415.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub


    Function CheckHostNameIn(ByVal s_machineName As String) As String
        Try
            Dim s_keepHostName As String
            Dim b_nameIsIn As Boolean = False
            If Len(s_machineName) > 0 Then

                Dim mySelectQuery, s_Website, s As String

                mySelectQuery = "SELECT HostName FROM Range where HostName ='" & s_machineName & "'"

                Dim myConnection1 As New OleDbConnection(ConnStringDNA())
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

                myConnection1.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                Dim Index As Integer

                While myReader.Read()
                    s_keepHostName = myReader("HostName").ToString
                    Exit While
                End While
                If Len(s_keepHostName) = 0 Then
                    CheckHostNameIn = s_keepHostName
                Else
                    CheckHostNameIn = "stop"
                End If
                myReader.Close()
                myConnection1.Close()

            End If

        Catch e_Async_457 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_457 --> " & e_Async_457.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Sub InsertAckNewConfigInAuthority(ByVal s_machineFrom, ByVal s_hostName)
        Dim myConnectionString, s As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_verb As String = "acknowNewMachinConfig"
        Dim b_continue As Boolean = True
        Dim s_so As String = "so"

        Try
            myConnectionString = "UPDATE RAN_BROADCAST SET DONE ='" & (s_so) & "'  where HOSTNAME = '" & Trim(s_machineFrom) & "' AND MACHINECHANGE = '" & Trim(s_hostName) & "'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

            Call CheckMachineToUpdateRange(s_hostName)

        Catch e_Async_484 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_484 --> " & e_Async_484.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub CheckMachineToUpdateRange(ByVal s_hostName)
        'Dim mySelectQuery, s, s_keepHostName As String
        'Dim myConnection As New OleDbConnection(ConnStringDNA())
        'Dim s_verb As String = "acknowNewMachinConfig"
        'Dim b_continue As Boolean = True
        'Dim s_so As String = "so"
        'Dim i_totalCount As Int64

        Try
            '    If s_hostName <> s_r_machineName Then
            '        mySelectQuery = "SELECT Count(MACHINECHANGE) as TOTALCOUNT FROM RAN_BROADCAST WHERE DONE <> '" & (s_so) & "'  AND HOSTNAME = '" & Trim(s_hostName) & "'"
            '        Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            '        myConnection.Open()

            '        Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            '        While myReader.Read()
            '            i_totalCount = myReader("TOTALCOUNT")
            '        End While

            '        If i_totalCount > 0 Then
            '            b_continue = False
            '        Else
            '            b_continue = True
            '        End If
            '        myReader.Close()

            '        If b_continue Then
            '            mySelectQuery = ""
            '            mySelectQuery = "UPDATE RANGE SET Authority = '" & (s_so) & "', SentOver = '" & (s_so) & "' WHERE HostName = '" & Trim(s_hostName) & "'"
            '            Dim myCommand2 As New OleDbCommand(mySelectQuery)
            '            myCommand2.Connection = myConnection
            '            myCommand2.ExecuteNonQuery()
            '            myCommand2.Connection.Close()
            '            myConnection.Close()
            '        End If
            '    End If

            Call CheckMachineToUpdateRange()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " async.611 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub


    Sub CheckMachineToUpdateRange()
        Dim mySelectQuery, s, s_keepHostName As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_verb As String = "acknowNewMachinConfig"
        Dim b_continue As Boolean = True
        Dim s_so As String = "so"
        Dim i_count As Int64

        Try

            mySelectQuery = "SELECT Count(MACHINECHANGE) as TOTALCOUNT FROM RAN_BROADCAST WHERE DONE <> '" & (s_so) & "'"
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                i_count = myReader("TOTALCOUNT")
            End While

            If i_count > 0 Then
                b_continue = False
            Else
                b_continue = True
            End If

            myReader.Close()

            If b_continue Then

                mySelectQuery = ""
                mySelectQuery = "UPDATE RANGE SET Authority='" & (s_so) & "',SentOver = '" & (s_so) & "'"
                Dim myCommand2 As New OleDbCommand(mySelectQuery)
                myCommand2.Connection = myConnection
                myCommand2.ExecuteNonQuery()
                myCommand2.Connection.Close()
                myConnection.Close()

            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " async.671 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub


    Sub InsertsaInRange(ByVal s_hostName)
        Dim mySelectQuery, s, s_keepHostName As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_sa As String = "sa"

        Try
            mySelectQuery = ""
            mySelectQuery = "UPDATE RANGE SET Authority='" & (s_sa) & "' WHERE HostName = '" & Trim(s_hostName) & "'"
            Dim myCommand2 As New OleDbCommand(mySelectQuery)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch e_Async_562 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_562 --> " & e_Async_562.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Sub InsertnoInRange(ByVal s_hostName)
        Dim mySelectQuery, s, s_keepHostName As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_no As String = "no"
        Dim s_sa As String = "sa"
        Try

            mySelectQuery = ""
            mySelectQuery = "UPDATE RANGE SET Authority='" & (s_no) & "' WHERE Authority ='" & s_sa & "'"
            Dim myCommand2 As New OleDbCommand(mySelectQuery)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch e_Async_562 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_562 --> " & e_Async_562.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub AcknowMachineNameReceived(ByVal s_machineFrom)

        Dim myConnectionString, s As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_so As String = "so"

        Try
            myConnectionString = "UPDATE RANGE SET SentOver='" & (s_so) & "' where HostName = '" & Trim(s_machineFrom) & "'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch e_Async_624 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_624 --> " & e_Async_624.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Public i_countAlphaRange As Int32

    Function DistributeAlphaRange()

        Dim i_totalCount, i_countOfZeroInRange As Int32

        Try
            If i_countAlphaRange = 0 Or i_countAlphaRange >= 50 Then
                i_countOfZeroInRange = TotalZeroCount()

                If i_countOfZeroInRange > 0 Then
                    i_totalCount = TotalCountFromRange()
                    DistributeRange(i_totalCount)
                End If

                i_countAlphaRange = 0

            Else
                i_countAlphaRange = i_countAlphaRange + 1
            End If

        Catch e_Async_766 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_766 --> " & e_Async_766.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Function TotalCountFromRange() As Int32

        Dim myConnectionString, s As String
        Dim s_so As String = "so"
        Dim SUrl As String = "no"
        Dim s_no As String = "n"
        Dim s_KeepDate As String = Now()
        Dim d_dateHour As Date
        Dim s_done As String
        Dim i_totalCount As Int32

        Try
            Dim mySelectQuery As String = "SELECT count(*) as TOTALCOUNT FROM RANGE"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer

            While myReader.Read()
                If IsNumeric(myReader("TOTALCOUNT")) Then
                    i_totalCount = myReader("TOTALCOUNT")
                Else
                    i_totalCount = 0
                End If
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection1.Close()

            TotalCountFromRange = i_totalCount

        Catch e_Async_719 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_719 --> " & e_Async_719.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function TotalZeroCount() As Int32
        Dim myConnectionString, s As String
        Dim s_so As String = "so"
        Dim SUrl As String = "no"
        Dim s_no As String = "n"
        Dim s_KeepDate As String = Now()
        Dim d_dateHour As Date
        Dim s_done As String
        Dim i_totalCount As Int32
        Dim s_zero As Integer = "0"

        Try
            Dim mySelectQuery As String = "SELECT count(*) as TOTALCOUNT FROM RANGE WHERE RANGEMINUS = '" & s_zero & "' and RANGEPLUS= '" & s_zero & "' "
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer

            While myReader.Read()
                If IsNumeric(myReader("TOTALCOUNT")) Then
                    i_totalCount = myReader("TOTALCOUNT")
                Else
                    i_totalCount = 0
                End If
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection1.Close()

            TotalZeroCount = i_totalCount

        Catch e_Async_741 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_741 --> " & e_Async_741.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Sub DistributeRange(ByVal i_totalCount As Object)
        Try
            Dim i_range As Double
            Dim i_rangeInt As Int32
            Dim i_sum As Int32
            Dim s_rangeMinus, s_rangePlus, s_stringRange As String

            Call RemoveAlphaRange()

            i_range = 676 / i_totalCount
            i_rangeInt = CInt(i_range)

            While i_sum <= 676

                If i_sum > 0 Then
                    i_sum = i_sum + 1
                End If
                If i_sum > 676 Then
                    Exit While
                End If
                s_rangeMinus = ConvertSumToString(i_sum)
                i_sum = i_sum + i_rangeInt
                If i_sum > 676 Then
                    i_sum = 676
                End If

                s_rangePlus = ConvertSumToString(i_sum)

                Call InsertAlphaRange(s_rangeMinus, s_rangePlus)

            End While

        Catch e_Async_767 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_767 --> " & e_Async_767.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Function ConvertSumToString(ByVal i_sum As Int32) As String
        Try
            Dim d_range As Double
            Dim i_howManyLetterDiz, i_howManyLetterUni As Int32

            If i_sum = 0 Then
                ConvertSumToString = CStr(i_sum)
            Else
                d_range = (i_sum) / 26
                If d_range > 1 Then
                    i_howManyLetterDiz = Fix(d_range)
                    If i_howManyLetterDiz = 26 Then
                        ConvertSumToString = "zz"
                        Exit Function
                    Else
                        d_range = d_range - CDbl(i_howManyLetterDiz)
                    End If
                    i_howManyLetterUni = CInt(d_range * 26)
                    ConvertSumToString = GiveLetterFromInt(i_howManyLetterDiz - 1) & GiveLetterFromInt(i_howManyLetterUni - 1)
                Else
                    i_howManyLetterUni = CInt(d_range * 26)
                    ConvertSumToString = GiveLetterFromInt(i_howManyLetterUni - 1)
                End If

            End If
        Catch e_Async_799 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_799 --> " & e_Async_799.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function GiveLetterFromInt(ByVal i_letter As Int32) As String
        Try
            If i_letter = -1 Then
                i_letter = 0
            End If

            Dim s_array() As String
            ReDim s_array(25)

            s_array(0) = "a"
            s_array(1) = "b"
            s_array(2) = "c"
            s_array(3) = "d"
            s_array(4) = "e"
            s_array(5) = "f"
            s_array(6) = "g"
            s_array(7) = "h"
            s_array(8) = "i"
            s_array(9) = "j"
            s_array(10) = "k"
            s_array(11) = "l"
            s_array(12) = "m"
            s_array(13) = "n"
            s_array(14) = "o"
            s_array(15) = "p"
            s_array(16) = "q"
            s_array(17) = "r"
            s_array(18) = "s"
            s_array(19) = "t"
            s_array(20) = "u"
            s_array(21) = "v"
            s_array(22) = "w"
            s_array(23) = "x"
            s_array(24) = "y"
            s_array(25) = "z"

            GiveLetterFromInt = s_array(i_letter)
        Catch e_Async_843 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_843 --> " & e_Async_843.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function GiveIntFromLetter(ByVal s_letter As String) As Int64
        Try
            Select Case s_letter
                Case "a"
                    GiveIntFromLetter = 1
                Case "b"
                    GiveIntFromLetter = 2
                Case "c"
                    GiveIntFromLetter = 3
                Case "d"
                    GiveIntFromLetter = 4
                Case "e"
                    GiveIntFromLetter = 5
                Case "f"
                    GiveIntFromLetter = 6
                Case "g"
                    GiveIntFromLetter = 7
                Case "h"
                    GiveIntFromLetter = 8
                Case "i"
                    GiveIntFromLetter = 9
                Case "j"
                    GiveIntFromLetter = 10
                Case "k"
                    GiveIntFromLetter = 11
                Case "l"
                    GiveIntFromLetter = 12
                Case "m"
                    GiveIntFromLetter = 13
                Case "n"
                    GiveIntFromLetter = 14
                Case "o"
                    GiveIntFromLetter = 15
                Case "p"
                    GiveIntFromLetter = 16
                Case "q"
                    GiveIntFromLetter = 17
                Case "r"
                    GiveIntFromLetter = 18
                Case "s"
                    GiveIntFromLetter = 19
                Case "t"
                    GiveIntFromLetter = 20
                Case "u"
                    GiveIntFromLetter = 21
                Case "v"
                    GiveIntFromLetter = 22
                Case "w"
                    GiveIntFromLetter = 23
                Case "x"
                    GiveIntFromLetter = 24
                Case "y"
                    GiveIntFromLetter = 25
                Case "z"
                    GiveIntFromLetter = 26
            End Select

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1007 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Sub RemoveAlphaRange()
        Dim myConnectionString, s As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_zero As String = "0"

        Try
            myConnectionString = "UPDATE RANGE SET RangeMinus='" & (s_zero) & "',RangePlus='" & (s_zero) & "' "
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch e_Async_1055 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1055 --> " & e_Async_1055.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub


    Function InsertAlphaRange(ByVal s_rangeMinus, ByVal s_rangePlus) As Boolean
        Dim myConnectionString, s As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_no As String = "n"
        Dim s_zero As String = "0"

        Try
            Dim b_changeDNSFOUND As Boolean = False
            Dim s_hostName As String
            Dim mySelectQuery As String = "SELECT TOP 1 HOSTNAME FROM RANGE WHERE RangeMinus = '" & s_zero & "' and RangePlus = '" & s_zero & "' ORDER BY HOSTNAME"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer
            While myReader.Read()
                s_hostName = myReader("HostName")
                'Console.WriteLine(s_hostName & " --- Aync.870")
            End While
            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection1.Close()

            If s_p_authority = "yes" And s_r_machineName = s_hostName Then
                If RefreshDNSFOUND(s_r_machineName, s_rangeMinus, s_rangePlus) Then
                    b_changeDNSFOUND = True
                End If
            End If

            myConnectionString = "UPDATE RANGE SET RangeMinus='" & (s_rangeMinus) & "',RangePlus='" & (s_rangePlus) & "',Authority='" & (s_no) & "' where HostName = '" & s_hostName & "'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

            If b_changeDNSFOUND Then
                Call UpdateDNSFOUNDWithNewRange(s_rangeMinus, s_rangePlus)
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1113 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Public b_p_doTheBigJobToSendNewConfiguration As Boolean = False

    Function InsertAlphaRange(ByVal s_rangeMinus, ByVal s_rangePlus, ByVal s_hostMachine) As Boolean
        Dim myConnectionString, s As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_no As String = "n"
        Dim s_so As String = "so"
        Dim s_zero As String = "0"

        Try

            b_p_doTheBigJobToSendNewConfiguration = True

            If s_p_authority = "yes" And s_r_machineName = s_hostMachine Then
                myConnectionString = "UPDATE RANGE SET RangeMinus='" & (s_rangeMinus) & "',RangePlus='" & (s_rangePlus) & "',Authority='" & (s_no) & "', SentOver = Authority='" & (s_so) & "' where HostName = '" & s_hostMachine & "'"
            Else
                myConnectionString = "UPDATE RANGE SET RangeMinus='" & (s_rangeMinus) & "',RangePlus='" & (s_rangePlus) & "',Authority='" & (s_no) & "', SentOver = Authority='" & (s_no) & "' where HostName = '" & s_hostMachine & "'"
            End If
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

            If s_p_authority = "yes" Then
                If s_r_machineName = s_hostMachine Then
                    'If RefreshDNSFOUND(s_r_machineName, s_rangeMinus, s_rangePlus) Then
                    Call UpdateDNSFOUNDWithNewRange(s_rangeMinus, s_rangePlus)
                    'End If
                End If
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1105 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Function InsertInRAN_BROADCAST(ByVal string_HOSTNAME As String) As Boolean
        Dim myConnectionString, s_hostName As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_no As String = "n"
        Dim s_zero As String = "0"
        Dim b_removeSOFromRange As Boolean

        Try

            Dim mySelectQuery As String = "SELECT HOSTNAME FROM RANGE WHERE HOSTNAME <> '" & s_r_machineName & "' ORDER BY HOSTNAME"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer
            While myReader.Read()
                s_hostName = myReader("HostName")
                Call PutInRAN_BROADCAST(s_hostName, string_HOSTNAME)
                b_removeSOFromRange = True
            End While
            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection1.Close()

            If b_removeSOFromRange Then
                Call RemoveSOFromRange()
            End If

        Catch e_Async_910 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_910 --> " & e_Async_910.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Function PutInRAN_BROADCAST(ByVal s_hostName, ByVal HOSTNAME) As Boolean
        Dim myConnectionString, s, s_queryName As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_no As String = "n"
        Dim s_zero As String = "0"
        Dim d_dateToday As DateTime = Now()
        Dim myInsertQuery As String
        Dim i_resursive As Int64 = i_p_countRecursive

        Try

            If s_p_authority = "yes" Then
                i_resursive = i_p_countRecursive
                If i_resursive > 0 Then
                    i_resursive = i_p_countRecursive
                End If
                myInsertQuery = "DELETE FROM RAN_BROADCAST WHERE HOSTNAME = '" & Trim(s_hostName) & "' AND MACHINECHANGE ='" & (HOSTNAME) & "'"
                Dim myCommand As New OleDbCommand(myInsertQuery)
                myCommand.Connection = myConnection
                myConnection.Open()
                myCommand.ExecuteNonQuery()

                myInsertQuery = ""
                myInsertQuery = "INSERT INTO RAN_BROADCAST (HOSTNAME,MACHINECHANGE, DONE, DDATE)" _
                 & "Values ('" & Trim(s_hostName) & "','" & (HOSTNAME) & "','" & (s_no) & "',#" & (d_dateToday) & "#)"
                Dim myCommand1 As New OleDbCommand(myInsertQuery)
                myCommand1.Connection = myConnection
                myCommand1.ExecuteNonQuery()
                myCommand1.Connection.Close()
                myConnection.Close()
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " async.954 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Sub Insert_In_BROADCAST_QUERY(ByVal s_queryIn)
        Dim myConnectionString, s_hostName As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_no As String = "n"
        Dim s_zero As String = "0"

        Try

            Dim mySelectQuery As String = "SELECT HOSTNAME FROM RANGE WHERE HOSTNAME <> '" & s_r_machineName & "' ORDER BY HOSTNAME"
            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
            myConnection1.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer
            While myReader.Read()
                s_hostName = myReader("HostName")
                Call Put_In_BROADCAST_QUERY(s_queryIn, s_hostName)
            End While
            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection1.Close()

        Catch e_Async_1054 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1054 --> " & e_Async_1054.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub Put_In_BROADCAST_QUERY(ByVal s_queryIn, ByVal s_hostName)
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_no As String = "n"
        Dim s_zero As String = "0"
        Dim myInsertQuery As String
        Dim d_dateToday As DateTime = Now()
        Dim s_queryMachineIn As Boolean

        Try

            Dim b_continue As Boolean = True
            Dim mySelectQuery As String = "SELECT s_Query FROM BROADCAST_QUERY WHERE s_Query = '" & Trim(s_queryIn) & "' and s_Machine = '" & Trim(s_hostName) & "'"
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_queryIn = myReader("s_Query")
                s_queryMachineIn = True
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            If s_queryMachineIn Then
                Dim s_yes As String = "yes"
                Dim myConnString = "UPDATE BROADCAST_QUERY SET s_Active ='" & (s_yes) & "', s_delete ='" & (s_no) & "' where s_Query = '" & Trim(s_queryIn) & "'"
                Dim myCommand2 As New OleDbCommand(myConnString)
                myCommand2.Connection = myConnection
                myConnection.Open()
                myCommand2.ExecuteNonQuery()
                myCommand2.Connection.Close()
                myConnection.Close()
            Else
                myInsertQuery = ""
                myInsertQuery = "INSERT INTO BROADCAST_QUERY (s_Query,s_Machine,d_Date)" _
                           & "Values ('" & Trim(s_queryIn) & "','" & Trim(s_hostName) & "',#" & (d_dateToday) & "#)"
                Dim myCommand1 As New OleDbCommand(myInsertQuery)
                myCommand1.Connection = myConnection
                myConnection.Open()
                myCommand1.ExecuteNonQuery()
                myCommand1.Connection.Close()
                myConnection.Close()
            End If
        Catch ex As System.Data.SqlClient.SqlException
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1150 --> " & ex.Message, EventLogEntryType.Information, 44)
        Catch eOX As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1154 --> " & eOX.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Public i_p_count_Query_SO As Int32

    Sub Check_BROADCAST_QUERY_SO()
        Try
            If i_p_count_Query_SO > 20 Or i_p_count_Query_SO = 0 Then

                i_p_count_Query_SO = 1
                Dim sSource As String = "AP_DENIS"
                Dim sLog As String = "Applo"
                Dim s_Verb As String
                Dim s_no As String = "n"
                Dim s_queryIn As String
                Dim b_continue As Boolean = True

                Call CleanBroadcastQuery()

                Dim mySelectQuery As String = "SELECT s_Query FROM DNAVALUE WHERE s_Authority = '" & s_no & "'"
                Dim myConnection As New OleDbConnection(ConnStringDNA())
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
                myConnection.Open()
                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

                While myReader.Read()
                    s_queryIn = myReader("s_Query")
                    If Check_BROADCAST_QUERY_SO(s_queryIn) = True Then
                        If Count_In_BROADCAST_QUERY_Not_Empty(s_queryIn) Then
                            Call Update_DNAVALUE_SO(s_queryIn)
                        End If
                    End If
                End While

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection.Close()
            Else
                i_p_count_Query_SO = i_p_count_Query_SO + 1
            End If

        Catch e_Async_1108 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1108 --> " & e_Async_1108.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub CleanBroadcastQuery()
        Try
            Dim s_queryIn As String
            Dim b_continue As Boolean = True
            Dim mySelectQuery As String = "SELECT distinct s_Query FROM BROADCAST_QUERY"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_queryIn = myReader("s_Query")
                Call Check_To_Delete_Query(s_queryIn)
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection.Close()
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1366 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub Check_To_Delete_Query(ByVal s_queryIn As String)
        Try
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            Dim s_Verb As String
            Dim s_no As String = "n"
            Dim b_queryIn As Boolean = False

            Dim mySelectQuery As String = "SELECT s_Query FROM DNAVALUE WHERE s_Query = '" & s_queryIn & "'"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_queryIn = myReader("s_Query")
                b_queryIn = True
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            If b_queryIn = False Then
                Call DeletequeryBroadcast(s_queryIn)
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1422 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub DeletequeryBroadcast(ByVal s_queryIn As String)
        Try
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myConnectionString, s, s_queryName As String
            Dim s_no As String = "n"
            Dim s_zero As String = "0"
            Dim d_dateToday As DateTime = Now()
            Dim myInsertQuery As String

            myInsertQuery = "DELETE FROM BROADCAST_QUERY WHERE s_Query = '" & Trim(s_queryIn) & "'"
            Dim myCommand1 As New OleDbCommand(myInsertQuery)
            myCommand1.Connection = myConnection
            myConnection.Open()
            myCommand1.ExecuteNonQuery()
            myCommand1.Connection.Close()
            myConnection.Close()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1437 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Function Check_BROADCAST_QUERY_SO(ByVal s_queryIn) As Boolean
        Try
            Check_BROADCAST_QUERY_SO = True
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            Dim s_Verb As String
            Dim s_so As String = "so"
            Dim s_sa As String = "sa"
            Dim s_upd As String = "u"
            Dim s_yes As String = "y"

            Dim b_continue As Boolean = True
            Dim mySelectQuery As String = "SELECT s_Query FROM BROADCAST_QUERY WHERE (s_Active<> '" & s_so & "' and s_Active<> '" & s_upd & "' and s_Active<> '" & s_sa & "') and (s_Delete<> '" & s_yes & "')"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_queryIn = myReader("s_Query")
                Check_BROADCAST_QUERY_SO = False
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection.Close()

        Catch e_Async_1108 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1108 --> " & e_Async_1108.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Sub Update_DNAVALUE_SO(ByVal s_queryIn)
        Try
            Dim s_so As String = "so"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myConnString = "UPDATE DNAVALUE SET s_Authority ='" & (s_so) & "'where s_Query = '" & (s_queryIn) & "'"

            Dim myCommand2 As New OleDbCommand(myConnString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch e_Async_1180 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1180 --> " & e_Async_1180.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Function Count_In_BROADCAST_QUERY_Not_Empty(ByVal s_queryIn) As Boolean
        Try
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            Dim s_Verb As String
            Dim i_totalQuery As Int32

            Dim b_continue As Boolean = True
            Dim mySelectQuery As String = "SELECT COUNT(s_Query) as TOTALQUERY FROM BROADCAST_QUERY WHERE s_Query = '" & s_queryIn & "'"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                i_totalQuery = myReader("TOTALQUERY")
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection.Close()
            If i_totalQuery > 0 Then
                Count_In_BROADCAST_QUERY_Not_Empty = True
            Else
                Count_In_BROADCAST_QUERY_Not_Empty = False
            End If

        Catch e_Async_1207 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1207 --> " & e_Async_1207.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Sub InsertInDeleted(ByVal URIToPass)
        Try
            Dim myConnection As New OleDbConnection(ConnStringDNA())

            If Len(URIToPass) > 0 Then
                Dim myInsertQuery As String = "INSERT INTO DELETED (S_QUERY) Values('" & Trim(URIToPass) & "')"
                Dim myCommand As New OleDbCommand(myInsertQuery)
                myCommand.Connection = myConnection
                myConnection.Open()
                myCommand.ExecuteNonQuery()
                myCommand.Connection.Close()
                myConnection.Close()
            End If

        Catch e_Async_1233 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1233 --> " & e_Async_1233.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub UpdateDeleted(ByVal URIToPass)
        Dim myConnectionString, s As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_yes As String = "y"

        Try
            myConnectionString = "UPDATE DELETED SET DONE='" & (s_yes) & "' WHERE S_QUERY = '" & Trim(URIToPass) & "'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch e_Async_1269 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1269 --> " & e_Async_1269.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Sub PutSAInDeleted(ByVal URIToPass)
        Dim myConnectionString, s As String
        Dim myConnection As New OleDbConnection(ConnStringDNA())
        Dim s_sa As String = "sa"

        Try
            myConnectionString = "UPDATE DELETED SET SENTOVER='" & (s_sa) & "' WHERE S_QUERY = '" & Trim(URIToPass) & "'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection
            myConnection.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            myConnection.Close()

        Catch e_Async_1269 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1269 --> " & e_Async_1269.Message, EventLogEntryType.Information, 44)
        End Try


    End Sub

    Sub DeleteQueryIfSOAll(ByVal URIToPass)

        Try
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            Dim s_Verb, s_queryIn As String
            Dim s_so As String = "so"
            Dim s_sa As String = "sa"
            Dim s_yes As String = "y"
            Dim b_deleteNow As Boolean = True

            Dim b_continue As Boolean = True
            Dim mySelectQuery As String = "SELECT s_Query FROM BROADCAST_QUERY WHERE s_Delete <> '" & s_so & "' and s_Query = '" & URIToPass & "'"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_queryIn = myReader("s_Query")
                b_deleteNow = False
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection.Close()

            If b_deleteNow Then

                Dim myConnectionString, s, s_queryName As String
                Dim s_no As String = "n"
                Dim s_zero As String = "0"
                Dim d_dateToday As DateTime = Now()
                Dim myInsertQuery As String

                myInsertQuery = "DELETE FROM BROADCAST_QUERY WHERE s_Query = '" & Trim(URIToPass) & "'"
                Dim myCommand1 As New OleDbCommand(myInsertQuery)
                myCommand1.Connection = myConnection
                myConnection.Open()
                myCommand1.ExecuteNonQuery()
                myCommand1.Connection.Close()
                myConnection.Close()

            End If

        Catch e_Async_1324 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1324 --> " & e_Async_1324.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Function RefreshDNSFOUND(ByVal s_machineName, ByVal RangeMinus, ByVal RangePlus) As Boolean
        Try
            Dim s_keepHostName, s_rangeMinus, s_rangePlus As String
            Dim b_nameIsIn As Boolean = False
            If Len(s_machineName) > 0 Then

                Dim mySelectQuery, s_Website, s As String
                mySelectQuery = "SELECT HostName,RangeMinus,RangePlus FROM Range where HostName ='" & s_machineName & "'"
                Dim myConnection1 As New OleDbConnection(ConnStringDNA())
                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)
                myConnection1.Open()

                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                Dim Index As Integer

                While myReader.Read()
                    s_rangeMinus = myReader("RangeMinus").ToString
                    s_rangePlus = myReader("RangePlus").ToString
                    Exit While
                End While

                myReader.Close()
                myConnection1.Close()

                If RangeMinus <> s_rangeMinus Or RangePlus <> s_rangePlus Then
                    RefreshDNSFOUND = True
                End If
            End If

        Catch e_Async_1324 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_Async_1324 --> " & e_Async_1324.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Sub UpdateDNSFOUNDWithNewRange(ByVal RangeMinus, ByVal RangePlus)
        Try
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            Dim s_Verb As String
            Dim s_no As String = "n"

            Dim s_queryIn As String
            Dim b_continue As Boolean = True
            Dim b_checkForOutside, b_checkForInside As Boolean

            Dim mySelectQuery As String = "SELECT s_Query FROM DNAVALUE"
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()
            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_queryIn = myReader("s_Query")
                b_checkForOutside = CheckForOutside(s_queryIn, RangeMinus, RangePlus)
                If b_checkForOutside Then
                    Call UpdateDNSFOUND_With_New_parameters_Out(s_queryIn, RangeMinus, RangePlus)
                End If
                b_checkForInside = CheckForInside(s_queryIn, RangeMinus, RangePlus)
                If b_checkForInside Then
                    Call UpdateDNSFOUND_With_New_parameters_In(s_queryIn, RangeMinus, RangePlus)
                End If
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1636 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Sub UpdateDNSFOUNDWithRangeEndOfCycle(ByVal s_databaseName)
        Try

            Dim s_rangeMinus As String = sr_RangeMinus
            Dim s_rangePlus As String = sr_RangePlus
            s_rangePlus = FindNextMinRange(s_rangePlus)

            Call UpdateDNSFOUND_With_New_parameters_Out_Of_Range(s_databaseName, s_rangeMinus, s_rangePlus)
            Call UpdateDNSFOUND_With_New_parameters_In_Out_Of_Range(s_databaseName, s_rangeMinus, s_rangePlus)

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1673 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Function CheckForOutside(ByVal s_queryIn, ByVal RangeMinus, ByVal RangePlus) As Boolean
        Try
            Dim mySelectQuery As String
            Dim s_no As String = "n"
            Dim s_yes As String = "y"
            Dim i_countURL As Integer
            Dim s_RangePlus As String = FindNextMinRange(RangePlus)

            mySelectQuery = "SELECT count(URLSite) as CountURL FROM DNSFOUND where (ShortURL < '" & (RangeMinus) & "' OR ShortURL > '" & (s_RangePlus) & "') AND (Done = '" & s_no & "' OR Done = '" & s_yes & "')"

            Dim myConnection As New OleDbConnection(ConnStringURLDNS(s_queryIn))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            Dim o_object As Object

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            While myReader.Read()
                i_countURL = myReader("CountURL")
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection.Close()

            If i_countURL > 0 Then
                CheckForOutside = True
            Else
                CheckForOutside = False
            End If

        Catch Async_1786 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async_1786 --> " & Async_1786.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Sub UpdateDNSFOUND_With_New_parameters_Out(ByVal s_queryIn, ByVal RangeMinus, ByVal RangePlus)

        Try
            Dim myConnectionString As String
            Dim s_so As String = "s"
            Dim s_nothing As String = ""
            RangePlus = FindNextMinRange(RangePlus)

            Dim objConn As New OleDbConnection(ConnStringURLDNS(s_queryIn))
            myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_so) & "', answer ='" & (s_nothing) & "'  where ShortURL < '" & (RangeMinus) & "' or ShortURL > '" & (RangePlus) & "'"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = objConn
            objConn.Open()
            myCommand2.ExecuteNonQuery()
            objConn.Close()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1657 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Sub UpdateDNSFOUND_With_New_parameters_Out_Of_Range(ByVal s_queryIn, ByVal RangeMinus, ByVal RangePlus)

        Try
            Dim myConnectionString As String
            Dim s_so As String = "s"
            Dim s_no As String = "n"
            Dim s_yes As String = "y"
            Dim s_nothing As String = ""
            Dim objConn As New OleDbConnection(ConnStringURLDNS(s_queryIn))
            myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_so) & "', answer ='" & (s_nothing) & "'  where (ShortURL < '" & (RangeMinus) & "' or ShortURL > '" & (RangePlus) & "') and (Done = '" & s_no & "' or Done = '" & s_yes & "')"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = objConn
            objConn.Open()
            myCommand2.ExecuteNonQuery()
            objConn.Close()
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1657 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Function CheckForInside(ByVal s_queryIn, ByVal RangeMinus, ByVal RangePlus) As Boolean
        Try
            Dim mySelectQuery As String
            Dim s_s As String = "s"
            Dim s_sa As String = "sa"
            Dim s_so As String = "so"
            Dim i_countURL As Integer

            mySelectQuery = "SELECT count(URLSite) as CountURL FROM DNSFOUND where ShortURL >= '" & (RangeMinus) & "' AND ShortURL <= '" & (RangePlus) & "' AND (Done = '" & s_s & "' OR Done = '" & s_sa & "' OR Done = '" & s_so & "')"

            Dim myConnection As New OleDbConnection(ConnStringURLDNS(s_queryIn))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            Dim o_object As Object

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            While myReader.Read()
                i_countURL = myReader("CountURL")
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If
            myConnection.Close()

            If i_countURL > 0 Then
                CheckForInside = True
            Else
                CheckForInside = False
            End If

        Catch Async_1786 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async_1786 --> " & Async_1786.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Sub UpdateDNSFOUND_With_New_parameters_In(ByVal s_queryIn, ByVal RangeMinus, ByVal RangePlus)

        Try
            Dim myConnectionString As String
            Dim s_no As String = "n"
            Dim s_s As String = "s"
            Dim s_sa As String = "sa"
            Dim s_so As String = "so"
            Dim s_nothing As String = ""
            RangePlus = FindNextMinRange(RangePlus)

            Dim objConn As New OleDbConnection(ConnStringURLDNS(s_queryIn))
            myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_no) & "', answer ='" & (s_nothing) & "'  where (ShortURL >= '" & (RangeMinus) & "' and ShortURL <= '" & (RangePlus) & "')"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = objConn
            objConn.Open()
            myCommand2.ExecuteNonQuery()
            objConn.Close()
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1678 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub


    Sub UpdateDNSFOUND_With_New_parameters_In_Out_Of_Range(ByVal s_queryIn, ByVal RangeMinus, ByVal RangePlus)

        Try
            Dim myConnectionString As String
            Dim s_no As String = "n"
            Dim s_s As String = "s"
            Dim s_sa As String = "sa"
            Dim s_so As String = "so"
            Dim s_nothing As String = ""
            Dim objConn As New OleDbConnection(ConnStringURLDNS(s_queryIn))
            myConnectionString = "UPDATE DNSFOUND SET Done='" & (s_no) & "', answer ='" & (s_nothing) & "'  where (ShortURL >= '" & (RangeMinus) & "' and ShortURL <= '" & (RangePlus) & "') and (Done = '" & s_s & "' or Done = '" & s_sa & "' or  Done = '" & s_so & "' )"
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = objConn
            objConn.Open()
            myCommand2.ExecuteNonQuery()
            objConn.Close()
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Async.1678 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Function GetGoodSpotFromURLFOUND(ByVal url, ByVal GoodSpot, ByVal sSendstrURI) As Int64

        Try

            Dim myConnString As String
            Dim sKeepString As String
            Dim s_se As String = "se"
            Dim i_countURL As Int32 = 0
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            Dim i_goodSpot As Int32

            Dim s_shortURL As String = Func_CleanUrl(url)

            Dim mySelectQuery As String = "SELECT Top 1 GoodSpot FROM URLFound where URL = '" & url & "' or ShortURL = '" & s_shortURL & "' ORDER BY Goodspot desc"
            Dim myConnection As New OleDbConnection(ConnStringURL(sSendstrURI))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                i_goodSpot = (myReader("GoodSpot"))
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            If i_goodSpot <= -900000 Then
                GetGoodSpotFromURLFOUND = -999999
            ElseIf i_goodSpot <= -10000 Then
                GetGoodSpotFromURLFOUND = -888888
            Else
                GetGoodSpotFromURLFOUND = GoodSpot
            End If

        Catch e_async_1521 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_async_1521 --> " & e_async_1521.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)

        End Try

    End Function

    Sub RemoveMachineFromRAN_Broadcast()
        Try
            Dim mySelectQuery, s, s_keepHostName As String
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim s_verb As String = "acknowNewMachinConfig"
            Dim b_continue As Boolean = True
            Dim s_so As String = "so"
            Dim i_count As Int64
            Dim s_machineName, s_hostName As String

            mySelectQuery = "SELECT DISTINCT HOSTNAME, MACHINECHANGE FROM RAN_BROADCAST"
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_hostName = myReader("HOSTNAME")
                Call DeleteMachine(s_hostName)
                s_machineName = myReader("MACHINECHANGE")
                If s_hostName <> s_machineName Then
                    Call DeleteMachine(s_machineName)
                End If
            End While

            myReader.Close()

        Catch e_async_1998 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_async_1998 --> " & e_async_1998.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)
        End Try
    End Sub

    Sub DeleteMachine(ByVal s_hostName As String)
        Try

            Dim s As String
            Dim s_so As String = "so"
            Dim s_verb As String = "newMachineConfig"
            Dim s_IPAddress As String = "nothing"
            Dim mySelectQuery As String
            Dim s_hostNameFound As String

            'Console.WriteLine(s_hostName & " --- search.959")

            mySelectQuery = "SELECT HOSTNAME FROM RANGE where HOSTNAME ='" & s_hostName & "'"

            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand1 As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()
            Dim myReader1 As OleDbDataReader = myCommand1.ExecuteReader()
            While myReader1.Read()
                s_hostNameFound = myReader1("HOSTNAME")
            End While
            If Not (myReader1.IsClosed) Then
                myReader1.Close()
            End If
            myConnection.Close()


            If Len(s_hostNameFound) = 0 Then
                Call DeleteFromRan_Broadcast(s_hostName)
            End If

        Catch e_async_2036 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_async_2036 --> " & e_async_2036.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)
        End Try
    End Sub

    Sub DeleteFromRan_Broadcast(ByVal s_hostName As String)
        Try
            Dim myConnectionString, s, s_queryName As String
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim s_no As String = "n"
            Dim s_zero As String = "0"
            Dim d_dateToday As DateTime = Now()
            Dim myInsertQuery As String
            Dim i_resursive As Int64 = i_p_countRecursive

            myInsertQuery = "DELETE FROM RAN_BROADCAST WHERE HOSTNAME = '" & s_hostName & "' OR MACHINECHANGE ='" & (s_hostName) & "'"
            Dim myCommand As New OleDbCommand(myInsertQuery)
            myCommand.Connection = myConnection
            myConnection.Open()
            myCommand.ExecuteNonQuery()

        Catch e_async_2073 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_async_2073 --> " & e_async_2073.Message, EventLogEntryType.Information, 45)
        End Try
    End Sub

End Module

