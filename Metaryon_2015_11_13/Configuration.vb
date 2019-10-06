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
Imports System.Threading
Imports System.Threading.ThreadPriority

Module Configuration

    'Public Class TestIt
    '    Shared Function test(ByVal sURI As String) As String
    '        test = sURI
    '    End Function
    'End Class

    'Function LookAlltheWeb_ToDelete(ByVal strURI As String) As String
    '    Dim SGetText As String
    '    Dim sGetHttp As String
    '    Dim sYahoo As String = "http://alltheweb.com/search?cat=web&cs=utf-8&l=any&q="
    '    Dim sFindNext20 As String = ">Next 20"
    '    Dim iEndUrlNext20 As Integer
    '    Dim iBeginNext20 As Integer
    '    Dim bContinue As Boolean = True
    '    Dim sFindQuote As String
    '    Dim iCounter As Integer = 10
    '    Dim sFindNext20Url As String
    '    Dim iLenghtNextUrl As Integer
    '    Dim sTakeCareOfString As String
    '    Dim iFirstTimeNext20 As Integer = 0
    '    Dim sOldsFindNext20Url As String
    '    Dim sKeepstrURI As String

    '    strURI = Trim(strURI)
    '    sKeepstrURI = Replace(strURI, " ", "+")
    '    sYahoo = sYahoo & strURI
    '    strURI = ""
    '    strURI = sYahoo
    '    Do While bContinue

    '        If Len(sFindNext20Url) > 10 Then
    '            If sFindNext20Url = sOldsFindNext20Url Then
    '                Exit Do
    '            End If
    '            strURI = sFindNext20Url
    '            sOldsFindNext20Url = sFindNext20Url
    '        End If

    '        MessageBox.Show(" form1.strURI.1145 -> " & strURI)

    '        Dim objURI As Uri = New Uri(strURI)
    '        Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
    '        Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
    '        Dim objStream As Stream = objWebResponse.GetResponseStream()
    '        Dim objStreamReader As StreamReader = New StreamReader(objStream)
    '        Dim strHTML As String = objStreamReader.ReadToEnd

    '        sTakeCareOfString = strHTML

    '        SGetText = strHTML

    '        If Len(SGetText) > 10 Then

    '            sGetHttp = "http://alltheweb.com/"

    '            LookAlltheWeb_ToDelete = strHTML
    '            If bExactPhrase Then
    '                sFindNext20Url = "http://alltheweb.com/search?q=" & sKeepstrURI & "&c=web&o=" & iCounter & "&l=any&cn=3&cs=utf-8&phrase=on"
    '            Else
    '                sFindNext20Url = "http://alltheweb.com/search?q=" & sKeepstrURI & "&c=web&o=" & iCounter & "&l=any&cn=3&cs=utf-8"
    '            End If

    '            iCounter = iCounter + 10
    '        End If

    '        If iCounter > 4000 Then
    '            Exit Do
    '        End If
    '    Loop
    'End Function


    ''Ne sert plus a rien***************************************************************************************


    Public Function InsertUrl(ByVal SUrlInsert As String) As String
        Dim sToInsert As String = SUrlInsert
        Dim myConnectionString As String
        ' If the connection string is null, use a default.
        Dim myConnection As New OleDbConnection(ConnString())
        Dim sDate As String = CStr(Now())
        Dim sNo As String = "n"

        If Len(sToInsert) > 0 Then

            Try
                Dim myInsertQuery As String = "INSERT INTO URLFOUND (url,DateHour,Done) Values('" & Trim(sToInsert) & "','" & (sDate) & "','" & (sNo) & "')"
                Dim myCommand As New OleDbCommand(myInsertQuery)
                myCommand.Connection = myConnection
                myConnection.Open()
                myCommand.ExecuteNonQuery()
                myCommand.Connection.Close()
                myConnection.Close()
            Catch e_form1_1202 As Exception
                Dim sSource As String = "AP_DENIS"
                Dim sLog As String = "Applo"
                EventLog.WriteEntry(sSource, " e_form1_1202--> " & e_form1_1202.Message, EventLogEntryType.Information, 44)
            Finally
                myConnection.Close()
            End Try
        End If
    End Function


    Sub CreateTable(ByVal sURIToPass As String, ByVal sMachine As String)

        Try
            Dim strPathToFile As String = Trim(sURIToPass)
            If Len(sURIToPass) > 2 Then
                strPathToFile = Replace(strPathToFile, " ", "_")
                strPathToFile = Replace(strPathToFile, "'", "_")
                Dim s_pathofURLFOUND As String = sAppPath() & strPathToFile & ".mdb"
                Dim s_pathOfDNSFOUND As String = sAppPath() & strPathToFile & "_DNS.mdb"

                If Not (File.Exists(s_pathofURLFOUND)) Then

                    Dim cat As New ADOX.Catalog
                    Dim strConn As String = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" & sAppPath() & strPathToFile & ".mdb"
                    cat.Create(strConn)

                    Dim myConnection As New OleDbConnection(strConn)
                    Dim myInsertQuery As String = "CREATE TABLE URLFOUND (IDCounter counter,ID Long default '0',URL varchar(250),DateHour Datetime,GoodSpot Long default '0',Done TEXT(2) NULL,Machine varchar(100) NULL,Signature Long default '0', Keywords Long default '0', ShortURL varchar(60) NULL, FreshPage TEXT(2) NULL, URL_PAGE MEMO DEFAULT 'n')"
                    Dim myCommand As New OleDbCommand(myInsertQuery)
                    myCommand.Connection = myConnection
                    myConnection.Open()
                    myCommand.ExecuteNonQuery()
                    myCommand.Connection.Close()
                    myConnection.Close()

                    Dim myInsertQuery1 As String = "CREATE INDEX IndexURL ON URLFOUND (IDCounter) with primary"
                    Dim myCommand1 As New OleDbCommand(myInsertQuery1)
                    myCommand1.Connection = myConnection
                    myConnection.Open()
                    myCommand1.ExecuteNonQuery()
                    myCommand1.Connection.Close()
                    myConnection.Close()

                End If

                If Not (File.Exists(s_pathOfDNSFOUND)) Then

                    Dim cat As New ADOX.Catalog
                    Dim strConn1 As String = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" & sAppPath() & strPathToFile & "_DNS.mdb"
                    cat.Create(strConn1)
                    Dim myConnectionDNS As New OleDbConnection(strConn1)
                    Dim myInsertQueryDNS As String = "CREATE TABLE DNSFOUND (ID counter, URLSite varchar(250),ILevel integer, DateHour Datetime,Done TEXT(2) NULL, iGoodSpot Long default '0', ShortURL varchar(60), Machine varchar(60) NULL, Answer varchar(60) NULL,NotFound_404 integer,SearchEngine varchar(60) NULL)"

                    Dim myCommandDNS As New OleDbCommand(myInsertQueryDNS)
                    'Try
                    myCommandDNS.Connection = myConnectionDNS
                    myConnectionDNS.Open()
                    myCommandDNS.ExecuteNonQuery()
                    myCommandDNS.Connection.Close()
                    myConnectionDNS.Close()

                    Dim myInsertQueryDNS1 As String = "CREATE INDEX IndexURLSite ON DNSFOUND (URLSite) with primary"
                    'MessageBox.Show(myInsertQuery, "myInsertQuery1 create DB")
                    Dim myCommandDNS1 As New OleDbCommand(myInsertQueryDNS1)
                    'Try
                    myCommandDNS1.Connection = myConnectionDNS
                    myConnectionDNS.Open()
                    myCommandDNS1.ExecuteNonQuery()
                    myCommandDNS1.Connection.Close()
                    myConnectionDNS.Close()

                    Dim myInsertQueryDNS2 As String = "CREATE INDEX IndexShortURL ON DNSFOUND (ShortURL)"
                    'MessageBox.Show(myInsertQuery, "myInsertQuery1 create DB")
                    Dim myCommandDNS2 As New OleDbCommand(myInsertQueryDNS2)
                    'Try
                    myCommandDNS2.Connection = myConnectionDNS
                    myConnectionDNS.Open()
                    myCommandDNS2.ExecuteNonQuery()
                    myCommandDNS2.Connection.Close()
                    myConnectionDNS.Close()

                    Dim myInsertQueryDNS3 As String = "CREATE INDEX IndexShortID ON DNSFOUND (ID)"
                    'MessageBox.Show(myInsertQuery, "myInsertQuery1 create DB")
                    Dim myCommandDNS3 As New OleDbCommand(myInsertQueryDNS3)
                    'Try
                    myCommandDNS3.Connection = myConnectionDNS
                    myConnectionDNS.Open()
                    myCommandDNS3.ExecuteNonQuery()
                    myCommandDNS3.Connection.Close()
                    myConnectionDNS.Close()

                    Dim myInsertQueryDNS4 As String = "CREATE INDEX IndexShortGoodSpot ON DNSFOUND (iGoodSpot)"
                    'MessageBox.Show(myInsertQuery, "myInsertQuery1 create DB")
                    Dim myCommandDNS4 As New OleDbCommand(myInsertQueryDNS4)
                    'Try
                    myCommandDNS4.Connection = myConnectionDNS
                    myConnectionDNS.Open()
                    myCommandDNS4.ExecuteNonQuery()
                    myCommandDNS4.Connection.Close()
                    myConnectionDNS.Close()

                    Dim myInsertQueryDNS5 As String = "CREATE INDEX IndexShortDateHour ON DNSFOUND (DateHour)"
                    Dim myCommandDNS5 As New OleDbCommand(myInsertQueryDNS5)
                    'Try
                    myCommandDNS5.Connection = myConnectionDNS
                    myConnectionDNS.Open()
                    myCommandDNS5.ExecuteNonQuery()
                    myCommandDNS5.Connection.Close()
                    myConnectionDNS.Close()
                    '---- Insert the query in QUERY database.
                End If
                '---- Create 1 table in database.

                Dim StrCon As String = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" & sAppPath() & "DNA\DNA.mdb"

                Dim myCon As New OleDbConnection(StrCon)
                Dim sDate As String = Now()
                Dim b_queryAlreadyIn As Boolean = CheckIfQueryAlreadyIn(sURIToPass)

                If b_queryAlreadyIn = False Then
                    Dim myInsert As String = "INSERT INTO QUERY (QuerySearch,DateInsert,Machine) Values('" & sURIToPass & "','" & (sDate) & "','" & (sMachine) & "')"
                    Dim myComm As New OleDbCommand(myInsert)
                    myComm.Connection = myCon
                    myCon.Open()
                    myComm.ExecuteNonQuery()
                    myComm.Connection.Close()
                    myCon.Close()
                End If

            End If

        Catch e_configuration_248 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_configuration_248--> " & e_configuration_248.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub


    'Sub InsertSite(ByVal SUrlInsert As String)

    '    Try
    '        Dim sToInsert As String = SUrlInsert
    '        Dim myConnectionString As String
    '        Dim myConnection As New OleDbConnection(ConnStringDNS())
    '        Dim sDate As String = Now()
    '        Dim iKeepSlash As Integer
    '        Dim sCheckString As String
    '        iKeepSlash = InStr(8, sToInsert, "/")

    '        If iKeepSlash > 0 Then
    '            sToInsert = Mid(sToInsert, 1, iKeepSlash)
    '        End If

    '        If Len(sToInsert) > 0 Then

    '            Dim myConnString As String
    '            Dim sKeepString As String
    '            Dim sDNSFound As String
    '            sToInsert = Replace(sToInsert, "'", "")
    '            Dim mySelectQuery As String = "SELECT URLSite FROM DNSFound where URLSite ='" & sToInsert & "' and Query ='" & sURIToPass & "'"

    '            Dim myCommand1 As New OleDbCommand(mySelectQuery, myConnection)
    '            myConnection.Open()
    '            Dim myReader As OleDbDataReader = myCommand1.ExecuteReader()

    '            sDNSFound = ""

    '            While myReader.Read()
    '                sDNSFound = sDNSFound & (myReader("URLSite")) & Chr(10)
    '            End While

    '            If Not (myReader.IsClosed) Then
    '                myReader.Close()
    '            End If

    '            myConnection.Close()

    '            If Len(sDNSFound) > 0 Then
    '                sKeepString = "stop"
    '            Else
    '                InsertDNSFound(sToInsert)
    '            End If

    '        End If

    '    Catch e_configuration_310 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " e_configuration_310--> " & e_configuration_310.Message, EventLogEntryType.Information, 44)
    '    End Try

    'End Sub


    'Sub InsertDNSFound_Copy(ByVal sToInsert As String)

    '    Try

    '        Dim myConnectionString As String
    '        ' If the connection string is null, use a default.
    '        Dim myConnection As New OleDbConnection(ConnStringDNS())
    '        Dim sDate As String = DateAdd(DateInterval.Minute, (-1 * i_p_Duration * 60), Now())
    '        Dim sCatchString As String = sToInsert
    '        Dim iLevel As Integer = 1
    '        Dim myInsertQuery As String

    '        Dim s_shortURL As String = Func_CleanUrl(sToInsert)
    '        If CheckRangeGoHere(sToInsert) = True Then
    '            Dim s_no As String = "n"
    '            myInsertQuery = "INSERT INTO DNSFound (URLSite,DateHour,Done,Query,ShortURL) Values('" & Trim(sToInsert) & "','" & (sDate) & "','" & (s_no) & "','" & (sURIToPass) & "','" & Trim(s_shortURL) & "')"
    '        Else
    '            Dim s_no As String = "s"
    '            myInsertQuery = "INSERT INTO DNSFound (URLSite,DateHour,Done,Query,ShortURL) Values('" & Trim(sToInsert) & "','" & (sDate) & "','" & (s_no) & "','" & (sURIToPass) & "','" & Trim(s_shortURL) & "')"
    '        End If

    '        Dim myCommand As New OleDbCommand(myInsertQuery)

    '        myCommand.Connection = myConnection
    '        myConnection.Open()
    '        myCommand.ExecuteNonQuery()
    '        myCommand.Connection.Close()
    '        myConnection.Close()

    '    Catch e_configuration_340 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " e_configuration_340--> " & e_configuration_340.Message, EventLogEntryType.Information, 44)
    '    End Try

    'End Sub


    'Function GetWebPageAsStringShort(ByVal strURI As String) As String

    '    Try

    '        Dim objURI As Uri = New Uri(strURI)
    '        Dim objWebRequest As WebRequest = WebRequest.Create(objURI)
    '        Dim objWebResponse As WebResponse = objWebRequest.GetResponse()
    '        Dim objStream As Stream = objWebResponse.GetResponseStream()
    '        Dim objStreamReader As StreamReader = New StreamReader(objStream)
    '        Dim strHTML As String = objStreamReader.ReadToEnd
    '        GetWebPageAsStringShort = strHTML

    '    Catch e_configuration_357 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " e_configuration_357--> " & e_configuration_357.Message, EventLogEntryType.Information, 44)
    '    End Try

    'End Function


    'Sub Delete_Click(Optional ByVal sURLToInsert = "", Optional ByVal s_machineNameAuthority = "")

    '    Dim sWhatToDelete As String = sURLToInsert
    '    Dim verb, strPathToFile, strPathToFile_DNS As String
    '    Dim strPathToFileLdb, strPathToFile_DNSLdb, strPathToFile1, strPathToFile_DNS1 As String
    '    Dim b_goBroadcast As Boolean = True

    '    strPathToFile = Trim(sWhatToDelete)
    '    strPathToFile = Replace(strPathToFile, " ", "_")
    '    strPathToFile1 = Replace(strPathToFile, "'", "_")
    '    strPathToFile = strPathToFile1 & ".mdb"
    '    strPathToFileLdb = strPathToFile1 & ".ldb"

    '    strPathToFile_DNS = sWhatToDelete
    '    strPathToFile_DNS = Replace(strPathToFile_DNS, " ", "_")
    '    strPathToFile_DNS1 = Replace(strPathToFile_DNS, "'", "_")
    '    strPathToFile_DNS = strPathToFile_DNS1 & "_DNS.mdb"
    '    strPathToFile_DNSLdb = strPathToFile_DNS1 & "_DNS.ldb"

    '    If (File.Exists(sAppPath() & strPathToFile)) Or (File.Exists(sAppPath() & strPathToFile_DNS)) Then

    '        Try

    '            Dim myConnection As New OleDbConnection(ConnStringDNA())
    '            Dim myInsertQuery As String = "Delete from DNAVALUE where s_Query = '" & sWhatToDelete & "'"

    '            Dim myCommand1 As New OleDbCommand(myInsertQuery)
    '            myCommand1.Connection = myConnection
    '            myConnection.Open()
    '            myCommand1.ExecuteNonQuery()
    '            myCommand1.Connection.Close()
    '            'myConnection = Nothing
    '            myConnection.Close()

    '            Dim myInsertQuery1 As String = "Delete from SCHEDULE where s_IDQuery = '" & sWhatToDelete & "'"

    '            Dim myCommand2 As New OleDbCommand(myInsertQuery1)
    '            myCommand2.Connection = myConnection
    '            myConnection.Open()
    '            myCommand2.ExecuteNonQuery()
    '            myCommand2.Connection.Close()
    '            'myConnection = Nothing
    '            myConnection.Close()
    '        Catch Service_Configuration_627 As Exception
    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " Service_Configuration_627 --> " & Service_Configuration_627.Message, EventLogEntryType.Information, 44)

    '        End Try

    '        Try
    '            If (File.Exists(sAppPath() & strPathToFile)) Then
    '                File.Delete(sAppPath() & strPathToFile)
    '            End If

    '            If (File.Exists(sAppPath() & strPathToFile_DNS)) Then
    '                File.Delete(sAppPath() & strPathToFile_DNS)
    '            End If

    '            If (File.Exists(sAppPath() & strPathToFileLdb)) Then
    '                File.Delete(sAppPath() & strPathToFileLdb)
    '            End If

    '            If (File.Exists(sAppPath() & strPathToFile_DNSLdb)) Then
    '                File.Delete(sAppPath() & strPathToFile_DNSLdb)
    '            End If

    '        Catch Service_configuration_426 As Exception
    '            b_goBroadcast = False
    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " Service_configuration_426 --> " & Service_configuration_426.Message, EventLogEntryType.Information, 44)
    '        End Try

    '    End If

    '    Try
    '        Dim s_verb As String = "AcknowledgeDelete"
    '        Dim s_insertDNA As String = s_r_machineName
    '        sURLToInsert = Replace(sURLToInsert, " ", "_")
    '        If b_goBroadcast Then
    '            Call BroadCastDNSFOUND(s_verb, sURLToInsert, s_insertDNA, s_machineNameAuthority)
    '        End If
    '    Catch Service_configuration_445 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " Service_configuration_445 --> " & Service_configuration_445.Message, EventLogEntryType.Information, 44)
    '    End Try

    'End Sub

End Module
