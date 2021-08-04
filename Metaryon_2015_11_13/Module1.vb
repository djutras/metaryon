
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
Imports System.Deployment

Public Module Module1
    Public sURIToPass As String
    Public bExactPhrase As Boolean
    Public iPercFree As Integer
    Public iMinFree As Integer

    Public Function ConnString() As String
        ConnString = "Provider=SQLOLEDB;Server=tcp:dbmetaryon.database.windows.net,1433;Database=Metaryon;Uid=djutras@dbmetaryon;Pwd=B1mjej86;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"
    End Function

    Public s_pathWithNameUrl As String
    Public s_oldDatabaseNameUrl As String

    Public Function ConnStringURL(ByVal sSendstrURI As String) As String
        Try

            ConnStringURL = "Provider=SQLOLEDB;Server=tcp:dbmetaryon.database.windows.net,1433;Database=Metaryon;Uid=djutras@dbmetaryon;Pwd=B1mjej86;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " module1.83 --> " & ex.Message, EventLogEntryType.Information, 45)
        End Try
    End Function

    Public s_pathWithName As String
    Public s_oldDatabaseName As String

    Public Function ConnStringURLDNS(ByVal sSendstrURI As String) As String
        Try

            ConnStringURLDNS = "Provider=SQLOLEDB;Server=tcp:dbmetaryon.database.windows.net,1433;Database=Metaryon;Uid=djutras@dbmetaryon;Pwd=B1mjej86;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " module1.92 --> " & ex.Message, EventLogEntryType.Information, 45)
        End Try
    End Function


    Public Function ConnStringDNS() As String
        'ConnStringDNS = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" & sAppPath() & "Viiger_DNS.mdb"
        ConnStringDNS = "Provider=SQLOLEDB;Server=tcp:dbmetaryon.database.windows.net,1433;Database=Metaryon;Uid=djutras@dbmetaryon;Pwd=B1mjej86;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"

    End Function

    Public Function ConnStringDNA() As String
        'ConnStringDNA = "Provider=Microsoft.Jet.OleDb.4.0;Data Source=" & sAppPath() & "DNA\DNA.mdb"
        ConnStringDNA = "Provider=SQLOLEDB;Server=tcp:dbmetaryon.database.windows.net,1433;Database=Metaryon;Uid=djutras@dbmetaryon;Pwd=B1mjej86;Encrypt=yes;TrustServerCertificate=no;Connection Timeout=30;"

    End Function

    Public s_pathOfApps As String

    Public Function sAppPath() As String
        Try
            If s_pathOfApps = "" Then
                Dim strt As String = System.Reflection.Assembly.GetExecutingAssembly().Location

                strt = Replace(strt, "\Debug\Metaryon_2015_11_13.exe", "")

                strt = strt & "\"
                s_pathOfApps = strt
                sAppPath = s_pathOfApps
            Else
                sAppPath = s_pathOfApps
            End If
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " module1.120 --> " & ex.Message, EventLogEntryType.Information, 45)
        End Try
    End Function

    Public Function CheckPeriod(ByVal bPeriod As Boolean) As Boolean
        Try
            Dim myConnString As String
            Dim dCheckPeriod As DateTime
            Dim mySelectQuery As String = "SELECT min(DateHour) as DateHour FROM DNSFOUND"
            'MessageBox.Show(mySelectQuery, "mySelectQuery time")
            Dim myConnection As New OleDbConnection(ConnStringDNS())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            'myConnection.Open()
            Try
                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
                While myReader.Read()
                    dCheckPeriod = myReader.GetDateTime(0)
                End While
                'MessageBox.Show(dCheckPeriod, "time try")
                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If
                ' always call Close when done reading.
                'myConnection = Nothing
                myConnection.Close()
            Catch
                dCheckPeriod = Now()
                'MessageBox.Show(dCheckPeriod, "time catch")
            End Try

            Dim iPeriodHour As Integer
            Dim mySelectQuery1 As String = "SELECT PeriodHour FROM DNA"
            Dim myConnection1 As New OleDbConnection(ConnStringDNS())
            Dim myCommand1 As New OleDbCommand(mySelectQuery1, myConnection1)

            Dim dHourDiff As Integer
            Dim dNow As DateTime = Now()

            dHourDiff = DateDiff(DateInterval.Minute, dCheckPeriod, dNow)

            If dHourDiff > iPeriodHour Then
                CheckPeriod = True
                'MessageBox.Show(dHourDiff, "dHourDiff True")
            Else
                CheckPeriod = False
                'MessageBox.Show(dHourDiff, "dHourDiff False ")
            End If
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " module1.164 --> " & ex.Message, EventLogEntryType.Information, 45)
        End Try
    End Function

    Public Sub CompacAccess1()

        'Dim jro As JRO.JetEngine
        'jro = New JRO.JetEngine()
        'Try
        'JRO.CompactDatabase("Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & sAppPath() & "Viiger_DNS.mdb", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sAppPath() & "Viiger_DNS1.mdb;Jet OLEDB:Engine Type=5")
        'If File.Exists(sAppPath() & "Viiger_DNS.mdb") Then
        'File.Move(sAppPath() & "Viiger_DNS.mdb", sAppPath() & "Viiger_DNS2.mdb")
        'File.Move(sAppPath() & "Viiger_DNS1.mdb", sAppPath() & "Viiger_DNS.mdb")
        'If File.Exists(sAppPath() & "Viiger_DNS.mdb") And File.Exists(sAppPath() & "Viiger_DNS2.mdb") Then
        'File.Delete(sAppPath() & "Viiger_DNS2.mdb")
        'End If
        'End If
        'Catch
        'End Try

    End Sub

    '    Public Sub CompacAccess1(ByVal sSendUri)
    '        Try
    '            Dim s_pathToDataDNS As String = ConnStringURLDNS(sSendUri)
    '            Dim sSendstrURI As String = Trim(sSendUri)
    '            sSendstrURI = sSendstrURI & "_DNS.mdb"
    '            sSendstrURI = Replace(sSendstrURI, " ", "_")

    '            Dim jro As JRO.JetEngine
    '            jro = New JRO.JetEngine

    '            If File.Exists("Viiger_DNS1.mdb") Then
    '                File.Delete(sAppPath() & "Viiger_DNS1.mdb")
    '            End If

    '            If File.Exists("Viiger_DNS2.mdb") Then
    '                File.Delete(sAppPath() & "Viiger_DNS2.mdb")
    '            End If

    '            jro.CompactDatabase("Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & s_pathToDataDNS & "", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sAppPath() & "Viiger_DNS1.mdb;Jet OLEDB:Engine Type=5")

    '            If File.Exists(sSendstrURI) Then
    '                File.Move(sSendstrURI, "Viiger_DNS2.mdb")
    '                File.Move("Viiger_DNS1.mdb", sSendstrURI)
    '                If File.Exists(sSendstrURI) And File.Exists("Viiger_DNS2.mdb") Then
    '                    File.Delete(sAppPath() & "Viiger_DNS2.mdb")
    '                End If
    '            End If

    '        Catch eOx As Exception
    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, "Module1.Compac_access_dns.183 --> " & eOx.Message, EventLogEntryType.Information, 44)
    '        End Try
    '    End Sub


    Public Sub CompacAccessUrlFound(ByVal sSendUri)
        Try
            'Dim s_pathToDataDNS As String = ConnStringURL(sSendUri)
            'Dim sSendstrURI As String = Trim(sSendUri)
            'sSendstrURI = sSendstrURI & ".mdb"
            'sSendstrURI = Replace(sSendstrURI, " ", "_")

            'Dim jro As JRO.JetEngine
            'jro = New JRO.JetEngine

            'If File.Exists("Viiger_DNS1.mdb") Then
            '    File.Delete(sAppPath() & "Viiger_DNS1.mdb")
            'End If

            'If File.Exists("Viiger_DNS2.mdb") Then
            '    File.Delete(sAppPath() & "Viiger_DNS2.mdb")
            'End If

            'jro.CompactDatabase("Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & s_pathToDataDNS & "", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sAppPath() & "Viiger_1.mdb;Jet OLEDB:Engine Type=5")

            'If File.Exists(sSendstrURI) Then
            '    File.Move(sSendstrURI, "Viiger_2.mdb")
            '    File.Move("Viiger_1.mdb", sSendstrURI)
            '    If File.Exists(sSendstrURI) And File.Exists("Viiger_2.mdb") Then
            '        File.Delete(sAppPath() & "Viiger_2.mdb")
            '    End If
            'End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, "Module1.Compac_access_dns.183 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub


    '    Public b_notDoneCompaction As Boolean = True
    '    Public i_p_counterForAccess As Integer

    '    Public Sub CompacAccess()

    '        Try

    '            Dim s_pathToDataDNA As String = ConnStringDNA()
    '            Dim s_path As String = sAppPath() & "DNA\"

    '            Dim jro As JRO.JetEngine
    '            jro = New JRO.JetEngine

    '            jro.CompactDatabase("Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & s_pathToDataDNA & "", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & s_path & "DNA_1.mdb;Jet OLEDB:Engine Type=5")

    '            If File.Exists("DNA\DNA.mdb") Then
    '                File.Move("DNA\DNA.mdb", "DNA\DNA_2.mdb")
    '                File.Move("DNA\DNA_1.mdb", "DNA\DNA.mdb")
    '                If File.Exists("DNA\DNA.mdb") And File.Exists("DNA\DNA_2.mdb") Then
    '                    File.Delete(s_path & "DNA_2.mdb")
    '                End If
    '            End If

    '            'Dim b_continue As Boolean = True
    '            'Dim i_day As Integer = DateDiff(DateInterval.Day, d_p_DateTime, Now())
    '            'Dim i_moduloDay = i_day Mod 5
    '            'Dim i_counter As Int32

    '            'Dim File As New FileInfo(sAppPath() & "DNA\DNA.mdb")
    '            'Dim i_lenDNA As Int64 = File.Length / 1024

    '            'If i_lenDNA >= 100000 Then

    '            'Dim myProcess As New Process
    '            'myProcess.StartInfo.FileName = "C_VIiger"
    '            'myProcess.StartInfo.WindowStyle = ProcessWindowStyle.Normal
    '            'myProcess.Start()

    '            'End If

    '        Catch e_module1_262 As Exception

    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " e_module1_262 --> " & e_module1_262.Message, EventLogEntryType.Information, 44)

    '        End Try


    '    End Sub

    '    <Serializable()>
    '    Public Class Voice
    '        Public Enum GenderType
    '            Male
    '            Female
    '        End Enum

    '        Private mstrName As String
    '        Private mGender As GenderType
    '        Private mdtBirthDate As Date

    '        Public Property Name() As String
    '            Get
    '                Return mstrName
    '            End Get
    '            Set(ByVal Value As String)
    '                mstrName = Value
    '            End Set
    '        End Property

    '        Public Property Gender() As GenderType
    '            Get
    '                Return mGender
    '            End Get
    '            Set(ByVal Value As GenderType)
    '                mGender = Value
    '            End Set
    '        End Property

    '        Public Property BirthDate() As Date
    '            Get
    '                Return mdtBirthDate
    '            End Get
    '            Set(ByVal Value As Date)
    '                mdtBirthDate = Value
    '            End Set
    '        End Property

    '        Public ReadOnly Property Age() As Integer
    '            Get
    '                Return DateDiff(DateInterval.Year, mdtBirthDate, Now)
    '            End Get
    '        End Property

    '        Public ReadOnly Property Tone() As Integer
    '            Get
    '                Select Case mGender
    '                    Case GenderType.Female
    '                        Return Age ^ 1.2 / 2
    '                    Case GenderType.Male
    '                        Return Age ^ 1.2
    '                End Select
    '            End Get
    '        End Property
    '    End Class

    Function FindSignature(ByVal s_FileToCrawlFromScratchpad, ByVal s_DatabaseName) As Int64
        Dim i_line As Int64
        Try
            s_FileToCrawlFromScratchpad = RemoveHtmlTag(s_FileToCrawlFromScratchpad)
            Dim s_firstKeywords, s_secondKeywords As String
            Dim s_firstKeywordsCheck, s_secondKeywordsCheck As String
            Dim i_whereFirstKeywords, i_whereSecondKeywords, i_counter _
            , i_beginBody, i_totalDifference, i_whereToLookAfter As Int64
            Dim b_continue As Boolean = True
            Dim i_numberOfKeywords As Int64 = 0
            Dim i_counterAsSafeguard As Int64
            Dim i_totalDifferenceBefore As Int64
            Dim b_newKeywords As Boolean = False

            s_firstKeywords = FirstKeyword(s_DatabaseName, s_firstKeywordsCheck)
            i_whereFirstKeywords = InStr(1, s_FileToCrawlFromScratchpad, s_firstKeywords, CompareMethod.Text)

            If i_whereFirstKeywords > 0 Then
                b_newKeywords = True
            End If

            i_whereToLookAfter = (i_whereFirstKeywords + Len(s_firstKeywords))

            Dim i_lenFirstKeywords As Int64 = Len(Trim(s_firstKeywords))

            'If i_whereFirstKeywords = 0 Or i_lenFirstKeywords = 0 Then
            '    b_continue = False
            'End If

            While b_continue

                If i_whereFirstKeywords > 0 Then
                    i_whereSecondKeywords = InStr(i_whereToLookAfter, s_FileToCrawlFromScratchpad, "e", CompareMethod.Text)
                    If i_whereSecondKeywords > 0 Then
                        i_totalDifference = i_totalDifference + (i_whereSecondKeywords - i_whereFirstKeywords)
                    Else
                        Exit While
                    End If
                Else
                    s_firstKeywordsCheck = s_firstKeywords
                    s_firstKeywords = FirstKeyword(s_DatabaseName, s_firstKeywordsCheck)
                    If s_firstKeywords = "stop" Or Len(Trim(s_firstKeywords)) = 0 Then
                        Exit While
                    Else
                        b_newKeywords = True
                    End If
                    i_whereSecondKeywords = 1
                End If

                If i_whereSecondKeywords > 0 Then
                    i_whereFirstKeywords = InStr(i_whereSecondKeywords, s_FileToCrawlFromScratchpad, s_firstKeywords, CompareMethod.Text)
                    If i_whereFirstKeywords > 0 Then
                        b_newKeywords = True
                    End If
                Else
                    i_whereFirstKeywords = 0
                End If

                If i_totalDifference <> i_totalDifferenceBefore And b_newKeywords Then
                    i_numberOfKeywords = i_numberOfKeywords + 1
                    b_newKeywords = False
                End If

                i_totalDifferenceBefore = i_totalDifference

                i_whereToLookAfter = (i_whereFirstKeywords + Len(s_firstKeywords) + 1)

                i_counterAsSafeguard = i_counterAsSafeguard + 1

                If i_counterAsSafeguard > 200 Then
                    Exit While
                End If

            End While

            Dim s_collectNumber As String

            s_collectNumber = CStr(i_numberOfKeywords) & "00000" & CStr(i_totalDifference)

            FindSignature = CDbl(s_collectNumber)

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " Module1.464 --> " & ex.Message & " -- ligne-->" & i_line, EventLogEntryType.Information, 45)
        End Try

    End Function



    Function FirstKeyword(ByVal s_DatabaseName, ByVal s_fKeywords) As String
        Try
            s_DatabaseName = Replace(s_DatabaseName, "(", "")
            s_DatabaseName = Replace(s_DatabaseName, ")", "")
            s_DatabaseName = Replace(s_DatabaseName, "  ", " ")
            s_DatabaseName = Replace(s_DatabaseName, "  ", " ")
            Dim s_keepArray() As String = Split(s_DatabaseName)
            Dim s_keepFirstKeywords As String = s_fKeywords
            Dim s_keep_secondKeywords As String = ""
            Dim i_whereFirstKeywords, i_whereSecondKeywords, i_counter As Int32
            Dim b_continue As Boolean = True
            Dim s_string As String

            i_counter = 0

            For i_counter = 0 To UBound(s_keepArray)
                s_string = s_keepArray(i_counter)
                If s_fKeywords = "" Then
                    s_string = s_string
                    Exit For
                Else
                    If s_fKeywords = s_string Then
                        If i_counter < UBound(s_keepArray) Then
                            s_string = s_keepArray(i_counter + 1)
                            If LCase(s_string) = "or" Then
                                If i_counter + 2 <= UBound(s_keepArray) Then
                                    s_string = s_keepArray(i_counter + 2)
                                    Exit For
                                End If
                            Else
                                Exit For
                            End If
                        Else
                            s_string = "stop"
                        End If
                    End If
                End If
            Next

            FirstKeyword = s_string

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " module1.525 --> " & ex.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)
        End Try


    End Function

    Function SecondKeyword(ByVal s_DatabaseName, ByVal s_sKeywords) As String

        Try
            Dim s_keepArray() As String = Split(s_DatabaseName)
            Dim s_keepFirstKeywords As String = ""
            Dim s_keep_secondKeywords As String = ""
            Dim i_whereFirstKeywords, i_whereSecondKeywords, i_counter As Int32
            Dim b_continue As Boolean = True

            If UBound(s_keepArray) = 0 Then
                'MsgBox("UBound(s_keepArray)_253) is " & UBound(s_keepArray) & " " & s_DatabaseName)
            Else

                While b_continue

                    If s_keepFirstKeywords = "" Then
                        If s_keepArray(i_counter) = "or" Or s_keepArray(i_counter) = "Or" Or
                        s_keepArray(i_counter) = "OR" Then
                            i_counter = (i_counter + 1)
                        Else
                            s_keepFirstKeywords = s_keepArray(i_counter)
                        End If
                    End If

                    If s_keep_secondKeywords = "" Then
                        If s_keepArray(i_counter + 1) = "or" Or s_keepArray(i_counter + 1) = "Or" Or
                        s_keepArray(i_counter + 1) = "OR" Then
                        Else
                            i_whereSecondKeywords = s_keepArray(i_counter + 1)
                        End If
                    End If

                End While

                'MsgBox("UBound(s_keepArray)_253) is " & UBound(s_keepArray) & " " & s_DatabaseName)
            End If

        Catch e_module1_353 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_353 --> " & e_module1_353.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)
        End Try

    End Function


    Function Funct_PageCounter(ByVal s_urlToCount, ByVal sDatabaseName, ByVal sNo) As String
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try
            If Len(s_urlToCount) > 0 And Len(sDatabaseName) > 0 Then
                If Func_Add(s_urlToCount, sNo, sDatabaseName) = "stop" Then
                    'Call Sub_BlockThatURLInDNSFOUND(s_urlToCount, sDatabaseName)
                    Funct_PageCounter = "stop"
                End If
            Else
                Funct_PageCounter = "stop"
                'EventLog.WriteEntry(sSource, " checker pourquoi ca arrive e_module1_425 s_urlToCount--> " & s_urlToCount & " sDatabaseName " & sDatabaseName, EventLogEntryType.Information, 49)
            End If

        Catch e_module1_399 As Exception

            EventLog.WriteEntry(sSource, " e_module1_399 --> " & e_module1_399.Message, EventLogEntryType.Information, 45)
        End Try
    End Function

    Function Func_CleanUrl(ByVal s_urlToCount) As String
        Try
            If Len(Trim(s_urlToCount)) > 2 Or InStr(s_urlToCount, ".", CompareMethod.Text) > 0 Then
                Dim s_array() As String
                Dim s_string As String
                Dim s_secondArray() As String
                Dim sb As New System.Text.StringBuilder
                Dim i_uBound As Integer

                s_urlToCount = Replace(s_urlToCount, "https://www.", "")
                s_urlToCount = Replace(s_urlToCount, "http://www.", "")
                s_urlToCount = Replace(s_urlToCount, "https://", "")
                s_urlToCount = Replace(s_urlToCount, "http://", "")

                s_array = Split(s_urlToCount, "/")

                s_string = s_array(0)

                s_secondArray = Split(s_string, ".")

                i_uBound = UBound(s_secondArray)

                If IsNumeric(s_secondArray(i_uBound)) Then
                    sb.Append(s_secondArray(i_uBound - 3))
                    sb.Append(".")
                    sb.Append(s_secondArray(i_uBound - 2))
                    sb.Append(".")
                    sb.Append(s_secondArray(i_uBound - 1))
                    sb.Append(".")
                    sb.Append(s_secondArray(i_uBound))
                    Func_CleanUrl = sb.ToString
                    Exit Function
                End If

                If i_uBound = 1 Then
                    sb.Append(s_secondArray(0))
                    sb.Append(".")
                    sb.Append(s_secondArray(1))
                    If Len(sb.ToString) > 60 Then
                        Dim i_keepCount As Integer
                        i_keepCount = Len(sb.ToString) - 60
                        Func_CleanUrl = Trim(Mid(sb.ToString, Len(sb.ToString) - i_keepCount, 60))
                    Else
                        Func_CleanUrl = sb.ToString
                    End If
                    Exit Function
                End If

                If Len(s_secondArray(i_uBound)) <= 3 Then
                    Dim b_dontDoIt As Boolean
                    If UBound(s_secondArray) = 0 Then
                        b_dontDoIt = True
                    End If
                    If Not (b_dontDoIt) Then
                        If Len(s_secondArray(i_uBound - 1)) <= 3 Then
                            If i_uBound >= 2 Then
                                sb.Append(s_secondArray(i_uBound - 2))
                                sb.Append(".")
                                sb.Append(s_secondArray(i_uBound - 1))
                                sb.Append(".")
                                sb.Append(s_secondArray(i_uBound))
                                If Len(sb.ToString) > 60 Then
                                    Dim i_keepCount As Integer
                                    i_keepCount = Len(sb.ToString) - 60
                                    Func_CleanUrl = Trim(Mid(sb.ToString, Len(sb.ToString) - i_keepCount, 60))
                                Else
                                    Func_CleanUrl = sb.ToString
                                End If
                                Exit Function
                            Else
                                sb.Append(s_secondArray(i_uBound - 1))
                                sb.Append(".")
                                sb.Append(s_secondArray(i_uBound))
                                If Len(sb.ToString) > 60 Then
                                    Dim i_keepCount As Integer
                                    i_keepCount = Len(sb.ToString) - 60
                                    Func_CleanUrl = Trim(Mid(sb.ToString, Len(sb.ToString) - i_keepCount, 60))
                                Else
                                    Func_CleanUrl = sb.ToString
                                End If
                                Exit Function
                            End If
                        Else
                            sb.Append(s_secondArray(i_uBound - 1))
                            sb.Append(".")
                            sb.Append(s_secondArray(i_uBound))
                            If Len(sb.ToString) > 60 Then
                                Dim i_keepCount As Integer
                                i_keepCount = Len(sb.ToString) - 60
                                Func_CleanUrl = Trim(Mid(sb.ToString, Len(sb.ToString) - i_keepCount, 60))
                            Else
                                Func_CleanUrl = sb.ToString
                            End If
                            Exit Function
                        End If
                    End If

                    If Len(sb.ToString) < 2 Then

                    End If


                End If

                If Len(sb.ToString) > 0 Then
                    Func_CleanUrl = sb.ToString
                Else
                    Func_CleanUrl = s_string
                End If
            End If

        Catch e_module1_471 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " module1.471 --> " & s_urlToCount & " -- " & e_module1_471.Message, EventLogEntryType.Information, 45)
        End Try
    End Function



    Function Func_Add(ByVal s_urlToCount, ByVal sNo, Optional ByVal sDatabaseName = "") As String

        Dim b_shortURLIsIn As Boolean = False
        Dim s_keepString, s_pageCounter As String
        Dim i_pageCounter As Long
        s_urlToCount = Func_CleanUrl(s_urlToCount)

        Try
            Dim mySelectQuery As String = "SELECT ShortURL,sCount FROM PAGECOUNTER where ShortURL ='" & Trim(s_urlToCount) & "' AND sQuery ='" & Trim(sDatabaseName) & "'"

            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                s_keepString = s_keepString & (myReader.GetString(0))
                If Not Convert.IsDBNull(myReader("sCount")) Then
                    s_pageCounter = myReader("sCount")
                Else
                    s_pageCounter = "1"
                End If
                i_pageCounter = CInt(s_pageCounter)
                b_shortURLIsIn = True
                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            'myConnection = Nothing

            myConnection.Close()

        Catch e_module1_496 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_496 --> " & e_module1_496.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)
        End Try

        Try
            If b_shortURLIsIn Then

                If i_pageCounter > i_p_MaxPage Then
                    Func_Add = "stop"
                    Exit Function
                End If

                i_pageCounter = i_pageCounter + 1

                Call Sub_UpdateCounter(s_keepString, i_pageCounter, sDatabaseName)

            Else

                If Len(s_urlToCount) > 0 Then
                    Dim myInsertQuery As String = "INSERT INTO PAGECOUNTER (ShortURL,Done,sQuery)" _
                    & "Values ('" & s_urlToCount & "','" & Trim(sNo) & "','" & sDatabaseName & "')"

                    Dim myConnection As New OleDbConnection(ConnStringDNA())
                    'MsgBox("myInsertQuery 674 " & myInsertQuery)
                    Dim myCommand As New OleDbCommand(myInsertQuery)
                    myCommand.Connection = myConnection
                    myConnection.Open()
                    myCommand.ExecuteNonQuery()
                    myCommand.Connection.Close()
                    'myConnection = Nothing
                    myConnection.Close()
                End If

            End If

        Catch e_module1_527 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_module1_527 --> " & e_module1_527.Message, EventLogEntryType.Information, 45)
            EventLog.WriteEntry(sSource, " e_s_urlToCount_527 --> " & s_urlToCount, EventLogEntryType.Information, 45)

            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)
        End Try

    End Function

    Sub Sub_UpdateCounter(ByVal s_keepString, ByVal i_pageCounter, ByVal sDatabaseName)
        Dim myInsertQuery As String
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try
            If IsNumeric(i_pageCounter) And Len(s_keepString) > 0 Then
                Dim myConnection As New OleDbConnection(ConnStringDNA())

                myInsertQuery = "Update PAGECOUNTER set sCount = '" & CStr(i_pageCounter) & "' where ShortURL = '" & s_keepString & "' and sQuery = '" & sDatabaseName & "'"

                Dim myCommand As New OleDbCommand(myInsertQuery)

                myCommand.Connection = myConnection
                myConnection.Open()
                myCommand.ExecuteNonQuery()
                myCommand.Connection.Close()
                'myConnection = Nothing
                myConnection.Close()
            Else
                'EventLog.WriteEntry(sSource, " e_module1_627  i_pageCounter--> " & i_pageCounter & " ShortURL " & s_keepString, EventLogEntryType.Information, 45)
            End If
        Catch e_module1_566 As Exception

            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_566 --> " & e_module1_566.Message, EventLogEntryType.Information, 45)
            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_566 myInsertQuery--> " & myInsertQuery, EventLogEntryType.Information, 45)

        End Try

    End Sub


    Sub Sub_BlockThatURLInDNSFOUND(ByVal s_urlToCount, ByVal sDatabaseName)
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try
            Dim s_shortUrl As String = Func_CleanUrl(s_urlToCount)

            If Len(s_urlToCount) > 0 And Len(sDatabaseName) > 0 And Len(s_shortUrl) > 0 Then

                Dim s_Yes As String = "y"

                Dim objConn As New OleDbConnection(ConnStringURLDNS(sDatabaseName))

                Dim myConnectionString As String =
                "UPDATE DNSFOUND SET Done='" & (s_Yes) & "' where ShortURL ='" & (s_shortUrl) & "' and sCurrentQuery='" & (sDatabaseName) & "'"

                Dim myCommand2 As New OleDbCommand(myConnectionString)
                myCommand2.Connection = objConn
                objConn.Open()
                myCommand2.ExecuteNonQuery()
                'objconn = Nothing
                objConn.Close()

            Else

                'EventLog.WriteEntry(sSource, " lookforfiles.e_module1_600 s_urlToCount--> " & s_urlToCount & "  sDatabaseName--> " & sDatabaseName & "  s_shortUrl--> " & s_shortUrl, EventLogEntryType.Information, 49)

            End If
        Catch e_module1_600 As Exception


            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_600 --> " & e_module1_600.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)

        End Try

    End Sub


    Sub Sub_DeleteRecordsScratchPad()

        Try
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myInsertQuery As String = "Delete from SCRATCHPAD"

            Dim myCommand1 As New OleDbCommand(myInsertQuery)
            myCommand1.Connection = myConnection
            myConnection.Open()
            myCommand1.ExecuteNonQuery()
            myCommand1.Connection.Close()
            'myConnection = Nothing
            myConnection.Close()

        Catch e_module1_635 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_635 --> " & e_module1_635.Message, EventLogEntryType.Information, 45)
        End Try

    End Sub


    '    Sub Sub_DeleteRecordsFromQuery(ByVal s_Query)

    '        Try
    '            Dim myConnection As New OleDbConnection(ConnStringDNA())
    '            Dim myInsertQuery As String = "Delete from Query where QuerySearch = '" & s_Query & "'"

    '            Dim myCommand1 As New OleDbCommand(myInsertQuery)
    '            myCommand1.Connection = myConnection
    '            myConnection.Open()
    '            myCommand1.ExecuteNonQuery()
    '            myCommand1.Connection.Close()
    '            'myConnection = Nothing
    '            myConnection.Close()
    '        Catch e_module1_679 As Exception
    '            Dim sSource As String = "AP_DENIS"
    '            Dim sLog As String = "Applo"
    '            EventLog.WriteEntry(sSource, " e_module1_679 --> " & e_module1_679.Message, EventLogEntryType.Information, 45)
    '        End Try

    '    End Sub

    Sub Sub_DeleteRecordsPageCounter(ByVal s_Query)

        Try
            Dim myConnection As New OleDbConnection(ConnStringDNA())
            Dim myInsertQuery As String = "Delete from PAGECOUNTER where sQuery='" & s_Query & "'"

            Dim myCommand1 As New OleDbCommand(myInsertQuery)
            myCommand1.Connection = myConnection
            myConnection.Open()
            myCommand1.ExecuteNonQuery()
            myCommand1.Connection.Close()
            'myConnection = Nothing
            myConnection.Close()
        Catch e_module1_645 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_645 --> " & e_module1_645.Message, EventLogEntryType.Information, 45)

        End Try

        'Try
        '    Dim jro As JRO.JetEngine
        '    jro = New JRO.JetEngine

        '    jro.CompactDatabase("Provider=Microsoft.Jet.OLEDB.4.0;Data source=" & sAppPath() & "DNA\DNA.mdb", "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sAppPath() & "DNA\DNA1.mdb;Jet OLEDB:Engine Type=5")

        '    If File.Exists(sAppPath() & "DNA\DNA.mdb") Then
        '        File.Move(sAppPath() & "DNA\DNA.mdb", sAppPath() & "DNA\DNA2.mdb")
        '        File.Move(sAppPath() & "DNA\DNA1.mdb", sAppPath() & "DNA\DNA.mdb")
        '        If File.Exists(sAppPath() & "DNA\DNA.mdb") And File.Exists(sAppPath() & "DNA\DNA2.mdb") Then
        '            File.Delete(sAppPath() & "DNA\DNA2.mdb")
        '        End If
        '    End If

        'Catch e_module1_673 As Exception

        '    Dim sSource As String = "AP_DENIS"
        '    Dim sLog As String = "Applo"
        '    EventLog.WriteEntry(sSource, " lookforfiles.e_module1_673 --> " & e_module1_673.Message, EventLogEntryType.Information, 45)

        'End Try

    End Sub


    Function Func_CheckMaxDNSToInsert(ByVal shortURL, ByVal sDatabaseName) As String

        Try
            Dim myConnString As String
            Dim sKeepString As String
            Dim s_se As String = "se"
            Dim i_countURL As Int32 = 0
            Dim s_shortUrl As String = Func_CleanUrl(shortURL)


            Dim mySelectQuery As String = "SELECT Count(*) as CountDNS FROM DNSFOUND where ShortURL ='" & s_shortUrl & "' and sCurrentQuery='" & sDatabaseName & "'"
            Dim myConnection As New OleDbConnection(ConnStringURLDNS(sDatabaseName))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                i_countURL = (myReader("CountDNS"))
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            'myConnection = Nothing
            myConnection.Close()

            If i_countURL > i_p_MaxPage Then
                Func_CheckMaxDNSToInsert = "stop"
            ElseIf shortURL = "" Or sDatabaseName = "" Then
                Func_CheckMaxDNSToInsert = "stop"
            Else
                Func_CheckMaxDNSToInsert = "go"
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " module1.1021 --> " & ex.Message, EventLogEntryType.Information, 45)
        End Try
    End Function


    Function F_CanMakePlace(ByVal sUrl As String, ByVal sDatabaseName As String, ByVal i_goodSpot As Int64)
        Try

            Dim myConnString As String
            Dim sKeepString As String
            Dim s_se As String = "se"
            Dim i_goodSpotLowest As Int64
            Dim s_shortUrl As String = Func_CleanUrl(sUrl)

            Dim mySelectQuery As String = "SELECT URLSite, iGoodSpot  FROM DNSFOUND where ShortURL ='" & s_shortUrl & "' ORDER BY iGoodSpot"
            Dim myConnection As New OleDbConnection(ConnStringURLDNS(sDatabaseName))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                i_goodSpotLowest = (myReader("iGoodSpot"))
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            If i_goodSpot < i_goodSpotLowest Then
                F_CanMakePlace = "stop"
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " module1.1034 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Function Sub_CheckURLInsertBySearchTools(ByVal s_SearchToolName, ByVal sDatabaseName)

        Try
            Dim myConnString As String
            Dim sKeepString As String
            Dim s_se As String = "se"
            Dim i_countURL As Int32 = 0
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"

            Dim mySelectQuery As String = "SELECT Count(*) as CountDNS FROM URLFound"
            Dim myConnection As New OleDbConnection(ConnStringURL(sDatabaseName))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                i_countURL = (myReader("CountDNS"))
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            'myConnection = Nothing
            myConnection.Close()

            If i_countURL > 0 Then
                'EventLog.WriteEntry(sSource, " e_module1_760 The search tool " & s_SearchToolName & " give " & i_countURL & " seeds ", EventLogEntryType.Information, 100)
            Else
                'EventLog.WriteEntry(sSource, " e_module1_760 This search tool doesn't give any answers! --> " & s_SearchToolName, EventLogEntryType.Information, 100)
            End If

        Catch e_module1_716 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.e_module1_716 --> " & e_module1_716.Message, EventLogEntryType.Information, 45)
            'MsgBox("The error (e_LookForFiles_174) is " & e_LookForFiles_174.Message)

        End Try
    End Function


    Sub Sub_InsertQueryFoundByKeywords(ByVal s_SearchToolName, ByVal s_p_datName)

        Try
            Dim myConnString As String
            Dim sKeepString As String
            Dim s_se As String = "se"
            Dim i_countURL As Int32 = 0
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"

            Dim mySelectQuery As String = "SELECT Count(*) as CountDNS FROM URLFound"
            Dim myConnection As New OleDbConnection(ConnStringURL(s_p_datName))
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

            myConnection.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

            While myReader.Read()
                i_countURL = (myReader("CountDNS"))
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            'myConnection = Nothing
            myConnection.Close()

            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myConnectionString As String = "UPDATE SUBQUERY SET URLFound=" & (i_countURL) & " where Query = '" & (s_p_datName) & "' and SubQuery ='" & (s_p_keywords) & "' "

            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = myConnection1
            myConnection1.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            ''myConnection1 = Nothing
            myConnection1.Close()

            If i_countURL > 0 Then
                'EventLog.WriteEntry(sSource, " e_module1_826 all the search tools together " & s_SearchToolName & " give " & i_countURL & " seeds ", EventLogEntryType.Information, 100)
            Else
                'EventLog.WriteEntry(sSource, " e_module1_828 all the search tools gave nothing, no answers! --> " & s_SearchToolName, EventLogEntryType.Information, 100)
            End If

        Catch e_module1_831 As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_module1_831 --> " & e_module1_831.Message, EventLogEntryType.Information, 45)

        End Try

    End Sub

    '    Sub InsertMemo(ByVal i_id, ByVal s_urlToFind, ByVal sSendstrURI)

    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"

    '        Try
    '            Dim s_title, s_description As String

    '            Dim myConnection As New OleDbConnection(ConnStringURL(sSendstrURI))
    '            Dim myInsertQuery As String = "INSERT INTO MEMO (ID,Title,Description)" _
    '            & "Values(" & i_id & ",'" & Trim(s_title) & "','" & (s_description) & "')"

    '            Dim myCommand As New OleDbCommand(myInsertQuery)
    '            myCommand.Connection = myConnection
    '            myConnection.Open()
    '            myCommand.ExecuteNonQuery()
    '            myCommand.Connection.Close()
    '            'myConnection = Nothing
    '            myConnection.Close()

    '        Catch e_module1_867 As Exception

    '            EventLog.WriteEntry(sSource, " e_module1_867 --> " & e_module1_867.Message, EventLogEntryType.Information, 45)

    '        End Try

    '    End Sub

    Sub Sub_SendHeartBeatUpdateDNS()

        Try
            Dim s_verb As String = "UpdateDNSFOUND"

            d_p_dateHeartBeat = Now()
            Call BroadCastDNSFOUND(s_verb, s_r_machineName)

            'Console.WriteLine(vbNewLine)
            'Console.WriteLine("***********************************************************************************")
            'Console.WriteLine(s_verb & s_r_machineName & " -- this machine sent an update to DNSFOUND!")
            'Console.WriteLine("***********************************************************************************")
            'Console.WriteLine(vbNewLine)
        Catch Module1_989 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.Module1_989 --> " & Module1_989.Message, EventLogEntryType.Information, 44)
        End Try

    End Sub

    Sub InsertDataInURLFOUND(ByVal sSendstrURI)

        Try
            If s_p_authority = "yes" Then
                '//---- Check in the URL is effectively in the authority database
                Dim b_theURLIsIn As Boolean = False
                Dim myConnString, s_pageContent As String
                Dim sKeepString As String = ""
                Dim sSendUri As String = sSendstrURI
                Dim s_yes As String = "y"
                Dim s_so As String = "so"
                Dim s_no As String = "n"
                Dim s_nothing As String = ""
                Dim s_stop As String = "stop"
                Dim i_idCounter As Int32
                Dim d_date As DateTime = Now.AddDays(-1)

                Dim myQuery As String = "SELECT TOP 1 IDCOUNTER, url, Datehour FROM URLFOUND where  GoodSpot <> -40004 and (Done = '" & s_no & "' or Done ='" & s_yes & "'  ) and (URL_PAGE = '" & s_no & "' or URL_PAGE = '" & s_nothing & "' or URL_PAGE = '" & s_stop & "') and DateHour >= #" & d_date & "# order by DateHour desc"
                Dim myConnection1 As New OleDbConnection(ConnStringURL(sSendUri))
                Dim myCommand1 As New OleDbCommand(myQuery, myConnection1)

                myConnection1.Open()

                Dim myReader As OleDbDataReader = myCommand1.ExecuteReader()
                While myReader.Read()
                    sKeepString = myReader("url")
                    i_idCounter = myReader("IDCounter")
                End While

                If Len(sKeepString) > 0 Then
                    b_theURLIsIn = True
                End If

                If Not (myReader.IsClosed) Then
                    myReader.Close()
                End If

                myConnection1.Close()

                '\\---- End of check 

                If Len(sKeepString) > 10 And Len(sKeepString) <= 250 Then

                    'If s_p_authority = "yes" Then
                    s_pageContent = f_FindAllTheHTMLFromTheURL(sKeepString, sSendstrURI)
                    s_pageContent = RemoveJSTag(s_pageContent)
                    s_pageContent = RemoveHtmlTag(s_pageContent)
                    s_pageContent = RemoveAccent(s_pageContent)
                    s_pageContent = keepASCIIOnly(s_pageContent)
                    'End If

                    If Len(s_pageContent) > 4 Then

                        Dim myInsertQuery As String = "UPDATE URLFOUND SET Done= '" & (s_so) & "',URL_PAGE = '" & s_pageContent & "', Machine = '" & CStr(Now()) & "' WHERE IDCounter =" & i_idCounter & ""
                        'Console.WriteLine("Module1.myInsertQuery.1033  -- " & myInsertQuery)
                        Dim myCommand As New OleDbCommand(myInsertQuery)
                        myCommand.Connection = myConnection1
                        myConnection1.Open()
                        myCommand.ExecuteNonQuery()
                        myCommand.Connection.Close()
                        myConnection1.Close()
                    Else
                        Dim s_sa As String = "sa"
                        Dim myInsertQuery As String = "UPDATE URLFOUND SET Done= '" & (s_sa) & "', Machine = '" & CStr(Now()) & "'  WHERE IDCounter =" & i_idCounter & ""

                        Dim myCommand As New OleDbCommand(myInsertQuery)
                        myCommand.Connection = myConnection1
                        myConnection1.Open()
                        myCommand.ExecuteNonQuery()
                        myCommand.Connection.Close()
                        myConnection1.Close()
                    End If
                End If

                If b_theURLIsIn = False Then
                    Dim s_sa As String = "sa"
                    Dim myInsertQuery As String = "UPDATE URLFOUND SET Done= '" & (s_no) & "', Machine = '" & CStr(Now()) & "'  WHERE Done= '" & (s_sa) & "' and DateHour >= #" & d_date & "#"

                    Dim myCommand As New OleDbCommand(myInsertQuery)
                    myCommand.Connection = myConnection1
                    myConnection1.Open()
                    myCommand.ExecuteNonQuery()
                    myCommand.Connection.Close()
                    myConnection1.Close()
                End If

            End If

        Catch Module1_1020 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " lookforfiles.Module1_1020 --> " & Module1_1020.Message, EventLogEntryType.Information, 44)
            'Console.WriteLine(" lookforfiles.Module1_1020 ******************--> " & Module1_1020.Message)
        End Try

    End Sub

    Public i_p_counterLoad As Int32

    'Sub SendLoadInfo()

    '    Try

    '        If s_p_authority <> "yes" Then

    '            Dim myConnString, mySelectQuery, s_query As String
    '            Dim sKeepString As String
    '            Dim date_InsertURL As DateTime = Now()
    '            Dim i_keepCumulative, i_totalCount, i_totalDNSCount, i_iterationEmpty, i_pageView As Int32
    '            Dim b_alreadyIn As Boolean = False
    '            Dim s_no As String = "n"

    '            If i_p_counterLoad = 0 Or i_p_counterLoad >= 1 Then

    '                mySelectQuery = "SELECT TOP 1 * FROM LOADBALANCING WHERE SENTTOAUTH = '" & s_no & "' ORDER BY ID DESC"
    '                Dim myConnection As New OleDbConnection(ConnStringDNA())
    '                Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)

    '                myConnection.Open()

    '                Dim myReader As OleDbDataReader = myCommand.ExecuteReader()

    '                While myReader.Read()
    '                    i_totalCount = myReader("TOTALCOUNT")
    '                    i_keepCumulative = myReader("CUMULATIVECOUNT")
    '                    i_totalDNSCount = myReader("TOTALDNSCOUNT")
    '                    i_iterationEmpty = myReader("ITERATIONEMPTY")
    '                    i_pageView = myReader("PAGEVIEW")
    '                    s_query = myReader("QUERY")
    '                    b_alreadyIn = True
    '                    Exit While
    '                End While

    '                If Not (myReader.IsClosed) Then
    '                    myReader.Close()
    '                End If

    '                myConnection.Close()

    '                If i_totalCount > 0 Or i_totalDNSCount > 0 Or i_iterationEmpty > 0 Then
    '                    SendHeartbeat(i_totalCount, i_totalDNSCount, i_iterationEmpty, i_pageView, s_query)
    '                End If

    '                i_p_counterLoad = 1

    '            Else
    '                i_p_counterLoad = i_p_counterLoad + 1
    '            End If

    '        End If


    '    Catch Module1_1133 As Exception
    '        Dim sSource As String = "AP_DENIS"
    '        Dim sLog As String = "Applo"
    '        EventLog.WriteEntry(sSource, " lookforfiles.Module1_1133 --> " & Module1_1133.Message, EventLogEntryType.Information, 44)
    '    End Try

    'End Sub

    Sub CheckDate()

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Try
            Dim s_datetime As DateTime = Now()
            Dim d_stopDate As DateTime = #1/1/2005 12:11:25 PM#
            If s_datetime > d_stopDate Then
                EventLog.WriteEntry(sSource, " lookforfiles.Module1_1163. Application Locked after 1/1/2004 please call Denis Jutras to de-lock it!!!!", EventLogEntryType.Information, 44444)
                System.Threading.Thread.Sleep(1000000)
            End If
        Catch Module1_1163 As Exception
            EventLog.WriteEntry(sSource, " lookforfiles.Module1_1163 --> " & Module1_1163.Message, EventLogEntryType.Information, 44)
        End Try
    End Sub

    Public i_p_idBefore As Int64

    Public Function f_TakeRacine_DNS_updated(ByVal sDNSName As String, ByVal sFoundRacine As String, ByVal Verb As String) As Object



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
        Dim s_KeepDate2 As DateTime = Now()
        Dim d_dateHour As Date
        Dim d_dateRatingDone As Date
        Dim b_go As Boolean = True
        Dim s_done As String
        Dim b_dnsFoundEmpty As Boolean = True
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim myConnection As New OleDbConnection(ConnStringURLDNS(sSendUri))
        Dim myReader As OleDbDataReader

        Try

            mySelectQuery = "SELECT top 1 ID,URLSite,Done,iGoodSpot,DateHour FROM DNSFOUND where iGoodSpot<>-40004 and sCurrentQuery='" & sDNSName & "' AND Done ='" & Trim(SUrl) & "' and URLSite <> '" & Trim(s_NA) & "' order by iGoodSpot desc,DateHour"

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

                d_p_DateTimeT = myReader("DateHour")

                If Len(sKeepString) > 4 Then
                    f_TakeRacine_DNS_updated = sKeepString
                Else
                    f_TakeRacine_DNS_updated = "stop"
                End If

                Call f_TakeRacine_DNS_updatedFU(f_TakeRacine_DNS_updated)

                Exit While
            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

            If b_dnsFoundEmpty Then

                i_p_countTimeAfterDNSFOUNDEmpty = i_p_countTimeAfterDNSFOUNDEmpty + 1
                'Console.WriteLine(i_p_countTimeAfterDNSFOUNDEmpty & "  -- lookforfiles.i_p_countTimeAfterDNSFOUNDEmpty.257")

                If i_p_countTimeAfterDNSFOUNDEmpty >= 1 Then
                    i_p_countTimeAfterDNSFOUNDEmpty = 0
                    'Call Sub_SendHeartBeatUpdateDNS()
                    Call Sub_checkDNSNotDone(sSendUri)
                    Call Sub_DeleteRecordsScratchPad()
                    Call sub_restartDNSFound(sSendUri)
                    Sub_DeleteRecordsPageCounter(sSendUri)
                End If

            End If

            Dim s_yes As String = "y"
            Dim objConn As New OleDbConnection(ConnStringURLDNS(sSendUri))
            Dim myConnectionString As String = "UPDATE DNSFOUND SET Done= '" & (s_yes) & "', Machine = '" & s_KeepDate2 & "' WHERE ID =" & i_ID & ""
            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = objConn
            objConn.Open()
            myCommand2.ExecuteNonQuery()
            objConn.Close()

            If i_ID = i_p_idBefore Then
                Try
                    Dim myInsertQuery As String = "Delete from DNSFOUND where ID =" & i_ID & ""
                    Dim myCommand1 As New OleDbCommand(myInsertQuery)
                    myCommand1.Connection = myConnection
                    myConnection.Open()
                    myCommand1.ExecuteNonQuery()
                    myCommand1.Connection.Close()
                    myConnection.Close()
                    Dim a, b, c As Int64
                    b = 1
                    c = 0
                    a = b / c
                Catch eOx As Exception
                    'CompacAccess1(sSendUri)
                    EventLog.WriteEntry(sSource, "Module1.1460 --> " & eOx.Message, EventLogEntryType.Information, 44)
                End Try
            End If

            i_p_idBefore = i_ID

        Catch eOx As Exception
            EventLog.WriteEntry(sSource, "Module1.1467 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Public Function f_TakeRacine_DNS_ID(ByVal sDNSName As String, ByVal sFoundRacine As String, ByVal Verb As String) As Int64

        Dim s_Verb As String = Verb
        Dim myConnString As String
        Dim sKeepString As String
        Dim sSendUri As String = sDNSName
        Dim s_NA = "n/a"
        Dim SUrl As String = "n"
        Dim i_ID As Int64
        Dim mySelectQuery As String
        Dim i_GoodSpot As Integer
        Dim s_KeepDate As String = Now()
        Dim d_dateHour As Date
        Dim d_dateRatingDone As Date
        Dim b_go As Boolean = True
        Dim s_done As String
        Dim b_dnsFoundEmpty As Boolean = True
        Try

            Dim myConnection As New OleDbConnection(ConnStringURLDNS(sSendUri))
            Dim myReader As OleDbDataReader

            mySelectQuery = "SELECT ID,URLSite,iGoodSpot,Done,DateHour FROM DNSFOUND where iGoodSpot <> -40004 and URLSite = '" & sFoundRacine & "' and sCurrentQuery='" & sDNSName & "'"

            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection)
            Dim o_object As Object
            Dim s_RetrieveIDDateHour As Array

            myConnection.Open()

            myReader = myCommand.ExecuteReader()

            While myReader.Read()

                sKeepString = sKeepString & (myReader("URLSite"))
                If IsNumeric(myReader("ID")) Then
                    i_ID = myReader("ID")
                Else
                    i_ID = 0
                End If
                i_GoodSpot = myReader("iGoodSpot")
                i_p_SendGoodSpotOverToUrlFound = myReader("iGoodSpot")
                s_done = myReader("Done")
                d_dateHour = myReader("DateHour")
                i_p_SendDateHourOverToUrlFound = myReader("DateHour")

                If s_Verb = "s_Racine" Then
                    d_p_DateTimeT = myReader("DateHour")
                End If

                f_TakeRacine_DNS_ID = i_ID

            End While

            If Not (myReader.IsClosed) Then
                myReader.Close()
            End If

            myConnection.Close()

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

            Dim objConn As New OleDbConnection(ConnStringURLDNS(sSendUri))
            Dim myConnectionString As String = "UPDATE DNSFOUND SET iGoodSpot='" & (i_GoodSpot) & "' WHERE ID =" & i_ID & ""

            Dim myCommand2 As New OleDbCommand(myConnectionString)
            myCommand2.Connection = objConn
            objConn.Open()
            myCommand2.ExecuteNonQuery()
            myCommand2.Connection.Close()
            objConn.Close()

        Catch ex As Exception

            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " module1.1542 --> " & ex.Message, EventLogEntryType.Information, 44)
            EventLog.WriteEntry(sSource, " module1.1543 sSendUri ---- mySelectQuery--> " & sSendUri & " ---- " & mySelectQuery, EventLogEntryType.Information, 44)

        End Try

    End Function

End Module

