



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


Module StartMeUp


    Public b_p_showS As Boolean = False
    Public i_p_lenghtMinimumFile As Int32 = 25000
    Public Form As New Form1

    Sub Main()


        'Call sub_restartDNSFound("trump")

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim sEvent As String

        Try

            Dim strt As String = System.Reflection.Assembly.GetExecutingAssembly().Location

            'Dim hostaddr As IPAddress
            'Dim strhost As String

            'Console.WriteLine("12/26/201")

            Dim YourIP As IPAddress = Dns.GetHostEntry("yahoo.com").AddressList(0)

            'Console.WriteLine("story.news.yahoo.com --> " & YourIP.ToString)



            Form1.TextBox1.Text = "build 1.18.11.03"

            ' Form1.ShowDialog()

            'Console.WriteLine("build 1.18.11.03")
            'Console.WriteLine(vbNewLine)
            'Console.WriteLine(System.Environment.MachineName())


            'System.Threading.Thread.Sleep(2000)

            'Call StartThread_t_ReceiveDNSFound()

            'Dim s_currentThreadName As String = Thread.CurrentThread.Name()
            'Thread.CurrentThread.Name = "Main_Thread"
            'Thread.CurrentThread.Priority = Normal

            'Call KillAndStartWatcher()

            If Form1.CheckSeed2.Checked = True Then
                Call FindSeeds1()
            End If

            Call StartTheScanningProcessFromConfiguration()

            sEvent = "aloha! -- " & CStr(Now())

            If Not EventLog.SourceExists(sSource) Then
                EventLog.CreateEventSource(sSource, sLog)
            End If

            'EventLog.WriteEntry(sSource, sEvent)

        Catch StartMeUp_41 As Exception

            'EventLog.WriteEntry(sSource, " StartMeUp_41 --> " & StartMeUp_41.Message, EventLogEntryType.Information, 44)

        End Try
    End Sub

    '-------------------------------------------------------------------------
    '-------------------------------------------------------------------------

    Public Sub fFindWhatToLookForFU()

        'Call sub_restartDNSFound("trump")

        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim sEvent As String

        Try

            Dim strt As String = System.Reflection.Assembly.GetExecutingAssembly().Location

            'Dim hostaddr As IPAddress
            'Dim strhost As String

            'Console.WriteLine("12/26/201")

            Dim YourIP As IPAddress = Dns.GetHostEntry("yahoo.com").AddressList(0)

            'Console.WriteLine("story.news.yahoo.com --> " & YourIP.ToString)


            Form1.TextBox1.Text = "build 1.18.11.03"
            Form1.Refresh()

            ' Form1.ShowDialog()

            'Console.WriteLine("build 1.18.11.03")
            'Console.WriteLine(vbNewLine)
            'Console.WriteLine(System.Environment.MachineName())


            'System.Threading.Thread.Sleep(2000)

            'Call StartThread_t_ReceiveDNSFound()

            'Dim s_currentThreadName As String = Thread.CurrentThread.Name()
            'Thread.CurrentThread.Name = "Main_Thread"
            'Thread.CurrentThread.Priority = Normal

            'Call KillAndStartWatcher()

            If Form1.CheckSeed2.Checked = True Then
                Call FindSeeds1()
            End If

            Call StartTheScanningProcessFromConfiguration()

            sEvent = "aloha! -- " & CStr(Now())

            If Not EventLog.SourceExists(sSource) Then
                EventLog.CreateEventSource(sSource, sLog)
            End If

            'EventLog.WriteEntry(sSource, sEvent)

        Catch fFindWhatToLookForFU_41 As Exception

            'EventLog.WriteEntry(sSource, " StartMeUp_41 --> " & StartMeUp_41.Message, EventLogEntryType.Information, 44)

        End Try
    End Sub

    Public Sub SelectFromDNSFoundFU()

    End Sub

    Public Function f_TakeRacine_DNS_updatedFU(f_TakeRacine_DNS_updated As String) As String

    End Function

    Public Sub StartScanningFU()

    End Sub

    Public Sub f_FindUrlToScanFU()

    End Sub

    Public Sub f_FetchRacineFU()

    End Sub

    Public Function InsertIntoSratchPadFU(s_myInsertQuery As String) As String

    End Function

    Public Function CheckDocumentInScratchpadFU(s_UrlFromScratchpad As String) As String

    End Function

    Public Function f_FindAllTheHTMLFromTheURLFU(s_html As String) As String

    End Function

    Public Sub ShowFunctionFU(ByVal s_function, ByVal i_lineNumber, ByVal b_showDebug)

    End Sub

    Public Function f_GetHtmlFromURIFU(s_GotHtml As String) As String

    End Function

    Public Function RemovePointBecauseOf404FU(myConnectionString As String) As String

    End Function

    Public Function InsertURLFOUNDFU(myInsertQuery As String) As String

    End Function


    '-------------------------------------------------------------------------
    '-------------------------------------------------------------------------


    Dim b_p_notfirstTime As Boolean = False

    Sub KillAndStartWatcher()
        Try
            Dim b_continue As Boolean = True
            Dim i_counter, i_countKill As Int16
            Dim KillWatcher As Boolean
            Dim b_continueToKill As Boolean = True

            'Console.WriteLine("Try to kill process Watcher!")
            'Console.WriteLine(vbNewLine)

            Dim myProcesses() As Process

            Do While b_continueToKill

                myProcesses = Process.GetProcessesByName("Watcher")

                If UBound(myProcesses) > 0 Then

                    myProcesses(0).Kill()
                    'Console.WriteLine("Process Watcher is kill!" & i_countKill)

                    If myProcesses(0).Responding Then
                        myProcesses(0).CloseMainWindow()
                    End If

                    If Not (myProcesses(0).Responding) Then
                        KillWatcher = True
                    End If
                    If i_countKill > 200 Then
                        Exit Do
                    End If

                Else
                    b_continueToKill = False
                    If b_p_notfirstTime = False Then
                        KillWatcher = True
                    End If
                    'Console.WriteLine("StartMeUp.Process Watcher is already closed, please check problem.86!")
                    'Console.WriteLine(vbNewLine)
                End If
            Loop



            i_counter = 0

            If KillWatcher Then
                b_p_notfirstTime = True

                Do While b_continue

                    Dim myProcess As New Process
                    myProcess.StartInfo.FileName = "Watcher.exe"
                    'Console.WriteLine("StartMeUp.87 -- Start Watcher.exe!")
                    'Console.WriteLine("")
                    myProcess.Start()
                    'Console.WriteLine("StartMeUp.87 -- Watcher.exe started with success!")
                    'Console.WriteLine("")
                    If myProcess.Responding Then
                        Exit Do
                    End If

                    If i_counter >= 100 Then
                        Exit Do
                    End If
                    i_counter = i_counter + 1

                Loop
            End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " StartMeUp.94 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try


    End Sub

End Module

Friend Class FormsName
End Class
