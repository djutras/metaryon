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

Module Mod1

    Public Sub VeryMain()
        'The 2 event handlers 
        'add an unhandled exceptions handler 
        Dim currentDomain As AppDomain = AppDomain.CurrentDomain
        'for regular unhandled stuff 
        AddHandler currentDomain.UnhandledException, AddressOf MYExceptionHandler
        'for threads behind forms 
        ''AddHandler Application.ThreadException, AddressOf MYThreadHandler

    End Sub

    Private Sub MYExceptionHandler(ByVal sender As Object, ByVal e As UnhandledExceptionEventArgs)
        Dim EX As Exception
        EX = e.ExceptionObject
        'Console.WriteLine(EX.StackTrace)
    End Sub

    Private Sub MYThreadHandler(ByVal sender As Object, ByVal e As Threading.ThreadExceptionEventArgs)
        'Console.WriteLine(e.Exception.StackTrace)
    End Sub


    Function VerifyFirstLetters(ByVal s_shortURL) As Boolean
        Try
            Dim s_allLetters, s_firstLetter, s_secondLetter As String
            Dim i_fLetter As Integer
            Dim i_sLetter As Integer
            s_allLetters = FindFirstTwoLettersOfSomething(LCase(s_shortURL))

            s_firstLetter = Mid(s_allLetters, 1, 1)
            s_secondLetter = Mid(s_allLetters, 2, 1)
            i_fLetter = GiveIntFromLetter(s_firstLetter)
            i_sLetter = GiveIntFromLetter(s_secondLetter)

            If ((i_fLetter >= 1 And i_fLetter <= 26) And (i_sLetter >= 1 And i_sLetter <= 26)) Then
                VerifyFirstLetters = True
            ElseIf (IsNumeric(s_firstLetter) And ((i_sLetter >= 1) And (i_sLetter <= 26))) Then
                VerifyFirstLetters = True
            ElseIf (((i_fLetter >= 1 And i_fLetter <= 26)) And IsNumeric(s_secondLetter)) Then
                VerifyFirstLetters = True
            ElseIf IsNumeric(s_firstLetter) And IsNumeric(s_secondLetter) Then
                VerifyFirstLetters = True
            Else
                VerifyFirstLetters = False
            End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " codefile1.63 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Function VerifyShortURL(ByVal s_shortURL) As Boolean
        Try
            Dim s_allLetters, s_firstLetter, s_secondLetter As String
            Dim i_positionOfDot, i_positionOfBracketFirst, i_positionOfBracketSecond As Integer

            i_positionOfDot = InStr(s_shortURL, ".", CompareMethod.Text)
            i_positionOfBracketFirst = InStr(s_shortURL, ">", CompareMethod.Text)
            i_positionOfBracketSecond = InStr(s_shortURL, "<", CompareMethod.Text)

            If i_positionOfDot > 0 And (i_positionOfBracketFirst = 0 And i_positionOfBracketSecond = 0) Then
                VerifyShortURL = True
            Else
                VerifyShortURL = False
            End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " codefile1.86 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Public b_p_dateFoundToday As Boolean

    Function GetDateFromContent(ByVal s_pageContent) As DateTime
        Try
            Dim d_timeNow As DateTime = Now()
            Dim d_date As DateTime
            Dim i_year As Integer = DateTime.Now.Year
            Dim i_month As Integer = DateTime.Now.Month
            Dim i_day As Integer = DateTime.Now.Day
            Dim s_monthEnglish As String = GetMonthIntToStringEnglish(i_month)
            Dim s_monthFrench As String = GetMonthIntToStringFrench(i_month)
            Dim b_foundRightYear, b_foundRightMonth, b_foundRightDay As Boolean
            Dim i_positionOfYear, i_positionOfYearStartAgain, i_positionOfMonthNumber, i_positionOfMonthEng, i_positionOfMonthFrench, i_positionOfDay As Int64
            Dim s_keepString As String
            Dim b_continue As Boolean = True
            Dim i_hoursToRemove As Integer = CInt(i_p_Duration * 0.75)
            Dim s_buildKeepString, s_buildKeepStringFirst, s_buildKeepSecond As String

            i_positionOfYearStartAgain = 1

            Do While b_continue

                i_positionOfYear = InStr(i_positionOfYearStartAgain, s_pageContent, CStr(i_year), CompareMethod.Text)

                If i_positionOfYear = 0 Then
                    GetDateFromContent = Now()
                    Exit Do
                End If

                s_keepString = Mid(s_pageContent, i_positionOfYear - 30, 60)
                s_keepString = Replace(s_keepString, CStr(i_year), " ")
                i_positionOfMonthNumber = InStr(s_keepString, CStr(i_month), CompareMethod.Text)
                i_positionOfMonthFrench = InStr(LCase(s_keepString), s_monthFrench, CompareMethod.Text)
                i_positionOfMonthEng = InStr(LCase(s_keepString), s_monthEnglish, CompareMethod.Text)

                If i_positionOfMonthNumber > 0 Or i_positionOfMonthFrench > 0 Or i_positionOfMonthEng > 0 Then
                    b_foundRightMonth = True
                End If

                If b_foundRightMonth And i_positionOfMonthNumber > 0 Then
                    Dim i_positionOfMonth As Integer = InStr(s_keepString, CStr(i_month), CompareMethod.Text)
                    Dim s_firstPart, s_secondPart As String
                    If i_positionOfMonth > 1 Then
                        s_firstPart = Mid(s_keepString, 1, i_positionOfMonth - 1)
                        s_secondPart = Mid(s_keepString, i_positionOfMonth + (Len(CStr(i_month)) + 1), Len(s_keepString) - (i_positionOfMonth + (Len(CStr(i_month)))))
                        s_keepString = s_firstPart & s_secondPart
                    End If
                End If

                i_positionOfMonthNumber = 0
                i_positionOfMonthFrench = 0
                i_positionOfMonthEng = 0

                If b_foundRightMonth Then
                    i_positionOfDay = InStr(s_keepString, CStr(i_day), CompareMethod.Text)
                    b_foundRightMonth = False
                End If

                If i_positionOfDay > 0 Then
                    b_foundRightDay = True
                    b_p_dateFoundToday = True
                    GetDateFromContent = Now()
                    i_positionOfDay = 0
                    Exit Function
                End If

                i_positionOfYearStartAgain = InStr(i_positionOfYear + 4, s_pageContent, CStr(i_year), CompareMethod.Text)

                If i_positionOfYearStartAgain = 0 Then
                    GetDateFromContent = Now()
                    Exit Do
                End If

            Loop

            Dim i_googleSearch As Int64 = InStr(s_pageContent, "Google Search", CompareMethod.Text)
            Dim i_dateCompare As DateTime

            If i_googleSearch > 0 Then
                GetDateFromContent = Now()
                b_p_dateFoundToday = True
            Else
                GetDateFromContent = Now()
            End If

        Catch eOx As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " codefile1.154 --> " & eOx.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function GetMonthIntToStringEnglish(ByVal i_month) As String
        Try
            Select Case i_month

                Case 1
                    GetMonthIntToStringEnglish = "jan"
                Case 2
                    GetMonthIntToStringEnglish = "feb"
                Case 3
                    GetMonthIntToStringEnglish = "mar"
                Case 4
                    GetMonthIntToStringEnglish = "apr"
                Case 5
                    GetMonthIntToStringEnglish = "may"
                Case 6
                    GetMonthIntToStringEnglish = "jun"
                Case 7
                    GetMonthIntToStringEnglish = "jul"
                Case 8
                    GetMonthIntToStringEnglish = "aug"
                Case 9
                    GetMonthIntToStringEnglish = "sep"
                Case 10
                    GetMonthIntToStringEnglish = "oct"
                Case 11
                    GetMonthIntToStringEnglish = "nov"
                Case 12
                    GetMonthIntToStringEnglish = "dec"

            End Select

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " codefile1.192 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function GetMonthIntToStringFrench(ByVal i_month) As String
        Try
            Select Case i_month

                Case 1
                    GetMonthIntToStringFrench = "jan"
                Case 2
                    GetMonthIntToStringFrench = "fev"
                Case 3
                    GetMonthIntToStringFrench = "mar"
                Case 4
                    GetMonthIntToStringFrench = "avr"
                Case 5
                    GetMonthIntToStringFrench = "mai"
                Case 6
                    GetMonthIntToStringFrench = "jun"
                Case 7
                    GetMonthIntToStringFrench = "jul"
                Case 8
                    GetMonthIntToStringFrench = "aou"
                Case 9
                    GetMonthIntToStringFrench = "sep"
                Case 10
                    GetMonthIntToStringFrench = "oct"
                Case 11
                    GetMonthIntToStringFrench = "nov"
                Case 12
                    GetMonthIntToStringFrench = "dec"

            End Select

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " codefile1.230 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function FixFirstSecondLetter(ByVal s_toCheck) As String

        Try
            Dim s_firstLetter, s_secondLetter As String
            Dim i_fLetter, i_sLetter As Int32

            s_firstLetter = Trim(Mid(Trim(s_toCheck), 1, 1))
            s_secondLetter = Trim(Mid(Trim(s_toCheck), 2, 1))

            i_fLetter = GiveIntFromLetter(s_firstLetter)
            i_sLetter = GiveIntFromLetter(s_secondLetter)

            If i_fLetter > 0 And i_fLetter <= 26 Then
                FixFirstSecondLetter = s_firstLetter
            Else
                FixFirstSecondLetter = "a"
            End If

            If i_sLetter > 0 And i_sLetter <= 26 Then
                FixFirstSecondLetter = FixFirstSecondLetter & s_secondLetter
            Else
                FixFirstSecondLetter = FixFirstSecondLetter & "a"
            End If

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " codefile1.261 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try

    End Function

    Function CheckIfQueryAlreadyIn(ByVal s_querySearch) As Boolean
        Try
            Dim s_keepHostName As String
            Dim b_nameIsIn As Boolean = False

            Dim mySelectQuery, s_Website, s As String

            mySelectQuery = "SELECT QuerySearch FROM Query where QuerySearch ='" & s_querySearch & "'"

            Dim myConnection1 As New OleDbConnection(ConnStringDNA())
            Dim myCommand As New OleDbCommand(mySelectQuery, myConnection1)

            myConnection1.Open()

            Dim myReader As OleDbDataReader = myCommand.ExecuteReader()
            Dim Index As Integer

            While myReader.Read()
                s_keepHostName = myReader("QuerySearch").ToString
                Exit While
            End While

            If Len(s_keepHostName) > 0 Then
                CheckIfQueryAlreadyIn = True
            Else
                CheckIfQueryAlreadyIn = False
            End If

            myReader.Close()
            myConnection1.Close()

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " codefile1.273 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function




End Module

