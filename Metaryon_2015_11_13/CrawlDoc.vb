
Imports System.Text
Imports System.IO
Imports System.IO.Stream
Imports System.IO.FileStream


Public Module CrawlDoc

    Public Function RemoveAccent(ByVal sFileToCrawl) As String

        Dim s_FileToCrawl As String = sFileToCrawl

        s_FileToCrawl = Replace(s_FileToCrawl, "nbsp;", " ")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#233;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#234;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#235;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#238;", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#239;", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#244;", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#251;", "u")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#231;", "ç")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#232;", "è")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#224;", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#226;", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#201;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#247;", "÷")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#215;", "×")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#39;", "'")

        s_FileToCrawl = Replace(s_FileToCrawl, "&#40;", "(")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#41;", ")")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#42;", "*")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#43;", "+")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#44;", ",")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#46;", ".")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#34;", """")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#38;", "&")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#36;", "$")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#37;", "%")

        s_FileToCrawl = Replace(s_FileToCrawl, "&#47;", "/ ")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#58;", ":")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#59;", ";")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#60;", "<")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#61;", "=")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#62;", ">")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#92;", "\")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#64;", "@")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#95;", "-")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#124;", "|")

        s_FileToCrawl = Replace(s_FileToCrawl, "&deg;", "°")
        s_FileToCrawl = Replace(s_FileToCrawl, "&eacute;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ecirc;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&euml;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&icirc;", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "&iuml;", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ocirc;", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ouml;", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ucirc;", "u")
        s_FileToCrawl = Replace(s_FileToCrawl, "&divide;", "÷")
        s_FileToCrawl = Replace(s_FileToCrawl, "&aelig;", "æ")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ccedil;", "c")
        s_FileToCrawl = Replace(s_FileToCrawl, "&egrave;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&agrave;", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "&Eacute;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&Ecirc;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&Ccedil;", "ç")
        s_FileToCrawl = Replace(s_FileToCrawl, "&copy;", "©")
        s_FileToCrawl = Replace(s_FileToCrawl, "&pound;", "£")
        s_FileToCrawl = Replace(s_FileToCrawl, "&yen;", "¥")
        s_FileToCrawl = Replace(s_FileToCrawl, "&reg;", "®")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ndash;", "–")
        s_FileToCrawl = Replace(s_FileToCrawl, "&mdash;", "—")
        s_FileToCrawl = Replace(s_FileToCrawl, "&cent;", "¢")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#64;", "@")

        s_FileToCrawl = Replace(s_FileToCrawl, "&frasl;", "/")
        s_FileToCrawl = Replace(s_FileToCrawl, "&lt;", "<")
        s_FileToCrawl = Replace(s_FileToCrawl, "&gt;", ">")

        s_FileToCrawl = Replace(s_FileToCrawl, "&quot;", """")
        s_FileToCrawl = Replace(s_FileToCrawl, "&amp;", "&")

        s_FileToCrawl = Replace(s_FileToCrawl, "&lsquo;", "‘")
        s_FileToCrawl = Replace(s_FileToCrawl, "&rsquo;", "’")

        s_FileToCrawl = Replace(s_FileToCrawl, "&trade;", "™")

        s_FileToCrawl = Replace(s_FileToCrawl, "é", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "è", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "ê", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "ë", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "ç", "c")
        s_FileToCrawl = Replace(s_FileToCrawl, "à", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "à", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "ä", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "â", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "å", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "ä", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "ï", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "î", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "ì", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "Ä", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "Å", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "É", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "ô", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "ö", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "ò", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "û", "u")
        s_FileToCrawl = Replace(s_FileToCrawl, "ù", "u")
        s_FileToCrawl = Replace(s_FileToCrawl, "ÿ", "y")
        s_FileToCrawl = Replace(s_FileToCrawl, "Ö", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "Ü", "u")

        RemoveAccent = s_FileToCrawl

    End Function


    Public Function ReplaceAccent(ByVal sFileToCrawl) As String

        Dim s_FileToCrawl As String = sFileToCrawl

        s_FileToCrawl = Replace(s_FileToCrawl, "&#233;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#234;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#235;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#238;", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#239;", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#244;", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#251;", "u")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#231;", "ç")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#232;", "è")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#224;", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#226;", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#201;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#247;", "÷")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#215;", "×")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#39;", "'")

        s_FileToCrawl = Replace(s_FileToCrawl, "&#40;", "(")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#41;", ")")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#42;", "*")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#43;", "+")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#44;", ",")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#46;", ".")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#34;", """")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#38;", "&")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#36;", "$")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#37;", "%")

        s_FileToCrawl = Replace(s_FileToCrawl, "&#47;", "/ ")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#58;", ":")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#59;", ";")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#60;", "<")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#61;", "=")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#62;", ">")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#92;", "\")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#64;", "@")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#95;", "-")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#124;", "|")

        s_FileToCrawl = Replace(s_FileToCrawl, "&deg;", "°")
        s_FileToCrawl = Replace(s_FileToCrawl, "&eacute;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ecirc;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&euml;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&icirc;", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "&iuml;", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ocirc;", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ouml;", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ucirc;", "u")
        s_FileToCrawl = Replace(s_FileToCrawl, "&divide;", "÷")
        s_FileToCrawl = Replace(s_FileToCrawl, "&aelig;", "æ")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ccedil;", "c")
        s_FileToCrawl = Replace(s_FileToCrawl, "&egrave;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&agrave;", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "&Eacute;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&Ecirc;", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "&Ccedil;", "ç")
        s_FileToCrawl = Replace(s_FileToCrawl, "&copy;", "©")
        s_FileToCrawl = Replace(s_FileToCrawl, "&pound;", "£")
        s_FileToCrawl = Replace(s_FileToCrawl, "&yen;", "¥")
        s_FileToCrawl = Replace(s_FileToCrawl, "&reg;", "®")
        s_FileToCrawl = Replace(s_FileToCrawl, "&ndash;", "–")
        s_FileToCrawl = Replace(s_FileToCrawl, "&mdash;", "—")
        s_FileToCrawl = Replace(s_FileToCrawl, "&cent;", "¢")
        s_FileToCrawl = Replace(s_FileToCrawl, "&#64;", "@")

        s_FileToCrawl = Replace(s_FileToCrawl, "&frasl;", "/")
        s_FileToCrawl = Replace(s_FileToCrawl, "&lt;", "<")
        s_FileToCrawl = Replace(s_FileToCrawl, "&gt;", ">")

        s_FileToCrawl = Replace(s_FileToCrawl, "&quot;", """")
        s_FileToCrawl = Replace(s_FileToCrawl, "&amp;", "&")

        s_FileToCrawl = Replace(s_FileToCrawl, "&lsquo;", "‘")
        s_FileToCrawl = Replace(s_FileToCrawl, "&rsquo;", "’")

        s_FileToCrawl = Replace(s_FileToCrawl, "&trade;", "™")

        s_FileToCrawl = Replace(s_FileToCrawl, "é", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "è", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "ê", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "ë", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "ç", "c")
        s_FileToCrawl = Replace(s_FileToCrawl, "à", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "à", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "ä", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "â", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "å", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "ä", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "ï", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "î", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "ì", "i")
        s_FileToCrawl = Replace(s_FileToCrawl, "Ä", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "Å", "a")
        s_FileToCrawl = Replace(s_FileToCrawl, "É", "e")
        s_FileToCrawl = Replace(s_FileToCrawl, "ô", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "ö", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "ò", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "û", "u")
        s_FileToCrawl = Replace(s_FileToCrawl, "ù", "u")
        s_FileToCrawl = Replace(s_FileToCrawl, "ÿ", "y")
        s_FileToCrawl = Replace(s_FileToCrawl, "Ö", "o")
        s_FileToCrawl = Replace(s_FileToCrawl, "Ü", "u")

        s_FileToCrawl = Replace(s_FileToCrawl, Chr(9), " ")
        s_FileToCrawl = Replace(s_FileToCrawl, Chr(10), " ")
        s_FileToCrawl = Replace(s_FileToCrawl, Chr(11), " ")
        s_FileToCrawl = Replace(s_FileToCrawl, Chr(13), " ")

        ReplaceAccent = LCase(s_FileToCrawl)

    End Function

    Dim b_p_notTheFirstTime As Boolean

    Public Function CheckForKeywords(ByVal sFileToCrawlFromScratchpad, ByVal sDatabaseName) As Boolean

        Dim s_FileToCrawl As String = sFileToCrawlFromScratchpad
        Dim s_DatabaseName As String = sDatabaseName
        Dim b_Continue As Boolean = True
        Dim s_KeywordsToFind As String
        Dim b_KeywordFind As Boolean = False
        Dim b_LookForOr As Boolean = False
        Dim i_CounterKeep As Integer = 0
        Dim i_counterAsSafeGuard As Int32

        Try
            If i_p_ExactPhrase = 1 Then
                CheckForKeywords = CheckExactPhrase(s_FileToCrawl, s_DatabaseName)
                Exit Function
            End If

            While b_Continue
                s_KeywordsToFind = f_KeywordsToFindA(s_DatabaseName)

                If b_LookForOr = True And LCase(s_KeywordsToFind) = "or" Then
                    b_LookForOr = False
                Else

                    If Len(s_p_DatabaseNameA) = 0 Then
                        i_CounterKeep = i_CounterKeep + 1
                    End If

                    If LCase(s_KeywordsToFind) = "or" And b_LookForOr = False And b_p_notTheFirstTime Then
                        CheckForKeywords = True
                        Exit While
                    ElseIf b_LookForOr = True Then
                        CheckForKeywords = False
                        Exit While
                    End If

                    If i_CounterKeep = 1 And b_LookForOr = True Then
                        CheckForKeywords = False
                        Exit While
                    End If

                    If f_CrawlFile(s_FileToCrawl, s_KeywordsToFind, sDatabaseName) = True Then
                        If b_LookForOr = False Then
                            b_LookForOr = False
                        End If
                    Else
                        b_LookForOr = True
                    End If

                    If Len(s_p_DatabaseNameA) = 0 And b_LookForOr = False Then
                        CheckForKeywords = True
                        Exit While
                    End If

                    If Len(s_p_DatabaseNameA) = 0 And b_LookForOr = True Then
                        CheckForKeywords = False
                        Exit While
                    End If

                End If

                b_p_notTheFirstTime = True

            End While

            s_p_DatabaseNameA = ""

        Catch e_crawl_161 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_crawl_161--> " & e_crawl_161.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Public Function CheckExactPhrase(ByVal sFileToCrawl, ByVal sDatabaseName) As Boolean
        Try
            Dim s_FileToCrawl As String = sFileToCrawl
            Dim s_DatabaseName As String = sDatabaseName
            Dim s_FoundExactPhrase As String
            Dim s_ExactPhraseToFind As String = sDatabaseName

            If InStr(s_FileToCrawl, sDatabaseName, CompareMethod.Text) > 0 Then
                CheckExactPhrase = True
            Else
                CheckExactPhrase = False
            End If
        Catch e_crawl_183 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_crawl_183--> " & e_crawl_183.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Public ReadOnly Property s_r_KeepKeywords() As String
        Get
            s_r_KeepKeywords = s_p_DatabaseName
        End Get
    End Property


    Public s_p_DatabaseName As String
    Public b_p_KeepKeywords As Boolean


    Function f_KeywordsToFind(ByVal sDatabaseName) As String

        Dim b_LookForOr As Boolean
        Dim s_DatabaseName As String
        Dim s_KeepKeywordsToSearch As String
        Dim s_KeywordsToFind As String
        Dim b_Continue As Boolean = True
        Dim i_KeepParOpen As Integer
        Dim i_RemoveKeywords As Integer
        Dim i_KeepParClose As Integer
        Dim s_KeepString As String
        Dim i_KeepLenght As Integer
        Dim s_ArrayKeywords() As String
        Dim s_RemoveKeywords As String
        Dim i_Ubound As Integer
        Dim b_KeepKeywords As Boolean = True
        Dim b_spaceExist As Boolean = True
        Dim i_doubleSpace As Integer

        Try

            If Len(Trim(s_r_KeepKeywords)) > 0 Then
                s_DatabaseName = s_r_KeepKeywords
            Else
                s_DatabaseName = sDatabaseName
            End If

            s_DatabaseName = Replace(s_DatabaseName, "(", "")
            s_DatabaseName = Replace(s_DatabaseName, ")", "")

            Do While b_spaceExist
                i_doubleSpace = InStr(s_DatabaseName, "  ", CompareMethod.Binary)
                If i_doubleSpace > 0 Then
                    s_DatabaseName = Replace(s_DatabaseName, "  ", " ")
                Else
                    b_spaceExist = False
                End If
            Loop

            s_ArrayKeywords = Split(Trim(s_DatabaseName))
            i_Ubound = UBound(s_ArrayKeywords)

            Do While b_Continue

                If UBound(s_ArrayKeywords) >= 0 Then
                    f_KeywordsToFind = s_ArrayKeywords(i_Ubound)
                    s_KeepString = s_ArrayKeywords(i_Ubound)
                    Exit Do
                Else

                    f_KeywordsToFind = "empty"
                    s_KeepString = "empty"
                    Exit Do

                End If
                i_Ubound = i_Ubound - 1

            Loop

            Dim i_KeepLenghtToCut As Integer = Len(Trim(s_DatabaseName)) - Len(Trim(s_KeepString))

            If i_KeepLenghtToCut > 0 Then
                s_KeepKeywordsToSearch = Mid(s_DatabaseName, 1, i_KeepLenghtToCut)
                s_p_DatabaseName = s_KeepKeywordsToSearch
            Else
                s_KeepKeywordsToSearch = ""
                s_p_DatabaseName = s_KeepKeywordsToSearch
            End If

        Catch e_crawl_283 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_crawl_283--> " & e_crawl_283.Message, EventLogEntryType.Information, 44)
        End Try

    End Function


    Public s_p_DatabaseNameA As String


    Function f_KeywordsToFindA(ByVal sDatabaseName) As String

        Dim b_LookForOr As Boolean
        Dim s_DatabaseName As String
        Dim s_KeepKeywordsToSearch As String
        Dim s_KeywordsToFind As String
        Dim b_Continue As Boolean = True
        Dim i_KeepParOpen As Integer
        Dim i_RemoveKeywords As Integer
        Dim i_KeepParClose As Integer
        Dim s_KeepString As String
        Dim i_KeepLenght As Integer
        Dim s_ArrayKeywords() As String
        Dim s_RemoveKeywords As String
        Dim i_Ubound As Integer
        Dim b_KeepKeywords As Boolean = True
        Dim b_spaceExist As Boolean = True
        Dim i_doubleSpace As Integer

        Try

            If Len(Trim(s_p_DatabaseNameA)) > 0 Then
                s_DatabaseName = s_p_DatabaseNameA
            Else
                s_DatabaseName = sDatabaseName
            End If

            s_DatabaseName = Replace(s_DatabaseName, "(", "")
            s_DatabaseName = Replace(s_DatabaseName, ")", "")

            Do While b_spaceExist
                i_doubleSpace = InStr(s_DatabaseName, "  ", CompareMethod.Binary)
                If i_doubleSpace > 0 Then
                    s_DatabaseName = Replace(s_DatabaseName, "  ", " ")
                Else
                    b_spaceExist = False
                End If
            Loop

            s_ArrayKeywords = Split(Trim(s_DatabaseName))
            i_Ubound = UBound(s_ArrayKeywords)

            Do While b_Continue

                If UBound(s_ArrayKeywords) >= 0 Then
                    f_KeywordsToFindA = s_ArrayKeywords(i_Ubound)
                    s_KeepString = s_ArrayKeywords(i_Ubound)
                    Exit Do
                Else

                    f_KeywordsToFindA = "empty"
                    s_KeepString = "empty"
                    Exit Do

                End If
                i_Ubound = i_Ubound - 1

            Loop

            Dim i_KeepLenghtToCut As Integer = Len(Trim(s_DatabaseName)) - Len(Trim(s_KeepString))

            If i_KeepLenghtToCut > 0 Then
                s_KeepKeywordsToSearch = Mid(s_DatabaseName, 1, i_KeepLenghtToCut)
                s_p_DatabaseNameA = s_KeepKeywordsToSearch
            Else
                s_KeepKeywordsToSearch = ""
                s_p_DatabaseNameA = s_KeepKeywordsToSearch
            End If

        Catch e_crawl_283 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_crawl_283--> " & e_crawl_283.Message, EventLogEntryType.Information, 44)
        End Try

    End Function



    Function F_CleanQuery(ByVal sDatabaseName) As String
        Try
            Dim s_DatabaseName As String = sDatabaseName

            s_DatabaseName = Replace(s_DatabaseName, " le ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " la ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " les ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " de ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " dans ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " des ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " du ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " un ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " une ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " à ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " l' ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " d' ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " s' ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " pour ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " sa ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " et ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " se ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " par ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " par ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " par ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " par ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " the ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " of ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " to ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " an ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " is ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " with ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " than ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " they ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " on ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " a ", " ")
            s_DatabaseName = Replace(s_DatabaseName, " and ", " ")

            F_CleanQuery = s_DatabaseName
        Catch e_crawl_332 As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " e_crawl_332--> " & e_crawl_332.Message, EventLogEntryType.Information, 44)
        End Try
    End Function


    Function f_CrawlFile(ByVal sFileToCrawl As String, ByVal sKeywordsToFind As String, ByVal s_databaseName As String) As Boolean
        Try
            Dim b_rightWords As Boolean
            Dim b_continue As Boolean = True

            If sFileToCrawl = "" Or sKeywordsToFind = "" Or s_databaseName = "" Then
                b_continue = False
            End If

            If b_continue Then

                If i_p_KeywordVisible = 1 Then
                    sFileToCrawl = RemoveHtmlTag(sFileToCrawl)
                End If

                sFileToCrawl = LCase(sFileToCrawl)

                Dim s_KeywordsToFind As String = ReplaceAccent(sKeywordsToFind)

                Dim b_KeywordInFile As Boolean = True

                If sFileToCrawl = "stop" Then
                    f_CrawlFile = False
                End If

                If InStr(Trim(s_KeywordsToFind), "-", CompareMethod.Text) = 1 Then
                    b_KeywordInFile = False
                    s_KeywordsToFind = Mid(s_KeywordsToFind, 2, Len(s_KeywordsToFind) - 1)
                End If

                If b_KeywordInFile Then
                    If InStr(sFileToCrawl, LCase(Trim(s_KeywordsToFind)), CompareMethod.Text) > 0 Then
                        b_rightWords = VerifyIfExactWord(sFileToCrawl, sKeywordsToFind, s_databaseName)
                        If b_rightWords Then
                            f_CrawlFile = True
                        Else
                            f_CrawlFile = False
                        End If
                    Else
                        f_CrawlFile = False
                    End If
                Else
                    If InStr(sFileToCrawl, LCase(Trim(s_KeywordsToFind)), CompareMethod.Text) > 0 Then
                        b_rightWords = VerifyIfExactWord(sFileToCrawl, sKeywordsToFind, s_databaseName)
                        If b_rightWords Then
                            f_CrawlFile = False
                        Else
                            f_CrawlFile = True
                        End If
                    Else
                        f_CrawlFile = True
                    End If
                End If
            Else
                f_CrawlFile = False
            End If
        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " CrawlDoc.508--> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

    Function VerifyIfExactWord(ByVal sFileToCrawl As String, ByVal sKeywordsToFind As String, ByVal s_databaseName As String) As Boolean
        Dim sSource As String = "AP_DENIS"
        Dim sLog As String = "Applo"
        Dim b_exactKeywords, b_continue As Boolean

        Try
            Dim b_findTheRightKeywords As Boolean
            Dim s_array() As String
            s_array = Split(s_databaseName)
            Dim i_counter As Integer
            Dim i_foundKeywords, i_open, i_close, i_lenDatabaseName As Integer

            Dim s_keywordsToCheck As String

            i_open = InStr(s_databaseName, "(", CompareMethod.Binary)
            i_close = InStr(s_databaseName, ")", CompareMethod.Binary)

            If i_open > 0 And i_close > 0 Then
                b_continue = True
            Else
                VerifyIfExactWord = True
                Exit Function
            End If

            If b_continue Then

                i_lenDatabaseName = Len(s_databaseName)

                If i_open = 1 And i_close = i_lenDatabaseName Then
                    b_exactKeywords = True
                End If

                If b_exactKeywords = False Then
                    For i_counter = 0 To UBound(s_array)
                        i_open = 0
                        i_close = 0
                        i_lenDatabaseName = 0

                        s_keywordsToCheck = s_array(i_counter)
                        i_foundKeywords = InStr(s_keywordsToCheck, sKeywordsToFind, CompareMethod.Text)
                        If i_foundKeywords > 0 Then
                            i_open = InStr(s_keywordsToCheck, "(", CompareMethod.Binary)
                            i_close = InStr(s_keywordsToCheck, ")", CompareMethod.Binary)
                            i_lenDatabaseName = Len(s_keywordsToCheck)

                            If (i_open = 1 And i_close = i_lenDatabaseName) Then
                                b_exactKeywords = True
                            End If
                        End If
                    Next

                End If
            End If

            If b_exactKeywords Then
                b_findTheRightKeywords = CheckLenghtOfKeywordsWithText(sFileToCrawl, sKeywordsToFind)
            Else
                VerifyIfExactWord = True
            End If

            If b_findTheRightKeywords Then
                VerifyIfExactWord = True
            Else
                VerifyIfExactWord = False
            End If


        Catch ex1 As Exception
            EventLog.WriteEntry(sSource, " CrawlDoc.562--> " & ex1.Message, EventLogEntryType.Information, 44)
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
            Dim b_continue As Boolean = True

            If s_keywordsToSearch = "" Then
                b_continue = False
            End If

            If b_continue Then
                For i_countA = 0 To i_uboundKeywords
                    For i_countB = 0 To i_uboundCutText

                        s_keywords = LCase(s_arrayKeywords(i_countA))
                        s_cText = LCase(s_arrayCutText(i_countB))

                        i_placeKeywords = InStr(s_cText, s_keywords, CompareMethod.Text)

                        If i_placeKeywords > 0 Then

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

                            Dim i_findS As Integer
                            i_findS = InStr(Len(s_cText), LCase(s_cText), "s", CompareMethod.Text)

                            If (i_lenKeywords = i_lenCutText) Or (((i_lenKeywords + 1) = i_lenCutText) And i_findS > 0) Then
                                CheckLenghtOfKeywordsWithText = True
                                Exit Function
                            End If
                        End If
                    Next
                Next
            End If

            CheckLenghtOfKeywordsWithText = False

        Catch ex As Exception
            Dim sSource As String = "AP_DENIS"
            Dim sLog As String = "Applo"
            EventLog.WriteEntry(sSource, " CrawlDoc.651 --> " & ex.Message, EventLogEntryType.Information, 44)
        End Try
    End Function

End Module
