Imports System.ComponentModel

Public Class Form1

    Private Declare Function GetPrivateProfileString Lib "kernel32" Alias _
            "GetPrivateProfileStringA" (ByVal lpApplicationName As String,
                ByVal lpKeyName As String,
                ByVal lpDefault As String,
                ByVal lpReturnedString As System.Text.StringBuilder,
                ByVal nSize As Integer, ByVal lpFileName As String) As Integer
    Private Declare Unicode Function WritePrivateProfileString Lib "kernel32" _
                Alias "WritePrivateProfileStringW" (ByVal lpApplicationName As String,
                ByVal lpKeyName As String, ByVal lpString As String,
                ByVal lpFileName As String) As Int32

    Private Function ReadIni(ByVal Section As String, ByVal Key As String, ByVal Deflt As String) As String
        Dim sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString(Section, Key, Deflt, sb, sb.Capacity, System.AppDomain.CurrentDomain.BaseDirectory() & "Resource.ini")
        ReadIni = sb.ToString
    End Function

    Private Sub WriteIni(ByVal Section As String, ByVal Key As String, ByVal Value As String)
        WritePrivateProfileString(Section, Key, Value, System.AppDomain.CurrentDomain.BaseDirectory() & "Resource.ini")
    End Sub

    Private Sub SetGender(ByVal Key As String, ByVal Value As String)
        WritePrivateProfileString("Gender", Key, Value, System.AppDomain.CurrentDomain.BaseDirectory() & "Gender.ini")
    End Sub

    Private Function GetGender(ByVal Key As String, ByVal Deflt As String) As String
        Dim sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString("Gender", Key, Deflt, sb, sb.Capacity, System.AppDomain.CurrentDomain.BaseDirectory() & "Gender.ini")
        GetGender = sb.ToString
    End Function

    Private Sub SetNotebook(ByVal Key As String, Optional ByVal Delete As Boolean = False)
        Dim Record As String = GetNotebook(Key)
        Record += 1
        If Delete Then Record = Nothing
        WritePrivateProfileString("Glossary", Key, Record, System.AppDomain.CurrentDomain.BaseDirectory() & "Notebook.ini")
    End Sub

    Private Function GetNotebook(ByVal Key As String) As String
        Dim sb = New System.Text.StringBuilder(500)
        GetPrivateProfileString("Glossary", Key, "0", sb, sb.Capacity, System.AppDomain.CurrentDomain.BaseDirectory() & "Notebook.ini")
        GetNotebook = sb.ToString
    End Function

    Private CurrentWord As String = ""
    Private EntryCount As Integer = 0
    Private SynonymCount As Integer = 0
    Private Analysing As Boolean = False
    Private Urls As String = "https://www.wikipedia.org"
    Private ReadOnly Shorten(,) As String = {{"ment", ""}, {"less", ""}, {"ied", "y"}, {"ed", ""}, {"ed", "e"},
                                    {"er", ""}, {"er", "e"}, {"est", ""}, {"est", "e"}, {"ly", ""},
                                    {"ily", "y"}, {"ing", ""}, {"ing", "e"}, {"tion", "t"}, {"tion", "te"},
                                    {"nce", "nt"}, {"ncy", "nt"}, {"ness", ""}, {"iness", "y"}, {"ful", ""},
                                    {"iful", "y"}, {"s", ""}, {"ses", "s"}, {"xes", "x"}, {"oes", "o"},
                                    {"ies", "y"}, {"ves", "f"}, {"le", "ility"}}
    Private ReadOnly Prefix As String() = {"a", "anti", "de", "dis", "il", "im", "in", "ir", "non", "un"}
    Private TimeRemain As Long = 0

    Private Function FormatTime(ByVal Sec As Long) As String
        If Sec >= 0 Then
            FormatTime = (Sec \ 3600) & ":" & ((Sec \ 60) Mod 60).ToString("00") & ":" & (Sec Mod 60).ToString("00")
        Else
            FormatTime = "-" & (-Sec \ 3600) & ":" & ((-Sec \ 60) Mod 60).ToString("00") & ":" & (-Sec Mod 60).ToString("00")
        End If
    End Function

    Private Function CountIni(ByVal Section As String) As Integer
        CountIni = -1
        Dim Temp As String = ""
        Dim Reached As Boolean = False
        On Error GoTo a
        Using r As System.IO.StreamReader = New System.IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory() & "Resource.ini")
            If Section <> "" Then 'count key number
                Do While r.EndOfStream = False
                    Temp = r.ReadLine()
                    If Temp = "[" & Section & "]" Then
                        Reached = True
                    Else
                        If Strings.Left(Temp, 1) = "[" Then Reached = False
                    End If
                    If Reached Then CountIni += 1
                Loop
            Else                'count section number
                CountIni = 0
                Do While r.EndOfStream = False
                    If Strings.Left(r.ReadLine(), 1) = "[" Then CountIni += 1
                Loop
            End If
        End Using
a:
        Return CountIni
    End Function

    Private Function RidPunct(ByVal Input As String) As String
        While Input.Length > 0 And InStr("`1234567890-=~!@#$%^&*()_+[]\;,./{}|:<>?' ", Strings.Left(Input, 1)) <> 0
            Input = Strings.Mid(Input, 2, Input.Length - 1)
        End While
        While Input.Length > 0 And InStr("`1234567890-=~!@#$%^&*()_+[]\;,./{}|:<>?' ", Strings.Right(Input, 1)) <> 0
            Input = Strings.Mid(Input, 1, Input.Length - 1)
        End While
        Return Input
    End Function

    Private Sub Parse(ByRef TrueStart As Integer, ByRef TrueEnd As Integer)
        If TextBox1.Text = "" Then Exit Sub
        Dim Article As String = TextBox1.Text.Replace(vbCr, " ").Replace(vbLf, " ").Replace("""", " ")
        If TextBox1.SelectionStart <> 0 Then
            TrueStart = InStrRev(Article, " ", TextBox1.SelectionStart)
            If TrueStart = 0 Then
                TrueStart = 1
            Else
                TrueStart += 1
            End If
        Else
            TrueStart = 1
        End If
        If TextBox1.SelectionStart <> TextBox1.TextLength Then
            TrueEnd = InStr(TextBox1.SelectionStart + 1, Article, " ")
            If TrueEnd = 0 Then
                TrueEnd = TextBox1.TextLength
            Else
                TrueEnd -= 1
            End If
        Else
            TrueEnd = TextBox1.TextLength
        End If
        If CurrentWord = Strings.Mid(Article, TrueStart, TrueEnd - TrueStart + 1) Then
            Exit Sub
        Else
            CurrentWord = Strings.Mid(Article, TrueStart, TrueEnd - TrueStart + 1)
        End If
    End Sub

    Private Sub UpdateTree(ByVal Root As String, ByVal CurrentW As String)
        If Root = "" Then Exit Sub
        Dim WordC As Integer = ReadIni(Root, 0, 0)
        Dim TempW As String = ""
        Dim CurrentNode As TreeNode = TreeView1.Nodes.Add(CurrentW)
        If WordC = 0 Then Exit Sub
        CurrentNode.Nodes.Add(Root)
        If CurrentW = Root Then CurrentNode.Nodes(0).NodeFont = New Font(TreeView1.Font, FontStyle.Bold)
        For i = 1 To WordC
            TempW = ReadIni(Root, i, 0)
            CurrentNode.Nodes.Add(TempW)
            If CurrentW = TempW Then CurrentNode.Nodes(i).NodeFont = New Font(TreeView1.Font, FontStyle.Bold)
        Next i
        For i = 0 To CurrentNode.GetNodeCount(False) - 1 'add numbers
            CurrentNode.Nodes(i).Text = i + 1 & ": " & CurrentNode.Nodes(i).Text
        Next i
    End Sub

    Private Function MakeGuessWord(ByVal CurrentW As String, ByVal SubtL As String, ByVal AddL As String) As String
        If Strings.StrComp(Strings.Right(CurrentW, SubtL.Length), SubtL, CompareMethod.Text) = 0 Then
            Return Strings.Mid(CurrentW, 1, CurrentW.Length - SubtL.Length) & AddL
        End If
        Return CurrentW
    End Function

    Private Function MakeGuessWord(ByVal CurrentW As String, ByVal Prefix As String) As String
        If Strings.StrComp(Strings.Left(CurrentW, Prefix.Length), Prefix, CompareMethod.Text) = 0 Then
            Return Strings.Mid(CurrentW, Prefix.Length + 1, CurrentW.Length - Prefix.Length)
        End If
        Return CurrentW
    End Function

    Private Function Derivatives(ByVal CurrentW As String) As String()
        Dim Guess As String = CurrentW
        Dim GuessOriginal As String = CurrentW
        Dim RootGuess As String = ""
        Dim Root As String = ReadIni("<Root>", CurrentW, "")
        Dim DCount As Integer = 1
        Dim DArray As String() = {CurrentW}
        If CurrentW = "" Then Return DArray
        For i = 0 To Shorten.GetLength(0) - 1
            Guess = MakeGuessWord(CurrentW, Shorten(i, 0), Shorten(i, 1))
            If Guess <> CurrentW Then
                ReDim Preserve DArray(DCount)
                DArray(DCount) = Guess
                DCount += 1
            End If
        Next i
        Dim TempCount As Integer = DCount
        For j = 0 To TempCount - 1
            For i = 0 To Prefix.GetLength(0) - 1
                Guess = MakeGuessWord(DArray(j), Prefix(i))
                If Guess <> DArray(j) Then
                    ReDim Preserve DArray(DCount)
                    DArray(DCount) = Guess
                    DCount += 1
                End If
            Next i
        Next j
        For j = 1 To DCount - 1
            RootGuess = ReadIni("<Root>", DArray(j), "")
            If RootGuess = "" Or RootGuess = Root Then
                DArray(j) = CurrentW   'remove empty
            End If
        Next j
        Dim TArray As String() = DArray.Distinct().ToArray
        For j = 0 To TArray.GetLength(0) - 1
            For i = 0 To Shorten.GetLength(0) - 1
                Guess = MakeGuessWord(TArray(j), Shorten(i, 1), Shorten(i, 0))
                If Guess <> TArray(j) Then
                    RootGuess = ReadIni("<Root>", Guess, "")
                    If RootGuess <> "" AndAlso RootGuess <> Root Then
                        ReDim Preserve DArray(DCount)
                        DArray(DCount) = Guess
                        DCount += 1
                    End If
                End If
            Next i
        Next j
        Return DArray.Distinct().ToArray
    End Function

    Private Sub FixW()
        Dim RStart As Integer = 0
        Dim REnd As Integer = 0
        Dim Gender As String = ""
        TreeView1.Nodes.Clear()
        Parse(RStart, REnd)
        CurrentWord = RidPunct(CurrentWord)                 'Get rid of punctuations
        Dim DArray As String() = Derivatives(CurrentWord)   'Guess possible tense
        For i = 0 To DArray.GetLength(0) - 1
            UpdateTree(ReadIni("<Root>", DArray(i), ""), DArray(i))
            If i = 0 AndAlso TreeView1.Nodes.Count = 0 Then TreeView1.Nodes.Add(DArray(i))
        Next i
        Dim NotebookEntry As String = GetNotebook(CurrentWord)
        Call TreeViewColor()
        If NotebookEntry <> 0 Then
            Dim NotebookNode As TreeNode = TreeView1.Nodes.Add(CurrentWord)
            NotebookNode.NodeFont = New Font(TreeView1.Font, FontStyle.Italic)
            NotebookNode.Nodes.Add(NotebookEntry & " ")
        End If
        TreeView1.ExpandAll()
    End Sub

    Private Sub TreeViewColor()
        Dim CurrentNode As TreeNode
        Dim FittedW As String = ""
        For i = 0 To TreeView1.Nodes.Count - 1
            For ii = 0 To TreeView1.Nodes(i).Nodes.Count
                If ii = TreeView1.Nodes(i).Nodes.Count Then
                    CurrentNode = TreeView1.Nodes(i)
                Else
                    CurrentNode = TreeView1.Nodes(i).Nodes(ii)
                End If
                FittedW = Strings.Right(CurrentNode.Text, CurrentNode.Text.Length - InStr(CurrentNode.Text, " "))
                Select Case GetGender(FittedW, "")
                    Case "C"
                        CurrentNode.ForeColor = Color.Green
                    Case "D"
                        CurrentNode.ForeColor = Color.Red
                    Case Else
                        CurrentNode.ForeColor = Color.Black
                End Select
            Next ii
        Next i
    End Sub

    Private Sub AddSynonymsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddSynonymsToolStripMenuItem.Click, TreeView1.MouseDoubleClick
        Dim Entry As String = RidPunct(InputBox("Enter the easy word: ", "Add Synonyms", CurrentWord))
        If Entry = "" Then Exit Sub
        Dim NewEntry As String = ""
        Dim WordN As Integer = 0
        Dim OldRecord As String = ""
        Dim EasyRoot As String = ReadIni("<Root>", Entry, "")
        Dim DiffRoot As String = ""
        If EasyRoot <> "" Then      'old root
            WordN = ReadIni(EasyRoot, 0, 0)     'new word
            For i = 1 To WordN
                OldRecord = OldRecord & i & ": " & ReadIni(EasyRoot, i, "") & vbCrLf
            Next i
            NewEntry = RidPunct(InputBox("Enter the difficult word: " & vbCrLf & "There are existing records: " & vbCrLf & OldRecord, "Add Synonyms of <" & EasyRoot & ">", ""))
            If NewEntry = "" Then Exit Sub
            DiffRoot = ReadIni("<Root>", NewEntry, "")
            If DiffRoot <> "" Then          'old word
                MsgBox("Word """ & NewEntry & """ has already been assigned to root <" & DiffRoot & ">",
                       MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Add Synonyms")
                Exit Sub
            End If
            WriteIni(EasyRoot, 0, WordN + 1)
            WriteIni(EasyRoot, WordN + 1, NewEntry)
            WriteIni("<Root>", NewEntry, EasyRoot)
            EntryCount += 1
        Else                    'new root
            NewEntry = RidPunct(InputBox("Enter the difficult word: ", "Add Synonyms", ""))  'new word
            If NewEntry = Entry Then
                MsgBox("Root and its synonyms cannot be the same!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Add Synonyms")
                Exit Sub
            ElseIf NewEntry = "" Then
                Exit Sub
            End If
            DiffRoot = ReadIni("<Root>", NewEntry, "")
            If DiffRoot <> "" Then          'old word
                MsgBox("Word """ & NewEntry & """ has already been assigned to root <" & DiffRoot & ">",
                       MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Add Synonyms")
                Exit Sub
            End If
            WriteIni(Entry, 0, 1)
            WriteIni(Entry, 1, NewEntry)
            WriteIni("<Root>", NewEntry, Entry)
            WriteIni("<Root>", Entry, Entry) 'create new root
            EntryCount += 2
            SynonymCount += 1
        End If
        CurrentWord = ""
        LabelEntryCount.Text = "Entry Count: " & SynonymCount & "/" & EntryCount
        Call FixW()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If Analysing Then
            Exit Sub
        End If
        Call FixW()
        Dim TempString() As String  'count words
        Dim CountN As Integer = 0
        TempString = Split(TextBox1.Text.Replace(vbCr, " ").Replace(vbLf, " "), " ")
        For i = 0 To TempString.Length - 1
            If RidPunct(TempString(i)) <> "" Then CountN += 1
        Next i
        LabelCount.Text = "Word Count: " & CountN
    End Sub

    Private Sub SelectAllToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectAllToolStripMenuItem.Click
        TextBox1.SelectAll()
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim SavePath As String = System.AppDomain.CurrentDomain.BaseDirectory() & "AutoSave.txt"
        If System.IO.File.Exists(SavePath) = True Then
            TextBox1.Text = My.Computer.FileSystem.ReadAllText(SavePath).ToString()
        Else
            My.Computer.FileSystem.WriteAllText(SavePath, "", False)
        End If
        With TextBox1
            .Dock = DockStyle.Fill
            .ContextMenuStrip = GenderMenuStrip
            .SelectionStart = .TextLength
            .ScrollToCaret()
        End With
        Me.Text = "English Enhancer - " & My.Application.Info.Version.ToString
        EntryCount = CountIni("<Root>")
        SynonymCount = CountIni("") - 1
        LabelEntryCount.Text = "Entry Count: " & SynonymCount & "/" & EntryCount
        With WebBrowser1
            .Dock = DockStyle.Fill
            .ScriptErrorsSuppressed = True
        End With
        ToolStripSeparator6.Visible = False
        With TreeView1
            .ContextMenuStrip = GenderMenuStrip
            .Dock = DockStyle.Fill
        End With
        With SplitContainer1
            .Dock = DockStyle.Fill
            .Panel2MinSize = 200
            .SplitterDistance = Me.Width * 0.8
        End With
        TSSpace.Text = ""
        'ToolTip1.SetToolTip(Ticker, "")
        Ticker.ToolTipText = "Left click to pause/resume. Right click to set. "
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Dim TrueStart As Integer = 0, TrueEnd As Integer = 0
        Parse(TrueStart, TrueEnd)
        If RidPunct(CurrentWord) <> "" And Asc(e.KeyChar) >= 48 And Asc(e.KeyChar) <= 58 Then '-=45, ==61
            Dim KeyNum As Integer = Asc(e.KeyChar) - 48
            Dim Root As String = ReadIni("<Root>", RidPunct(CurrentWord), "")
            If Root = "" Then Exit Sub
            Dim WordC As Integer = ReadIni(Root, 0, 0)
            If KeyNum - 1 > WordC Then Exit Sub
            Dim ReplaceW As String = ""
            Dim PreviousZero As Integer = 0
            Dim Position As Integer = TrueEnd
            While Strings.Mid(TextBox1.Text, Position, 1) = "0"
                PreviousZero += 1
                Position -= 1
            End While
            If KeyNum = 1 And PreviousZero = 0 Then
                ReplaceW = Root
            Else
                ReplaceW = ReadIni(Root, KeyNum - 1 + PreviousZero * 10, "")
            End If
            If ReplaceW <> "" Then
                TextBox1.Text = Strings.Left(TextBox1.Text, TrueStart - 1) & ReplaceW & Strings.Right(TextBox1.Text, TextBox1.TextLength - TrueEnd)
                TextBox1.SelectionStart = TrueStart + ReplaceW.Length - 1
                e.KeyChar = vbBack
                Call FixW()
            End If
        End If
    End Sub

    Private Sub CombineRootsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CombineRootsToolStripMenuItem.Click
        Dim EntryA As String = RidPunct(InputBox("Enter the first (host) root or its synonym: ", "Combine Roots"))
        If EntryA = "" Then
            Exit Sub
        End If
        Dim RootA As String = ReadIni("<Root>", EntryA, "")
        If RootA = "" Then
            MsgBox("There is no record of word <" & EntryA & ">.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Combine Roots")
            Exit Sub
        End If
        Dim EntryB As String = RidPunct(InputBox("Enter the second root or its synonym: ", "Combine Roots"))
        If EntryB = "" Then Exit Sub
        Dim RootB As String = ReadIni("<Root>", EntryB, "")
        If RootB = "" Then
            MsgBox("There is no record of word <" & EntryB & ">.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Combine Roots")
            Exit Sub
        ElseIf RootB = RootA Then
            MsgBox("Two roots cannot be the same!", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Combine Roots")
            Exit Sub
        End If
        'confirm to continue
        Dim AddN As Integer = ReadIni(RootB, 0, 0)
        Dim OriginalN As Integer = ReadIni(RootA, 0, 0)
        Dim SynA As String = ""
        Dim SynB As String = ""
        For i = 1 To OriginalN
            SynA = SynA & " " & ReadIni(RootA, i, "") & ","
        Next i
        For i = 1 To AddN
            SynB = SynB & " " & ReadIni(RootB, i, "") & ","
        Next i
        Dim Consent As MsgBoxResult = MsgBox("Please confirm that roots: " & vbCrLf _
                                             & "<" & RootA & ">:" & SynA & vbCrLf _
                                             & "<" & RootB & ">:" & SynB & vbCrLf _
                                             & "will be combined. " & vbCrLf _
                                             , MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Combine Roots")
        If Consent = vbNo Then
            Exit Sub
        End If
        'start combine
        WriteIni("<Root>", RootB, RootA)
        Dim TempW As String = ""
        For i = 1 To AddN
            TempW = ReadIni(RootB, i, "")
            WriteIni(RootA, OriginalN + i, TempW)
            WriteIni("<Root>", TempW, RootA)
        Next i
        WriteIni(RootA, OriginalN + AddN + 1, RootB)
        WriteIni(RootA, 0, OriginalN + AddN + 1)
        WriteIni(RootB, Nothing, Nothing)
        MsgBox("Roots combined successfully!", MsgBoxStyle.Exclamation + MsgBoxStyle.OkOnly, "Combine Roots")
        SynonymCount -= 1
        LabelEntryCount.Text = "Entry Count: " & SynonymCount & "/" & EntryCount
        Call FixW()
    End Sub

    Private Sub Form1_Closing(sender As Object, e As CancelEventArgs) Handles Me.Closing
        Dim SavePath As String = System.AppDomain.CurrentDomain.BaseDirectory() & "AutoSave.txt"
        If System.IO.File.Exists(SavePath) = True Then
            My.Computer.FileSystem.WriteAllText(SavePath, TextBox1.Text, False)
        Else
            My.Computer.FileSystem.WriteAllText(SavePath, "", False)
        End If
    End Sub

    Private Sub DeleteEntryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteEntryToolStripMenuItem.Click, DeleteEntryToolStripMenuItem1.Click
        Dim Entry As String = RidPunct(InputBox("Enter the entry to be deleted: ", "Delete Entry", CurrentWord))
        If Entry = "" Then Exit Sub
        Dim Root As String = ReadIni("<Root>", Entry, "")
        If Root = "" Then
            MsgBox("There is no record of word <" & Entry & ">.", MsgBoxStyle.Critical + MsgBoxStyle.OkOnly, "Delete Entry")
            Exit Sub
        End If
        Dim WordN As Integer = ReadIni(Root, 0, 0)
        Dim EntryN As Integer = 0
        Dim Consent As MsgBoxResult = MsgBox("Please confirm that word """ & Entry & """ will be deleted from root <" & Root & ">.",
                                             MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Delete Entry")
        If Consent = MsgBoxResult.No Then
            Exit Sub
        End If
        If Root <> Entry Then
            For i = 1 To WordN
                If ReadIni(Root, i, "") = Entry Then
                    EntryN = i
                    Exit For
                End If
            Next i
            For i = EntryN To WordN
                If i = WordN Then
                    WriteIni(Root, i, Nothing)
                Else
                    WriteIni(Root, i, ReadIni(Root, i + 1, ""))
                End If
            Next i
            WriteIni(Root, 0, WordN - 1)
            WriteIni("<Root>", Entry, Nothing)
            If WordN = 1 Then
                WriteIni("<Root>", Root, Nothing)
                WriteIni(Root, Nothing, Nothing)
                EntryCount -= 1
                SynonymCount -= 1
            End If
            EntryCount -= 1
        Else    ' entry is a root
            WriteIni("<Root>", Root, Nothing)
            If WordN = 1 Then
                WriteIni("<Root>", ReadIni(Root, 1, ""), Nothing)
                WriteIni(Root, Nothing, Nothing)
                EntryCount -= 1
                SynonymCount -= 1
            Else
                Dim NewRoot As String = ReadIni(Root, 1, "")
                Dim TempEntry As String = ""
                WriteIni(NewRoot, 0, WordN - 1)
                For i = 2 To WordN
                    TempEntry = ReadIni(Root, i, "")
                    WriteIni(NewRoot, i - 1, TempEntry)
                    WriteIni("<Root>", TempEntry, NewRoot)
                Next i
                WriteIni("<Root>", NewRoot, NewRoot)
                WriteIni(Root, Nothing, Nothing)
            End If
            EntryCount -= 1
        End If
        CurrentWord = ""
        LabelEntryCount.Text = "Entry Count: " & SynonymCount & "/" & EntryCount
        Call FixW()
    End Sub

    Private Sub SaveAsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveAsToolStripMenuItem.Click
        If SaveFileDialog1.ShowDialog() = DialogResult.OK Then
            Dim Path As String = SaveFileDialog1.FileName
            My.Computer.FileSystem.WriteAllText(Path, TextBox1.Text, False)
            SaveFileDialog1.FileName = ""
        End If
    End Sub

    Private Sub Form1_Resize(sender As Object, e As EventArgs) Handles Me.Resize
        ToolStripTextBox1.Width = Me.Width - WindowToolStripMenuItem.Width - SynonymsToolStripMenuItem.Width -
        NotebookToolStripMenuItem.Width - GenderMenuStrip.Width - ControlToolStripMenuItem.Width + 150
    End Sub

    Private Sub AnalysisToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalysisToolStripMenuItem.Click
        Dim Temp As String = ""
        Dim ResultN() As Integer = Nothing
        Const MaxNum As Integer = 20
        For i = 0 To MaxNum Step 1
            ReDim Preserve ResultN(i)
            ResultN(i) = 0
        Next i
        'On Error GoTo a
        Using r As System.IO.StreamReader = New System.IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory() & "Resource.ini")
            Do While r.EndOfStream = False
                Temp = r.ReadLine()
                If Strings.Left(Temp, 2) = "0=" Then
                    ResultN(Strings.Right(Temp, Strings.Len(Temp) - 2) + 1) += 1
                End If
            Loop
        End Using
        Dim ResultS As String = ""
        Dim SumN As Integer = 0
        For i = 2 To MaxNum
            SumN += ResultN(i)
        Next i
        For i = 2 To MaxNum Step 1
            If ResultN(i) <> 0 Then
                ResultS = ResultS & i & ": "
                For ii = 1 To Fix(ResultN(i) / SumN * 100)
                    ResultS = ResultS & "■"
                Next ii
                ResultS = ResultS & " (" & Int(ResultN(i) / SumN * 1000) / 10 & "%)" & vbCrLf
            End If
        Next i
        MsgBox("Group Size Distribution: " & vbCrLf & ResultS &
                "Group Count: " & SumN & vbCrLf &
                "Total Entries: " & EntryCount, MsgBoxStyle.OkOnly + MsgBoxStyle.Exclamation, "Analysis")
a:
    End Sub

    Private Sub TreeView1_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles TreeView1.AfterSelect
        CurrentWord = Strings.Right(TreeView1.SelectedNode.Text, TreeView1.SelectedNode.Text.Length - InStr(TreeView1.SelectedNode.Text, " "))
        TextBox1.Select(TextBox1.SelectionStart + TextBox1.SelectionLength, 0)
    End Sub

    Private Sub SetAsAComplimentToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetAsAComplimentToolStripMenuItem.Click, SetAsAComplimentToolStripMenuItem1.Click
        If CurrentWord <> "" Then
            'Dim Consent As MsgBoxResult = MsgBox("Would you assign word """ & CurrentWord & """ as a compliment?", vbYesNo + vbQuestion + vbDefaultButton2, "Set as a Compliment")
            'If Consent = vbNo Then 
            '   Exit Sub
            'End If
            SetGender(CurrentWord, "C")
            Call TreeViewColor()
        End If
    End Sub

    Private Sub SetAsADerogatoryToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetAsADerogatoryToolStripMenuItem.Click, SetAsADerogatoryToolStripMenuItem1.Click
        If CurrentWord <> "" Then
            'Dim Consent As MsgBoxResult = MsgBox("Would you assign word """ & CurrentWord & """ as a derogatory?", vbYesNo + vbQuestion + vbDefaultButton2, "Set as a Compliment")
            'If Consent = vbNo Then 
            '   Exit Sub
            'End If
            SetGender(CurrentWord, "D")
            Call TreeViewColor()
        End If
    End Sub

    Private Sub SetAsANeuterToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SetAsANeuterToolStripMenuItem.Click, SetAsANeuterToolStripMenuItem1.Click
        If CurrentWord <> "" Then
            'Dim Consent As MsgBoxResult = MsgBox("Would you assign word """ & CurrentWord & """ as a neuter" & vbCrLf & "which consequently resets the record? ", vbYesNo + vbQuestion + vbDefaultButton2, "Set as a Compliment")
            'If Consent = vbNo Then 
            '   Exit Sub
            'End If
            SetGender(CurrentWord, Nothing)
            Call TreeViewColor()
        End If
    End Sub

    Private Sub ListToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles ListToolStripMenuItem1.Click
        CurrentWord = ""
        Dim Temp As String = ""
        Dim RootN As Integer = 0
        Dim RootR() As String = Nothing
        'On Error GoTo a
        Using r As System.IO.StreamReader = New System.IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory() & "Resource.ini")
            Do While r.EndOfStream = False
                Temp = r.ReadLine()
                If Strings.Left(Temp, 1) = "[" And Temp <> "[<Root>]" Then
                    ReDim Preserve RootR(RootN)
                    RootR(RootN) = Strings.Mid(Temp, 2, Temp.Length - 2)
                    RootN += 1
                    r.ReadLine()
                End If
            Loop
        End Using
        TreeView1.Nodes.Clear()
        System.Array.Sort(RootR)
        For i = 0 To RootN - 1
            TreeView1.Nodes.Add(RootR(i))
            For ii = 0 To ReadIni(RootR(i), 0, 0) - 1
                TreeView1.Nodes(i).Nodes.Add(ReadIni(RootR(i), ii + 1, 0))
            Next ii
        Next i
        Call TreeViewColor()
a:
    End Sub

    Private Sub CopyToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CopyToolStripMenuItem.Click
        If TextBox1.SelectedText <> "" Then
            Clipboard.SetText(TextBox1.SelectedText)
        Else
            If CurrentWord <> "" Then
                Clipboard.SetText(CurrentWord)
            End If
        End If
    End Sub

    Private Sub WebBrowserToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles WebBrowserToolStripMenuItem.Click
        WebBrowser1.Visible = True
        TextBox1.Visible = False
        ToolStripTextBox1.Text = Urls
        ToolStripTextBox1.ForeColor = Color.Black
        WebBrowserToolStripMenuItem.Checked = True
        TextEditorToolStripMenuItem.Checked = False
        RefreshToolStripMenuItem.Visible = True
        RefreshToolStripMenuItem.Enabled = True
        '==========
        SaveAsToolStripMenuItem.Visible = False
        SaveAsToolStripMenuItem.Enabled = False
        SelectAllToolStripMenuItem.Visible = False
        SelectAllToolStripMenuItem.Enabled = False
        ToolStripSeparator7.Visible = False
        AnalyzeTextToolStripMenuItem.Visible = False
        AnalyzeTextToolStripMenuItem.Enabled = False
        AnalyzeTextIncludingToolStripMenuItem.Visible = False
        AnalyzeTextIncludingToolStripMenuItem.Enabled = False
    End Sub

    Private Sub TextEditorToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TextEditorToolStripMenuItem.Click, MyBase.Load
        WebBrowser1.Visible = False
        TextBox1.Visible = True
        ToolStripTextBox1.ForeColor = Color.Gray
        ToolStripTextBox1.Text = "Search..."
        WebBrowserToolStripMenuItem.Checked = False
        TextEditorToolStripMenuItem.Checked = True
        RefreshToolStripMenuItem.Visible = False
        RefreshToolStripMenuItem.Enabled = False
        '=========
        SaveAsToolStripMenuItem.Visible = True
        SaveAsToolStripMenuItem.Enabled = True
        SelectAllToolStripMenuItem.Visible = True
        SelectAllToolStripMenuItem.Enabled = True
        ToolStripSeparator7.Visible = True
        AnalyzeTextToolStripMenuItem.Visible = True
        AnalyzeTextToolStripMenuItem.Enabled = True
        AnalyzeTextIncludingToolStripMenuItem.Visible = True
        AnalyzeTextIncludingToolStripMenuItem.Enabled = True
    End Sub

    Private Sub SwitchWindowToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SwitchWindowToolStripMenuItem.Click
        If TextEditorToolStripMenuItem.Checked = True Then
            Call WebBrowserToolStripMenuItem_Click(sender, e)
        Else
            Call TextEditorToolStripMenuItem_Click(sender, e)
        End If
    End Sub

    Private Sub ToolStripTextBox1_KeyDown(sender As Object, e As KeyEventArgs) Handles ToolStripTextBox1.KeyDown
        'If e.KeyCode = Keys.A AndAlso e.Control = True Then 'ctrl+A
        '   ToolStripTextBox1.SelectAll()
        'End If
        If WebBrowserToolStripMenuItem.Checked = False Then
            Exit Sub
        End If
        If e.KeyCode = Keys.Enter Then
            WebBrowser1.Navigate(ToolStripTextBox1.Text)
        End If
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        WebBrowser1.Refresh()
    End Sub

    Private Sub WebBrowser1_DocumentCompleted(sender As Object, e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowser1.DocumentCompleted
        ToolStripTextBox1.Text = WebBrowser1.Url.ToString
        Urls = WebBrowser1.Url.ToString
    End Sub

    Private Sub AddToNotebookToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AddToNotebookToolStripMenuItem.Click
        If CurrentWord <> "" Then
            SetNotebook(CurrentWord)
            Call FixW()
        End If
    End Sub

    Private Sub DeleteNotebookToolStripMenuItem1_Click(sender As Object, e As EventArgs) Handles DeleteNotebookToolStripMenuItem1.Click, DeleteEntryofNotebookToolStripMenuItem.Click
        If CurrentWord = "" Then
            Exit Sub
        End If
        If GetNotebook(CurrentWord) = 0 Then
            Exit Sub
        End If
        Dim Consent As MsgBoxResult = MsgBox("Are you sure you want to delete notebook entry: """ & CurrentWord & """? ",
                    MsgBoxStyle.Question + MsgBoxStyle.YesNo + MsgBoxStyle.DefaultButton2, "Delete Selected Entry")
        If Consent = MsgBoxResult.Yes Then
            SetNotebook(CurrentWord, True)
        End If
        'Call ListAllEntriesToolStripMenuItem_Click(sender, e)
        Call FixW()
    End Sub

    Private Sub ListAllEntriesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ListAllEntriesToolStripMenuItem.Click
        TreeView1.Nodes.Clear()
        CurrentWord = ""
        Dim Temp As String = ""
        Dim Reached As Boolean = False
        Dim CurrentR As String = ""
        Dim Value As String = "0"
        On Error GoTo a
        Using r As System.IO.StreamReader = New System.IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory() & "Notebook.ini")
            Do While r.EndOfStream = False
                Temp = r.ReadLine()
                If Reached Then
                    CurrentR = Strings.Left(Temp, InStr(Temp, "=") - 1)
                    Value = Strings.Right(Temp, Temp.Length - InStr(Temp, "="))
                    TreeView1.Nodes.Add(Value & ": " & CurrentR).Tag = Value
                End If
                If Temp = "[Glossary]" Then
                    Reached = True
                End If
            Loop
        End Using
        Call TreeViewColor()
a:
    End Sub

    Private Sub TextBox1_Click(sender As Object, e As EventArgs) Handles TextBox1.Click
        If Not Analysing Then Call FixW()
    End Sub

    Private Sub AnalyzeTextToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AnalyzeTextToolStripMenuItem.Click, AnalyzeTextIncludingToolStripMenuItem.Click
        Dim Paragraph As String = TextBox1.Text.Replace(vbCr, " ").Replace(vbLf, " ").Replace("""", " ")
        Dim TempString() As String = Split(Paragraph, " ")
        Dim CurrentW As String = ""
        Dim Position As Integer = 0
        Analysing = True
        If TextBox1.Text.Last <> " " And TextBox1.Text.Last <> vbCr And TextBox1.Text.Last <> vbLf Then
            TextBox1.Text = TextBox1.Text & " "
            TextBox1.Select(TextBox1.TextLength - 1, 1)
            TextBox1.SelectionFont = New Font(TextBox1.SelectionFont, FontStyle.Regular)
        End If
        TextBox1.SelectAll()
        TextBox1.SelectionFont = New Font(TextBox1.Font, FontStyle.Regular)
        For i = 0 To TempString.Length - 1
            CurrentW = RidPunct(TempString(i))
            Position += TempString(i).Length + 1
            If CurrentW <> "" Then
                If ReadIni("<Root>", CurrentW, "") <> "" Or
                    (sender Is AnalyzeTextIncludingToolStripMenuItem AndAlso Derivatives(CurrentW).GetLength(0) > 1) Then
                    TextBox1.Select(Position - TempString(i).Length - 1, TempString(i).Length)
                    TextBox1.SelectionFont = New Font(TextBox1.SelectionFont, TextBox1.SelectionFont.Style + FontStyle.Underline)
                End If
                If GetNotebook(CurrentW) <> 0 Then
                    TextBox1.Select(Position - TempString(i).Length - 1, TempString(i).Length)
                    TextBox1.SelectionFont = New Font(TextBox1.SelectionFont, TextBox1.SelectionFont.Style + FontStyle.Bold)
                End If
            End If
            Application.DoEvents()
        Next i
        TextBox1.Select(TextBox1.TextLength, 0)
        Analysing = False
    End Sub

    Private Sub CurrentWordDisplayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CurrentWordDisplayToolStripMenuItem.Click
        If CurrentWord <> "" Then
            TreeView1.Nodes.Clear()
            Call UpdateTree(ReadIni("<Root>", CurrentWord, ""), CurrentWord)
            Dim NotebookEntry As String = GetNotebook(CurrentWord)
            Call TreeViewColor()
            If NotebookEntry <> 0 Then
                Dim NotebookNode As TreeNode = TreeView1.Nodes.Add(CurrentWord)
                NotebookNode.NodeFont = New Font(TreeView1.Font, FontStyle.Italic)
                NotebookNode.Nodes.Add(NotebookEntry & " ")
            End If
            TreeView1.ExpandAll()
        End If
    End Sub

    Private Sub GenderMenuStrip_Opening(sender As Object, e As CancelEventArgs) Handles GenderMenuStrip.Opening
        Dim CurrentW As String = RidPunct(CurrentWord)
        If CurrentW = "" Then
            CurrentWordDisplayToolStripMenuItem.Text = """No Selected Word"""
            CurrentWordDisplayToolStripMenuItem.Enabled = False
        Else
            CurrentWordDisplayToolStripMenuItem.Text = CurrentW
            CurrentWordDisplayToolStripMenuItem.Enabled = True
            If GetNotebook(CurrentW) = 0 Then
                DeleteEntryofNotebookToolStripMenuItem.Enabled = False
            Else
                DeleteEntryofNotebookToolStripMenuItem.Enabled = True
            End If
            If ReadIni("<Root>", CurrentW, "") = "" Then
                DeleteEntryToolStripMenuItem1.Enabled = False
            Else
                DeleteEntryToolStripMenuItem1.Enabled = True
            End If
        End If
    End Sub

    Private Sub ToolStripTextBox1_GotFocus(sender As Object, e As EventArgs) Handles ToolStripTextBox1.GotFocus
        If WebBrowserToolStripMenuItem.Checked = True Then Exit Sub
        ToolStripTextBox1.ForeColor = Color.Black
        If ToolStripTextBox1.Text = "Search..." Then
            ToolStripTextBox1.Text = ""
        Else
            Call ToolStripTextBox1_TextChanged(sender, e)
        End If
    End Sub

    Private Sub ToolStripTextBox1_LostFocus(sender As Object, e As EventArgs) Handles ToolStripTextBox1.LostFocus
        If WebBrowserToolStripMenuItem.Checked = True Then Exit Sub
        ToolStripTextBox1.ForeColor = Color.Gray
        If ToolStripTextBox1.Text = "" Then
            ToolStripTextBox1.Text = "Search..."
        End If
    End Sub

    Private Sub ToolStripTextBox1_TextChanged(sender As Object, e As EventArgs) Handles ToolStripTextBox1.TextChanged
        If WebBrowserToolStripMenuItem.Checked = True Then Exit Sub
        Dim KeyWord As String = ToolStripTextBox1.Text
        Dim Temp As String = ""
        Dim Reached As Boolean = False
        Dim AArray() As String = Nothing  'abc???
        Dim AIndex As Integer = 0
        Dim RArray() As String = Nothing   '??abc??
        Dim RIndex As Integer = 0
        Dim Keys As String = ""
        If ToolStripTextBox1.Text = "" Or ToolStripTextBox1.Text = "Search..." Then Exit Sub
        Using r As System.IO.StreamReader = New System.IO.StreamReader(System.AppDomain.CurrentDomain.BaseDirectory() & "Resource.ini")
            Do While r.EndOfStream = False
                Temp = r.ReadLine()
                If Temp = "[<Root>]" Then
                    Reached = True
                    Temp = r.ReadLine()
                End If
                If Reached Then
                    If Strings.Left(Temp, 1) = "[" Then Exit Do
                    Keys = Strings.Left(Temp, InStr(Temp, "=") - 1)
                    If InStr(Temp, "=") > KeyWord.Length And InStr(Keys, KeyWord, CompareMethod.Text) <> 0 Then
                        If Strings.Left(Temp, KeyWord.Length) = KeyWord Then
                            ReDim Preserve AArray(AIndex)
                            AArray(AIndex) = Keys
                            AIndex += 1
                        Else
                            ReDim Preserve RArray(RIndex)
                            RArray(RIndex) = Keys
                            RIndex += 1
                        End If
                    End If
                End If
            Loop
        End Using
        TreeView1.Nodes.Clear()
        Call UpdateSearch(AArray, AIndex)
        Call UpdateSearch(RArray, RIndex)
        TreeView1.ExpandAll()
        Call TreeViewColor()
    End Sub

    Private Sub UpdateSearch(ByVal ResultArray As String(), ByVal ArrayIndex As Integer)
        Dim CurrentNode As TreeNode
        If ArrayIndex <> 0 Then
            System.Array.Sort(ResultArray)
            CurrentNode = TreeView1.Nodes.Add(ResultArray(0))
            For i = 0 To ArrayIndex - 1
                CurrentNode.Nodes.Add(ResultArray(i))
            Next i
        End If
    End Sub

    Private Sub AboutUsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutUsToolStripMenuItem.Click
        MsgBox(Me.Text & vbNewLine _
       & vbNewLine _
       & "--English Version--" & vbNewLine _
       & "English learning companion and handy notebook" & vbNewLine _
       & "zhuoheliu@outlook.com" & vbNewLine _
       & vbNewLine _
       & "© 2016 - 2021 Zhuohe Liu", MsgBoxStyle.OkOnly + MsgBoxStyle.Information, "About Us")
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        TimeRemain -= 1
        Ticker.Text = FormatTime(TimeRemain)
    End Sub

    Private Sub Ticker_MouseDown(sender As Object, e As MouseEventArgs) Handles Ticker.MouseDown
        If e.Button = MouseButtons.Left Then    'pause and start
            Timer1.Enabled = Not Timer1.Enabled
            If Timer1.Enabled Then
                Ticker.ForeColor = Color.DarkGreen
            Else
                Ticker.ForeColor = Color.Gray
            End If
        ElseIf e.Button = MouseButtons.Right Then 'setting timer
            TimeRemain = Val(InputBox("Please set the timer in minutes: " & vbCrLf & "Type ""0"" to count forward. ", "Timer", "5")) * 60
            Ticker.Text = FormatTime(TimeRemain)
            If TimeRemain < 0 Then Exit Sub
        End If
    End Sub

End Class
