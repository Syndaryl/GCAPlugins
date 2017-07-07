Imports GCA5Engine
Imports System.Drawing
Imports System.Reflection
Imports GCA5.Interfaces

'Any such DLL needs to add References to:
'
'   GCA5Engine
'   GCA5.Interfaces.DLL
'   System.Drawing  (System.Drawing; v4.X) 'for colors and anything drawaing related.
'
'in order to work as a print sheet.
'
Public Class SimpleTextExport
    Implements GCA5.Interfaces.IExportSheet


    Public Event RequestRunSpecificOptions(sender As GCA5.Interfaces.IExportSheet, e As GCA5.Interfaces.DialogOptions_RequestedOptions) Implements GCA5.Interfaces.IExportSheet.RequestRunSpecificOptions

    Private MyOptions As GCA5Engine.SheetOptionsManager
    Private OwnedItemText As String = "* = item is owned by another, its point value is included in the other item."
    Private ShowOwnedMessage As Boolean
    ''' <summary>
    ''' If multiple characters might be printed, this is the value of the Always Ask Me option
    ''' </summary>
    ''' <remarks></remarks>
    Private AlwaysAskMe As Integer

    '******************************************************************************************
    '* All Interface Implementations
    '******************************************************************************************
    Public Sub CreateOptions(Options As GCA5Engine.SheetOptionsManager) Implements GCA5.Interfaces.IExportSheet.CreateOptions
        'This is the routine where all the Options we want to use are created,
        'and where the UI for the Preferences dialog is filled out.
        '
        'This is equivalent to CharacterSheetOptions from previous implementations

        Dim ok As Boolean
        Dim newOption As GCA5Engine.SheetOption

        Dim descFormat As New SheetOptionDisplayFormat
        descFormat.BackColor = SystemColors.Info
        descFormat.CaptionLocalBackColor = SystemColors.Info

        '* Description block at top *
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_Description"
        newOption.Type = GCA5Engine.OptionType.Header
        newOption.UserPrompt = Name & " " & Version
        newOption.DisplayFormat = descFormat
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Description"
        newOption.Type = GCA5Engine.OptionType.Caption
        newOption.UserPrompt = Description
        newOption.DisplayFormat = descFormat
        ok = Options.AddOption(newOption)

        '******************************
        '* Characters 
        '******************************
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_Characters"
        newOption.Type = GCA5Engine.OptionType.Header
        newOption.UserPrompt = "Printing Characters"
        ok = Options.AddOption(newOption)

        'NOTE: Because List is now a 0-based Array, the number of the 
        'DefaultValue and the selected Value is 0-based!
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "OutputCharacters"
        newOption.Type = GCA5Engine.OptionType.ListNumber
        newOption.UserPrompt = "When exporting, how do you want to handle exporting when multiple characters are loaded?"
        newOption.DefaultValue = 0 'first item
        newOption.List = {"Export just the current character", "Export all the characters to the file", "Always ask me what to do"}
        ok = Options.AddOption(newOption)
        AlwaysAskMe = 2

        'NOTE: Because List is now a 0-based Array, the number of the 
        'DefaultValue and the selected Value is 0-based!
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "CharacterSeparator"
        newOption.Type = GCA5Engine.OptionType.ListNumber
        newOption.UserPrompt = "Please select how you'd like to mark the break between characters when printing multiple characters to the file."
        newOption.DefaultValue = 1 'second item
        newOption.List = {"Do nothing", "Print a line of *", "Print a line of =", "Print a line of -", "Use HTML to indicate a horizontal rule"}
        ok = Options.AddOption(newOption)


        '******************************
        '* Included Sections 
        '******************************
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_TextBlocks"
        newOption.Type = GCA5Engine.OptionType.Header
        newOption.UserPrompt = "Sections to Include"
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "ShowMovementBlock"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "Include a block showing all the Environment Move rates?"
        newOption.DefaultValue = True
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "ShowMovementZero"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "When printing the block above, include movement rates even when they're at zero?"
        newOption.DefaultValue = False
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "ShowAllAdditionalStats"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "Include a block showing all additional attributes that aren't hidden?"
        newOption.DefaultValue = False
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "ShowAdditionalStatsAtZero"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "When printing the block above, include attributes even when they're at zero?"
        newOption.DefaultValue = False
        ok = Options.AddOption(newOption)

        '******************************
        '* Other Stuff 
        '******************************
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_Other"
        newOption.Type = GCA5Engine.OptionType.Header
        newOption.UserPrompt = "Other Options"
        ok = Options.AddOption(newOption)

        'NOTE: Because List is now a 0-based Array, the number of the 
        'DefaultValue and the selected Value is 0-based!
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "HeadingStyle"
        newOption.Type = GCA5Engine.OptionType.ListNumber
        newOption.UserPrompt = "Please select the way you'd like to differentiate section headers from their various items."
        newOption.DefaultValue = 1 'second item
        newOption.List = {"Do nothing", "Use a row of dashes under the header", "Use BBCode to mark the header as bold", "Use HTML to mark the header as bold"}
        ok = Options.AddOption(newOption)

        'NOTE: Because List is now a 0-based Array, the number of the 
        'DefaultValue and the selected Value is 0-based!
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "BonusLineStyle"
        newOption.Type = GCA5Engine.OptionType.ListNumber
        newOption.UserPrompt = "Please select the way you'd like to differentiate bonus lines ('Includes: +X from Z') from their related items."
        newOption.DefaultValue = 1 'second item
        newOption.List = {"Do nothing", "Use a tab character preceding them", "Use BBCode to mark them in italic", "Use HTML to mark them in italic", "Tab character and BBCode", "Tab character and HTML"}
        ok = Options.AddOption(newOption)

        'NOTE: Because List is now a 0-based Array, the number of the 
        'DefaultValue and the selected Value is 0-based!
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "PointsLineStyle"
        newOption.Type = GCA5Engine.OptionType.ListNumber
        newOption.UserPrompt = "Please select the way you'd like to differentiate the Point Summary line from the surrounding text."
        newOption.DefaultValue = 1 'second item
        newOption.List = {"Do nothing", "Print it as multiple lines with a header", "Use BBCode to mark them in bold", "Use HTML to mark them in bold", "Use BBCode to mark them in italic", "Use HTML to mark them in italic", "Use BBCode to mark them in bold and italic", "Use HTML to mark them in bold and italic"}
        ok = Options.AddOption(newOption)

    End Sub

    Public Sub UpgradeOptions(Options As GCA5Engine.SheetOptionsManager) Implements GCA5.Interfaces.IExportSheet.UpgradeOptions
        'This is called only when a particular plug-in is loaded the first time,
        'and before SetOptions.

        'I don't do anything with this.
    End Sub

    Public ReadOnly Property Description As String Implements GCA5.Interfaces.IExportSheet.Description
        Get
            Return "Exports a simple text file from currently loaded GCA5 characters."
        End Get
    End Property

    Public ReadOnly Property Name As String Implements GCA5.Interfaces.IExportSheet.Name
        Get
            Return "Simple Text Export"
        End Get
    End Property

    Public Function SupportedFileTypeFilter() As String Implements GCA5.Interfaces.IExportSheet.SupportedFileTypeFilter
        Return "Text files (*.txt)|*.txt"
    End Function

    Public ReadOnly Property Version As String Implements GCA5.Interfaces.IExportSheet.Version
        Get
            Return AutoFindVersion()
        End Get
    End Property

    Public Function GenerateExport(Party As GCA5Engine.Party, TargetFilename As String, Options As GCA5Engine.SheetOptionsManager) As Boolean Implements GCA5.Interfaces.IExportSheet.GenerateExport
        Dim PrintMults As Boolean = False

        'This creates the export file on disk.

        'Set our Options to the stored values we've just been given
        MyOptions = Options

        'set our default PrintMults
        'newOption.List = {"Export just the current character", "Export all the characters to the file", "Always ask me what to do"}
        If MyOptions.Value("OutputCharacters") = 1 Then
            PrintMults = True
        End If

        'Here, if you needed it, you'd create the RunSpecificOptions that you need the user to set,
        'and then you'd raise the event to get those options from the user.
        'Dim RunSpecificOptions As New GCA5Engine.SheetOptionsManager("RunSpecificOptions For " & Name)
        'RaiseEvent RequestRunSpecificOptions(RunSpecificOptions)

        'if there are multiple characters...
        If Party.Characters.Count > 1 Then
            '... and if we're supposed to ask what to do...
            If MyOptions.Value("OutputCharacters") = AlwaysAskMe Then
                '... ask what to do.

                'In this case, this would actually be better served with a message box, and that's what's commented out below this block,
                'but this shows the idea behind using the RunSpecificOptions.

                '*****
                '* We need to get more options
                '*****
                Dim ok As Boolean
                Dim newOption As GCA5Engine.SheetOption
                Dim RunSpecificOptions As New GCA5Engine.SheetOptionsManager("RunSpecificOptions For " & Name)

                'Create the options
                newOption = New GCA5Engine.SheetOption
                newOption.Name = "Header_Characters"
                newOption.Type = GCA5Engine.OptionType.Header
                newOption.UserPrompt = "Printing Characters"
                ok = RunSpecificOptions.AddOption(newOption)

                newOption = New GCA5Engine.SheetOption
                newOption.Name = "OutputCharacters"
                newOption.Type = GCA5Engine.OptionType.ListNumber
                newOption.UserPrompt = "When exporting, how do you want to handle exporting when multiple characters are loaded?"
                newOption.DefaultValue = 0 'first item
                newOption.List = {"Export just the current character", "Export all the characters to the file"}
                ok = RunSpecificOptions.AddOption(newOption)

                'Create the event object that will carry our options and tell us later if the dialog was canceled
                Dim e As New GCA5.Interfaces.DialogOptions_RequestedOptions
                e.RunSpecificOptions = RunSpecificOptions

                'Raise the event to get the user input
                RaiseEvent RequestRunSpecificOptions(Me, e)

                'If user canceled, we abort
                If e.Canceled Then
                    Return False
                End If

                If RunSpecificOptions.Value("OutputCharacters") = 1 Then
                    'print just one character
                    PrintMults = True
                End If

                '*****
                '* The commented block below could replace the code above, if the input needed is as simple as the example used here.
                '*****
                'Select Case MsgBox("By default, only the Current character will be exported. Do you want to export ALL the loaded characters instead?", MsgBoxStyle.YesNoCancel, "Export All Characters")
                '    Case MsgBoxResult.Cancel
                '        'cancel, abort export
                '        Return False
                '    Case MsgBoxResult.No
                '        'No, print just the one character
                '        PrintMults = False
                '    Case Else
                '        'Yes, print all characters
                '        PrintMults = True
                'End Select
            End If
        Else
            PrintMults = False
        End If


        'This is defined in GCA5Engine, and is a text file writer that outputs UTF-8 text files.
        Dim fw As New FileWriter

        'Creates a string buffer for the file, but doesn't actually open and write it until FileClose is called.
        fw.FileOpen(TargetFilename)

        If PrintMults Then
            'Export every character to this file
            For Each CurChar As GCACharacter In Party.Characters
                PrintCharacter(CurChar, fw)

                'newOption.List = {"Do nothing", "Print a line of *", "Print a line of =", "Print a line of -", "Use HTML to indicate a horizontal rule"}
                Select Case MyOptions.Value("CharacterSeparator")
                    Case 1 'line of *
                        fw.Paragraph(StrDup(60, "*"))
                    Case 2 'line of =
                        fw.Paragraph(StrDup(60, "="))
                    Case 3 'line of -
                        fw.Paragraph(StrDup(60, "-"))
                    Case 4 'html
                        fw.Paragraph("<hr />")
                    Case Else 'do nothing
                End Select
                fw.Paragraph("")
            Next
        Else
            'just print Current character
            PrintCharacter(Party.Current, fw)
        End If


        'Save all we've written to the file and quit.
        Try
            fw.FileClose()
        Catch ex As Exception
            'problem encountered
            Notify(Name & ": " & Err.Number & ": " & ex.Message & vbCrLf & "Stack Trace: " & vbCrLf & ex.StackTrace, Priority.Red)
            Return False
        End Try

        'all good
        Return True
    End Function



    '******************************************************************************************
    '* All Internal Routines
    '******************************************************************************************
    Public Function AutoFindVersion() As String
        Dim longFormVersion As String = ""

        Dim currentDomain As AppDomain = AppDomain.CurrentDomain
        'Provide the current application domain evidence for the assembly.
        'Load the assembly from the application directory using a simple name.
        currentDomain.Load("SimpleTextExport")

        'Make an array for the list of assemblies.
        Dim assems As [Assembly]() = currentDomain.GetAssemblies()

        'List the assemblies in the current application domain.
        'Echo("List of assemblies loaded in current appdomain:")
        Dim assem As [Assembly]
        'Dim co As New ArrayList
        For Each assem In assems
            If assem.FullName.StartsWith("SimpleTextExport") Then
                Dim parts(0) As String
                parts = assem.FullName.Split(",")
                'name and version are the first two parts
                longFormVersion = parts(1)
                'Version=1.2.3.4
                parts = longFormVersion.Split("=")
                Return parts(1)
            End If
        Next assem

        Return longFormVersion
    End Function

    Private Sub PrintCharacter(CurChar As GCACharacter, fw As FileWriter)
        'This does all the work of printing each block of the current character
        Dim ListNum As Integer

        fw.Paragraph("Name: " & CurChar.Name)
        fw.Paragraph("Player: " & CurChar.Player)
        fw.Paragraph("Race: " & CurChar.Race)
        If CurChar.Appearance <> "" Then
            fw.Paragraph("Appearance: " & CurChar.Appearance)
        End If
        If CurChar.Height <> "" Then
            fw.Paragraph("Height: " & CurChar.Height)
        End If
        If CurChar.Weight <> "" Then
            fw.Paragraph("Weight: " & CurChar.Weight)
        End If
        If CurChar.Age <> "" Then
            fw.Paragraph("Age: " & CurChar.Age)
        End If
        fw.Paragraph("")

        PrintAttributes(CurChar, fw) 'includes bonuses
        PrintLiftAndDamage(CurChar, fw)
        PrintMovement(CurChar, fw)

        If MyOptions.Value("ShowAllAdditionalStats") Then
            PrintAdditionalAttributes(CurChar, fw)
        End If

        PrintSocialBackground(CurChar, fw)

        For ListNum = Ads To Spells
            PrintTraitList(ListNum, CurChar, fw)

            If ListNum = Quirks Then
                'we're breaking order to print templates after quirks
                PrintTraitList(Templates, CurChar, fw)
            End If
        Next
        If ShowOwnedMessage Then
            fw.Paragraph("")
            fw.Paragraph(OwnedItemText)
        End If

        PrintPointSummary(CurChar, fw)

        PrintMeleeAttacks(CurChar, fw)
        PrintRangedAttacks(CurChar, fw)

        PrintTraitList(Equipment, CurChar, fw)

        PrintDescription(CurChar, fw)
        PrintNotes(CurChar, fw)

    End Sub

    Sub PrintAttributes(CurChar As GCACharacter, fw As FileWriter)
        'Dim tmp, work As String
        Dim ListName As String
        Dim curItem As GCATrait

        ListName = "Attributes [" & CurChar.Cost(Stats) & "]"
        PrintHeading(ListName, fw)

        Dim StatNames As New Collection
        StatNames.Add("ST")
        StatNames.Add("DX")
        StatNames.Add("IQ")
        StatNames.Add("HT")

        For Each s As String In StatNames
            curItem = CurChar.ItemByNameAndExt(s, Stats)
            If curItem IsNot Nothing Then
                PrintTrait(curItem, fw, 0)
            End If
        Next

        fw.Paragraph("")

        StatNames.Clear()
        StatNames.Add("Hit Points")
        StatNames.Add("Will")
        StatNames.Add("Perception")
        StatNames.Add("Fatigue Points")

        For Each s As String In StatNames
            curItem = CurChar.ItemByNameAndExt(s, Stats)
            If curItem IsNot Nothing Then
                PrintTrait(curItem, fw, 0)
            End If
        Next

        fw.Paragraph("")

    End Sub

    Sub PrintLiftAndDamage(CurChar As GCACharacter, fw As FileWriter)
        Dim tmp As String
        Dim curItem As GCATrait

        Dim StatNames As New Collection
        StatNames.Add("Basic Lift")

        For Each s As String In StatNames
            curItem = CurChar.ItemByNameAndExt(s, Stats)
            If curItem IsNot Nothing Then
                tmp = curItem.DisplayName & " " & curItem.DisplayScore

                If curItem.TagItem("points") <> 0 Then
                    tmp = tmp & " [" & curItem.TagItem("points") & "]"
                End If

                If curItem.TagItem("bonuslist") <> "" Then
                    tmp = tmp & " ("
                    tmp = tmp & "Includes: "
                    tmp = tmp & curItem.TagItem("bonuslist")
                    tmp = tmp & ")"
                End If
                fw.Paragraph(tmp)
            End If
        Next

        fw.Paragraph("Damage " & CurChar.BaseTH & "/" & CurChar.BaseSW)

        fw.Paragraph("")
    End Sub

    Sub PrintMovement(CurChar As GCACharacter, fw As FileWriter)
        Dim tmp As String
        Dim curItem As GCATrait
        Dim BaseScore, Score As Double

        Dim StatNames As New Collection
        StatNames.Add("Basic Speed")
        StatNames.Add("Basic Move")

        For Each s As String In StatNames
            curItem = CurChar.ItemByNameAndExt(s, Stats)
            If curItem IsNot Nothing Then
                tmp = curItem.DisplayName & " " & curItem.DisplayScore

                If curItem.TagItem("points") <> 0 Then
                    tmp = tmp & " [" & curItem.TagItem("points") & "]"
                End If

                If curItem.TagItem("bonuslist") <> "" Then
                    tmp = tmp & " ("
                    tmp = tmp & "Includes: "
                    tmp = tmp & curItem.TagItem("bonuslist")
                    tmp = tmp & ")"
                End If
                fw.Paragraph(tmp)
            End If
        Next

        fw.Paragraph("")

        If MyOptions.Value("ShowMovementBlock") Then
            StatNames = New Collection
            StatNames.Add("Air Move")
            StatNames.Add("Ground Move")
            StatNames.Add("Space Move")
            StatNames.Add("Tunneling Move")
            StatNames.Add("Water Move")

            For Each s As String In StatNames
                curItem = CurChar.ItemByNameAndExt(s, Stats)
                If curItem IsNot Nothing Then
                    BaseScore = curItem.TagItem("basescore")
                    Score = curItem.TagItem("score")

                    If (BaseScore <> 0 AndAlso Score <> 0) OrElse MyOptions.Value("ShowMovementZero") = True Then

                        If BaseScore = Score Then
                            tmp = curItem.DisplayName & " " & Score
                        Else
                            tmp = curItem.DisplayName & " " & BaseScore & "/" & Score
                        End If

                        If curItem.TagItem("bonuslist") <> "" Then
                            tmp = tmp & " ("
                            tmp = tmp & "Includes: "
                            tmp = tmp & curItem.TagItem("bonuslist")
                            tmp = tmp & ")"
                        End If
                        fw.Paragraph(tmp)

                    End If
                End If
            Next

            fw.Paragraph("")
        End If
    End Sub

    Sub PrintAdditionalAttributes(CurChar As GCACharacter, fw As FileWriter)
        Dim tmp As String
        Dim i As Integer
        Dim curItem As GCATrait

        For i = 1 To CurChar.Items.Count
            curItem = CurChar.Items.Item(i)

            If curItem.ItemType = Stats Then
                If curItem.TagItem("display") = "" Then
                    If curItem.TagItem("hide") = "" Then

                        If (curItem.Score <> 0) OrElse MyOptions.Value("ShowAdditionalStatsAtZero") = True Then

                            tmp = curItem.DisplayName & " " & curItem.DisplayScore

                            If curItem.TagItem("points") <> 0 Then
                                tmp = tmp & " [" & curItem.TagItem("points") & "]"
                            End If

                            If curItem.TagItem("bonuslist") <> "" Then
                                tmp = tmp & " ("
                                tmp = tmp & "Includes: "
                                tmp = tmp & curItem.TagItem("bonuslist")
                                tmp = tmp & ")"
                            End If
                            fw.Paragraph(tmp)

                        End If

                    End If
                End If
            End If
        Next

        fw.Paragraph("")
    End Sub

    Sub PrintSocialBackground(CurChar As GCACharacter, fw As FileWriter)
        Dim tmp As String
        Dim curItem As GCATrait

        PrintHeading("Social Background", fw)

        '* TL
        Dim StatNames As New Collection
        StatNames.Add("Tech Level")

        fw.AddText("TL: ")
        For Each s As String In StatNames
            curItem = CurChar.ItemByNameAndExt(s, Stats)
            If curItem IsNot Nothing Then
                tmp = curItem.DisplayScore

                If curItem.TagItem("points") <> 0 Then
                    tmp = tmp & " [" & curItem.TagItem("points") & "]"
                End If

                fw.Paragraph(tmp)
            End If
        Next

        '* Cultures
        If CurChar.Count(Cultures) > 0 Then
            tmp = ""
            For Each curItem In CurChar.ItemsByType(Cultures)
                If curItem.TagItem("hide") = "" Then 'not hidden
                    If tmp = "" Then
                        tmp = curItem.DisplayName
                    Else
                        tmp = tmp & "; " & curItem.DisplayName
                    End If
                    tmp = tmp & " [" & curItem.TagItem("points") & "]"
                End If
            Next
            fw.Paragraph("Cultural Familiarities: " & tmp & ".")
        End If

        '* Languages
        If CurChar.Count(Languages) > 0 Then
            tmp = ""
            For Each curItem In CurChar.ItemsByType(Languages)
                If curItem.TagItem("hide") = "" Then 'not hidden
                    If tmp = "" Then
                        tmp = curItem.DisplayName
                    Else
                        tmp = tmp & "; " & curItem.DisplayName
                    End If
                    tmp = tmp & " [" & curItem.TagItem("points") & "]"
                End If
            Next
            fw.Paragraph("Languages: " & tmp & ".")
        End If
        fw.Paragraph("")

    End Sub
    Sub PrintTraitList(ListNum As Integer, CurChar As GCACharacter, fw As FileWriter)
        Dim ListName As String

        If CurChar.Count(ListNum) <= 0 Then Return

        ListName = ReturnListNameNew(ListNum)
        If ListNum = Equipment Then
            ListName = ListName & " [" & Format(CurChar.Cost(ListNum), "Currency") & "]"
        Else
            ListName = ListName & " [" & CurChar.Cost(ListNum) & "]"
        End If
        PrintHeading(ListName, fw)

        For Each curItem In CurChar.ItemsByType(ListNum)
            If curItem.ParentKey <> "" Then
                'don't show this one, because it's a child of another trait
                Continue For
            End If

            'don't show items that are components
            If curItem.TagItem("keep") <> "" Then
                Continue For
            End If

            PrintTrait(curItem, fw, 0)
        Next

        fw.Paragraph("")
    End Sub
    Private Sub PrintTrait(curItem As GCATrait, fw As FileWriter, IndentLevel As Integer)
        Dim tmp, work As String
        Dim sep As String

        sep = "; " ' vbTab '"  "
        Dim indent As String = ""
        If IndentLevel > 0 Then
            indent = StrDup(IndentLevel * 3, " ")
        End If

        Select Case curItem.ItemType
            Case TraitTypes.Attributes
                tmp = curItem.DisplayName & " " & curItem.DisplayScore

                If curItem.TagItem("points") <> 0 Then
                    tmp = tmp & " [" & curItem.TagItem("points") & "]"
                End If

                'we print bonuses inline with attributes only
                If curItem.TagItem("bonuslist") <> "" Then
                    work = ""
                    work = work & "("
                    work = work & "Includes: "
                    work = work & curItem.TagItem("bonuslist")
                    work = work & ")"

                    Select Case MyOptions.Value("BonusLineStyle")
                        Case 1 'tab
                            work = vbTab & work
                        Case 2 'bbcode italic
                            work = " [i]" & work & "[/i]"
                        Case 3 'html italic
                            work = " <em>" & work & "</em>"
                        Case 4 'tab and bbcode
                            work = vbTab & "[i]" & work & "[/i]"
                        Case 5 'tab and html
                            work = vbTab & "<em>" & work & "</em>"
                        Case Else 'do nothing
                            work = " " & work
                    End Select
                    tmp = tmp & work
                End If

                fw.Paragraph(indent & tmp)

            Case TraitTypes.Skills
                tmp = curItem.DisplayName

                If curItem.Mods.Count > 0 Then
                    tmp = tmp & curItem.ExpandedModCaptions(False)
                End If

                tmp = tmp & " " & curItem.TagItem("type")
                tmp = tmp & " - "

                Select Case curItem.TagItem("sd")
                    Case "1"
                        'technique
                        tmp = tmp & " " & curItem.Level

                    Case "2"
                        'combo
                        tmp = tmp & " " & curItem.TagItem("combolevel")

                    Case Else
                        'normal skill
                        work = curItem.TagItem("stepoff")
                        If work <> "" Then
                            tmp = tmp & work
                            work = curItem.TagItem("step")
                            If work <> "" Then
                                tmp = tmp & work
                            Else
                                tmp = tmp & "?"
                            End If
                        Else
                            tmp = tmp & "?+?"
                        End If
                        tmp = tmp & " " & curItem.Level

                End Select

                tmp = tmp & " [" & curItem.TagItem("points")
                If curItem.TagItem("owned") = "yes" Then
                    tmp = tmp & "*"
                    ShowOwnedMessage = True
                End If
                tmp = tmp & "]"

                fw.Paragraph(indent & tmp)

            Case TraitTypes.Spells
                tmp = curItem.DisplayName

                tmp = tmp & " " & curItem.TagItem("type")
                tmp = tmp & " - "

                work = curItem.TagItem("stepoff")
                If work <> "" Then
                    tmp = tmp & work
                    work = curItem.TagItem("step")
                    If work <> "" Then
                        tmp = tmp & work
                    Else
                        tmp = tmp & "?"
                    End If
                Else
                    tmp = tmp & "?+?"
                End If
                tmp = tmp & " " & curItem.Level

                tmp = tmp & " [" & curItem.TagItem("points")
                If curItem.TagItem("owned") = "yes" Then
                    tmp = tmp & "*"
                    ShowOwnedMessage = True
                End If
                tmp = tmp & "]"

                fw.Paragraph(indent & tmp)

            Case TraitTypes.Templates
                tmp = curItem.DisplayName

                tmp = tmp & " [" & curItem.TagItem("points")

                If curItem.TagItem("owned") = "yes" Then
                    tmp = tmp & "*"
                    ShowOwnedMessage = True
                End If

                tmp = tmp & "]"

                fw.Paragraph(indent & tmp)

            Case TraitTypes.Equipment
                tmp = curItem.DisplayName
                tmp = tmp & sep & "Qty:" & curItem.TagItem("count")
                tmp = tmp & sep & "Wgt:" & curItem.TagItem("weight")
                tmp = tmp & sep & Format(curItem.TagItem("cost"), "Currency")

                If curItem.TagItem("owned") = "yes" Then
                    tmp = tmp & "*"
                    ShowOwnedMessage = True
                End If

                'tmp = tmp & "]"

                fw.Paragraph(indent & tmp)

            Case Else
                tmp = curItem.DisplayName

                tmp = tmp & " [" & curItem.TagItem("points")

                If curItem.TagItem("owned") = "yes" Then
                    tmp = tmp & "*"
                    ShowOwnedMessage = True
                End If

                tmp = tmp & "]"

                fw.Paragraph(indent & tmp)

        End Select

        'print bonuses
        Select Case curItem.ItemType
            Case TraitTypes.Attributes
                'handled inline above

            Case TraitTypes.Spells
                'don't print spell bonuses

            Case Else
                If curItem.TagItem("bonuslist") <> "" Then
                    Select Case MyOptions.Value("BonusLineStyle")
                        Case 1 'tab
                            fw.Paragraph(vbTab & "Includes: " & curItem.TagItem("bonuslist"))
                        Case 2 'bbcode italic
                            fw.Paragraph("[i]Includes: " & curItem.TagItem("bonuslist") & "[/i]")
                        Case 3 'html italic
                            fw.Paragraph("<em>Includes: " & curItem.TagItem("bonuslist") & "</em>")
                        Case 4 'tab and bbcode
                            fw.Paragraph(vbTab & "[i]Includes: " & curItem.TagItem("bonuslist") & "[/i]")
                        Case 5 'tab and html
                            fw.Paragraph(vbTab & "<em>Includes: " & curItem.TagItem("bonuslist") & "</em>")
                        Case Else 'do nothing
                            fw.Paragraph("Includes: " & curItem.TagItem("bonuslist"))
                    End Select
                End If
        End Select

        'always include children
        'check for children
        If curItem.Children.Count > 0 Then
            For Each childItem As GCATrait In curItem.Children
                PrintTrait(childItem, fw, IndentLevel + 1)
            Next
        End If
    End Sub


    Sub PrintMeleeAttacks(CurChar As GCACharacter, fw As FileWriter)
        Dim curItem As GCATrait
        Dim CurMode, ModeCount As Integer
        Dim tmp, work, sep, DamageText, NotesText As String

        PrintHeading("Melee Attacks", fw)

        For i = 1 To CurChar.Items.Count
            curItem = CurChar.Items.Item(i)
            If curItem.DamageModeTagItemCount("charreach") > 0 Then
                If curItem.TagItem("hide") = "" Then 'not hidden

                    'base info on a line, then mode info on other lines, unless just one mode then all one line
                    ModeCount = curItem.DamageModeTagItemCount("charreach")
                    tmp = ""
                    sep = "; " ' vbTab '"  "

                    'print the name
                    tmp = tmp & curItem.DisplayName

                    If ModeCount > 1 Then
                        'we're doing separate lines for each mode
                        fw.Paragraph(tmp)
                    End If
                    'fw.Paragraph(tmp)

                    '* Now do the modes
                    CurMode = curItem.DamageModeTagItemAt("charreach")
                    Do
                        'this mode is hand!

                        DamageText = curItem.DamageModeTagItem(CurMode, "chardamage")
                        If curItem.DamageModeTagItem(CurMode, "chararmordivisor") <> "" Then
                            work = curItem.DamageModeTagItem(CurMode, "chararmordivisor")
                            If work = "!" Then
                                work = ChrW(8734)
                            End If
                            DamageText = DamageText & " (" & work & ")"
                        End If
                        DamageText = DamageText & " " & curItem.DamageModeTagItem(CurMode, "chardamtype")
                        If curItem.DamageModeTagItem(CurMode, "charradius") <> "" Then
                            DamageText = DamageText & " (" & curItem.DamageModeTagItem(CurMode, "charradius") & ")"
                        End If

                        If ModeCount > 1 Then
                            'we're doing separate lines for each mode
                            tmp = "   " 'vbTab '"      "

                            'print the mode name
                            tmp = tmp & curItem.DamageModeName(CurMode)
                        End If

                        'print the damage
                        tmp = tmp & sep & "Dam:" & DamageText

                        'print the reach
                        tmp = tmp & sep & "Reach:" & curItem.DamageModeTagItem(CurMode, "charreach")

                        'print the skill
                        tmp = tmp & sep & "Skill:" & curItem.DamageModeTagItem(CurMode, "charskillused")

                        'print the level
                        tmp = tmp & sep & "Level:" & curItem.DamageModeTagItem(CurMode, "charskillscore")

                        'print the parry
                        If curItem.DamageModeTagItem(CurMode, "charparryscore") <> "" Then
                            tmp = tmp & sep & "Parry:" & curItem.DamageModeTagItem(CurMode, "charparryscore")
                        End If

                        'print the ST
                        If curItem.DamageModeTagItem(CurMode, "charminst") <> "" Then
                            tmp = tmp & sep & "ST:" & curItem.DamageModeTagItem(CurMode, "charminst")
                        End If

                        'print the lc
                        tmp = tmp & sep & "LC:" & curItem.DamageModeTagItem(CurMode, "lc")

                        'print the notes
                        NotesText = curItem.DamageModeTagItem(CurMode, "notes")
                        If NotesText <> "" Then
                            tmp = tmp & sep & "Notes:" & NotesText
                        End If

                        fw.Paragraph(tmp)

                        If curItem.Modes.Mode(CurMode).ItemNotesText <> "" Then
                            fw.Paragraph(vbTab & "Notes: " & curItem.Modes.Mode(CurMode).ItemNotesText)
                        End If

                        CurMode = curItem.DamageModeTagItemAt("charreach", CurMode + 1)
                    Loop While CurMode > 0

                End If
            End If
        Next

        fw.Paragraph("")
    End Sub
    Sub PrintRangedAttacks(CurChar As GCACharacter, fw As FileWriter)
        Dim curItem As GCATrait
        Dim CurMode, ModeCount As Integer
        Dim tmp, work, sep, DamageText, RangeText, NotesText As String

        PrintHeading("Ranged Attacks", fw)

        For i = 1 To CurChar.Items.Count
            curItem = CurChar.Items.Item(i)
            If curItem.DamageModeTagItemCount("charrangemax") > 0 Then
                If curItem.TagItem("hide") = "" Then 'not hidden

                    'base info on a line, then mode info on other lines, unless just one mode then all one line
                    ModeCount = curItem.DamageModeTagItemCount("charrangemax")
                    tmp = ""
                    sep = "; " ' vbTab '"  "

                    'print the name
                    tmp = tmp & curItem.DisplayName

                    If ModeCount > 1 Then
                        'we're doing separate lines for each mode
                        fw.Paragraph(tmp)
                    End If
                    'fw.Paragraph(tmp)

                    '* Now do the modes
                    CurMode = curItem.DamageModeTagItemAt("charrangemax")
                    Do
                        'this mode is hand!

                        DamageText = curItem.DamageModeTagItem(CurMode, "chardamage")
                        If curItem.DamageModeTagItem(CurMode, "chararmordivisor") <> "" Then
                            work = curItem.DamageModeTagItem(CurMode, "chararmordivisor")
                            If work = "!" Then
                                work = ChrW(8734)
                            End If
                            DamageText = DamageText & " (" & work & ")"
                        End If
                        DamageText = DamageText & " " & curItem.DamageModeTagItem(CurMode, "chardamtype")
                        If curItem.DamageModeTagItem(CurMode, "charradius") <> "" Then
                            DamageText = DamageText & " (" & curItem.DamageModeTagItem(CurMode, "charradius") & ")"
                        End If

                        RangeText = curItem.DamageModeTagItem(CurMode, "charrangehalfdam")
                        If RangeText = "" Then
                            RangeText = curItem.DamageModeTagItem(CurMode, "charrangemax")
                        Else
                            RangeText = RangeText & " / " & curItem.DamageModeTagItem(CurMode, "charrangemax")
                        End If


                        If ModeCount > 1 Then
                            'we're doing separate lines for each mode
                            tmp = "   " 'vbTab '"      "

                            'print the mode name
                            tmp = tmp & curItem.DamageModeName(CurMode)
                        End If

                        'print the damage
                        tmp = tmp & sep & "Dam:" & DamageText

                        'print the acc
                        tmp = tmp & sep & "Acc:" & curItem.DamageModeTagItem(CurMode, "characc")

                        'print the range
                        tmp = tmp & sep & "Range:" & RangeText

                        'print the rof
                        tmp = tmp & sep & "RoF:" & curItem.DamageModeTagItem(CurMode, "charrof")

                        'print the shots
                        tmp = tmp & sep & "Shots:" & curItem.DamageModeTagItem(CurMode, "charshots")

                        'print the level
                        tmp = tmp & sep & "Level:" & curItem.DamageModeTagItem(CurMode, "charskillscore")

                        'print the ST
                        If curItem.DamageModeTagItem(CurMode, "charminst") <> "" Then
                            tmp = tmp & sep & "ST:" & curItem.DamageModeTagItem(CurMode, "charminst")
                        End If

                        'print the bulk
                        If curItem.DamageModeTagItem(CurMode, "bulk") <> "" Then
                            tmp = tmp & sep & "ST:" & curItem.DamageModeTagItem(CurMode, "bulk")
                        End If

                        'print the rcl
                        If curItem.DamageModeTagItem(CurMode, "charrcl") <> "" Then
                            tmp = tmp & sep & "ST:" & curItem.DamageModeTagItem(CurMode, "charrcl")
                        End If

                        'print the lc
                        tmp = tmp & sep & "LC:" & curItem.DamageModeTagItem(CurMode, "lc")

                        'print the notes
                        NotesText = curItem.DamageModeTagItem(CurMode, "notes")
                        If NotesText <> "" Then
                            tmp = tmp & sep & "Notes:" & NotesText
                        End If

                        fw.Paragraph(tmp)

                        If curItem.Modes.Mode(CurMode).ItemNotesText <> "" Then
                            fw.Paragraph(vbTab & "Notes: " & curItem.Modes.Mode(CurMode).ItemNotesText)
                        End If

                        CurMode = curItem.DamageModeTagItemAt("charrangemax", CurMode + 1)
                    Loop While CurMode > 0

                End If
            End If
        Next

        fw.Paragraph("")
    End Sub
    Sub PrintDescription(CurChar As GCACharacter, fw As FileWriter)

        If CurChar.Description = "" Then Return

        PrintHeading("Description", fw)

        fw.Paragraph(PlainText(CurChar.Description))
        fw.Paragraph("")

    End Sub
    Sub PrintNotes(CurChar As GCACharacter, fw As FileWriter)

        If CurChar.Notes = "" Then Return

        PrintHeading("Notes", fw)

        fw.Paragraph(PlainText(CurChar.Notes))
        fw.Paragraph("")

    End Sub
    Sub PrintPointSummary(CurChar As GCACharacter, fw As FileWriter)
        Dim tmp, sep As String

        tmp = ""
        sep = ""
        If MyOptions.Value("PointsLineStyle") = 1 Then
            sep = vbCrLf
            tmp = "Points Summary"

            Select Case MyOptions.Value("HeadingStyle")
                Case 1 'dashes
                    tmp = tmp & sep
                    tmp = tmp & StrDup(40, "-") & sep
                Case 2 'bbcode bold
                    tmp = "[b]" & tmp & "[/b]" & sep
                Case 3 'html bold
                    tmp = "<strong>" & tmp & "</strong>" & sep
                Case Else 'do nothing
                    tmp = tmp & sep
            End Select
        End If

        tmp = tmp & "Attributes/Secondary Characteristics [" & CurChar.Cost(Stats) & "] " & sep
        tmp = tmp & "Advantages/Perks/TL/Languages/Cultural Familiarities [" & (CurChar.Cost(Ads) + CurChar.Cost(Perks) + CurChar.Cost(Templates)) & "] " & sep
        tmp = tmp & "Disadvantages/Quirks [" & (CurChar.Cost(Disads) + CurChar.Cost(Quirks)) & "] " & sep
        tmp = tmp & "Skills/Techniques/Spells [" & (CurChar.Cost(Skills) + CurChar.Cost(Spells)) & "] " & sep
        tmp = tmp & "= Total [" & CurChar.TotalCost & "] "

        Select Case MyOptions.Value("PointsLineStyle")
            Case 1 'mult lines w/header
            Case 2 'bbcode bold
                tmp = "[b]" & tmp & "[/b]"
            Case 3 'html bold
                tmp = "<strong>" & tmp & "</strong>"
            Case 4 'bbcode italic
                tmp = "[i]" & tmp & "[/i]"
            Case 5 'html italic
                tmp = "<em>" & tmp & "</em>"
            Case 6 'bbcode bold italic
                tmp = "[b][i]" & tmp & "[/i][/b]"
            Case 7 'html bold italic
                tmp = "<strong><em>" & tmp & "</em></strong>"
            Case Else 'do nothing
        End Select

        fw.Paragraph(tmp)
        fw.Paragraph("")
    End Sub

    Sub PrintHeading(Heading As String, fw As FileWriter)
        Select Case MyOptions.Value("HeadingStyle")
            Case 1 'dashes
                fw.Paragraph(Heading)
                fw.Paragraph(StrDup(40, "-"))
            Case 2 'bbcode bold
                fw.Paragraph("[b]" & Heading & "[/b]")
            Case 3 'html bold
                fw.Paragraph("<strong>" & Heading & "</strong>")
            Case Else 'do nothing
                fw.Paragraph(Heading)
        End Select
    End Sub

    Public Function PreferredFilterIndex() As Integer Implements IExportSheet.PreferredFilterIndex
        Throw New NotImplementedException()
    End Function

    Public Function PreviewOptions(Options As SheetOptionsManager) As Boolean Implements IExportSheet.PreviewOptions
        Throw New NotImplementedException()
    End Function
End Class
