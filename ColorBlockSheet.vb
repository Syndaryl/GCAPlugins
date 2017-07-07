Imports GCA5Engine
Imports C1.C1Preview
Imports System.Drawing
Imports System.Reflection

'Any such DLL needs to add References to:
'
'   System.Drawing  (System.Drawing; v4.X)
'   C1.C1Report.4   (ComponentOne Reports; v4.X)
'   GCA5Engine.DLL
'   GCA5.Interfaces.DLL
'
'in order to work as a print sheet.
'
Public Class ColorBlockSheet
    Implements GCA5.Interfaces.IPrinterSheet

    Private Const Hand As Integer = 1
    Private Const Ranged As Integer = 2

    Private NewPage As Boolean = True
    Private ShadeAltLines As Boolean = True
    Private LineBefore As Boolean = True

    Private MarginLeft, MarginRight, MarginTop, MarginBottom As Double
    Private PageWidth, PageHeight As Double

    Private MyOptions As GCA5Engine.SheetOptionsManager
    Private ShowHidden() As Boolean
    Private ShowComponents() As Boolean
    Private BonusesAsFootnotes() As Boolean
    Private BonusesInline() As Boolean

    Private Fields As ArrayList '0 based

    Dim IconColLeft As Double
    Dim IconColWidth As Double
    Dim IndentStepSize As Double = 0.125
    Dim ComponentDrawCol As Integer

    Dim FooterHeight As Double = 0.25

    Dim pointBoxWidth As Double = 7 / 16

    Dim numCols As Integer = 0
    Dim PointsCol As Integer = 0
    Dim ColWidths(0) As Double
    Dim ColAligns(0) As AlignHorzEnum
    Dim AltColLefts(0) As Double 'for combos
    Dim AltColWidths(0) As Double 'for combos

    Dim SubHeads(0) As String


    Private MyChar As GCACharacter
    Public Property Character As GCACharacter
        Get
            Return MyChar
        End Get
        Set(value As GCACharacter)
            MyChar = value
        End Set
    End Property

    Private WithEvents MyPrintDoc As C1.C1Preview.C1PrintDocument
    Public Property PrintDoc As C1.C1Preview.C1PrintDocument
        Get
            Return MyPrintDoc
        End Get
        Set(value As C1.C1Preview.C1PrintDocument)
            MyPrintDoc = value
        End Set
    End Property

    Public ReadOnly Property Name() As String Implements GCA5.Interfaces.IPrinterSheet.Name
        Get
            Return "Color Block Sheet"
        End Get
    End Property
    Public ReadOnly Property Description() As String Implements GCA5.Interfaces.IPrinterSheet.Description
        Get
            Return "Prints a sheet in the style of the blocks used in the Compact View of GCA5."
        End Get
    End Property
    Public ReadOnly Property Version() As String Implements GCA5.Interfaces.IPrinterSheet.Version
        Get
            'Return "1.0.0.13"
            Return AutoFindVersion()
        End Get
    End Property
    Public Function AutoFindVersion() As String
        Dim longFormVersion As String = ""

        Dim currentDomain As AppDomain = AppDomain.CurrentDomain
        'Provide the current application domain evidence for the assembly.
        'Load the assembly from the application directory using a simple name.
        currentDomain.Load("ColorBlockSheet")

        'Make an array for the list of assemblies.
        Dim assems As [Assembly]() = currentDomain.GetAssemblies()

        'List the assemblies in the current application domain.
        'Echo("List of assemblies loaded in current appdomain:")
        Dim assem As [Assembly]
        'Dim co As New ArrayList
        For Each assem In assems
            If assem.FullName.StartsWith("ColorBlockSheet") Then
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


    Public Sub New()
        'For testing, I'm creating an Options object with different
        'values than those I'm using for the defaults for the 
        'IPrinterSheet routines.

        MyOptions = New GCA5Engine.SheetOptionsManager(Name)
        CreateOptions(MyOptions)

        MyOptions.Value("SheetFont") = New Font("Arial Narrow", 10)
    End Sub

    Public Sub UpgradeOptions(Options As GCA5Engine.SheetOptionsManager) Implements GCA5.Interfaces.IPrinterSheet.UpgradeOptions
        'This is called only when a particular plug-in is loaded the first time,
        'and before SetOptions.

        'I don't do anything with this.

    End Sub
    Public Sub CreateOptions(Options As GCA5Engine.SheetOptionsManager) Implements GCA5.Interfaces.IPrinterSheet.CreateOptions
        'This is the routine where all the Options we want to use are created,
        'and where the UI for the Preferences dialog is filled out.
        '
        'This is equivalent to CharacterSheetOptions from previous implementations
        Dim i As Integer
        Dim ok As Boolean
        Dim newOption As GCA5Engine.SheetOption

        Dim descFormat As New SheetOptionDisplayFormat
        descFormat.BackColor = Color.Gold
        descFormat.ForeColor = Color.Black
        'descFormat.CaptionLocalBackColor = SystemColors.Info
        descFormat.BoxDescriptionForeColor = Color.Black
        descFormat.BoxDescriptionBackColor = Color.LightGoldenrodYellow

        Dim boxFormat As New SheetOptionDisplayFormat
        boxFormat.ForeColor = Color.White
        boxFormat.BackColor = Color.Gray
        boxFormat.BoxDescriptionForeColor = Color.Black
        boxFormat.BoxDescriptionBackColor = Color.LightGray


        '**************************************************
        '* Description block at top *
        '**************************************************
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Box_Description"
        newOption.Type = GCA5Engine.OptionType.Box
        newOption.Value = Name & " " & Version()
        newOption.UserPrompt = Description
        newOption.DisplayFormat = descFormat
        ok = Options.AddOption(newOption)


        '**************************************************
        '* Fonts *
        '**************************************************
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Box_Fonts" '"Header_Fonts"
        newOption.Type = GCA5Engine.OptionType.Box  ' GCA5Engine.OptionType.Header
        newOption.Value = "Font Options"
        newOption.UserPrompt = "There is a bug in Windows which will allow you to select fonts that are not actually supported for display. We are sorry about that. If that happens here, GCA will post an error to the Log, and the font value will not change."
        newOption.DisplayFormat = New SheetOptionDisplayFormat(boxFormat)
        newOption.DisplayFormat.BoxDescriptionBackColor = Color.MistyRose
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "SheetFont"
        newOption.Type = GCA5Engine.OptionType.Font
        newOption.UserPrompt = "Font for the form text items"
        newOption.DefaultValue = New Font("Arial", 10)
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "FootnotesFont"
        newOption.Type = GCA5Engine.OptionType.Font
        newOption.UserPrompt = "Font for the various footnote items."
        newOption.DefaultValue = New Font("Arial Narrow", 8)
        ok = Options.AddOption(newOption)


        '**************************************************
        '* Trait Options *
        '**************************************************
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Box_Traits"
        newOption.Type = GCA5Engine.OptionType.Box
        newOption.Value = "Miscellaneous Trait Options"
        newOption.DisplayFormat = boxFormat
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "ShowWhere"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "Do you want to show 'Location' or 'Where' information for traits that have it?"
        newOption.DefaultValue = True
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "WhereNote"
        newOption.Type = GCA5Engine.OptionType.Caption
        newOption.UserPrompt = "'Location' or 'Where' will be included in the footnote font directly under the applicable item."
        newOption.DisplayFormat.BackColor = SystemColors.Info
        newOption.DisplayFormat.ForeColor = SystemColors.InfoText
        newOption.DisplayFormat.CaptionLocalBackColor = SystemColors.Info
        ok = Options.AddOption(newOption)


        '**************************************************
        '* Options by ItemType *
        '**************************************************
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Box_ByType" '"Header_Hidden"
        newOption.Type = GCA5Engine.OptionType.Box  ' GCA5Engine.OptionType.Header
        newOption.Value = "Options by Trait Type"
        newOption.DisplayFormat = boxFormat
        ok = Options.AddOption(newOption)



        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_Hidden"
        newOption.Type = GCA5Engine.OptionType.Header
        newOption.UserPrompt = "Show Hidden Traits"
        ok = Options.AddOption(newOption)

        Dim a1(LastItemType - 1) As Boolean
        Dim t1(LastItemType - 1) As String
        For i = 1 To LastItemType
            a1(i - 1) = False
            t1(i - 1) = ReturnListNameNew(i)
        Next
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "ShowHidden"
        newOption.Type = OptionType.ListArray
        newOption.UserPrompt = "Check the sections for which you'd like to display traits that are normally hidden."
        newOption.DefaultValue = a1
        newOption.List = t1
        ok = Options.AddOption(newOption)



        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_Components"
        newOption.Type = GCA5Engine.OptionType.Header
        newOption.UserPrompt = "Show Component Traits"
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "CompNote"
        newOption.Type = GCA5Engine.OptionType.Caption
        newOption.UserPrompt = "Components are the traits that make up meta-traits and templates. If you would like to have them listed under their owning trait, turn that on here."
        newOption.DisplayFormat.CaptionLocalBackColor = Color.LightGray
        newOption.DisplayFormat.ForeColor = Color.Black
        newOption.DisplayFormat.BackColor = Color.LightGray
        ok = Options.AddOption(newOption)

        Dim a2(LastItemType - 1) As Boolean
        Dim t2(LastItemType - 1) As String
        For i = 1 To LastItemType
            a2(i - 1) = False
            t2(i - 1) = ReturnListNameNew(i)
        Next
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "ShowComponents"
        newOption.Type = OptionType.ListArray
        newOption.UserPrompt = "Check the sections for which you'd like to display any component traits."
        newOption.DefaultValue = a2
        newOption.List = t2
        ok = Options.AddOption(newOption)



        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_BonusFootnotes"
        newOption.Type = GCA5Engine.OptionType.Header
        newOption.UserPrompt = "Show Bonuses to Traits"
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "BonusNote"
        newOption.Type = GCA5Engine.OptionType.Caption
        newOption.UserPrompt = "If you don't want a section to display bonuses (common for spells), then make sure Spells is not checked for either section. If you have sections checked for both Footnotes and Inline, then both versions will display. These bonus options use the Footnote font."
        newOption.DisplayFormat.CaptionLocalBackColor = Color.LightGray
        newOption.DisplayFormat.ForeColor = Color.Black
        newOption.DisplayFormat.BackColor = Color.LightGray
        ok = Options.AddOption(newOption)

        Dim a3(LastItemType - 1) As Boolean
        Dim t3(LastItemType - 1) As String
        For i = 1 To LastItemType
            a3(i - 1) = True
            t3(i - 1) = ReturnListNameNew(i)
        Next
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "BonusesAsFootnotes"
        newOption.Type = OptionType.ListArray
        newOption.UserPrompt = "Check the sections for which you'd like to display all bonuses to traits as footnotes at the end of the listings."
        newOption.DefaultValue = a3
        newOption.List = t3
        ok = Options.AddOption(newOption)

        Dim a4(LastItemType - 1) As Boolean
        Dim t4(LastItemType - 1) As String
        For i = 1 To LastItemType
            a4(i - 1) = False
            t4(i - 1) = ReturnListNameNew(i)
        Next
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "BonusesInline"
        newOption.Type = OptionType.ListArray
        newOption.UserPrompt = "Check the sections for which you'd like to display bonuses to traits as 'inline' notes under each trait item."
        newOption.DefaultValue = a4
        newOption.List = t4
        ok = Options.AddOption(newOption)


        '**************************************************
        '* Color Options *
        '**************************************************
        '* Description block at top *
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Box_Colors" ' "Header_Colors"
        newOption.Type = GCA5Engine.OptionType.Box  'GCA5Engine.OptionType.Header
        newOption.Value = "Color Options"
        newOption.UserPrompt = "Set the colors for the various blocks here."
        newOption.DisplayFormat = boxFormat
        ok = Options.AddOption(newOption)


        newOption = New GCA5Engine.SheetOption
        newOption.Name = "EncumbranceLevelColor"
        newOption.Type = GCA5Engine.OptionType.Color
        newOption.UserPrompt = "The Encumbrance block is drawn in the Attributes colors selected below. However, please select the color used to highlight the current encumbrance level (a border will be drawn around those values in this color)."
        newOption.DefaultValue = Color.ForestGreen
        ok = Options.AddOption(newOption)


        '* Colors *
        For i = Stats To LastItemType
            newOption = New GCA5Engine.SheetOption
            newOption.Name = i & "Colors"
            newOption.Type = GCA5Engine.OptionType.Header
            newOption.UserPrompt = ReturnListNameNew(i) & " Block Colors"
            ok = Options.AddOption(newOption)

            newOption = New GCA5Engine.SheetOption
            newOption.Name = i & "BackColor"
            newOption.Type = GCA5Engine.OptionType.Color
            newOption.UserPrompt = "Color for the background."
            newOption.DefaultValue = Color.White
            ok = Options.AddOption(newOption)

            newOption = New GCA5Engine.SheetOption
            newOption.Name = i & "BorderColor"
            newOption.Type = GCA5Engine.OptionType.Color
            newOption.UserPrompt = "Color for the border lines and header background."
            Select Case i
                Case Attributes, Cultures, Languages
                    newOption.DefaultValue = Color.Maroon
                Case Ads, Perks
                    newOption.DefaultValue = Color.Blue
                Case Disads, Quirks
                    newOption.DefaultValue = Color.Brown
                Case Skills
                    newOption.DefaultValue = Color.Purple
                Case Spells
                    newOption.DefaultValue = Color.SaddleBrown
                Case Templates
                    newOption.DefaultValue = Color.Red
                Case Equipment
                    newOption.DefaultValue = Color.DeepSkyBlue
            End Select
            ok = Options.AddOption(newOption)

            newOption = New GCA5Engine.SheetOption
            newOption.Name = i & "ShadeColor"
            newOption.Type = GCA5Engine.OptionType.Color
            newOption.UserPrompt = "Color for the background of the alternate lines."
            Select Case i
                Case Attributes, Cultures, Languages
                    newOption.DefaultValue = Color.Linen
                Case Ads, Perks
                    newOption.DefaultValue = Color.AliceBlue
                Case Disads, Quirks
                    newOption.DefaultValue = Color.Linen
                Case Skills
                    newOption.DefaultValue = Color.LavenderBlush
                Case Spells
                    newOption.DefaultValue = Color.OldLace
                Case Templates
                    newOption.DefaultValue = Color.LavenderBlush
                Case Equipment
                    newOption.DefaultValue = Color.Azure
            End Select
            ok = Options.AddOption(newOption)

            newOption = New GCA5Engine.SheetOption
            newOption.Name = i & "TextColor"
            newOption.Type = GCA5Engine.OptionType.Color
            newOption.UserPrompt = "Color for the main text (printed against the background and alternate colors)."
            newOption.DefaultValue = Color.Black
            ok = Options.AddOption(newOption)

            newOption = New GCA5Engine.SheetOption
            newOption.Name = i & "HeaderTextColor"
            newOption.Type = GCA5Engine.OptionType.Color
            newOption.UserPrompt = "Color for the header text (printed against the border color)."
            newOption.DefaultValue = Color.White
            ok = Options.AddOption(newOption)
        Next


        '*  *
        '* Testing Only Options *
        '*  *
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Box_Testing" ' "Header_Testing"
        newOption.Type = GCA5Engine.OptionType.Box  ' GCA5Engine.OptionType.Header
        newOption.Value = "Test Options"
        newOption.UserPrompt = "Test Options for Testing Purposes"
        newOption.DisplayFormat = New SheetOptionDisplayFormat(boxFormat)
        newOption.DisplayFormat.BackColor = Color.Purple
        newOption.DisplayFormat.ForeColor = Color.White
        newOption.DisplayFormat.BoxDescriptionBackColor = Color.LavenderBlush
        newOption.DisplayFormat.BoxDescriptionForeColor = Color.Black
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "TestList"
        newOption.Type = GCA5Engine.OptionType.List
        newOption.UserPrompt = "TestList: Please pick an item from the list of items."
        newOption.DefaultValue = "Charlie"
        newOption.List = {"Alpha", "Bravo", "Charlie", "Delta"}
        ok = Options.AddOption(newOption)

        'NOTE: Because List is now a 0-based Array, the number of the 
        'DefaultValue and the selected Value is 0-based!
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "TestListNumber"
        newOption.Type = GCA5Engine.OptionType.ListNumber
        newOption.UserPrompt = "TestListNumber: Please pick an item from the list of items."
        newOption.DefaultValue = 3
        newOption.List = {"Alpha", "Bravo", "Charlie", "Delta"}
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "TestListFlag"
        newOption.Type = GCA5Engine.OptionType.ListFlag
        newOption.UserPrompt = "TestListFlag: Please pick the items you like from the list of available items."
        newOption.DefaultValue = 8
        newOption.List = {"Alpha", "Bravo", "Charlie", "Delta"}
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "TestListArray"
        newOption.Type = GCA5Engine.OptionType.ListArray
        newOption.UserPrompt = "TestListArray: Please pick the items you like from the list of available items."
        Dim defA() As Boolean = {True, False, True, False}
        newOption.DefaultValue = defA
        newOption.List = {"Alpha", "Bravo", "Charlie", "Delta"}
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "TestListOrdered"
        newOption.Type = GCA5Engine.OptionType.ListOrdered
        newOption.UserPrompt = "TestListOrdered: Please arrange these items into the order that you prefer."
        newOption.DefaultValue = ""
        newOption.List = {"Alpha", "Bravo", "Charlie", "Delta", "Echo", "Foxtrot", "Golf", "Hotel", "India"}
        ok = Options.AddOption(newOption)


        newOption = New GCA5Engine.SheetOption
        newOption.Name = "TestListText"
        newOption.Type = GCA5Engine.OptionType.Text
        newOption.UserPrompt = "TestListText: Please enter some text here."
        newOption.DefaultValue = "The quick brown fox jumped over a lazy dog."
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "TestListFlag2"
        newOption.Type = GCA5Engine.OptionType.ListFlag
        newOption.UserPrompt = "TestListFlag2: Please pick the items you like from the list of available items."
        newOption.DefaultValue = 0
        Dim defS(63) As String
        For i = 0 To 63
            defS(i) = i + 1
        Next
        newOption.List = defS
        ok = Options.AddOption(newOption)

    End Sub
    Public Function GeneratePrinterDocument(GCAParty As Party, PageSettings As C1PageSettings, Options As GCA5Engine.SheetOptionsManager) As C1.C1Preview.C1PrintDocument Implements GCA5.Interfaces.IPrinterSheet.GeneratePrinterDocument
        'Set our Options to the stored values we've just been given
        MyOptions = Options
        MyChar = GCAParty.Current

        'initialize some working values


        'Initialize some page values
        'Notify(Name & " " & Version, Priority.Green) 'GCA should now print this to the log itself in SheetView.
        'Notify("Home: " & Options.SheetHomeFolder, Priority.Blue)

        MyPrintDoc = New C1.C1Preview.C1PrintDocument


        MyPrintDoc.TagOpenParen = "[["
        MyPrintDoc.TagCloseParen = "]]"


        'I think in inches. Note that some values returned from the engine won't be in the value set here
        MyPrintDoc.DefaultUnit = C1.C1Preview.UnitTypeEnum.Inch
        MyPrintDoc.ResolvedUnit = UnitTypeEnum.Inch


        MyPrintDoc.Style.Font = MyOptions.Value("SheetFont") 'New Font("Times New Roman", 10)
        'MyPrintDoc.Style.TextColor = MyOptions.Value("SheetTextColor")


        ' Create page layout.
        Dim pl As New PageLayout
        pl.PageSettings = New C1.C1Preview.C1PageSettings()

        pl.PageSettings.PaperKind = PageSettings.PaperKind
        pl.PageSettings.Landscape = PageSettings.Landscape
        pl.PageSettings.TopMargin = PageSettings.TopMargin ' 0.5
        pl.PageSettings.BottomMargin = PageSettings.BottomMargin  '0.5
        pl.PageSettings.LeftMargin = PageSettings.LeftMargin  '0.5
        pl.PageSettings.RightMargin = PageSettings.RightMargin  '0.75

        pl.Columns.Add()
        pl.Columns.Add()

        MyPrintDoc.PageLayouts.Default = pl


        'The document settings for these values will be in inches for US page types, and in
        'millimeters for international types, and possibly other values for ANSI types and whatnot,
        'so we need to get the values in the measurement system we want, and we do that here.
        MarginLeft = MyPrintDoc.PageLayout.PageSettings.LeftMargin.ConvertUnit(UnitTypeEnum.Inch)
        MarginRight = MyPrintDoc.PageLayout.PageSettings.RightMargin.ConvertUnit(UnitTypeEnum.Inch)
        MarginTop = MyPrintDoc.PageLayout.PageSettings.TopMargin.ConvertUnit(UnitTypeEnum.Inch)
        MarginBottom = MyPrintDoc.PageLayout.PageSettings.BottomMargin.ConvertUnit(UnitTypeEnum.Inch)
        PageWidth = MyPrintDoc.PageLayout.PageSettings.Width.ConvertUnit(UnitTypeEnum.Inch)
        PageHeight = MyPrintDoc.PageLayout.PageSettings.Height.ConvertUnit(UnitTypeEnum.Inch)

        If MyChar Is Nothing Then
            Dim boxWidth, boxHeight, boxLeft, boxTop As Double
            Dim ld As C1.C1Preview.LineDef

            MyPrintDoc.StartDoc()

            boxLeft = MarginLeft
            boxTop = MarginTop
            boxWidth = PageWidth - MarginRight - boxLeft
            boxHeight = PageHeight - MarginBottom - boxTop

            ld = New C1.C1Preview.LineDef(0.06, Color.AliceBlue, Drawing2D.DashStyle.Dash)
            MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

            Dim EmptyPageStyle As Style = MyPrintDoc.Style.Children.Add()
            EmptyPageStyle.Font = New Font(MyPrintDoc.Style.Font.Name, 24, FontStyle.Bold)
            EmptyPageStyle.TextAlignVert = AlignVertEnum.Center
            EmptyPageStyle.TextAlignHorz = AlignHorzEnum.Center

            MyPrintDoc.RenderDirectText(boxLeft, boxTop, "There are no characters loaded.", boxWidth, boxHeight, EmptyPageStyle)
            MyPrintDoc.EndDoc()

            Return MyPrintDoc
        End If


        CreateStyles()
        CreateHeader()
        CreateFooter()

        ShowHidden = MyOptions.Value("ShowHidden")
        ShowComponents = MyOptions.Value("ShowComponents")
        BonusesAsFootnotes = MyOptions.Value("BonusesAsFootnotes")
        BonusesInline = MyOptions.Value("BonusesInline")

        'Start working
        'MyPrintDoc.StartDoc()
        PrintBlocks()

        'MyPrintDoc.EndDoc()
        Return MyPrintDoc
    End Function
    Private Sub PrintBlocks()
        Dim i As Integer

        PrintGeneralInfoBlock()
        PrintSpacer()

        For i = 1 To LastItemType
            If i = TraitTypes.Advantages Then
                'before we print ads, print the Enc & Move block
                PrintEncumbranceBlock()
                PrintSpacer()
            End If

            SetTraitColumns(i, 0)
            CreateFields(i)

            If Fields.Count > 0 Then
                SetFieldSizes()

                'print
                PrintTraitsBlock(i)
                PrintSpacer()
            End If
        Next

        PrintProtectionBlock()
        PrintSpacer()

        'Dim pl As New PageLayout
        'pl.AssignFrom(MyPrintDoc.PageLayouts.Default)
        'pl.Columns.RemoveAt(1)
        'MyPrintDoc.PageLayouts.Default = pl
        'MyPrintDoc.NewPage(pl)

        SetWeaponColumns(Hand, 0)
        CreateWeaponFields(Hand)
        If Fields.Count > 0 Then
            SetFieldSizes()

            'print
            PrintWeaponsBlock(Hand)
            PrintSpacer()
        End If
        SetWeaponColumns(Ranged, 0)
        CreateWeaponFields(Ranged)
        If Fields.Count > 0 Then
            SetFieldSizes()

            'print
            PrintWeaponsBlock(Ranged)
            PrintSpacer()
        End If


        MyPrintDoc.Generate()
    End Sub
    Private Sub PrintSpacer()
        Dim rt As RenderText = New RenderText(vbCrLf)

        MyPrintDoc.Body.Children.Add(rt)
    End Sub
    Private Sub PrintEncumbranceBlock()
        Dim Title As String = "Encumbrance & Move"
        Dim ItemType As Integer = Attributes
        Dim o As RenderText
        Dim i As Integer

        '8 columns: icon & name & none & light & med & hvy & x-hvy & icon
        Dim curCol As Integer
        Dim curRow As Integer

        numCols = 7
        ReDim ColAligns(numCols)
        ReDim ColWidths(numCols)
        ReDim SubHeads(numCols)
        PointsCol = 0
        ComponentDrawCol = 1

        IconColLeft = 0 'LeftOffset
        IconColWidth = 1 / 32 ' 0

        curCol = 0
        ColWidths(curCol) = IconColWidth
        'icon col

        curCol += 1
        SubHeads(curCol) = "Name"
        ColAligns(curCol) = AlignHorzEnum.Left
        ColWidths(curCol) = 0

        curCol += 1
        SubHeads(curCol) = "None"
        ColAligns(curCol) = AlignHorzEnum.Right
        ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding 

        curCol += 1
        SubHeads(curCol) = "Light"
        ColAligns(curCol) = AlignHorzEnum.Right
        ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding 

        curCol += 1
        SubHeads(curCol) = "Med"
        ColAligns(curCol) = AlignHorzEnum.Right
        ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding 

        curCol += 1
        SubHeads(curCol) = "Hvy"
        ColAligns(curCol) = AlignHorzEnum.Right
        ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding 

        curCol += 1
        SubHeads(curCol) = "X-Hvy"
        ColAligns(curCol) = AlignHorzEnum.Right
        ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding 

        curCol += 1
        'icon col
        ColWidths(curCol) = IconColWidth

        '** Create Data Table
        'set styles
        Dim rtab As New RenderTable

        rtab.Style.BackColor = MyOptions.Value(ItemType & "BackColor")

        rtab.CellStyle.Padding.Left = 1 / 32
        rtab.CellStyle.Padding.Right = 1 / 32
        rtab.Style.Borders.All = New LineDef(0.02, MyOptions.Value(ItemType & "BorderColor"))
        'rtab.Style.GridLines.All = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Horz = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Vert = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)

        'create top two header rows
        rtab.Rows(0).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
        rtab.Rows(0).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")
        rtab.Rows(1).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
        rtab.Rows(1).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")

        'print header area
        rtab.Cells(0, 0).Style.TextAlignHorz = AlignHorzEnum.Left
        rtab.Cells(0, 0).SpanCols = numCols - 1
        Dim rt1 As New RenderText(Title, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), AlignHorzEnum.Left)
        rtab.Cells(0, 0).Area.Children.Add(rt1)


        'set col widths
        For i = 0 To UBound(ColWidths)
            'leave 0 widths to the table layout engine to fit in
            If ColWidths(i) > 0 Then
                rtab.Cols(i).Width = ColWidths(i)
            End If
        Next

        'print sub-heads
        curRow = 1
        For i = 1 To UBound(SubHeads)
            rtab.Cells(curRow, i).Text = SubHeads(i)
            rtab.Cols(i).Style.TextAlignHorz = ColAligns(i)
        Next

        rtab.RowGroups(0, 2).Header = TableHeaderEnum.Page

        'Add table to sheet
        MyPrintDoc.Body.Children.Add(rtab)


        '** Create Data
        Dim tmpName As String
        Dim curTrait As GCATrait

        Dim NeededStats(8) As String
        NeededStats(0) = "Basic Lift"
        NeededStats(1) = "Air Move"
        NeededStats(2) = "Brachiation Move"
        NeededStats(3) = "Ground Move"
        NeededStats(4) = "Space Move"
        NeededStats(5) = "TK Move"
        NeededStats(6) = "Tunneling Move"
        NeededStats(7) = "Water Move"
        NeededStats(8) = "Dodge"

        curRow = 2
        If MyChar.Items.Count > 0 Then
            'look for the move scores we need.
            For Each tmpName In NeededStats
                i = MyChar.ItemPositionByNameAndExt(tmpName, Stats)
                If i > 0 Then
                    curTrait = MyChar.Items.Item(i)

                    If curTrait.TagItem("hide") = "" Then
                        'okay, print this one.

                        Dim Values(numCols) As String

                        'Get Trait Values
                        Values(1) = curTrait.Name
                        Values(2) = curTrait.Score
                        Values(3) = Int(curTrait.Score * 0.8)
                        Values(4) = Int(curTrait.Score * 0.6)
                        Values(5) = Int(curTrait.Score * 0.4)
                        Values(6) = Int(curTrait.Score * 0.2)

                        'Min 1, unless score is 0
                        For i = 3 To numCols
                            If Val(Values(i)) < 1 Then
                                If curTrait.Score >= 1 Then
                                    Values(i) = 1
                                Else
                                    Values(i) = 0
                                End If
                            End If
                        Next

                        'Special Cases!
                        If curTrait.LCaseName = "basic lift" Then
                            Values(1) = "Carry (lbs)"
                            Values(3) = curTrait.Score * 2
                            Values(4) = curTrait.Score * 3
                            Values(5) = curTrait.Score * 6
                            Values(6) = curTrait.Score * 10
                        ElseIf curTrait.LCaseName = "dodge" Then
                            Values(3) = curTrait.Score - 1
                            Values(4) = curTrait.Score - 2
                            Values(5) = curTrait.Score - 3
                            Values(6) = curTrait.Score - 4
                        Else
                            Values(1) = Values(1) & " (yds)"
                        End If


                        'Print the column values 
                        For curCol = 0 To numCols
                            o = New RenderText

                            o.Text = Values(curCol)
                            rtab.Cells(curRow, curCol).RenderObject = o
                        Next

                        '*****
                        '* Alt color for basic lift and dodge
                        Select Case tmpName
                            Case "Basic Lift", "Dodge"
                                rtab.Rows(curRow).Style.BackColor = MyOptions.Value(ItemType & "ShadeColor")
                        End Select
                        '*****

                        '*****
                        '* Highlight cur enc level col
                        curCol = MyChar.EncumbranceLevel + 2 '+2 to skip the icon and name cols
                        If curCol <= 7 Then
                            rtab.Cells(curRow, curCol).Style.GridLines.Left = New LineDef(0.02, MyOptions.Value("EncumbranceLevelColor"), System.Drawing.Drawing2D.DashStyle.Solid)
                            rtab.Cells(curRow, curCol).Style.GridLines.Right = New LineDef(0.02, MyOptions.Value("EncumbranceLevelColor"), System.Drawing.Drawing2D.DashStyle.Solid)
                        End If
                        '*****

                        curRow += 1
                    End If
                End If
            Next

            '*****
            '* Highlight cur enc level col
            curCol = MyChar.EncumbranceLevel + 2 '+2 to skip the icon and name cols
            If curCol <= 7 Then
                'rtab.Cols(curCol).Style.GridLines.Left = New LineDef(0.02, Color.ForestGreen, System.Drawing.Drawing2D.DashStyle.Solid)
                'rtab.Cols(curCol).Style.GridLines.Right = New LineDef(0.02, Color.ForestGreen, System.Drawing.Drawing2D.DashStyle.Solid)
                rtab.Cells(1, curCol).Style.GridLines.Top = New LineDef(0.02, MyOptions.Value("EncumbranceLevelColor"), System.Drawing.Drawing2D.DashStyle.Solid)
                rtab.Cells(curRow - 1, curCol).Style.GridLines.Bottom = New LineDef(0.02, MyOptions.Value("EncumbranceLevelColor"), System.Drawing.Drawing2D.DashStyle.Solid)

                rtab.Cells(1, curCol).Style.GridLines.Left = New LineDef(0.02, MyOptions.Value("EncumbranceLevelColor"), System.Drawing.Drawing2D.DashStyle.Solid)
                rtab.Cells(1, curCol).Style.GridLines.Right = New LineDef(0.02, MyOptions.Value("EncumbranceLevelColor"), System.Drawing.Drawing2D.DashStyle.Solid)
            End If

            '*****
            '* print our loadout
            o = New RenderText
            o.Text = "Loadout"
            rtab.Cells(curRow, 1).RenderObject = o
            rtab.Cells(curRow, 1).Style.TextAlignHorz = AlignHorzEnum.Left

            o = New RenderText
            o.Text = MyChar.CurrentLoadout
            rtab.Cells(curRow, 2).RenderObject = o
            rtab.Cells(curRow, 2).Style.TextAlignHorz = AlignHorzEnum.Left
            rtab.Cells(curRow, 2).SpanCols = 4

            o = New RenderText
            o.Text = Format(MyChar.CurrentLoad, "#,0.0") '& " lbs"
            rtab.Cells(curRow, 6).RenderObject = o
            rtab.Cells(curRow, 6).Style.TextAlignHorz = AlignHorzEnum.Right

            rtab.Rows(curRow).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
            rtab.Rows(curRow).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")
            '*****
        End If

    End Sub
    Private Sub PrintProtectionBlock()
        Dim Title As String
        Dim ItemType As Integer = Equipment

        Dim rtab As New RenderTable

        rtab.Style.BackColor = MyOptions.Value(ItemType & "BackColor")

        rtab.CellStyle.Padding.Left = 1 / 32
        rtab.CellStyle.Padding.Right = 1 / 32
        rtab.Style.Borders.All = New LineDef(0.02, MyOptions.Value(ItemType & "BorderColor"))
        rtab.Style.GridLines.All = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)

        'create header row
        rtab.Rows(0).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
        rtab.Rows(0).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")

        'print header area
        Title = "Protection"

        rtab.Cells(0, 0).Style.TextAlignHorz = AlignHorzEnum.Left
        rtab.Cells(0, 0).Text = Title

        rtab.RowGroups(0, 1).Header = TableHeaderEnum.Page

        'print contents
        Dim ri As New RenderImage

        Dim Settings As New ProtectionPaperDollSettings
        Settings.Character = MyChar
        Settings.TextFont = New Font(MyPrintDoc.Style.Font.Name, 14)
        Settings.TextColor = MyOptions.Value(ItemType & "TextColor") 'NEW 2015 09 11 'Color.Black
        Settings.TextColorAlt = Color.Black
        Settings.ShadeColor = MyOptions.Value(ItemType & "ShadeColor")
        Settings.BackColor = Color.White
        Settings.BorderColor = MyOptions.Value(ItemType & "BorderColor")

        ri.Image = GetProtectionPaperDoll(Settings)
        ri.Width = "100%"
        rtab.Cells(1, 0).Area.Children.Add(ri)

        'Add table to sheet
        MyPrintDoc.Body.Children.Add(rtab)



        'NEW 2015 03 31
        '*************************
        'Print included armor items; by layers, if applicable
        '*************************

        'NEW 2015 05 29 safety check
        If MyChar.CurrentLoadout = "" Then Return
        If MyChar.LoadOuts.Count = 0 Then Return
        'END NEW 2015 05 29

        Dim MyCurrentLoadOut As LoadOut = MyChar.LoadOuts(MyChar.CurrentLoadout)

        Const IconCol = 0
        Const LetterCol = 1
        Const DRCol = 2
        Const LocCol = 3

        rtab = New RenderTable

        rtab.Style.BackColor = MyOptions.Value(ItemType & "BackColor")

        rtab.CellStyle.Padding.Left = 1 / 32
        rtab.CellStyle.Padding.Right = 1 / 32
        rtab.Style.Borders.All = New LineDef(0.02, MyOptions.Value(ItemType & "BorderColor"))
        rtab.Style.GridLines.Horz = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Vert = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)

        rtab.Cols(IconCol).Width = GetSize("[MM]").Width
        rtab.Cols(LetterCol).Width = GetSize("[MM]").Width
        rtab.Cols(LocCol).Width = GetSize(" Overall DR Bonus ").Width '0.5
        rtab.Cols(LocCol).Style.TextAlignHorz = AlignHorzEnum.Right

        'create header row
        rtab.Rows(0).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
        rtab.Rows(0).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")

        'print header area
        If MyCurrentLoadOut.UserOrderedLayers Then
            Title = "Layered Armor" 'MyChar.CurrentLoadout '"Protection"
        Else
            Title = "Armor" 'MyChar.CurrentLoadout '"Protection"
        End If

        rtab.Cells(IconCol, 0).Style.TextAlignHorz = AlignHorzEnum.Left
        rtab.Cells(IconCol, 0).Text = Title
        rtab.Cells(IconCol, 0).SpanCols = 4

        rtab.RowGroups(0, 1).Header = TableHeaderEnum.Page

        'print contents
        Dim alt As Boolean = False
        Dim curRow As Integer = 0
        Dim armorValue As String = ""
        Dim armorLoc As String = ""
        Dim tmp As String = ""

        Dim curLayer As LayerItem
        Dim curItem As GCATrait

        If Not MyCurrentLoadOut.UserOrderedLayers OrElse MyCurrentLoadOut.OrderedLayers Is Nothing Then
            'straight item list
            rtab.Cols(IconCol).Width = 0

            For Each curItem In MyCurrentLoadOut.ArmorItems

                'don't print the shields here
                If Not MyCurrentLoadOut.ShieldItems.Contains(curItem.CollectionKey) Then
                    'FIRST ROW
                    curRow += 1

                    'Name
                    rtab.Cells(curRow, LetterCol).Text = curItem.DisplayName
                    rtab.Cells(curRow, LetterCol).SpanCols = 3

                    'SECOND ROW
                    curRow += 1

                    'Reference Symbol 
                    'not used here

                    'Armor Values
                    armorValue = ""
                    If curItem.TagItem("chardr") <> "" Then
                        armorValue = "DR: " & curItem.TagItem("chardr")
                    Else
                        'no value, maybe DR stat
                        If curItem.ItemType = TraitTypes.Stats Then
                            armorValue = "DR: " & curItem.Score.ToString
                        End If
                    End If
                    If curItem.TagItem("chardb") <> "" Then
                        armorValue = "DB: " & curItem.TagItem("chardb") & " " & armorValue
                    End If

                    tmp = ""
                    If curItem.TagItem("chardeflect") <> "" Then
                        tmp = "Def: " & curItem.TagItem("chardeflect")
                    End If
                    If curItem.TagItem("charfortify") <> "" Then
                        If tmp = "" Then
                            tmp = "Fort: " & curItem.TagItem("charfortify")
                        Else
                            tmp = tmp & " " & "Fort: " & curItem.TagItem("charfortify")
                        End If
                    End If
                    If tmp <> "" Then
                        armorValue = armorValue & " (" & tmp & ")"
                    End If

                    rtab.Cells(curRow, LetterCol).Text = armorValue
                    rtab.Cells(curRow, LetterCol).SpanCols = 2

                    'Locations
                    armorLoc = ""
                    If curItem.TagItem("locationcoverage") <> "" Then 'NEW 2015 04 17 changed 'charlocation' to 'locationcoverage'
                        armorLoc = curItem.TagItem("locationcoverage")
                    Else
                        If curItem.TagItem("location") <> "" Then
                            armorLoc = curItem.TagItem("location")
                        Else
                            'no location, maybe DR stat
                            If curItem.ItemType = TraitTypes.Stats Then
                                armorLoc = "overall DR bonus"
                            End If
                        End If
                    End If
                    rtab.Cells(curRow, LocCol).Text = armorLoc

                    'Shading
                    rtab.Rows(curRow - 1).Style.TextColor = MyOptions.Value(ItemType & "TextColor") 'NEW 2015 09 11
                    rtab.Rows(curRow).Style.TextColor = MyOptions.Value(ItemType & "TextColor") 'NEW 2015 09 11
                    If alt Then
                        rtab.Rows(curRow - 1).Style.BackColor = MyOptions.Value(ItemType & "ShadeColor")
                        rtab.Rows(curRow).Style.BackColor = MyOptions.Value(ItemType & "ShadeColor")
                    End If

                    alt = Not alt

                End If
            Next
        Else
            'layered order

            If MyCurrentLoadOut.OrderedLayers.Count > 0 Then
                For Each curLayer In MyCurrentLoadOut.OrderedLayers
                    curItem = curLayer.Item

                    'FIRST ROW
                    curRow += 1

                    'draw our icons
                    ri = New RenderImage
                    Dim bm As Bitmap

                    If curLayer.IsFlexible Then
                        bm = New Bitmap(My.Resources.Resources.alpha_24_layer_flexible)
                    Else
                        bm = New Bitmap(My.Resources.Resources.alpha_24_layer_rigid)
                    End If
                    ri.Image = bm
                    rtab.Cells(curRow, IconCol).Area.Children.Add(ri)
                    rtab.Cells(curRow, IconCol).SpanRows = 2

                    'Name
                    rtab.Cells(curRow, LetterCol).Text = curItem.DisplayName
                    rtab.Cells(curRow, LetterCol).SpanCols = 3

                    'SECOND ROW
                    curRow += 1

                    'Reference Symbol
                    rtab.Cells(curRow, LetterCol).Text = curLayer.FootnoteSymbol

                    'Armor Values
                    If curItem.TagItem("chardr") <> "" Then
                        armorValue = "DR: " & curItem.TagItem("chardr")
                    Else
                        'no value, maybe DR stat
                        If curItem.ItemType = TraitTypes.Stats Then
                            armorValue = "DR: " & curItem.Score.ToString
                        End If
                    End If
                    If curItem.TagItem("chardb") <> "" Then
                        armorValue = "DB: " & curItem.TagItem("chardb") & " " & armorValue
                    End If

                    tmp = ""
                    If curItem.TagItem("chardeflect") <> "" Then
                        tmp = "Def: " & curItem.TagItem("chardeflect")
                    End If
                    If curItem.TagItem("charfortify") <> "" Then
                        If tmp = "" Then
                            tmp = "Fort: " & curItem.TagItem("charfortify")
                        Else
                            tmp = tmp & " " & "Fort: " & curItem.TagItem("charfortify")
                        End If
                    End If
                    If tmp <> "" Then
                        armorValue = armorValue & " (" & tmp & ")"
                    End If

                    rtab.Cells(curRow, DRCol).Text = armorValue

                    'Locations
                    If curItem.TagItem("locationcoverage") <> "" Then 'NEW 2015 04 17 changed 'charlocation' to 'locationcoverage'
                        armorLoc = curItem.TagItem("locationcoverage")
                    Else
                        If curItem.TagItem("location") <> "" Then
                            armorLoc = curItem.TagItem("location")
                        Else
                            'no location, maybe DR stat
                            If curItem.ItemType = TraitTypes.Stats Then
                                armorLoc = "overall DR bonus"
                            End If
                        End If
                    End If
                    rtab.Cells(curRow, LocCol).Text = armorLoc

                    'Shading
                    rtab.Rows(curRow - 1).Style.TextColor = MyOptions.Value(ItemType & "TextColor") 'NEW 2015 09 11
                    rtab.Rows(curRow).Style.TextColor = MyOptions.Value(ItemType & "TextColor") 'NEW 2015 09 11
                    If alt Then
                        rtab.Rows(curRow - 1).Style.BackColor = MyOptions.Value(ItemType & "ShadeColor")
                        rtab.Rows(curRow).Style.BackColor = MyOptions.Value(ItemType & "ShadeColor")
                    End If

                    alt = Not alt
                Next
            Else
                'Name on row alone
                rtab.Rows(curRow).Style.TextColor = MyOptions.Value(ItemType & "TextColor") 'NEW 2015 09 11
                rtab.Cells(curRow, 0).Text = "No applicable traits."
            End If
        End If

        'Add table to sheet
        MyPrintDoc.Body.Children.Add(rtab)



        '*************************
        'Print included shield items; by layers, if applicable
        '*************************
        rtab = New RenderTable

        rtab.Style.BackColor = MyOptions.Value(ItemType & "BackColor")

        rtab.CellStyle.Padding.Left = 1 / 32
        rtab.CellStyle.Padding.Right = 1 / 32
        rtab.Style.Borders.All = New LineDef(0.02, MyOptions.Value(ItemType & "BorderColor"))
        rtab.Style.GridLines.Horz = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Vert = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)

        rtab.Cols(IconCol).Width = 0
        rtab.Cols(LetterCol).Width = GetSize("[MM]").Width
        rtab.Cols(LocCol).Width = GetSize(" Overall DR Bonus ").Width '0.5
        rtab.Cols(LocCol).Style.TextAlignHorz = AlignHorzEnum.Right

        'create header row
        rtab.Rows(0).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
        rtab.Rows(0).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")

        'print header area
        Title = "Shields" 'MyChar.CurrentLoadout '"Protection"

        rtab.Cells(IconCol, 0).Style.TextAlignHorz = AlignHorzEnum.Left
        rtab.Cells(IconCol, 0).Text = Title
        rtab.Cells(IconCol, 0).SpanCols = 4

        rtab.RowGroups(0, 1).Header = TableHeaderEnum.Page

        'print contents
        alt = False
        curRow = 0
        armorValue = ""
        armorLoc = ""
        tmp = ""
        If MyCurrentLoadOut.ShieldItems.Count > 0 Then
            'active shields
            For Each curItem In MyCurrentLoadOut.ShieldItems

                'FIRST ROW
                curRow += 1

                'Name
                rtab.Cells(curRow, LetterCol).Text = curItem.DisplayName
                rtab.Cells(curRow, LetterCol).SpanCols = 3

                'SECOND ROW
                curRow += 1

                'Reference Symbol 
                'not used here

                'Armor Values
                armorValue = ""
                If curItem.TagItem("chardr") <> "" Then
                    armorValue = "DR: " & curItem.TagItem("chardr")
                Else
                    'no value, maybe DR stat
                    If curItem.ItemType = TraitTypes.Stats Then
                        armorValue = "DR: " & curItem.Score.ToString
                    End If
                End If
                If curItem.TagItem("chardb") <> "" Then
                    armorValue = "DB: " & curItem.TagItem("chardb") & " " & armorValue
                End If

                tmp = ""
                If curItem.TagItem("chardeflect") <> "" Then
                    tmp = "Def: " & curItem.TagItem("chardeflect")
                End If
                If curItem.TagItem("charfortify") <> "" Then
                    If tmp = "" Then
                        tmp = "Fort: " & curItem.TagItem("charfortify")
                    Else
                        tmp = tmp & " " & "Fort: " & curItem.TagItem("charfortify")
                    End If
                End If
                If tmp <> "" Then
                    armorValue = armorValue & " (" & tmp & ")"
                End If

                rtab.Cells(curRow, LetterCol).Text = armorValue
                rtab.Cells(curRow, LetterCol).SpanCols = 2

                armorLoc = MyCurrentLoadOut.ShieldArcs(curItem.CollectionKey)

                rtab.Cells(curRow, LocCol).Text = armorLoc

                'Shading
                rtab.Rows(curRow - 1).Style.TextColor = MyOptions.Value(ItemType & "TextColor") 'NEW 2015 09 11
                rtab.Rows(curRow).Style.TextColor = MyOptions.Value(ItemType & "TextColor") 'NEW 2015 09 11
                If alt Then
                    rtab.Rows(curRow - 1).Style.BackColor = MyOptions.Value(ItemType & "ShadeColor")
                    rtab.Rows(curRow).Style.BackColor = MyOptions.Value(ItemType & "ShadeColor")
                End If

                alt = Not alt
            Next
        Else
            'no active shields

            'FIRST ROW
            curRow += 1

            'Name
            rtab.Rows(curRow).Style.TextColor = MyOptions.Value(ItemType & "TextColor") 'NEW 2015 09 11
            rtab.Cells(curRow, LetterCol).Text = "No active shields."
            rtab.Cells(curRow, LetterCol).SpanCols = 3
        End If


        'Add table to sheet
        MyPrintDoc.Body.Children.Add(rtab)

    End Sub
    Private Sub PrintWeaponsBlock(WeaponType As Integer)
        Dim Title As String
        Dim i As Integer
        Dim FieldIndex As Integer
        Dim curRow, curCol As Integer
        Dim myHeight As Single
        Dim ComponentIndent As Single

        Dim curField As clsDisplayField
        Dim curItem As GCATrait
        Dim rtab As New RenderTable

        Dim ItemType As Integer = Equipment


        'rtab.Width = "100%"
        rtab.Style.BackColor = MyOptions.Value(ItemType & "BackColor")

        rtab.CellStyle.Padding.Left = 1 / 32
        rtab.CellStyle.Padding.Right = 1 / 32
        rtab.Style.Borders.All = New LineDef(0.02, MyOptions.Value(ItemType & "BorderColor"))
        'rtab.Style.GridLines.All = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Horz = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Vert = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)

        'create top two header rows
        rtab.Rows(0).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
        rtab.Rows(0).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")
        rtab.Rows(1).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
        rtab.Rows(1).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")

        'print header area
        Select Case WeaponType
            Case Hand
                Title = "Melee Attacks"

                'Change Layout to use full page width
                'If anything needs to be two-column after the Weapons blocks,
                'that layout will need to be restored.
                Dim pl As New PageLayout
                pl.AssignFrom(MyPrintDoc.PageLayouts.Default)
                pl.Columns.RemoveAt(1)

                Dim nl As New LayoutChangeNewPage(pl)

                rtab.LayoutChangeBefore = nl

            Case Else
                Title = "Ranged Attacks"
        End Select


        rtab.Cells(0, 0).Style.TextAlignHorz = AlignHorzEnum.Left
        rtab.Cells(0, 0).SpanCols = numCols - 1
        Dim rt1 As New RenderText(Title, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), AlignHorzEnum.Left)
        'Dim rt2 As New RenderText(MyChar.Cost(ItemType).ToString, AlignHorzEnum.Right)
        rtab.Cells(0, 0).Area.Children.Add(rt1)
        'rtab.Cells(0, numCols - 1).Area.Children.Add(rt2)

        'set col widths
        For i = 0 To UBound(ColWidths)
            'leave 0 widths to the table layout engine to fit in
            If ColWidths(i) > 0 Then
                rtab.Cols(i).Width = ColWidths(i)
            End If
        Next

        'print sub-heads
        For i = 1 To UBound(SubHeads)
            rtab.Cells(1, i).Text = SubHeads(i)
            rtab.Cols(i).Style.TextAlignHorz = ColAligns(i)
        Next

        rtab.RowGroups(0, 2).Header = TableHeaderEnum.Page

        'print contents

        For FieldIndex = 0 To Fields.Count - 1
            curRow = FieldIndex + 2

            curField = Fields(FieldIndex)
            curItem = curField.Data
            myHeight = curField.Size.Height

            ComponentIndent = 0
            If curField.IsComponent > 0 Then
                ComponentIndent = (IndentStepSize * curField.IsComponent)
            End If

            rtab.Rows(curRow).Style.TextColor = MyOptions.Value(ItemType & "TextColor")

            'Shading
            If curItem.TagItem("highlight") <> "" Then
                rtab.Rows(curRow).Style.BackColor = Color.Yellow
            Else
                If curField.Alt And ShadeAltLines Then
                    rtab.Rows(curRow).Style.BackColor = MyOptions.Value(ItemType & "ShadeColor")
                End If
            End If


            'Print the column values 
            For curCol = 0 To numCols
                If curCol = ComponentDrawCol AndAlso curField.IsComponent > 0 Then
                    rtab.Cells(curRow, curCol).CellStyle.Padding.Left = 1 / 32 + ComponentIndent
                End If

                rtab.Cells(curRow, curCol).Text = curField.Values(curCol)
            Next

        Next

        'Add table to sheet
        MyPrintDoc.Body.Children.Add(rtab)
    End Sub
    Private Sub PrintTraitsBlock(ItemType As Integer)
        Dim Title As String
        Dim i As Integer
        Dim FieldIndex As Integer
        Dim curRow, curCol As Integer
        Dim myHeight As Single
        Dim ComponentIndent As Single

        Dim IsPlaceholder As Boolean

        Dim curField As clsDisplayField
        Dim curItem As GCATrait
        Dim rtab As New RenderTable

        Dim hasBonusList As Boolean
        Dim tmpBonusText As String
        Dim FootnoteMarker As String
        Dim Footnotes As New FootnoteManager()
        Footnotes.FootnoteStyle = FootnoteMarkerStyle.Symbol

        rtab.Style.BackColor = MyOptions.Value(ItemType & "BackColor")

        rtab.CellStyle.Padding.Left = 1 / 32
        rtab.CellStyle.Padding.Right = 1 / 32
        rtab.Style.Borders.All = New LineDef(0.02, MyOptions.Value(ItemType & "BorderColor"))
        'rtab.Style.GridLines.All = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Horz = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Vert = New LineDef(0.01, MyOptions.Value(ItemType & "BorderColor"), System.Drawing.Drawing2D.DashStyle.Dot)

        'create top two header rows
        rtab.Rows(0).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
        rtab.Rows(0).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")
        rtab.Rows(1).Style.BackColor = MyOptions.Value(ItemType & "BorderColor")
        rtab.Rows(1).Style.TextColor = MyOptions.Value(ItemType & "HeaderTextColor")

        'print header area
        Title = ReturnLongListName(ItemType)

        rtab.Cells(0, 0).Style.TextAlignHorz = AlignHorzEnum.Left
        rtab.Cells(0, 0).SpanCols = numCols - 1
        Dim rt1 As New RenderText(Title, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), AlignHorzEnum.Left)
        Dim rt2 As New RenderText(MyChar.Cost(ItemType).ToString, AlignHorzEnum.Right)
        rtab.Cells(0, 0).Area.Children.Add(rt1)
        rtab.Cells(0, numCols - 1).Area.Children.Add(rt2)

        'set col widths
        For i = 0 To UBound(ColWidths)
            'leave 0 widths to the table layout engine to fit in
            If ColWidths(i) > 0 Then
                rtab.Cols(i).Width = ColWidths(i)
            End If
        Next

        'print sub-heads
        For i = 1 To UBound(SubHeads)
            rtab.Cells(1, i).Text = SubHeads(i)
            rtab.Cols(i).Style.TextAlignHorz = ColAligns(i)
        Next

        rtab.RowGroups(0, 2).Header = TableHeaderEnum.Page

        'print contents
        For FieldIndex = 0 To Fields.Count - 1
            curRow = FieldIndex + 2

            curField = Fields(FieldIndex)
            curItem = curField.Data
            myHeight = curField.Size.Height

            'NEW TESTING 2015 09 06
            IsPlaceholder = False
            If curItem.ItemType = TraitTypes.None Then IsPlaceholder = True
            'END NEW

            hasBonusList = False
            FootnoteMarker = ""
            tmpBonusText = ""
            If curItem.TagItem("bonuslist") <> "" Then
                hasBonusList = True
                tmpBonusText = curItem.TagItem("bonuslist")
            End If
            If curItem.TagItem("conditionallist") <> "" Then
                hasBonusList = True
                If tmpBonusText <> "" Then
                    tmpBonusText = tmpBonusText & vbCrLf & "Conditional: "
                Else
                    tmpBonusText = "Conditional: "
                End If
                tmpBonusText = tmpBonusText & curItem.TagItem("conditionallist")
            End If
            If BonusesAsFootnotes(ItemType - 1) Then
                If hasBonusList Then
                    FootnoteMarker = Footnotes.Add(tmpBonusText)
                End If
            End If

            ComponentIndent = 0
            If curField.IsComponent > 0 Then
                ComponentIndent = (IndentStepSize * curField.IsComponent)
            End If

            rtab.Rows(curRow).Style.TextColor = MyOptions.Value(ItemType & "TextColor")
            'NEW TESTUNG 2015 09 06
            If IsPlaceholder Then
                rtab.Rows(curRow).Style.FontBold = True
            End If
            'END NEW

            'Shading
            If curItem.TagItem("highlight") <> "" Then
                rtab.Rows(curRow).Style.BackColor = Color.Yellow
            Else
                If curField.Alt And ShadeAltLines Then
                    rtab.Rows(curRow).Style.BackColor = MyOptions.Value(ItemType & "ShadeColor")
                End If
            End If

            'Print the column values 
            For curCol = 0 To numCols

                '*****
                'For Sheet View inside GCA
                '*****
                'Helpfully, we aren't allowed to assign UserData to individual cells,
                'we we have to assign a renderobject just so we can do that, which
                'makes our life more difficult.
                Dim o As New RenderParagraph 'Using RenderParagraph instead of RenderText allows us to AddText() different fonts and styles within the block down below.
                Dim actor As SheetActionField
                '*****
                '*****

                If curCol = ComponentDrawCol AndAlso curField.IsComponent > 0 Then
                    rtab.Cells(curRow, curCol).CellStyle.Padding.Left = 1 / 32 + ComponentIndent
                End If

                If curCol = ComponentDrawCol AndAlso hasBonusList = True Then
                    o.Content.AddText(curField.Values(curCol) & FootnoteMarker)
                Else
                    o.Content.AddText(curField.Values(curCol))
                End If
                rtab.Cells(curRow, curCol).RenderObject = o


                '*****
                'Helpfully, we aren't allowed to assign UserData to individual cells,
                'so we have to assign a renderobject just so we can do that, which
                'is kind of annoying.
                If curField.SheetActionFields.ContainsKey(curCol) Then
                    actor = curField.SheetActionFields(curCol)
                    o.UserData = actor
                End If
                '*****
                'For Sheet View inside GCA
                '*****


                If curCol = ComponentDrawCol Then
                    'Only add these items when we're in the Name column.

                    If MyOptions.Value("ShowWhere") Then
                        If curItem.TagItem("location") <> "" Then
                            o.Content.AddText(vbCrLf & "Location: " & curItem.TagItem("location"), MyOptions.Value("FootnotesFont"))
                        End If
                        If curItem.TagItem("where") <> "" Then
                            o.Content.AddText(vbCrLf & "Where: " & curItem.TagItem("where"), MyOptions.Value("FootnotesFont"))
                        End If
                    End If
                    If BonusesInline(ItemType - 1) Then
                        o.Content.AddText(vbCrLf & tmpBonusText, MyOptions.Value("FootnotesFont"))
                    End If
                End If
            Next
        Next

        'Print footnotes, if there are any
        If Footnotes.Count > 0 Then
            'each footnoot as a new cell
            For i = 1 To Footnotes.Count
                curRow = curRow + 1
                rt1 = New RenderText(Footnotes.FootnoteWithMarker(i, FootnoteEnclosureStyle.None), MyOptions.Value("FootnotesFont"), AlignHorzEnum.Left)

                rtab.Cells(curRow, 1).RenderObject = rt1
                rtab.Cells(curRow, 1).SpanCols = numCols - 1

                'rtab.Cells(curRow, 1).Text = Footnotes.FootnoteWithMarker(i, FootnoteEnclosureStyle.None)
            Next
            'Old single block here
            'curRow = curRow + 1
            'rtab.Cells(curRow, 1).SpanCols = numCols - 1
            'rtab.Cells(curRow, 1).Text = Footnotes.FootnoteBlock
        End If


        'Add table to sheet
        MyPrintDoc.Body.Children.Add(rtab)
    End Sub
    Private Sub PrintGeneralInfoBlock()
        Dim rtab As New RenderTable(8, 2)

        'TESTING
        'Helpfully, we aren't allowed to assign UserData to individual cells,
        'we we have to assign a renderobject just so we can do that, which
        'makes our life more difficult.
        Dim o As RenderText
        Dim actor As SheetActionField
        'END TESTING

        rtab.CellStyle.Padding.Left = 1 / 32
        rtab.CellStyle.Padding.Right = 1 / 32
        rtab.Style.Borders.All = New LineDef(0.02, Color.Gray)
        'rtab.Style.GridLines.All = New LineDef(0.01, Color.Gray, System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Horz = New LineDef(0.01, Color.Gray, System.Drawing.Drawing2D.DashStyle.Dot)
        rtab.Style.GridLines.Vert = New LineDef(0.01, Color.Gray, System.Drawing.Drawing2D.DashStyle.Dot)

        rtab.Cols(0).Width = 1
        rtab.Cols(0).Style.TextAlignHorz = AlignHorzEnum.Right

        rtab.Cells(0, 0).Style.TextAlignHorz = AlignHorzEnum.Left
        rtab.Cells(0, 0).Style.BackColor = Color.Gray
        rtab.Cells(0, 0).Style.TextColor = Color.White
        rtab.Cells(0, 0).SpanCols = 2
        rtab.Cells(0, 0).Text = "General Information"

        rtab.Cells(1, 0).Text = "Name"
        'rtab.Cells(1, 1).Text = MyChar.Name
        '****************************************
        'TESTING for SheetActionFields
        actor = New SheetActionField
        actor.Name = "char Name"
        actor.ActionType = SheetAction.EditPlainText
        actor.Dock = PanelDockStyle.Fill
        actor.ActionObject = MyChar
        actor.ActionTag = "Name"
        actor.FamilyGroup = True

        o = New RenderText
        o.Text = MyChar.Name
        o.UserData = actor

        rtab.Cells(1, 1).RenderObject = o
        'END TESTING 
        '****************************************

        rtab.Cells(2, 0).Text = "Player"
        'rtab.Cells(2, 1).Text = MyChar.Player
        '****************************************
        'TESTING for SheetActionFields
        actor = New SheetActionField
        actor.Name = "char Player"
        actor.ActionType = SheetAction.EditPlainText
        actor.Dock = PanelDockStyle.Fill
        actor.ActionObject = MyChar
        actor.ActionTag = "Player"
        actor.FamilyGroup = True

        o = New RenderText
        o.Text = MyChar.Player
        o.UserData = actor

        rtab.Cells(2, 1).RenderObject = o
        'END TESTING 
        '****************************************

        rtab.Cells(3, 0).Text = "Race"
        'rtab.Cells(3, 1).Text = MyChar.Race
        '****************************************
        'TESTING for SheetActionFields
        actor = New SheetActionField
        actor.Name = "char Race"
        actor.ActionType = SheetAction.EditPlainText
        actor.Dock = PanelDockStyle.Fill
        actor.ActionObject = MyChar
        actor.ActionTag = "Race"
        actor.FamilyGroup = True

        o = New RenderText
        o.Text = MyChar.Race
        o.UserData = actor

        rtab.Cells(3, 1).RenderObject = o
        'END TESTING 
        '****************************************

        rtab.Cells(4, 0).Text = "Appearance"
        'rtab.Cells(4, 1).Text = MyChar.Appearance
        '****************************************
        'TESTING for SheetActionFields
        actor = New SheetActionField
        actor.Name = "char Appearance"
        actor.ActionType = SheetAction.EditPlainText
        actor.Dock = PanelDockStyle.Fill
        actor.ActionObject = MyChar
        actor.ActionTag = "Appearance"
        actor.FamilyGroup = True

        o = New RenderText
        o.Text = MyChar.Appearance
        o.UserData = actor

        rtab.Cells(4, 1).RenderObject = o
        'END TESTING 
        '****************************************

        rtab.Cells(5, 0).Text = "Height"
        'rtab.Cells(5, 1).Text = MyChar.Height
        '****************************************
        'TESTING for SheetActionFields
        actor = New SheetActionField
        actor.Name = "char Height"
        actor.ActionType = SheetAction.EditPlainText
        actor.Dock = PanelDockStyle.Fill
        actor.ActionObject = MyChar
        actor.ActionTag = "Height"
        actor.FamilyGroup = True

        o = New RenderText
        o.Text = MyChar.Height
        o.UserData = actor

        rtab.Cells(5, 1).RenderObject = o
        'END TESTING 
        '****************************************

        rtab.Cells(6, 0).Text = "Weight"
        'rtab.Cells(6, 1).Text = MyChar.Weight
        '****************************************
        'TESTING for SheetActionFields
        actor = New SheetActionField
        actor.Name = "char Weight"
        actor.ActionType = SheetAction.EditPlainText
        actor.Dock = PanelDockStyle.Fill
        actor.ActionObject = MyChar
        actor.ActionTag = "Weight"
        actor.FamilyGroup = True

        o = New RenderText
        o.Text = MyChar.Weight
        o.UserData = actor

        rtab.Cells(6, 1).RenderObject = o
        'END TESTING 
        '****************************************

        rtab.Cells(7, 0).Text = "Age"
        'rtab.Cells(7, 1).Text = MyChar.Age
        '****************************************
        'TESTING for SheetActionFields
        actor = New SheetActionField
        actor.Name = "char Age"
        actor.ActionType = SheetAction.EditPlainText
        actor.Dock = PanelDockStyle.Fill
        actor.ActionObject = MyChar
        actor.ActionTag = "Age"
        actor.FamilyGroup = True

        o = New RenderText
        o.Text = MyChar.Age
        o.UserData = actor

        rtab.Cells(7, 1).RenderObject = o
        'END TESTING 
        '****************************************

        'Add table to sheet
        MyPrintDoc.Body.Children.Add(rtab)
    End Sub










    Private Sub CreateHeader()
        Dim ra As New RenderArea
        Dim theader As New C1.C1Preview.RenderTable(MyPrintDoc)

        theader.Cols(0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left

        If MyOptions.InSheetView Then
            theader.Cols(TraitTypes.LastItemType - 1).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Right

            theader.Cells(0, 0).Text = MyChar.FullName
            theader.Cells(0, 0).SpanCols = TraitTypes.LastItemType - 2
            theader.Cells(0, TraitTypes.LastItemType - 1).Text = MyChar.TotalCost & " pts"
        Else
            theader.Cols(1).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Right

            theader.Cells(0, 0).Text = MyChar.FullName
            theader.Cells(0, 1).Text = MyChar.TotalCost & " pts"
        End If

        theader.Style.Borders.Bottom = New C1.C1Preview.LineDef("0.5mm", Color.Black)
        theader.Rows(0).Style.Font = New Font("Times New Roman", 6)

        '****************************************
        'TESTING for SheetActionFields
        If MyOptions.InSheetView Then
            Dim o As RenderText
            Dim actor As SheetActionField
            Dim s As String

            For i = TraitTypes.Stats To TraitTypes.LastItemType
                s = ReturnListNameNew(i)

                actor = New SheetActionField
                actor.Name = s & " button"
                actor.ActionType = SheetAction.PermanentButton
                actor.Dock = PanelDockStyle.Fill
                actor.ActionObject = MyChar
                actor.ActionTag = "[" & s & "]"
                actor.Text = s

                o = New RenderText
                o.Text = " "
                o.UserData = actor

                theader.Cells(1, i - 1).RenderObject = o
            Next

        End If
        'END TESTING 
        '****************************************


        ra.Children.Add(theader)

        'without the next bit, the sheet body begins directly on the base of the header, with no whitespace, which I don't like
        ra.Children.Add(New RenderText(vbCrLf, New Font("Times New Roman", 6)))

        MyPrintDoc.PageLayout.PageHeader = ra
        'MyPrintDoc.PageLayout.PageHeader = theader
    End Sub
    Private Sub CreateFooter()
        Dim ra As New RenderArea
        Dim theader As New C1.C1Preview.RenderTable(MyPrintDoc)

        theader.Cols(0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        theader.Cols(1).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Center
        theader.Cols(2).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Right

        theader.Style.Borders.Top = New C1.C1Preview.LineDef("0.5mm", Color.Black)
        theader.Rows(0).Style.Font = New Font("Times New Roman", 6)

        If MyChar.QuickRulesOptionsCode() <> "" Then
            theader.Cells(0, 0).Text = "GURPS Character Assistant 5 (" & MyChar.QuickRulesOptionsCode() & ")"
        Else
            theader.Cells(0, 0).Text = "GURPS Character Assistant 5"
        End If
        theader.Cells(0, 1).Text = "Page [[PageNo]] of [[PageCount]]."
        theader.Cells(0, 2).Text = "http://www.sjgames.com/gurps/characterassistant/"

        'without the next bit, the sheet body begins directly on the base of the header, with no whitespace, which I don't like
        ra.Children.Add(New RenderText(vbCrLf, New Font("Times New Roman", 6)))

        ra.Children.Add(theader)

        MyPrintDoc.PageLayout.PageFooter = ra
        'MyPrintDoc.PageLayout.PageFooter = theader
    End Sub
    Private Sub CreateStyles()


    End Sub

    Public Function ReturnRTFHeader() As String
        'Remember: this header + RTF text + "}"

        Dim RTFHeader As String
        RTFHeader = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033"
        RTFHeader = RTFHeader & "{\fonttbl{\f0\fnil\fcharset0 Times New Roman;}"
        'RTFHeader = RTFHeader & "{\f1\fnil\fcharset0 Renfrew;}"
        RTFHeader = RTFHeader & "}"
        'RTFHeader = RTFHeader & "{\colortbl ;\red255\green0\blue0;\red0\green77\blue187;\red0\green176\blue80;\red156\green133\blue192;}"
        RTFHeader = RTFHeader & "{\colortbl ;\red255\green0\blue0;\red0\green255\blue0;\red0\green0\blue255;}"
        RTFHeader = RTFHeader & "\viewkind4\uc1\pard\sa200\sl276\slmult1\lang9\fs20 "

        Return RTFHeader
    End Function
    Private Function ValidTrait(ByVal curItem As GCATrait, ByVal ItemType As Integer) As Boolean
        If curItem.TagItem("hide") <> "" AndAlso ShowHidden(ItemType - 1) = False Then
            Return False
        End If
        Select Case ItemType
            Case Stats
                If curItem.TagItem("mainwin") = "" Then
                    Return False
                End If
        End Select

        Return True
    End Function
    Private Function ValidWeaponTrait(ByVal curItem As GCATrait, ByVal WeaponType As Integer) As Boolean
        If curItem.TagItem("hide") <> "" Then
            'hidden, okay for non-stats and non-equipment
            'If curItem.ItemType = Equipment OrElse curItem.ItemType = Stats Then
            Return False
            'End If
        End If
        Select Case WeaponType
            Case Hand
                If curItem.DamageModeTagItemCount("charreach") > 0 Then
                    Return True
                End If
            Case Ranged
                If curItem.DamageModeTagItemCount("charrangemax") > 0 Then
                    Return True
                End If
        End Select

        Return False
    End Function

    Private Sub CreateFields(ItemType As Integer)
        'This routine creates all the fields we'll have in the list,
        'for all the valid traits, and if ShowComponents, then
        'for all children of valid traits, too.
        Dim curItem As GCATrait
        Dim ok As Boolean = True
        Dim tmp As String = ""

        Dim curTrait As Integer = 0

        Fields = New ArrayList '0 based




        '*****
        'NEW TESTING 2015 09 06
        '*****
        Dim DoingOrderBy As Boolean = False
        Dim GroupByCat As Boolean = False
        Dim GroupByTag As Boolean = False
        Dim SpecifiedTag As String = ""
        Dim IncludeTagPartInHeader As Boolean = True
        Dim SpecifiedCatsOnly As Boolean = False
        Dim GroupsAtEnd As Boolean = False
        Dim ValuesToGroupBy As New Collection
        Dim GroupedTraits As New Dictionary(Of String, Collection)



        '* Grouping
        'This object will use the character's specified grouping options
        'to create grouped data for us to use in printing, so we don't
        'have to do it all ourselves in each sheet.
        Dim gtlb As New GroupedTraitListBuilder
        gtlb.Character = MyChar 'set the character to use
        gtlb.ItemType = ItemType 'set the trait type, so it knows what grouping options to use
        If ItemType = Stats Then
            'order by order specified in mainwin() tag
            gtlb.OrderBy = "mainwin+#" 'set the ordering options; we only order attributes here
        End If
        gtlb.ShowComponents = ShowComponents(ItemType - 1) 'honor the sheet option for whether to show component traits
        gtlb.ShowHiddenTraits = ShowHidden(ItemType - 1) 'honor the sheet option for whether to show hidden traits
        'build grouped data
        gtlb.BuildGroupedTraits() 'this builds our grouped data lists, so we can get them using the functions below
        'get grouped data
        ValuesToGroupBy = gtlb.ValuesToGroupBy 'the various items by which data is grouped; also the keys to get the collections of traits in GroupedTraits
        GroupedTraits = gtlb.GroupedTraits 'a dictionary of collections of traits, keyed by the ValuesToGroupBy


        'and store the used grouping data for use below
        GroupByCat = False
        GroupByTag = False
        SpecifiedCatsOnly = False
        GroupsAtEnd = False
        SpecifiedTag = ""
        IncludeTagPartInHeader = True
        With gtlb.GroupingOptions 'this is the grouping options object used to group our data above.
            If .GroupingType = TraitGroupingType.ByCategory Then
                GroupByCat = True
                If .SpecifiedValuesOnly Then
                    SpecifiedCatsOnly = True

                    If .GroupsAtEnd Then GroupsAtEnd = True
                End If
            End If
            If .GroupingType = TraitGroupingType.ByTag Then
                GroupByTag = True
                If .SpecifiedValuesOnly Then
                    SpecifiedCatsOnly = True
                    SpecifiedTag = .SpecifiedTag
                    IncludeTagPartInHeader = .IncludeTagPartInHeader

                    If .GroupsAtEnd Then GroupsAtEnd = True
                End If
            End If
        End With


        '* Allow grouping by category
        If GroupByCat Then
            If ValuesToGroupBy.Count = 0 Then GroupByCat = False
        End If
        '* Allow grouping by tag
        If GroupByTag Then
            If ValuesToGroupBy.Count = 0 Then GroupByTag = False
        End If


        If GroupByTag OrElse GroupByCat Then
            'Debug.Print("BoxTraitList.CreateFields(" & ReturnListName(MyItemType) & ") GROUPED")

            Dim tmpTagValue As String
            Dim curItems As Collection

            If SpecifiedCatsOnly Then
                'only listing certain groups
                If GroupsAtEnd Then
                    'do ungrouped first, then groups
                    tmpTagValue = ""

                    If ValuesToGroupBy.Contains(tmpTagValue) Then
                        'we have some to print
                        curItems = GroupedTraits(tmpTagValue)

                        If curItems.Count > 0 Then
                            'print placeholder
                            Dim Placeholder As New GCATrait
                            Placeholder.Name = "Ungrouped " & ReturnLongListName(ItemType)
                            Placeholder.Owner = MyChar
                            Placeholder.IDKey = 0 'MyChar.NewKey '0
                            Placeholder.ItemType = TraitTypes.None  '-1
                            AddField(Placeholder, ItemType)

                            'print all items
                            For Each curItem In curItems
                                AddField(curItem, ItemType)
                            Next
                        End If
                    End If

                    'do all other groups
                    For Each s As String In ValuesToGroupBy
                        'don't include the ungrouped ones
                        If s <> "" Then
                            curItems = GroupedTraits(s)

                            If curItems.Count > 0 Then
                                'print placeholder
                                Dim Placeholder As New GCATrait
                                If GroupByTag AndAlso IncludeTagPartInHeader Then
                                    Placeholder.Name = SpecifiedTag & " = " & s
                                Else
                                    Placeholder.Name = s
                                End If
                                Placeholder.Owner = MyChar
                                Placeholder.IDKey = 0 'MyChar.NewKey '0
                                Placeholder.ItemType = TraitTypes.None  '-1
                                AddField(Placeholder, ItemType)

                                'print all items
                                For Each curItem In curItems
                                    AddField(curItem, ItemType)
                                Next
                            End If

                        End If
                    Next

                Else
                    'do groups, then ungrouped

                    'do all other groups
                    For Each s As String In ValuesToGroupBy
                        'don't include the ungrouped ones
                        If s <> "" Then
                            curItems = GroupedTraits(s)

                            If curItems.Count > 0 Then
                                'print placeholder
                                Dim Placeholder As New GCATrait
                                If GroupByTag AndAlso IncludeTagPartInHeader Then
                                    Placeholder.Name = SpecifiedTag & " = " & s
                                Else
                                    Placeholder.Name = s
                                End If
                                Placeholder.Owner = MyChar
                                Placeholder.IDKey = 0 'MyChar.NewKey '0
                                Placeholder.ItemType = TraitTypes.None  '-1
                                AddField(Placeholder, ItemType)

                                'print all items
                                For Each curItem In curItems
                                    AddField(curItem, ItemType)
                                Next
                            End If

                        End If
                    Next

                    'do ungrouped
                    tmpTagValue = ""

                    If ValuesToGroupBy.Contains(tmpTagValue) Then
                        'we have some to print
                        curItems = GroupedTraits(tmpTagValue)

                        If curItems.Count > 0 Then
                            'print placeholder
                            Dim Placeholder As New GCATrait
                            Placeholder.Name = "Ungrouped " & ReturnLongListName(ItemType)
                            Placeholder.Owner = MyChar
                            Placeholder.IDKey = 0 'MyChar.NewKey '0
                            Placeholder.ItemType = TraitTypes.None  '-1
                            AddField(Placeholder, ItemType)

                            'print all items
                            For Each curItem In curItems
                                AddField(curItem, ItemType)
                            Next
                        End If
                    End If



                End If
            Else
                'break out all groups
                For Each s As String In ValuesToGroupBy
                    curItems = GroupedTraits(s)

                    If curItems.Count > 0 Then
                        'print placeholder
                        Dim Placeholder As New GCATrait
                        If s = "" Then
                            If GroupByTag AndAlso IncludeTagPartInHeader Then
                                Placeholder.Name = SpecifiedTag & " = (no value)"
                            Else
                                Placeholder.Name = "(no value)"
                            End If
                        Else
                            If GroupByTag AndAlso IncludeTagPartInHeader Then
                                Placeholder.Name = SpecifiedTag & " = " & s
                            Else
                                Placeholder.Name = s
                            End If
                        End If
                        Placeholder.Owner = MyChar
                        Placeholder.IDKey = 0 'MyChar.NewKey '0
                        Placeholder.ItemType = TraitTypes.None  '-1
                        AddField(Placeholder, ItemType)

                        'print all items
                        For Each curItem In curItems
                            AddField(curItem, ItemType)
                        Next
                    End If
                Next
            End If

        Else
            'ORIGINAL CODE IN THIS ELSE BLOCK
            Dim MyTraits As SortedTraitCollection = MyChar.ItemsByType(ItemType)

            For curTrait = 1 To MyTraits.Count
                If ItemType = Stats Then
                    curItem = MyTraits.OrderedItem(curTrait)
                Else
                    curItem = MyTraits.Item(curTrait)
                End If

                If Not ValidTrait(curItem, ItemType) Then
                    'don't add this item
                    Continue For
                End If

                ok = True
                If curItem.ParentKey <> "" Then
                    'If MyChar.Items.Item(curItem.ParentKey).ItemType = ItemType Then
                    'don't show this one, because it's a child of another trait
                    ok = False
                    'End If
                End If
                If ShowComponents(ItemType - 1) Then
                    'don't show items that are components
                    'at this stage, as they'll be shown with
                    'their parents
                    If ok Then
                        If curItem.TagItem("keep") <> "" Then
                            tmp = "k" & curItem.TagItem("keep")
                            If MyChar.Items.Item(tmp).ItemType = ItemType Then
                                'don't show this one, because it's a component of another trait
                                'that is being displayed in this list
                                ok = False
                            End If
                        End If
                    End If
                End If
                If ok Then
                    'This trait is valid, so include it.
                    AddField(curItem, ItemType)
                End If
            Next


        End If

        '*****
        'END NEW TESTING
        '*****


        'If ItemType = Stats Then
        '    'order by order specified in mainwin() tag
        '    MyTraits.OrderBy("mainwin+#")
        '    MyTraits.OrderItems()
        'End If

        'For curTrait = 1 To MyTraits.Count
        '    If ItemType = Stats Then
        '        curItem = MyTraits.OrderedItem(curTrait)
        '    Else
        '        curItem = MyTraits.Item(curTrait)
        '    End If

        '    If Not ValidTrait(curItem, ItemType) Then
        '        'don't add this item
        '        Continue For
        '    End If

        '    ok = True
        '    If curItem.ParentKey <> "" Then
        '        'If MyChar.Items.Item(curItem.ParentKey).ItemType = ItemType Then
        '        'don't show this one, because it's a child of another trait
        '        ok = False
        '        'End If
        '    End If
        '    If ShowComponents(ItemType - 1) Then
        '        'don't show items that are components
        '        'at this stage, as they'll be shown with
        '        'their parents
        '        If ok Then
        '            If curItem.TagItem("keep") <> "" Then
        '                tmp = "k" & curItem.TagItem("keep")
        '                If MyChar.Items.Item(tmp).ItemType = ItemType Then
        '                    'don't show this one, because it's a component of another trait
        '                    'that is being displayed in this list
        '                    ok = False
        '                End If
        '            End If
        '        End If
        '    End If
        '    If ok Then
        '        'This trait is valid, so include it.
        '        AddField(curItem, ItemType)
        '    End If
        'Next

        'Now, all the fields have been created, so we now need to determine
        'whether each field will need to draw the connector lines for components
        'that may be above and below them.

        If Fields.Count = 0 Then Exit Sub

        'We need to go through backwards, because that way we will know if we're in
        'an indented zone before we go into a deeper zone - if that's the case, we'll need
        'to draw the outer lines for the existing shallower zones.
        Dim i As Integer
        Dim curField As clsDisplayField

        Dim curIndentLevel As Integer = 0
        Dim curColFieldMask As Integer = 0

        'If EngineConfig.ShowComponents(ItemType) Then
        'don't waste our time if we're not even going to use the component stuff

        curIndentLevel = Fields(Fields.Count - 1).IsComponent
        curColFieldMask = (2 ^ curIndentLevel)

        'Our OutsideComponentLinesToDraw *includes* the current indent level
        For i = Fields.Count - 1 To 0 Step -1
            curField = Fields(i)

            If curField.IsComponent > curIndentLevel Then
                'we've moved deeper, so if any lines
                'are being drawn on the outside, we need to 
                'include those on our field.

                'set the new indent level
                curIndentLevel = curField.IsComponent

                'take the level from the item we just came in from!
                If i + 1 < Fields.Count Then
                    curColFieldMask = Fields(i + 1).OutsideComponentLinesToDraw Or (2 ^ curField.IsComponent)
                End If

            ElseIf curField.IsComponent < curIndentLevel Then
                'backed out a level

                'exclude the existing level in our mask
                curColFieldMask = curColFieldMask Xor (2 ^ curIndentLevel)

                'set the new indent level
                curIndentLevel = curField.IsComponent

                'be sure the current level is included
                curColFieldMask = curColFieldMask Or (2 ^ curIndentLevel)
            End If

            curField.OutsideComponentLinesToDraw = curColFieldMask
        Next
        'End If


    End Sub
    Private Sub CreateWeaponFields(WeaponType As Integer)
        'This routine creates all the fields we'll have in the list,
        'for all the valid traits, and if ShowComponents, then
        'for all children of valid traits, too.
        Dim curItem As GCATrait
        Dim tmp As String = ""

        Dim curTrait As Integer = 0

        Fields = New ArrayList '0 based

        For curTrait = 1 To MyChar.Items.Count
            curItem = MyChar.Items.Item(curTrait)

            If ValidWeaponTrait(curItem, WeaponType) Then
                'This trait is valid, so include it.
                AddWeaponField(curItem, WeaponType)
            End If
        Next

    End Sub
    Private Sub AddField(curItem As GCATrait, ItemType As Integer, Optional SubComponentLevel As Integer = 0, Optional IsLastComponent As Boolean = False, Optional ParentField As clsDisplayField = Nothing)
        Dim curField As clsDisplayField
        Dim tmp As String
        Dim kidKey As String
        Dim i As Integer
        Dim childItem As GCATrait

        'create the field
        curField = New clsDisplayField

        If ItemType = Skills Then
            If curItem.TagItem("combolevel") <> "" Then
                curField.IsCombo = True
            Else
                curField.IsCombo = False
            End If
        End If
        curField.IsComponent = SubComponentLevel
        If SubComponentLevel > 0 Then
            curField.ComponentParent = ParentField
            curField.IsLastComponent = IsLastComponent
        End If


        'add the trait
        curField.Data = curItem
        curField.LinePen = New Pen(Brushes.LightGray) 'New Pen(MyColorProfile.TraitTextColor)

        Fields.Add(curField)

        '****************************************
        'TESTING for SheetActionFields

        curField.SheetActionFields = New Dictionary(Of Integer, SheetActionField)

        'We're going to get extra fancy, just to show what you can do,
        'but we're currently going to offer editing only in the Name column.
        Dim rootSAF As SheetActionField
        Dim extraAF As SheetActionField
        Dim panelAF As SheetActionField

        rootSAF = New SheetActionField
        rootSAF.Name = "display name label in center"
        rootSAF.ActionType = SheetAction.Label
        rootSAF.Dock = PanelDockStyle.Fill
        rootSAF.ActionObject = curItem
        rootSAF.ActionTag = "displayname"

        'These are all additional action objects that will
        'activate at the same time, but we have to store this in the 
        'ExtraActionFields list of our first item.
        '
        'If you do something like this, be sure that you don't specify
        'a Panel as the root item, because all ExtraActionFields on a
        'Panel are built into the Panel, which has some different results.
        '
        'We build a panel farther down this example.
        '
        'For these, PanelDockStyle tells GCA which side of the field
        'area you want the control to be set on, and EdgeAlign says
        'whether you want it inside the field, or outside of it.

        extraAF = New SheetActionField
        extraAF.Name = "title label above"
        extraAF.ActionType = SheetAction.Label
        extraAF.ActionObject = rootSAF.ActionObject 'Always be sure to specify this for every Action, as it is the Trait we're working on.
        extraAF.Dock = PanelDockStyle.Top
        extraAF.EdgeAlign = EdgeAlign.Outside
        extraAF.Text = "Be Happy - Edit Stuff" 'for a Label, specify text fills that text. To fill with the tag of the trait, specify ActionTag instead. 
        extraAF.BackColor = Color.Black
        extraAF.ForeColor = Color.White
        rootSAF.ExtraActionFields.Add(extraAF)

        extraAF = New SheetActionField
        extraAF.Name = "delete button on left"
        extraAF.ActionType = SheetAction.DeleteButton
        extraAF.ActionObject = rootSAF.ActionObject
        extraAF.Dock = PanelDockStyle.Left
        extraAF.EdgeAlign = EdgeAlign.Outside
        rootSAF.ExtraActionFields.Add(extraAF)

        extraAF = New SheetActionField
        extraAF.Name = "edit button on right"
        extraAF.ActionType = SheetAction.EditButton
        extraAF.ActionObject = rootSAF.ActionObject
        extraAF.Dock = PanelDockStyle.Right
        extraAF.EdgeAlign = EdgeAlign.Outside
        rootSAF.ExtraActionFields.Add(extraAF)

        'Create a panel

        panelAF = New SheetActionField
        panelAF.Name = "panel on bottom"
        panelAF.ActionType = SheetAction.Panel
        panelAF.ActionObject = rootSAF.ActionObject
        panelAF.Dock = PanelDockStyle.Bottom
        panelAF.BackColor = Color.Black
        panelAF.EdgeAlign = EdgeAlign.Outside
        rootSAF.ExtraActionFields.Add(panelAF) 'still adding it to our root item.


        'Panel Objects
        'All of these are added to the ExtraActionFields of the panel
        'we just added above.
        '
        'For a panel, the DockStyle works like controls in .Net,
        'docking themselves to that edge. You'll need to use Zorder
        'to make sure that things end up in the order you expect, if
        'you mix and match docking sides with a Fill. A higher
        'Zorder will end up in the front, so use that for any Fill item.

        extraAF = New SheetActionField
        extraAF.Name = "delete button inside panel"
        extraAF.ActionType = SheetAction.DeleteButton
        extraAF.ActionObject = rootSAF.ActionObject
        extraAF.Dock = PanelDockStyle.Right
        extraAF.ZOrder = 1
        panelAF.ExtraActionFields.Add(extraAF)

        extraAF = New SheetActionField
        extraAF.Name = "mods button inside panel"
        extraAF.ActionType = SheetAction.ModifiersButton
        extraAF.ActionObject = rootSAF.ActionObject
        extraAF.Dock = PanelDockStyle.Right
        extraAF.ZOrder = 2
        panelAF.ExtraActionFields.Add(extraAF)

        extraAF = New SheetActionField
        extraAF.Name = "inc button inside panel"
        extraAF.ActionType = SheetAction.IncButton
        extraAF.ActionObject = rootSAF.ActionObject
        extraAF.Dock = PanelDockStyle.Right
        extraAF.ZOrder = 3
        panelAF.ExtraActionFields.Add(extraAF)

        extraAF = New SheetActionField
        extraAF.Name = "dec button inside panel"
        extraAF.ActionType = SheetAction.DecButton
        extraAF.ActionObject = rootSAF.ActionObject
        extraAF.Dock = PanelDockStyle.Right
        extraAF.ZOrder = 4
        panelAF.ExtraActionFields.Add(extraAF)

        'END TESTING
        '****************************************

        curField.Values = New ArrayList
        curField.Values.Add("") 'icon col
        Select Case ItemType
            Case Languages
                '4 columns: icon & name & level & points
                curField.Values.Add(curItem.FullNameTL)
                curField.Values.Add(curItem.LevelName)
                curField.Values.Add(curItem.Points)

                '****************************************
                'TESTING for SheetActionFields
                If curItem.IDKey <> 0 Then curField.SheetActionFields.Add(1, rootSAF)
                'END TESTING
                '****************************************

            Case Stats
                '4 columns: icon & name & score & points
                curField.Values.Add(curItem.DisplayName)
                curField.Values.Add(curItem.Score)
                curField.Values.Add(curItem.Points)

                '****************************************
                'TESTING for SheetActionFields
                If curItem.IDKey <> 0 Then
                    curField.SheetActionFields.Add(1, rootSAF)


                    'decided to also allow editing the SCORE here
                    Dim tmpAF As SheetActionField
                    tmpAF = New SheetActionField
                    tmpAF.Name = "change the score of this attribute"
                    tmpAF.ActionType = SheetAction.ChangeScoreDecimal
                    tmpAF.Dock = PanelDockStyle.Fill
                    tmpAF.ActionObject = curItem
                    tmpAF.ActionTag = "score"
                    curField.SheetActionFields.Add(2, tmpAF)

                End If
                'END TESTING
                '****************************************

            Case Skills
                '5 columns: icon & name & level & rel level & points
                curField.Values.Add(curItem.DisplayName)
                If curField.IsCombo Then
                    curField.Values.Add(curItem.TagItem("combolevel"))
                    'curField.Values.Add(curItem.RelativeLevel)
                Else
                    curField.Values.Add(curItem.Level)
                    'curField.Values.Add(curItem.RelativeLevel)
                End If
                curField.Values.Add(curItem.Points)

                '****************************************
                'TESTING for SheetActionFields
                If curItem.IDKey <> 0 Then
                    curField.SheetActionFields.Add(1, rootSAF)

                    'decided to also allow editing the LEVEL here
                    Dim tmpAF As SheetActionField
                    tmpAF = New SheetActionField
                    tmpAF.Name = "change the level of this skill"
                    tmpAF.ActionType = SheetAction.ChangeScoreInteger
                    tmpAF.Dock = PanelDockStyle.Fill
                    tmpAF.ActionObject = curItem
                    tmpAF.ActionTag = "level"
                    curField.SheetActionFields.Add(2, tmpAF)
                End If
                'END TESTING
                '****************************************

            Case Spells
                '4 columns: icon & name & level & points
                curField.Values.Add(curItem.DisplayName)
                curField.Values.Add(curItem.Level)
                curField.Values.Add(curItem.Points)

                '****************************************
                'TESTING for SheetActionFields
                If curItem.IDKey <> 0 Then
                    curField.SheetActionFields.Add(1, rootSAF)

                    'decided to also allow editing the LEVEL here
                    Dim tmpAF As SheetActionField
                    tmpAF = New SheetActionField
                    tmpAF.Name = "change the level of this spell"
                    tmpAF.ActionType = SheetAction.ChangeScoreInteger
                    tmpAF.Dock = PanelDockStyle.Fill
                    tmpAF.ActionObject = curItem
                    tmpAF.ActionTag = "level"
                    curField.SheetActionFields.Add(2, tmpAF)
                End If
                'END TESTING
                '****************************************

            Case Equipment
                '7 columns: icon & name & qty & cost & weight
                curField.Values.Add(curItem.DisplayName)
                curField.Values.Add(curItem.TagItem("count"))
                curField.Values.Add(curItem.TagItem("cost"))
                curField.Values.Add(curItem.TagItem("weight"))

                '****************************************
                'TESTING for SheetActionFields
                If curItem.IDKey <> 0 Then curField.SheetActionFields.Add(1, rootSAF)
                'END TESTING
                '****************************************

            Case Else
                '3 columns: icon & name & points
                curField.Values.Add(curItem.DisplayName)
                curField.Values.Add(curItem.Points)

                '****************************************
                'TESTING for SheetActionFields
                If curItem.IDKey <> 0 Then curField.SheetActionFields.Add(1, rootSAF)
                'END TESTING
                '****************************************
        End Select
        curField.Values.Add("") 'icon col

        'NEW TESTING 2015 09 06
        If curItem.IDKey = 0 Then
            'PLACEHOLDER
            'only print name of the item, but leave the blanks for the other columns
            For curCol = 2 To curField.Values.Count - 1
                curField.Values(curCol) = ""
            Next
        End If
        'END NEW


        'if ShowComponents is True, then we need to also include
        'all of the item's component traits
        If ShowComponents(ItemType - 1) Then
            'check for kids
            tmp = Trim(curItem.TagItem("pkids"))
            'Add any children it has
            If tmp <> "" Then
                'there are kids, referenced by IDKey only
                Dim Kids(0) As String
                Kids = ParseArray(tmp, ",", 0)

                For i = 1 To UBound(Kids)
                    kidKey = "k" & Kids(i)
                    childItem = MyChar.Items.Item(kidKey)
                    If childItem IsNot Nothing Then
                        AddField(childItem, ItemType, SubComponentLevel + 1, i = UBound(Kids), curField)
                    End If
                Next
            End If
        End If

        'always include children
        'check for children
        tmp = Trim(curItem.ChildKeyList)
        'Add any children it has
        If tmp <> "" Then
            'there are kids, referenced by CollectionKey
            Dim Kids(0) As String
            Kids = ParseArray(tmp, ",", 0)

            For i = 1 To UBound(Kids)
                childItem = MyChar.Items.Item(Kids(i))
                If childItem IsNot Nothing Then
                    AddField(childItem, ItemType, SubComponentLevel + 1, i = UBound(Kids), curField)
                End If
            Next
        End If

    End Sub
    Private Sub AddWeaponField(curItem As GCATrait, WeaponType As Integer)
        Dim curField, ParentField As clsDisplayField
        Dim tmp As String
        Dim curMode, ModeCount As Integer

        Dim DamageText, RangeText As String

        Select Case WeaponType
            Case Hand
                'Add all modes
                ModeCount = curItem.DamageModeTagItemCount("charreach")
                curMode = curItem.DamageModeTagItemAt("charreach")

                If ModeCount = 1 Then
                    'all on one line

                    curField = New clsDisplayField
                    curField.Data = curItem
                    Fields.Add(curField)

                    curField.Values = New ArrayList

                    'add values
                    curField.Values.Add("") 'icon col

                    curField.Values.Add(curItem.FullNameTL)

                    DamageText = curItem.DamageModeTagItem(curMode, "chardamage")
                    If curItem.DamageModeTagItem(curMode, "chararmordivisor") <> "" Then
                        tmp = curItem.DamageModeTagItem(curMode, "chararmordivisor")
                        If tmp = "!" Then
                            tmp = ChrW(8734)
                        End If
                        DamageText = DamageText & " (" & tmp & ")"
                    End If
                    DamageText = DamageText & " " & curItem.DamageModeTagItem(curMode, "chardamtype")
                    If curItem.DamageModeTagItem(curMode, "charradius") <> "" Then
                        DamageText = DamageText & " (" & curItem.DamageModeTagItem(curMode, "charradius") & ")"
                    End If
                    curField.Values.Add(DamageText)

                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "charreach"))

                    'If Options.Value("ShowSkillParry") Then
                    'tmp = curItem.DamageModeTagItem(curMode, "charskillscore") & " (" & curItem.DamageModeTagItem(curMode, "charparryscore") & ")"
                    'Else
                    tmp = curItem.DamageModeTagItem(curMode, "charparry")
                    'End If
                    curField.Values.Add(tmp)

                    'If Options.Value("ShowMinST") Then
                    '    curField.Values.Add(curItem.DamageModeTagItem(curMode, "charminst"))
                    'End If

                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "notes"))

                    curField.Values.Add("") 'icon col
                Else
                    'base info on one line, mode info on separate lines

                    'build base line
                    curField = New clsDisplayField
                    curField.Data = curItem
                    Fields.Add(curField)
                    ParentField = curField

                    curField.Values = New ArrayList

                    'add values
                    curField.Values.Add("") 'icon col

                    curField.Values.Add(curItem.FullNameTL)

                    curField.Values.Add("")
                    curField.Values.Add("")
                    curField.Values.Add("")
                    'curField.Values.Add("")
                    curField.Values.Add("")

                    curField.Values.Add("") 'for the broken box column
                    curField.Values.Add(curItem.TagItem("cost"))
                    curField.Values.Add(curItem.TagItem("weight"))
                    curField.Values.Add("") 'icon col

                    'now build mode lines
                    curMode = curItem.DamageModeTagItemAt("charreach")
                    Do
                        curField = New clsDisplayField
                        curField.Data = curItem
                        Fields.Add(curField)

                        curField.IsComponent = 1
                        curField.ComponentParent = ParentField

                        curField.Values = New ArrayList

                        'add values
                        curField.Values.Add("") 'icon col
                        curField.Values.Add(curItem.DamageModeName(curMode))

                        DamageText = curItem.DamageModeTagItem(curMode, "chardamage")
                        If curItem.DamageModeTagItem(curMode, "chararmordivisor") <> "" Then
                            tmp = curItem.DamageModeTagItem(curMode, "chararmordivisor")
                            If tmp = "!" Then
                                tmp = ChrW(8734)
                            End If
                            DamageText = DamageText & " (" & tmp & ")"
                        End If
                        DamageText = DamageText & " " & curItem.DamageModeTagItem(curMode, "chardamtype")
                        If curItem.DamageModeTagItem(curMode, "charradius") <> "" Then
                            DamageText = DamageText & " (" & curItem.DamageModeTagItem(curMode, "charradius") & ")"
                        End If
                        curField.Values.Add(DamageText)

                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "charreach"))

                        'If Options.Value("ShowSkillParry") Then
                        'tmp = curItem.DamageModeTagItem(curMode, "charskillscore") & " (" & curItem.DamageModeTagItem(curMode, "charparryscore") & ")"
                        'Else
                        tmp = curItem.DamageModeTagItem(curMode, "charparry")
                        'End If
                        curField.Values.Add(tmp)

                        'If Options.Value("ShowMinST") Then
                        '    curField.Values.Add(curItem.DamageModeTagItem(curMode, "charminst"))
                        'End If

                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "notes"))

                        curField.Values.Add("") 'icon col

                        'find the next applicable mode
                        curMode = curItem.DamageModeTagItemAt("charreach", curMode + 1)

                        If curMode = 0 Then
                            curField.IsLastComponent = True
                        End If
                    Loop While curMode > 0
                End If

            Case Ranged
                'Add all modes
                ModeCount = curItem.DamageModeTagItemCount("charrangemax")
                curMode = curItem.DamageModeTagItemAt("charrangemax")

                If ModeCount = 1 Then
                    'all on one line

                    curField = New clsDisplayField
                    curField.Data = curItem
                    Fields.Add(curField)

                    curField.Values = New ArrayList

                    'add values
                    curField.Values.Add("") 'icon col

                    curField.Values.Add(curItem.FullNameTL)

                    DamageText = curItem.DamageModeTagItem(curMode, "chardamage")
                    If curItem.DamageModeTagItem(curMode, "chararmordivisor") <> "" Then
                        tmp = curItem.DamageModeTagItem(curMode, "chararmordivisor")
                        If tmp = "!" Then
                            tmp = ChrW(8734)
                        End If
                        DamageText = DamageText & " (" & tmp & ")"
                    End If
                    DamageText = DamageText & " " & curItem.DamageModeTagItem(curMode, "chardamtype")
                    If curItem.DamageModeTagItem(curMode, "charradius") <> "" Then
                        DamageText = DamageText & " (" & curItem.DamageModeTagItem(curMode, "charradius") & ")"
                    End If
                    curField.Values.Add(DamageText)

                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "characc"))

                    RangeText = curItem.DamageModeTagItem(curMode, "charrangehalfdam")
                    If RangeText = "" Then
                        RangeText = curItem.DamageModeTagItem(curMode, "charrangemax")
                    Else
                        RangeText = RangeText & " / " & curItem.DamageModeTagItem(curMode, "charrangemax")
                    End If
                    curField.Values.Add(RangeText)

                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "charrof"))
                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "charshots"))
                    'If Options.Value("ShowRangedLevel") Then
                    '    curField.Values.Add(curItem.DamageModeTagItem(curMode, "charskillscore"))
                    'End If
                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "charminst"))
                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "bulk"))
                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "charrcl"))
                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "lc"))
                    curField.Values.Add(curItem.DamageModeTagItem(curMode, "notes"))

                    curField.Values.Add("") 'icon col
                Else
                    'base info on one line, mode info on separate lines

                    'build base line
                    curField = New clsDisplayField
                    curField.Data = curItem
                    Fields.Add(curField)
                    ParentField = curField

                    curField.Values = New ArrayList

                    'add values
                    curField.Values.Add("") 'icon col

                    curField.Values.Add(curItem.FullNameTL)

                    curField.Values.Add("")
                    curField.Values.Add("")
                    curField.Values.Add("")

                    curField.Values.Add("")
                    curField.Values.Add("")
                    'curField.Values.Add("")
                    curField.Values.Add("")
                    curField.Values.Add("")
                    curField.Values.Add("")
                    curField.Values.Add("")
                    curField.Values.Add("")

                    curField.Values.Add("") 'icon col

                    'now build mode lines
                    curMode = curItem.DamageModeTagItemAt("charrangemax")
                    Do
                        curField = New clsDisplayField
                        curField.Data = curItem
                        Fields.Add(curField)

                        curField.IsComponent = 1
                        curField.ComponentParent = ParentField

                        curField.Values = New ArrayList

                        'add values
                        curField.Values.Add("") 'icon col
                        curField.Values.Add(curItem.DamageModeName(curMode))

                        DamageText = curItem.DamageModeTagItem(curMode, "chardamage")
                        If curItem.DamageModeTagItem(curMode, "chararmordivisor") <> "" Then
                            tmp = curItem.DamageModeTagItem(curMode, "chararmordivisor")
                            If tmp = "!" Then
                                tmp = ChrW(8734)
                            End If
                            DamageText = DamageText & " (" & tmp & ")"
                        End If
                        DamageText = DamageText & " " & curItem.DamageModeTagItem(curMode, "chardamtype")
                        If curItem.DamageModeTagItem(curMode, "charradius") <> "" Then
                            DamageText = DamageText & " (" & curItem.DamageModeTagItem(curMode, "charradius") & ")"
                        End If
                        curField.Values.Add(DamageText)

                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "characc"))

                        RangeText = curItem.DamageModeTagItem(curMode, "charrangehalfdam")
                        If RangeText = "" Then
                            RangeText = curItem.DamageModeTagItem(curMode, "charrangemax")
                        Else
                            RangeText = RangeText & " / " & curItem.DamageModeTagItem(curMode, "charrangemax")
                        End If
                        curField.Values.Add(RangeText)

                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "charrof"))
                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "charshots"))
                        'If Options.Value("ShowRangedLevel") Then
                        '    curField.Values.Add(curItem.DamageModeTagItem(curMode, "charskillscore"))
                        'End If
                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "charminst"))
                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "bulk"))
                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "charrcl"))
                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "lc"))
                        curField.Values.Add(curItem.DamageModeTagItem(curMode, "notes"))

                        curField.Values.Add("") 'icon col

                        'find the next applicable mode
                        curMode = curItem.DamageModeTagItemAt("charrangemax", curMode + 1)

                        If curMode = 0 Then
                            curField.IsLastComponent = True
                        End If
                    Loop While curMode > 0
                End If
        End Select

    End Sub
    Private Sub SetFieldSizes()
        '*****
        '* Size and store all our trait value lines
        '*
        Dim AltLine As Boolean = False
        Dim curField As clsDisplayField
        Dim FieldIndex As Integer = 0

        For FieldIndex = 0 To Fields.Count - 1
            'Set our Field values
            curField = Fields(FieldIndex)

            curField.Alt = AltLine

            'Adjust for next lines
            AltLine = Not AltLine
        Next

    End Sub
    Private Sub SetWeaponColumns(WeaponType As Integer, LeftOffset As Double)
        Dim miniColWidth As Double = 0
        Dim minColWidth As Double = 0
        Dim medColWidth As Double = 0
        Dim medplusColWidth As Double = 0
        Dim wideColWidth As Double = 0
        Dim curCol As Integer

        PointsCol = 0

        miniColWidth = 0.25
        minColWidth = 0.375
        medColWidth = 0.5
        medplusColWidth = 0.75 '0.625
        wideColWidth = 0.875

        IconColLeft = LeftOffset
        IconColWidth = 1 / 32 ' 0

        ComponentDrawCol = 1

        '0 based columns
        Select Case WeaponType
            Case Hand
                'If Options.Value("ShowMinST") Then
                '    numCols = 11
                'else
                numCols = 6
                'End If
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                SubHeads(curCol) = "Weapon"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = 0

                curCol += 1
                SubHeads(curCol) = "Damage"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'wideColWidth
                If ColWidths(curCol) < 1 Then ColWidths(curCol) = 1

                curCol += 1
                SubHeads(curCol) = "Reach"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'medColWidth

                curCol += 1
                'If Options.Value("ShowSkillParry") Then
                'SubHeads(curCol) = "Lvl(Pry)"
                'Else
                SubHeads(curCol) = "Parry"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'medColWidth
                'End If
                'If Options.Value("ShowMinST") Then
                'curCol += 1
                'SubHeads(curCol) = "Min ST"
                'ColAligns(curCol) = AlignHorzEnum.Left
                'ColWidths(curCol) = minColWidth
                'End If

                curCol += 1
                SubHeads(curCol) = "Notes"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'medColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth


            Case Ranged
                'If Options.Value("ShowRangedLevel") Then
                '    numCols = 17
                'else
                numCols = 12
                'End If

                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                SubHeads(curCol) = "Weapon"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = 0

                curCol += 1
                SubHeads(curCol) = "Damage"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = wideColWidth

                curCol += 1
                SubHeads(curCol) = "Acc"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = minColWidth

                curCol += 1
                SubHeads(curCol) = "Range"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = wideColWidth

                curCol += 1
                SubHeads(curCol) = "RoF"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = minColWidth

                curCol += 1
                SubHeads(curCol) = "Shots"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = medColWidth

                'If Options.Value("ShowRangedLevel") Then
                'curCol += 1
                'SubHeads(curCol) = "Lvl"
                'ColAligns(curCol) = AlignHorzEnum.Left
                'ColWidths(curCol) = minColWidth
                'End If

                curCol += 1
                SubHeads(curCol) = "ST"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = minColWidth

                curCol += 1
                SubHeads(curCol) = "Bulk"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = minColWidth

                curCol += 1
                SubHeads(curCol) = "Rcl"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = minColWidth

                curCol += 1
                SubHeads(curCol) = "LC"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = miniColWidth

                curCol += 1
                SubHeads(curCol) = "Notes"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = medColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth



        End Select

    End Sub
    Private Sub SetTraitColumns(ItemType As Integer, LeftOffset As Double)
        '*
        '* Set Column Positions
        '* 
        Dim miniColWidth As Double = 0
        Dim medColWidth As Double = 0
        Dim altColWidth As Double = 0
        Dim ptsColWidth As Double = 0
        Dim minColWidth As Double = 0
        Dim medplusColWidth As Double = 0
        Dim wideColWidth As Double = 0

        ptsColWidth = 0.5 '0.375
        minColWidth = ptsColWidth
        miniColWidth = 0.25
        medColWidth = 0.5
        medplusColWidth = 0.625
        wideColWidth = 0.75

        IconColLeft = 0 'LeftOffset
        IconColWidth = 1 / 32 ' 0

        ComponentDrawCol = 1

        '0 based columns 
        Select Case ItemType
            Case Languages
                '5 columns: icon & name & level & points & icon
                Dim curCol As Integer

                numCols = 4
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                PointsCol = 0
                ComponentDrawCol = 1

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                SubHeads(curCol) = "Language"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = 0

                curCol += 1
                SubHeads(curCol) = "Level"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = 1

                curCol += 1
                SubHeads(curCol) = "Points"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'medColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth

            Case Stats
                '5 columns: icon & name & score & points & icon
                Dim curCol As Integer

                numCols = 4
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                PointsCol = 0
                ComponentDrawCol = 1

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                SubHeads(curCol) = "Attribute"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = 0

                curCol += 1
                SubHeads(curCol) = "Score"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'ptsColWidth

                curCol += 1
                SubHeads(curCol) = "Points"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'medColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth

            Case Skills
                'combos can have pretty wide levels, so we need to allow more room for them
                altColWidth = 0.75

                Dim curCol As Integer

                numCols = 4
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim AltColWidths(numCols) 'for combos
                ReDim SubHeads(numCols)
                PointsCol = 0
                ComponentDrawCol = 1

                curCol = 0
                ColWidths(curCol) = IconColWidth
                AltColWidths(curCol) = ColWidths(curCol)
                'icon col

                curCol += 1
                SubHeads(curCol) = "Skill"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = 0
                AltColWidths(curCol) = ColWidths(curCol)

                curCol += 1
                SubHeads(curCol) = "Level"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'ptsColWidth
                AltColWidths(curCol) = altColWidth

                curCol += 1
                SubHeads(curCol) = "Points"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'medColWidth
                AltColWidths(curCol) = ColWidths(curCol)

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth

            Case Spells
                '5 columns: icon & name & level & points & icon
                Dim curCol As Integer

                numCols = 4
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                PointsCol = 0
                ComponentDrawCol = 1

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                SubHeads(curCol) = "Spell"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = 0

                curCol += 1
                SubHeads(curCol) = "Level"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'ptsColWidth

                curCol += 1
                SubHeads(curCol) = "Points"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'medColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth

            Case Equipment
                '6 columns: icon & name & qty & cost & weight & icon
                Dim curCol As Integer

                numCols = 5
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                PointsCol = 0
                ComponentDrawCol = 1

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                SubHeads(curCol) = "Item"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = 0

                curCol += 1
                SubHeads(curCol) = "Qty"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'ptsColWidth

                curCol += 1
                SubHeads(curCol) = "Cost"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'medplusColWidth

                curCol += 1
                SubHeads(curCol) = "Weight"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding  'medplusColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth


            Case Else
                '4 columns: icon & name & points & icon
                Dim curCol As Integer

                numCols = 3
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                PointsCol = 0
                ComponentDrawCol = 1

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                Select Case ItemType
                    Case Perks
                        SubHeads(curCol) = "Perk"
                    Case Disads
                        SubHeads(curCol) = "Disadvantage"
                    Case Quirks
                        SubHeads(curCol) = "Quirk"
                    Case Templates
                        SubHeads(curCol) = "Template"
                    Case Cultures
                        SubHeads(curCol) = "Culture"

                    Case Else
                        SubHeads(curCol) = "Advantage"
                End Select

                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = 0

                curCol += 1
                SubHeads(curCol) = "Points"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = GetSize(SubHeads(curCol)).Width + 1 / 8 'to allow for padding 'ptsColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth
        End Select
    End Sub
    Private Function GetSize(Text As String) As SizeD
        'Gets size based on normal text, 
        'with no given size restrictions

        Dim szD As SizeD
        Dim rText As RenderText

        rText = New RenderText(Text, MyPrintDoc.Style)

        'get size
        MyPrintDoc.Body.Children.Add(rText)
        'measure
        szD = rText.CalcSize(Unit.Auto, Unit.Auto)
        MyPrintDoc.Body.Children.Remove(rText)

        Return szD
    End Function


    ''' <summary>
    ''' This class is for tracking where in a display area a field of information is.
    ''' </summary>
    ''' <remarks></remarks>
    Public Class clsDisplayField
        '*
        '* Variables
        '*

        ''' <summary>
        ''' The action field (by column key) of data in the display area
        ''' </summary>
        ''' <remarks></remarks>
        Public SheetActionFields As Dictionary(Of Integer, SheetActionField)
        ''' <summary>
        ''' The value for each column of data in the display area
        ''' </summary>
        ''' <remarks></remarks>
        Public Values As ArrayList

        'These are absolute positions
        Private pTop As Double = -1
        Private pBottom As Double = -1
        Private pLeft As Double = -1
        Private pRight As Double = -1

        'These are relative
        Private pWidth As Double = -1
        Private pHeight As Double = -1

        'useful info
        Private pAlt As Boolean = False

        'These are data items
        Private pObject As Object

        '*
        '* Properties
        '*
        Public Property IsTruncated As Boolean = False
        Public Property IsSelected As Boolean = False
        Public Property IsCombo As Boolean = False
        Public Property ObjectsDisplayed As Boolean = False
        Public Property IsComponent As Integer = 0 'Incremented for each sub-component
        Public Property ComponentParent As clsDisplayField
        Public Property IsLastComponent As Boolean = False
        'Flags for whether we need to draw the lines for components
        'below us to their parents above us inside our row
        Public Property OutsideComponentLinesToDraw As Integer = 0

        'Public Property LineColor As Color = Color.Black
        Public Property LinePen As New Pen(Brushes.Black)

        Public Property Top As Double
            Get
                Return pTop
            End Get
            Set(ByVal value As Double)
                pTop = value

                If pHeight < 0 And pBottom > 0 Then
                    pHeight = pBottom - pTop
                ElseIf pBottom < 0 And pHeight > 0 Then
                    pBottom = pTop + pHeight
                End If
            End Set
        End Property
        Public Property Bottom As Double
            Get
                Return pBottom
            End Get
            Set(ByVal value As Double)
                pBottom = value

                If pHeight < 0 And pTop > 0 Then
                    pHeight = pBottom - pTop
                ElseIf pTop < 0 And pHeight > 0 Then
                    pTop = pBottom - pHeight
                End If
            End Set
        End Property
        Public Property Left As Double
            Get
                Return pLeft
            End Get
            Set(ByVal value As Double)
                pLeft = value

                If pWidth < 0 And pRight > 0 Then
                    pWidth = pRight - pLeft
                ElseIf pRight < 0 And pWidth > 0 Then
                    pRight = pLeft + pWidth
                End If
            End Set
        End Property
        Public Property Right As Double
            Get
                Return pRight
            End Get
            Set(ByVal value As Double)
                pRight = value

                If pWidth < 0 And pLeft > 0 Then
                    pWidth = pRight - pLeft
                ElseIf pLeft < 0 And pWidth > 0 Then
                    pLeft = pWidth - pRight
                End If
            End Set
        End Property
        Public Property Width As Double
            Get
                Return pWidth
            End Get
            Set(ByVal value As Double)
                pWidth = value

                If pRight < 0 And pLeft > 0 Then
                    pRight = pLeft + pWidth
                ElseIf pLeft < 0 And pRight > 0 Then
                    pLeft = pRight - pWidth
                End If
            End Set
        End Property
        Public Property Height As Double
            Get
                Return pHeight
            End Get
            Set(ByVal value As Double)
                pHeight = value

                If pBottom < 0 And pTop > 0 Then
                    pBottom = pTop + pHeight
                ElseIf pTop < 0 And pBottom > 0 Then
                    pTop = pBottom - pHeight
                End If
            End Set
        End Property
        Public Property Point As PointF
            Get
                Return New PointF(pLeft, pTop)
            End Get
            Set(ByVal value As PointF)
                pLeft = value.X
                pTop = value.Y
            End Set
        End Property
        Public Property Size As SizeF
            Get
                Return New SizeF(pWidth, pHeight)
            End Get
            Set(ByVal value As SizeF)
                pWidth = value.Width
                pHeight = value.Height

                pRight = pLeft + pWidth
                pBottom = pTop + pHeight
            End Set
        End Property

        Public Property Alt As Boolean
            Get
                Return pAlt
            End Get
            Set(ByVal value As Boolean)
                pAlt = value
            End Set
        End Property
        Public Property Data As Object
            Get
                Return pObject
            End Get
            Set(ByVal value As Object)
                pObject = value
            End Set
        End Property

    End Class

End Class
