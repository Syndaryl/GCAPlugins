Imports GCA5Engine
Imports C1.C1Preview
Imports System.Drawing
Imports System.Reflection

'Any such DLL needs to add References to:
'
'   System.Drawing  (System.Drawing; v4.X)
'   C1.C1Report.4   (ComponentOne Reports; v4.X)
'   GCA5Engine
'   GCA5.Interfaces.DLL
'
'in order to work as a print sheet.
'
Public Class OfficialCharacterSheet
    Implements GCA5.Interfaces.IPrinterSheet

    Private Const Hand As Integer = 1
    Private Const Ranged As Integer = 2

    Private ShadeAltLines As Boolean = True
    Private LineBefore As Boolean = True

    Private MarginLeft, MarginRight, MarginTop, MarginBottom As Double
    Private PageWidth, PageHeight As Double

    Private MyStyles As Collection
    Private MyOptions As GCA5Engine.SheetOptionsManager
    Private ShowHidden() As Boolean
    Private ShowComponents() As Boolean

    Private Fields As ArrayList '0 based

    Private IconColLeft As Double
    Private IconColWidth As Double
    Private IndentStepSize As Double = 0.125
    Private ComponentDrawCol As Integer
    Private BrokenEffectCol As Integer

    Private FooterHeight As Double = 0.25

    Private SpacerHeight As Double = 1 / 16
    Private PointsBoxHeight = 1.25
    Private SpeedRangeTableHeight As Double = 3.25
    Private pointBoxWidth As Double = 7 / 16

    Private numCols As Integer = 0
    Private PointsCol As Integer = 0
    Private ColLefts(0) As Double
    Private ColWidths(0) As Double
    Private ColAligns(0) As AlignHorzEnum
    Private AltColLefts(0) As Double 'for combos
    Private AltColWidths(0) As Double 'for combos

    Private MaxWidth As Double = 0
    Private SubHeads(0) As String

    Private HasOverflow As Boolean = False
    Private HasOverflowWeapons As Boolean = False
    Private OverflowFrom(LastItemType) As Integer
    Private WeaponOverflowFrom(Ranged) As Integer
    Private GrimoireOverflowFrom As Integer

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
            Return "GURPS 4th Edition Official Character Sheet"
        End Get
    End Property
    Public ReadOnly Property Description() As String Implements GCA5.Interfaces.IPrinterSheet.Description
        Get
            Dim tmp As String

            tmp = "Prints a close approximation of the official character sheet released for GURPS 4th Edition."
            tmp = tmp & vbCrLf & vbCrLf
            tmp = tmp & "Note: On A4 paper, this sheet will not print correctly with left and right margins greater than about 20mm. Use the smallest margins you can to get the best results."

            Return tmp
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
        currentDomain.Load("OfficialCharacterSheet")

        'Make an array for the list of assemblies.
        Dim assems As [Assembly]() = currentDomain.GetAssemblies()

        'List the assemblies in the current application domain.
        'Echo("List of assemblies loaded in current appdomain:")
        Dim assem As [Assembly]
        'Dim co As New ArrayList
        For Each assem In assems
            If assem.FullName.StartsWith("OfficialCharacterSheet") Then
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
        MyOptions.Value("SheetTextColor") = Color.Red
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

        Dim ok As Boolean
        Dim newOption As GCA5Engine.SheetOption
        Dim i As Integer

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


        '* Fonts *
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_Fonts"
        newOption.Type = GCA5Engine.OptionType.Header
        newOption.UserPrompt = "Font and Text Options"
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "SheetFont"
        newOption.Type = GCA5Engine.OptionType.Font
        newOption.UserPrompt = "Font for the form text items"
        newOption.DefaultValue = New Font("Times New Roman", 10)
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "SheetTextColor"
        newOption.Type = GCA5Engine.OptionType.Color
        newOption.UserPrompt = "Color for the form text items"
        newOption.DefaultValue = Color.Black
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "UserFont"
        newOption.Type = GCA5Engine.OptionType.Font
        newOption.UserPrompt = "Font for the user text items"
        newOption.DefaultValue = New Font("Arial Narrow", 10)
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "UserTextColor"
        newOption.Type = GCA5Engine.OptionType.Color
        newOption.UserPrompt = "Color for the user text items"
        newOption.DefaultValue = Color.Green
        ok = Options.AddOption(newOption)

        'NEW 2015 05 25, 2015 05 27
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "UserBonusesFont"
        newOption.Type = GCA5Engine.OptionType.Font
        newOption.UserPrompt = "Font for the lines showing 'Included' and 'Conditional' bonuses in trait listings"
        newOption.DefaultValue = New Font("Times New Roman", 8)
        ok = Options.AddOption(newOption)
        'END NEW

        'NEW 2015 05 25
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "UserBonusesColor"
        newOption.Type = GCA5Engine.OptionType.Color
        newOption.UserPrompt = "Color for the lines showing 'Included' and 'Conditional' bonuses in trait listings"
        newOption.DefaultValue = Color.Black
        ok = Options.AddOption(newOption)
        'END NEW


        '* Item Display *
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_Display"
        newOption.Type = GCA5Engine.OptionType.Header
        newOption.UserPrompt = "Display Options"
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "LineBefore"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "Draw a dotted line before each new item in a listing of items."
        newOption.DefaultValue = True
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "ShadeAltLines"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "Shade alternate lines in listings of items."
        newOption.DefaultValue = True
        ok = Options.AddOption(newOption)

        'Weapon tables 
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "AllowDynamicSizingOfWeaponBlocks"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "Allow Dynamic Sizing. Sacrifice any free space from one weapon table to allow more weapon modes to appear in an overflowing weapon table."
        newOption.DefaultValue = True
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "SacrificeNotesForWeapons"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "If more room is needed for the weapon tables, sacrifice the Character Notes block for it. (This will only be used if Allow Dynamic Sizing is also being used.)"
        newOption.DefaultValue = True
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "ProtectionNotEquipment"
        newOption.Type = GCA5Engine.OptionType.YesNo
        newOption.UserPrompt = "Print the Protection paper-doll on page 2, bumping Equipment to the Overflow pages."
        newOption.DefaultValue = False
        ok = Options.AddOption(newOption)



        '* Hidden Items *
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

        '* Component Items *
        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Header_Components"
        newOption.Type = OptionType.Header
        newOption.UserPrompt = "Show Component Traits"
        ok = Options.AddOption(newOption)

        newOption = New GCA5Engine.SheetOption
        newOption.Name = "Caption_Components"
        newOption.Type = OptionType.Caption
        newOption.UserPrompt = "Components are the traits that make up meta-traits and templates. If you would like to have them listed under their owning trait, turn that on here."
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


    End Sub
    Public Function GeneratePrinterDocument(GCAParty As Party, PageSettings As C1PageSettings, Options As GCA5Engine.SheetOptionsManager) As C1.C1Preview.C1PrintDocument Implements GCA5.Interfaces.IPrinterSheet.GeneratePrinterDocument
        Dim i As Integer

        'Notify(Name & " " & Version, Priority.Green)

        'Set our Options to the stored values we've just been given
        MyOptions = Options

        MyChar = GCAParty.Current
        MyPrintDoc = New C1.C1Preview.C1PrintDocument

        MyPrintDoc.TagOpenParen = "[["
        MyPrintDoc.TagCloseParen = "]]"

        'Initialize some page values

        'I think in inches. Note that some values returned from the engine won't be in the units set here
        MyPrintDoc.DefaultUnit = C1.C1Preview.UnitTypeEnum.Inch
        MyPrintDoc.ResolvedUnit = UnitTypeEnum.Inch

        'this is required, because I manually place everything, and generate my own page breaks
        MyPrintDoc.AllowNonReflowableDocs = True

        MyPrintDoc.Style.Font = MyOptions.Value("SheetFont")
        MyPrintDoc.Style.TextColor = MyOptions.Value("SheetTextColor")

        MyPrintDoc.PageLayout.PageSettings.PaperKind = PageSettings.PaperKind
        MyPrintDoc.PageLayout.PageSettings.Landscape = PageSettings.Landscape
        MyPrintDoc.PageLayout.PageSettings.TopMargin = PageSettings.TopMargin ' 0.5
        MyPrintDoc.PageLayout.PageSettings.BottomMargin = PageSettings.BottomMargin  '0.5
        MyPrintDoc.PageLayout.PageSettings.LeftMargin = PageSettings.LeftMargin  '0.5
        MyPrintDoc.PageLayout.PageSettings.RightMargin = PageSettings.RightMargin  '0.75

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

        'initialize some working values
        MyStyles = New Collection
        HasOverflow = False
        HasOverflowWeapons = False
        'since we'll be using a list of display fields, which are 0 based,
        'rather than our normal trait lists which are 1 based, we can't
        'use the default 0 value as an indicator of No value, so we need
        'to set our No value to -1 instead
        For i = 1 To LastItemType
            OverflowFrom(i) = -1
        Next
        For i = 1 To Ranged
            WeaponOverflowFrom(i) = -1
        Next
        GrimoireOverflowFrom = -1

        ShadeAltLines = MyOptions.Value("ShadeAltLines")
        LineBefore = MyOptions.Value("LineBefore")

        ShowHidden = MyOptions.Value("ShowHidden")
        ShowComponents = MyOptions.Value("ShowComponents")


        CreateStyles()
        CreateFooter()

        'Start working
        MyPrintDoc.StartDoc()


        PrintPage1()
        MyPrintDoc.NewPage()
        PrintPage2()

        If HasOverflow Then
            MyPrintDoc.NewPage()
            PrintPageOverflowTraits()
        End If
        If HasOverflowWeapons Then
            MyPrintDoc.NewPage()
            PrintPageOverflowWeapons()
        End If

        If MyChar.ItemsByType(Spells).Count > 0 Then
            MyPrintDoc.NewPage()
            PrintPageGrimoire()
        End If


        MyPrintDoc.EndDoc()

        Return MyPrintDoc
    End Function
    Private Sub PrintPage1()
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim col2Left, col2part2Left As Double

        'Dim ld As C1.C1Preview.LineDef

        boxLeft = MarginLeft
        boxTop = MarginTop
        boxWidth = PageWidth - MarginRight - boxLeft
        boxHeight = PageHeight - MarginBottom - boxTop

        'ld = New C1.C1Preview.LineDef(0.06, Color.AliceBlue, Drawing2D.DashStyle.Dash)
        'MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        '***** ***** ***** ***** ***** ***** ***** ***** ***** ***** ***** ***** 
        '***** Top Half of Page
        '***** ***** ***** ***** ***** ***** ***** ***** ***** ***** ***** ***** 

        '*
        '* Print Graphic Header Block
        '*
        PrintSectionPage1Header()

        '***** 
        '***** Left side of page
        '***** 
        boxTop = PrintSectionAttributes()


        '*
        '* Print Basics
        '*
        boxTop = boxTop + SpacerHeight
        boxTop = PrintSectionBasics(boxTop)

        '*
        '* Print Encumbrance & Move
        '*
        boxTop = boxTop + SpacerHeight
        boxTop = PrintSectionMove(boxTop)

        'ld = New C1.C1Preview.LineDef(0.06, Color.AliceBlue, Drawing2D.DashStyle.Dash)
        'MyPrintDoc.RenderDirectLine(MarginLeft, boxTop, PageWidth - MarginRight, boxTop, ld)


        '***** 
        '***** Right side of page
        '***** 

        '*
        '* Print Languages
        '*
        col2Left = MarginLeft + 3.625
        boxLeft = col2Left

        boxTop = PrintSectionLanguages(boxLeft)

        '*
        '* Print DR
        '*
        boxTop = boxTop + SpacerHeight
        PrintSectionDR(boxLeft, boxTop)

        '*
        '* Print TL & Cultures
        '*
        boxLeft = boxLeft + 11 / 16 + SpacerHeight
        col2part2Left = boxLeft
        boxTop = PrintSectionCultures(boxLeft, boxTop)



        '*
        '* Print Parry & Block
        '*
        boxLeft = col2Left
        boxTop = boxTop + SpacerHeight

        PrintSectionParryBlock(boxLeft, boxTop)

        '*
        '* Print Reactions
        '*
        boxLeft = col2part2Left
        boxTop = PrintSectionReactions(boxLeft, boxTop)



        '***** ***** ***** ***** ***** ***** ***** ***** ***** ***** ***** ***** 
        '***** Bottom Half of Page
        '***** ***** ***** ***** ***** ***** ***** ***** ***** ***** ***** ***** 

        PrintSectionAdsDisads(boxTop)
        PrintSectionSkills(col2Left, boxTop)



        ''TESTING
        ''The purpose of this block is to create a section that mimics the direct rendering
        ''done in the Stats block above, but with specific render objects, so that it's possible
        ''to have access to the object being rendered.
        ''
        ''This is necessary so we can set the UserData for the render object, which will allow
        ''us to track and find data fields, so that we can use edit controls on the Sheet View.
        'Dim Col1AttributeBoxWidth, Col2AttributeBoxWidth, Col1AttributeOffset, Col2AttributeOffset As Double
        'Dim stat As String = ""
        'Dim i As Integer = 0
        'Dim tmp As String = ""

        'Dim AttributeCenterStyle As Style = MyStyles("AttributeCenterStyle")
        'ld = New C1.C1Preview.LineDef(0.01, Color.Green, Drawing2D.DashStyle.Solid)

        'boxWidth = 7.5 / 16 '0.25
        'boxHeight = 7.5 / 16 + 1 / 32 '0.25
        'Col1AttributeBoxWidth = 0.5
        'Col2AttributeBoxWidth = 0.625
        'Col1AttributeOffset = 1 / 16
        'Col2AttributeOffset = 0.675

        ''* ST line
        'boxLeft = MarginLeft + Col1AttributeBoxWidth  '0.5
        'boxTop = MarginTop + 14 / 16 '13 / 16 '2

        'stat = "ST"
        'i = MyChar.ItemPositionByNameAndExt(stat, Stats)


        'Dim rr As New RenderRectangle(boxWidth, boxHeight, ld, Brushes.CornflowerBlue)
        'rr.UserData = "Testing"
        'MyPrintDoc.RenderDirect(boxLeft, boxTop, rr)

        'Dim rt As New RenderText(MyChar.Items.Item(i).Score, AttributeCenterStyle)
        'rt.UserData = "Testing2"
        'MyPrintDoc.RenderDirect(boxLeft, boxTop, rt, boxWidth, boxHeight)
        ''END TESTING


    End Sub
    Private Sub PrintPage2()
        Dim curTop As Double
        Dim col2Left As Double
        Dim curBottom As Double

        'left side for Hand block is indented to allow for logo
        Dim HandBlockLeft As Double = MarginLeft + 1 + 11 / 16
        Dim RangedBlockLeft As Double = MarginLeft

        'set our default heights, and get our desired heights, for our weapon blocks
        Dim DefaultHeightHandBlock As Double = 1 + 1 / 16 + 1 / 8
        Dim DefaultHeightRangedBlock As Double = 2 + 1 / 16
        Dim UseThisHandBlockHeight As Double = DefaultHeightHandBlock
        Dim UseThisRangedBlockHeight As Double = DefaultHeightRangedBlock

        'if we can use the notes block, this will have how much room we get from that
        Dim ExtraAvailableSpace As Double = 0
        Dim NotesSacrificed As Boolean = False

        If MyOptions.Value("AllowDynamicSizingOfWeaponBlocks") Then
            'hand block can't be too short because of the logo on the left
            Dim MinHandBlockHeight As Double = 0.5 + 0.1875 + 0.1875

            Dim DesiredHeightHandBlock As Double = GetWeaponSectionDesiredHeight(Hand, HandBlockLeft, MinHandBlockHeight)
            Dim DesiredHeightRangedBlock As Double = GetWeaponSectionDesiredHeight(Ranged, RangedBlockLeft)

            'now, see if we can fit all weapons, or at least overflow fewer, by adjusting block heights
            If DesiredHeightHandBlock < DefaultHeightHandBlock AndAlso DesiredHeightRangedBlock < DefaultHeightRangedBlock Then
                'the defaults are fine, don't need to do anything else
            Else
                'it doesn't all fit in the given space, see what we can move around

                If DesiredHeightHandBlock < DefaultHeightHandBlock Then
                    'we have room that we can take from the Hand block
                    UseThisRangedBlockHeight = UseThisRangedBlockHeight + (DefaultHeightHandBlock - DesiredHeightHandBlock)
                    UseThisHandBlockHeight = DesiredHeightHandBlock

                    If UseThisRangedBlockHeight < DesiredHeightRangedBlock Then
                        'still need more room

                        If MyOptions.Value("SacrificeNotesForWeapons") Then
                            ExtraAvailableSpace = PageHeight - MarginBottom - FooterHeight
                            ExtraAvailableSpace = ExtraAvailableSpace - (MarginTop + DefaultHeightHandBlock + SpacerHeight + DefaultHeightRangedBlock + SpacerHeight + SpeedRangeTableHeight + SpacerHeight + PointsBoxHeight)
                            If ExtraAvailableSpace < 0 Then ExtraAvailableSpace = 0
                            NotesSacrificed = True
                        End If

                        UseThisRangedBlockHeight = UseThisRangedBlockHeight + ExtraAvailableSpace
                    End If

                ElseIf DesiredHeightRangedBlock < DefaultHeightRangedBlock Then
                    'we have room that we can take from the Ranged block
                    UseThisHandBlockHeight = UseThisHandBlockHeight + (DefaultHeightRangedBlock - DesiredHeightRangedBlock)
                    UseThisRangedBlockHeight = DesiredHeightRangedBlock

                    If UseThisHandBlockHeight < DesiredHeightHandBlock Then
                        'still need more room

                        If MyOptions.Value("SacrificeNotesForWeapons") Then
                            ExtraAvailableSpace = PageHeight - MarginBottom - FooterHeight
                            ExtraAvailableSpace = ExtraAvailableSpace - (MarginTop + DefaultHeightHandBlock + SpacerHeight + DefaultHeightRangedBlock + SpacerHeight + SpeedRangeTableHeight + SpacerHeight + PointsBoxHeight)
                            If ExtraAvailableSpace < 0 Then ExtraAvailableSpace = 0
                            NotesSacrificed = True
                        End If

                        UseThisHandBlockHeight = UseThisHandBlockHeight + ExtraAvailableSpace
                    End If
                Else
                    'we don't have any spare room in either block

                    If MyOptions.Value("SacrificeNotesForWeapons") Then
                        ExtraAvailableSpace = PageHeight - MarginBottom - FooterHeight
                        ExtraAvailableSpace = ExtraAvailableSpace - (MarginTop + DefaultHeightHandBlock + SpacerHeight + DefaultHeightRangedBlock + SpacerHeight + SpeedRangeTableHeight + SpacerHeight + PointsBoxHeight)
                        If ExtraAvailableSpace < 0 Then ExtraAvailableSpace = 0
                        NotesSacrificed = True
                    End If

                    'see if one benefits entirely from the extra space we have.
                    If ExtraAvailableSpace > 0 Then
                        If DesiredHeightHandBlock <= DefaultHeightHandBlock + ExtraAvailableSpace Then
                            'we have enough room for this one.
                            UseThisHandBlockHeight = DesiredHeightHandBlock
                            'leave any left over for ranged
                            ExtraAvailableSpace = (DefaultHeightHandBlock + ExtraAvailableSpace) - UseThisHandBlockHeight
                            UseThisRangedBlockHeight = UseThisRangedBlockHeight + ExtraAvailableSpace

                        ElseIf DesiredHeightRangedBlock <= DefaultHeightRangedBlock + ExtraAvailableSpace Then
                            'we have enough room for this one.
                            UseThisRangedBlockHeight = DesiredHeightRangedBlock
                            'leave any left over for ranged
                            ExtraAvailableSpace = (DefaultHeightRangedBlock + ExtraAvailableSpace) - UseThisRangedBlockHeight
                            UseThisHandBlockHeight = UseThisHandBlockHeight + ExtraAvailableSpace

                        Else
                            'nothing fits well, just split it.
                            UseThisRangedBlockHeight = UseThisRangedBlockHeight + ExtraAvailableSpace / 2
                            'leave any left over for ranged
                            ExtraAvailableSpace = (DefaultHeightRangedBlock + ExtraAvailableSpace) - UseThisRangedBlockHeight
                            UseThisHandBlockHeight = UseThisHandBlockHeight + ExtraAvailableSpace

                        End If
                    End If

                End If
            End If

        End If


        'widths of speed/range + hit locations
        col2Left = MarginLeft + (1.75) + 1 / 16 + (1 + 4 / 16) + 1 / 16

        PrintSectionPage2Header(UseThisHandBlockHeight)

        curTop = PrintSectionHandWeapons(HandBlockLeft, MarginTop, UseThisHandBlockHeight)
        curTop = curTop + SpacerHeight

        curTop = PrintSectionRangedWeapons(RangedBlockLeft, curTop, UseThisRangedBlockHeight)
        curTop = curTop + SpacerHeight



        If MyOptions.Value("ProtectionNotEquipment") Then
            'print protection here instead of equipment
            HasOverflow = True
            OverflowFrom(Equipment) = 0

            PrintSectionProtection(col2Left, curTop)
        Else
            'print equipment
            PrintSectionEquipment(col2Left, curTop)
        End If




        PrintSectionSpeedRangeTable(MarginLeft, curTop)
        curTop = PrintSectionHitLocationTable(MarginLeft + 1.75 + 1 / 16, curTop)

        curTop = curTop + SpacerHeight
        curTop = PrintSectionInfoBox(MarginLeft + 1.75 + 1 / 16, curTop)

        curBottom = PrintSectionPointsSummary()
        curBottom = curBottom - SpacerHeight

        If NotesSacrificed Then Return 'Don't print notes, we took away the space for it

        curTop = curTop + SpacerHeight
        PrintSectionCharacterNotes(curTop, curBottom - curTop)
    End Sub
    Private Sub PrintPageOverflowTraits()
        Dim i As Integer

        Dim curTop As Double
        Dim col1Left As Double
        Dim col2Left As Double
        Dim ColWidth As Double
        Dim maxHeight As Double
        Dim workHeight As Double

        Dim PrintingColumn1 As Boolean = True

        Dim LeftSide As Double

        workHeight = PageHeight - MarginBottom - FooterHeight - MarginTop
        curTop = MarginTop
        ColWidth = ((PageWidth - MarginRight - MarginLeft) / 2) - 1 / 32
        col1Left = MarginLeft
        col2Left = MarginLeft + ColWidth + 1 / 16

        maxHeight = workHeight

        For i = 1 To LastItemType
            If OverflowFrom(i) > -1 Then
                Do
                    If PrintingColumn1 Then LeftSide = col1Left Else LeftSide = col2Left

                    curTop = PrintOverflowTraits(i, LeftSide, curTop, ColWidth, maxHeight, OverflowFrom(i))
                    curTop = curTop + 1 / 16
                    maxHeight = workHeight - (curTop - MarginTop)

                    If OverflowFrom(i) > -1 Then
                        'still didn't finish
                        If Not PrintingColumn1 Then MyPrintDoc.NewPage()

                        PrintingColumn1 = Not PrintingColumn1
                        curTop = MarginTop
                        maxHeight = workHeight
                    Else
                        'finished
                        Exit Do
                    End If
                Loop
            End If
        Next

    End Sub
    Private Sub PrintPageOverflowWeapons()
        Dim i As Integer

        Dim curTop As Double
        Dim maxHeight As Double
        Dim workHeight As Double

        workHeight = PageHeight - MarginBottom - FooterHeight - MarginTop
        curTop = MarginTop

        maxHeight = workHeight

        For i = 1 To Ranged
            If WeaponOverflowFrom(i) > -1 Then
                Do
                    curTop = PrintOverflowWeapons(i, curTop, maxHeight, WeaponOverflowFrom(i))
                    curTop = curTop + 1 / 16
                    maxHeight = workHeight - (curTop - MarginTop)

                    If WeaponOverflowFrom(i) > -1 Then
                        'still didn't finish
                        MyPrintDoc.NewPage()
                        curTop = MarginTop
                        maxHeight = workHeight
                    Else
                        'finished
                        Exit Do
                    End If
                Loop
            End If
        Next

    End Sub
    Private Sub PrintPageGrimoire()
        Dim workHeight As Double

        workHeight = PageHeight - MarginBottom - FooterHeight - MarginTop

        GrimoireOverflowFrom = 0
        Do
            PrintOverflowGrimoire(MarginTop, workHeight, GrimoireOverflowFrom)

            If GrimoireOverflowFrom > -1 Then
                'still didn't finish
                MyPrintDoc.NewPage()
            Else
                'finished
                Exit Do
            End If
        Loop

    End Sub

    Private Function PrintSectionSpeedRangeTable(LeftSide As Double, TopSide As Double) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim curRow As Integer
        Dim maxHeight As Double
        Dim rowHeight As Double

        Dim BaseFont As New Font("Times New Roman", 10)
        Dim HeaderFont As New Font("Times New Roman", 10, FontStyle.Bold)
        Dim SubHeaderFont As New Font("Times New Roman", 9)
        Dim SubHeaderBoldFont As New Font("Times New Roman", 9, FontStyle.Bold)

        Dim TableCenterStyle As Style = MyPrintDoc.Style.Children.Add()
        TableCenterStyle.Font = BaseFont
        TableCenterStyle.TextAlignHorz = AlignHorzEnum.Center

        Dim TableRightStyle As Style = MyPrintDoc.Style.Children.Add()
        TableRightStyle.Font = BaseFont
        TableRightStyle.TextAlignHorz = AlignHorzEnum.Right

        Dim TableLeftStyle As Style = MyPrintDoc.Style.Children.Add()
        TableLeftStyle.Font = BaseFont
        TableLeftStyle.TextAlignHorz = AlignHorzEnum.Left

        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)
        Dim ldShade As New C1.C1Preview.LineDef(0.01, Color.LightGray, Drawing2D.DashStyle.Solid)
        Dim ldLine As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Dot)


        boxLeft = LeftSide
        boxWidth = 1.75
        boxTop = TopSide
        boxHeight = SpeedRangeTableHeight

        maxHeight = boxHeight
        rowHeight = GetHeight("X", HeaderFont)

        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, curTop, "SPEED/RANGE TABLE", boxWidth, HeaderFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        curTop = curTop + rowHeight
        maxHeight = boxHeight - (curTop - boxTop)

        MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, curTop, "For complete table, see p. 550.", boxWidth, SubHeaderFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        curTop = curTop + GetHeight("X", SubHeaderFont)
        maxHeight = boxHeight - (curTop - boxTop)

        curTop = curTop + rowHeight / 2
        maxHeight = boxHeight - (curTop - boxTop)


        Dim Mods() As Integer = {0, -1, -2, -3, -4, -5, -6, -7, -8, -9, -10, -11, -12, -13, -14, -15}
        Dim Vals() As Integer = {2, 3, 5, 7, 10, 15, 20, 30, 50, 70, 100, 150, 200, 300, 500, 700}
        Dim Shades() As Integer = {1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0, 1, 1, 0, 0}

        Dim Widths() As Double = {0.65, 5 / 8, boxWidth - 0.65 - 5 / 8}

        rowHeight = GetHeight("X", SubHeaderBoldFont)

        MyPrintDoc.RenderDirectText(boxLeft, curTop, "Speed/", Widths(0), SubHeaderBoldFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        MyPrintDoc.RenderDirectText(boxLeft + Widths(0), curTop, "Linear", boxWidth - Widths(0), SubHeaderBoldFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        curTop = curTop + rowHeight
        maxHeight = boxHeight - (curTop - boxTop)

        MyPrintDoc.RenderDirectText(boxLeft, curTop, "Range", Widths(0), SubHeaderBoldFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        MyPrintDoc.RenderDirectText(boxLeft + Widths(0), curTop, "Measurement", boxWidth - Widths(0), SubHeaderBoldFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        curTop = curTop + rowHeight
        maxHeight = boxHeight - (curTop - boxTop)

        MyPrintDoc.RenderDirectText(boxLeft, curTop, "Modifier", Widths(0), SubHeaderBoldFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        MyPrintDoc.RenderDirectText(boxLeft + Widths(0), curTop, "(range/speed)", boxWidth - Widths(0), SubHeaderBoldFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        curTop = curTop + rowHeight
        maxHeight = boxHeight - (curTop - boxTop)

        rowHeight = GetHeight("X", TableCenterStyle)
        For curRow = 0 To 15
            If Shades(curRow) Then
                MyPrintDoc.RenderDirectRectangle(boxLeft, curTop, boxWidth, rowHeight, ldShade, Color.LightGray)
            End If
            If LineBefore Then
                MyPrintDoc.RenderDirectLine(boxLeft, curTop, boxLeft + boxWidth, curTop, ldLine)
            End If

            If curRow = 0 Then
                MyPrintDoc.RenderDirectText(boxLeft + Widths(0) + Widths(1) + 1 / 32, curTop, "or less", Widths(2) - 1 / 32, rowHeight, TableLeftStyle)
            End If
            MyPrintDoc.RenderDirectText(boxLeft, curTop, Mods(curRow), Widths(0) / 2, rowHeight, TableRightStyle)
            MyPrintDoc.RenderDirectText(boxLeft + Widths(0), curTop, Vals(curRow) & " yd", Widths(1), rowHeight, TableRightStyle)

            curTop = curTop + rowHeight
            maxHeight = boxHeight - (curTop - boxTop)
        Next



        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionHitLocationTable(LeftSide As Double, TopSide As Double) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim curRow As Integer
        Dim maxHeight As Double
        Dim rowHeight As Double

        Dim BaseFont As New Font("Times New Roman", 10)
        Dim HeaderFont As New Font("Times New Roman", 10, FontStyle.Bold)
        Dim SubHeaderFont As New Font("Times New Roman", 9)
        Dim SubHeaderBoldFont As New Font("Times New Roman", 9, FontStyle.Bold)

        Dim TableCenterStyle As Style = MyPrintDoc.Style.Children.Add()
        TableCenterStyle.Font = BaseFont
        TableCenterStyle.TextAlignHorz = AlignHorzEnum.Center

        Dim TableRightStyle As Style = MyPrintDoc.Style.Children.Add()
        TableRightStyle.Font = BaseFont
        TableRightStyle.TextAlignHorz = AlignHorzEnum.Right

        Dim TableLeftStyle As Style = MyPrintDoc.Style.Children.Add()
        TableLeftStyle.Font = BaseFont
        TableLeftStyle.TextAlignHorz = AlignHorzEnum.Left

        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)
        Dim ldShade As New C1.C1Preview.LineDef(0.01, Color.LightGray, Drawing2D.DashStyle.Solid)
        Dim ldLine As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Dot)


        boxLeft = LeftSide
        boxWidth = 1 + 4 / 16
        boxTop = TopSide
        boxHeight = 2

        maxHeight = boxHeight
        rowHeight = GetHeight("X", HeaderFont)

        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, curTop, "HIT LOCATION", boxWidth, HeaderFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        curTop = curTop + rowHeight
        maxHeight = boxHeight - (curTop - boxTop)

        curTop = curTop + rowHeight / 2 - 1 / 32
        maxHeight = boxHeight - (curTop - boxTop)

        Dim Mods() As Integer = {0, -2, -3, -4, -5, -5, -7}
        Dim Vals() As String = {"Torso", "Arm/Leg", "Groin", "Hand", "Face", "Neck", "Skull"}
        Dim Shades() As Integer = {1, 1, 0, 0, 1, 1, 0}

        Dim Widths() As Double = {10 / 16, boxWidth - 10 / 16}

        rowHeight = GetHeight("X", SubHeaderBoldFont)

        MyPrintDoc.RenderDirectText(boxLeft, curTop, "Modifier", Widths(0), SubHeaderBoldFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        MyPrintDoc.RenderDirectText(boxLeft + Widths(0), curTop, "Location", boxWidth - Widths(0), SubHeaderBoldFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        curTop = curTop + rowHeight
        maxHeight = boxHeight - (curTop - boxTop)

        rowHeight = GetHeight("X", TableCenterStyle)
        For curRow = 0 To 6
            If Shades(curRow) Then
                MyPrintDoc.RenderDirectRectangle(boxLeft, curTop, boxWidth, rowHeight, ldShade, Color.LightGray)
            End If
            If LineBefore Then
                MyPrintDoc.RenderDirectLine(boxLeft, curTop, boxLeft + boxWidth, curTop, ldLine)
            End If

            MyPrintDoc.RenderDirectText(boxLeft, curTop, Mods(curRow), Widths(0) / 2, rowHeight, TableRightStyle)
            MyPrintDoc.RenderDirectText(boxLeft + Widths(0), curTop, Vals(curRow), Widths(1), rowHeight, TableLeftStyle)

            curTop = curTop + rowHeight
            maxHeight = boxHeight - (curTop - boxTop)
        Next

        curTop = curTop + rowHeight / 2
        maxHeight = boxHeight - (curTop - boxTop)

        Dim rtf As String
        rtf = ReturnRTFHeader() & "\fs18 {\i Imp} or {\i Pi} attacks can target vitals at -3 or eyes at -9." & "}"
        MyPrintDoc.RenderDirectRichText(boxLeft + 1 / 32, curTop, rtf, boxWidth - 2 / 32, maxHeight, TableLeftStyle)


        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionInfoBox(LeftSide As Double, TopSide As Double) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double

        Dim BaseFont As New Font("Times New Roman", 10)

        Dim TableCenterStyle As Style = MyPrintDoc.Style.Children.Add()
        TableCenterStyle.Font = BaseFont
        TableCenterStyle.TextAlignHorz = AlignHorzEnum.Center

        Dim TableRightStyle As Style = MyPrintDoc.Style.Children.Add()
        TableRightStyle.Font = BaseFont
        TableRightStyle.TextAlignHorz = AlignHorzEnum.Right

        Dim TableLeftStyle As Style = MyPrintDoc.Style.Children.Add()
        TableLeftStyle.Font = BaseFont
        TableLeftStyle.TextAlignHorz = AlignHorzEnum.Left

        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        boxLeft = LeftSide
        boxWidth = 1 + 4 / 16
        boxTop = TopSide
        boxHeight = 1.25 - 1 / 16

        Dim rtf As String
        rtf = ReturnRTFHeader()
        rtf = rtf & "\fs12 This sheet printed from {\b GURPS Character Assistant}. "
        rtf = rtf & "This and other {\b GURPS} forms may also be downloaded at {\b www.sjgames.com\\gurps\\resources}.\par "
        rtf = rtf & "Copyright © 2004 Steve Jackson Games Incorporated. All rights reserved."
        rtf = rtf & "}"

        MyPrintDoc.RenderDirectRichText(boxLeft + 1 / 32, boxTop + 1 / 16, rtf, boxWidth - 2 / 32, boxHeight - 1 / 8, TableCenterStyle)


        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionEquipment(LeftSide As Double, TopSide As Double) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim maxHeight As Double
        Dim rowHeight, tmpHeight As Double

        'Dim ld As New C1.C1Preview.LineDef(0.01, Color.Green, Drawing2D.DashStyle.Solid)
        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        '* Generate Equipment
        boxLeft = LeftSide
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = TopSide
        boxHeight = PageHeight - MarginBottom - boxTop - FooterHeight

        MaxWidth = boxWidth
        rowHeight = GetHeight("X", CType(MyStyles("SheetTextLeftStyle"), Style))
        tmpHeight = GetHeight("X", CType(MyStyles("UserTextLeftStyle"), Style))
        If tmpHeight > rowHeight Then rowHeight = tmpHeight
        
        maxHeight = boxHeight - rowHeight 'leave room at bottom for Totals
        
        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft, curTop, "ARMOR & POSSESSIONS", boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        curTop = curTop + 0.1875
        maxHeight = boxHeight - (curTop - boxTop) - rowHeight 'leave room at bottom for Totals
        
        '***** Equipment
        'set the column widths we'll need
        SetTraitColumns(Equipment, boxLeft)

        'set the list of items we'll be printing
        CreateFields(Equipment)

        'get the sizes of all our data lines
        SetFieldSizes()

        'print sub-heads
        curTop = PrintSubHeads(boxLeft, curTop)
        maxHeight = boxHeight - (curTop - boxTop) - rowHeight 'leave room at bottom for Totals

        'and now print the fields
        curTop = PrintFields(Equipment, boxLeft, curTop, maxHeight)
        ''maxHeight = boxHeight - (curTop - boxTop)

        'print the Totals
        Dim tmpStyle1, tmpStyle2, tmpStyle3 As Style
        tmpStyle1 = MyStyles("SheetTextRightBoldStyle")
        tmpStyle2 = MyStyles("UserTextLeftStyle")
        tmpStyle3 = MyStyles("UserTextRightStyle")

        MyPrintDoc.RenderDirectText(ColLefts(BrokenEffectCol - 1), boxTop + boxHeight - rowHeight, "Totals:", ColWidths(BrokenEffectCol - 1), rowHeight, tmpStyle1)
        MyPrintDoc.RenderDirectText(ColLefts(numCols - 2), boxTop + boxHeight - rowHeight, MyChar.Cost(Equipment), ColWidths(numCols - 2), rowHeight, tmpStyle2)
        MyPrintDoc.RenderDirectText(ColLefts(numCols - 1), boxTop + boxHeight - rowHeight, MyChar.TotalWeight, ColWidths(numCols - 1), rowHeight, tmpStyle3)

        'MyPrintDoc.RenderDirectText(ColLefts(BrokenEffectCol - 1), boxTop + boxHeight - rowHeight, "Totals:", ColWidths(BrokenEffectCol - 1), New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), Color.Black, AlignHorzEnum.Right)
        'MyPrintDoc.RenderDirectText(ColLefts(numCols - 2), boxTop + boxHeight - rowHeight, MyChar.Cost(Equipment), ColWidths(numCols - 2), New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize), Color.Blue, AlignHorzEnum.Left)
        'MyPrintDoc.RenderDirectText(ColLefts(numCols - 1), boxTop + boxHeight - rowHeight, MyChar.TotalWeight, ColWidths(numCols - 1), New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize), Color.Blue, AlignHorzEnum.Right)

        '*
        '* Draw the broken boxes around the info
        '*
        Dim mainboxWidth As Double = ColLefts(BrokenEffectCol) - boxLeft
        Dim subboxLeft As Double = ColLefts(BrokenEffectCol)
        Dim subboxRight As Double = ColLefts(numCols) + ColWidths(numCols)

        'draw the right box
        MyPrintDoc.RenderDirectRectangle(subboxLeft, boxTop, subboxRight - subboxLeft, boxHeight, ld)

        'clear the white area
        Dim ldWhite As New C1.C1Preview.LineDef(0.01, Color.White, Drawing2D.DashStyle.Solid)
        MyPrintDoc.RenderDirectRectangle(ColLefts(BrokenEffectCol), boxTop - 1 / 32, ColWidths(BrokenEffectCol), boxHeight + 1 / 16, ldWhite, Color.White)

        'draw the left box
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, mainboxWidth, boxHeight, ld)

        Return curTop
    End Function
    Private Function GetWeaponSectionDesiredHeight(WeaponType As Integer, LeftSide As Double, Optional ByVal MinimumAllowedHeight As Double = 0) As Double
        Dim boxLeft As Double
        Dim maxHeight As Double
        Dim FieldIndex As Integer
        Dim curField As clsDisplayField


        '* Generate Weapons
        boxLeft = LeftSide
        MaxWidth = PageWidth - MarginRight - boxLeft


        maxHeight = 0 'boxTop
        'MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, curTop, "HAND WEAPONS", boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        maxHeight = maxHeight + 0.1875

        'set the column widths we'll need
        SetWeaponColumns(WeaponType, boxLeft)

        'set the list of modes we'll be printing
        CreateWeaponFields(WeaponType)

        'get the sizes of all our data lines
        SetFieldSizes()

        'print sub-heads
        maxHeight = maxHeight + 0.1875 'PrintSubHeads(boxLeft, curTop)

        'and now print the fields
        For FieldIndex = 0 To Fields.Count - 1
            curField = Fields(FieldIndex)

            maxHeight = maxHeight + curField.Height
        Next
        '*****

        If MinimumAllowedHeight > maxHeight Then maxHeight = MinimumAllowedHeight

        Return maxHeight
    End Function
    Private Function PrintSectionHandWeapons(LeftSide As Double, TopSide As Double, UseThisHandBlockHeight As Double) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim maxHeight As Double

        'Dim ld As New C1.C1Preview.LineDef(0.01, Color.Green, Drawing2D.DashStyle.Solid)
        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        '* Generate Weapons
        boxLeft = LeftSide
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = TopSide
        'boxHeight = 1 + 1 / 16 + 1 / 8 'PageHeight - MarginBottom - boxTop - FooterHeight
        boxHeight = UseThisHandBlockHeight

        MaxWidth = boxWidth
        maxHeight = boxHeight

        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, curTop, "HAND WEAPONS", boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        curTop = curTop + 0.1875
        maxHeight = boxHeight - (curTop - boxTop)

        'set the column widths we'll need
        SetWeaponColumns(Hand, boxLeft)

        'set the list of modes we'll be printing
        CreateWeaponFields(Hand)

        'get the sizes of all our data lines
        SetFieldSizes()

        'print sub-heads
        curTop = PrintSubHeads(boxLeft, curTop)
        maxHeight = boxHeight - (curTop - boxTop)

        'and now print the fields
        curTop = PrintFields(Equipment, boxLeft, curTop, maxHeight, 0, True, Hand)
        maxHeight = boxHeight - (curTop - boxTop)

        '*
        '* Draw the broken boxes around the info
        '*
        Dim mainboxWidth As Double = ColLefts(BrokenEffectCol) - boxLeft
        Dim subboxLeft As Double = ColLefts(BrokenEffectCol)
        Dim subboxRight As Double = ColLefts(numCols) + ColWidths(numCols)

        'draw the right box
        MyPrintDoc.RenderDirectRectangle(subboxLeft, boxTop, subboxRight - subboxLeft, boxHeight, ld)

        'clear the white area
        Dim ldWhite As New C1.C1Preview.LineDef(0.01, Color.White, Drawing2D.DashStyle.Solid)
        MyPrintDoc.RenderDirectRectangle(ColLefts(BrokenEffectCol), boxTop - 1 / 32, ColWidths(BrokenEffectCol), boxHeight + 1 / 16, ldWhite, Color.White)

        'draw the left box
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, mainboxWidth, boxHeight, ld)

        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionRangedWeapons(LeftSide As Double, TopSide As Double, UseThisRangedBlockHeight As Double) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim maxHeight As Double

        'Dim ld As New C1.C1Preview.LineDef(0.01, Color.Green, Drawing2D.DashStyle.Solid)
        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        '* Generate Weapons
        boxLeft = LeftSide
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = TopSide
        'boxHeight = 2 + 1 / 16  'PageHeight - MarginBottom - boxTop - FooterHeight
        boxHeight = UseThisRangedBlockHeight

        MaxWidth = boxWidth
        maxHeight = boxHeight

        'MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld, Brushes.Orange)

        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, curTop, "RANGED WEAPONS", boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        curTop = curTop + 0.1875
        maxHeight = boxHeight - (curTop - boxTop)

        'set the column widths we'll need
        SetWeaponColumns(Ranged, boxLeft)

        'set the list of modes we'll be printing
        CreateWeaponFields(Ranged)

        'get the sizes of all our data lines
        SetFieldSizes()

        'print sub-heads
        curTop = PrintSubHeads(boxLeft, curTop)
        maxHeight = boxHeight - (curTop - boxTop)

        'and now print the fields
        curTop = PrintFields(Equipment, boxLeft, curTop, maxHeight, 0, True, Ranged)
        maxHeight = boxHeight - (curTop - boxTop)

        '*
        '* Draw the broken boxes around the info
        '*
        Dim mainboxWidth As Double = ColLefts(BrokenEffectCol) - boxLeft
        Dim subboxLeft As Double = ColLefts(BrokenEffectCol)
        Dim subboxRight As Double = ColLefts(numCols) + ColWidths(numCols)

        'draw the right box
        MyPrintDoc.RenderDirectRectangle(subboxLeft, boxTop, subboxRight - subboxLeft, boxHeight, ld)

        'clear the white area
        Dim ldWhite As New C1.C1Preview.LineDef(0.01, Color.White, Drawing2D.DashStyle.Solid)
        MyPrintDoc.RenderDirectRectangle(ColLefts(BrokenEffectCol), boxTop - 1 / 32, ColWidths(BrokenEffectCol), boxHeight + 1 / 16, ldWhite, Color.White)

        'draw the left box
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, mainboxWidth, boxHeight, ld)

        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionAttributes() As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double

        Dim stat As String = ""
        Dim i As Integer = 0
        Dim tmp As String = ""
        Dim ld As C1.C1Preview.LineDef

        Dim AttributePointsStyle As Style = MyStyles("AttributePointsStyle")
        Dim AttributeCenterStyle As Style = MyStyles("AttributeCenterStyle")
        Dim AttributeLeftStyle As Style = MyStyles("AttributeLeftStyle")
        Dim AttributeRightStyle As Style = MyStyles("AttributeRightStyle")
        Dim TinyTextStyle As Style = MyStyles("SheetTinyTextStyle")

        '*
        '* Print Attributes
        '*

        ld = New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)
        Dim Col1AttributeBoxWidth, Col2AttributeBoxWidth, Col1AttributeOffset, Col2AttributeOffset As Double

        boxWidth = 7.5 / 16 '0.25
        boxHeight = 7.5 / 16 + 1 / 32 '0.25
        Col1AttributeBoxWidth = 0.5
        Col2AttributeBoxWidth = 0.625
        Col1AttributeOffset = 1 / 16
        Col2AttributeOffset = 0.675

        '* ST line
        boxLeft = MarginLeft + Col1AttributeBoxWidth  '0.5
        boxTop = MarginTop + 14 / 16 '13 / 16 '2

        stat = "ST"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Items.Item(i).Score, boxWidth, boxHeight, AttributeCenterStyle)
        MyPrintDoc.RenderDirectText(boxLeft - Col1AttributeBoxWidth, boxTop, stat, Col1AttributeBoxWidth - Col1AttributeOffset, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "[", pointBoxWidth, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "]", pointBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, MyChar.Items.Item(i).Points, pointBoxWidth, boxHeight, AttributePointsStyle)

        boxLeft = boxLeft + boxWidth + 19 / 16

        stat = "Hit Points"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Items.Item(i).Score, boxWidth, boxHeight, AttributeCenterStyle)
        MyPrintDoc.RenderDirectText(boxLeft - Col2AttributeOffset, boxTop, "HP", Col2AttributeBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))

        boxLeft = boxLeft + boxWidth
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop - 1 / 8, "CURRENT", boxWidth, 1 / 8, TinyTextStyle)

        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "[", pointBoxWidth, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "]", pointBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, MyChar.Items.Item(i).Points, pointBoxWidth, boxHeight, AttributePointsStyle)

        '* DX line

        boxLeft = MarginLeft + Col1AttributeBoxWidth  '0.5
        boxTop = boxTop + boxHeight

        stat = "DX"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Items.Item(i).Score, boxWidth, boxHeight, AttributeCenterStyle)
        MyPrintDoc.RenderDirectText(boxLeft - Col1AttributeBoxWidth, boxTop, stat, Col1AttributeBoxWidth - Col1AttributeOffset, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "[", pointBoxWidth, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "]", pointBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, MyChar.Items.Item(i).Points, pointBoxWidth, boxHeight, AttributePointsStyle)

        boxLeft = boxLeft + boxWidth + 19 / 16

        stat = "Will"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Items.Item(i).Score, boxWidth, boxHeight, AttributeCenterStyle)
        MyPrintDoc.RenderDirectText(boxLeft - Col2AttributeOffset, boxTop, "Will", Col2AttributeBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))


        boxLeft = boxLeft + boxWidth

        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "[", pointBoxWidth, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "]", pointBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, MyChar.Items.Item(i).Points, pointBoxWidth, boxHeight, AttributePointsStyle)

        '* IQ line

        boxLeft = MarginLeft + Col1AttributeBoxWidth  '0.5
        boxTop = boxTop + boxHeight

        stat = "IQ"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Items.Item(i).Score, boxWidth, boxHeight, AttributeCenterStyle)
        MyPrintDoc.RenderDirectText(boxLeft - Col1AttributeBoxWidth, boxTop, stat, Col1AttributeBoxWidth - Col1AttributeOffset, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "[", pointBoxWidth, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "]", pointBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, MyChar.Items.Item(i).Points, pointBoxWidth, boxHeight, AttributePointsStyle)

        boxLeft = boxLeft + boxWidth + 19 / 16

        stat = "Perception"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Items.Item(i).Score, boxWidth, boxHeight, AttributeCenterStyle)
        MyPrintDoc.RenderDirectText(boxLeft - Col2AttributeOffset, boxTop, "Per", Col2AttributeBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))

        boxLeft = boxLeft + boxWidth

        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "[", pointBoxWidth, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "]", pointBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, MyChar.Items.Item(i).Points, pointBoxWidth, boxHeight, AttributePointsStyle)

        '* HT line

        boxLeft = MarginLeft + Col1AttributeBoxWidth  '0.5
        boxTop = boxTop + boxHeight

        stat = "HT"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Items.Item(i).Score, boxWidth, boxHeight, AttributeCenterStyle)
        MyPrintDoc.RenderDirectText(boxLeft - Col1AttributeBoxWidth, boxTop, stat, Col1AttributeBoxWidth - Col1AttributeOffset, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "[", pointBoxWidth, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "]", pointBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, MyChar.Items.Item(i).Points, pointBoxWidth, boxHeight, AttributePointsStyle)

        boxLeft = boxLeft + boxWidth + 19 / 16

        stat = "Fatigue Points"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Items.Item(i).Score, boxWidth, boxHeight, AttributeCenterStyle)
        MyPrintDoc.RenderDirectText(boxLeft - Col2AttributeOffset, boxTop, "FP", Col2AttributeBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))

        boxLeft = boxLeft + boxWidth
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop - 1 / 8, "CURRENT", boxWidth, 1 / 8, TinyTextStyle)

        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "[", pointBoxWidth, boxHeight, MyStyles("SheetAttributeLeftStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, "]", pointBoxWidth, boxHeight, MyStyles("SheetAttributeRightStyle"))
        MyPrintDoc.RenderDirectText(boxLeft + boxWidth + 1 / 16, boxTop, MyChar.Items.Item(i).Points, pointBoxWidth, boxHeight, AttributePointsStyle)



        boxTop = boxTop + boxHeight
        Return boxTop
    End Function
    Private Sub PrintSectionPage1Header()
        '*
        '* Print Graphic Header Block
        '*
        Dim boxLeft, boxTop As Double
        Dim stat As String = ""
        Dim i As Integer = 0
        Dim tmp As String = ""

        Dim ia As New C1.C1Preview.ImageAlign(C1.C1Preview.ImageAlignHorzEnum.Center, C1.C1Preview.ImageAlignVertEnum.Top, True, True, True, False, False)
        Dim bm As New Bitmap(EngineConfig.BaseSystemDataFolder & "graphics\gurps logo.bmp") ' GCA5.My.Resources.Resources.gca_g_icon_image_32x32)

        boxLeft = MarginLeft
        boxTop = MarginTop

        MyPrintDoc.RenderDirectImage(boxLeft, boxTop, bm, 1.75, 9 / 16, ia)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop + 9 / 16, "CHARACTER SHEET", 1.75, New Font(MyPrintDoc.Style.Font.Name, 12), MyOptions.Value("SheetTextColor"), C1.C1Preview.AlignHorzEnum.Center)

        '*
        '* Print Text stuff
        '*
        Dim blockLeft, blockWidth As Double
        blockLeft = boxLeft + 1.75 + 1 / 16
        blockWidth = PageWidth - MarginRight - blockLeft

        Dim rt As New RenderTable(3, 20)
        rt.Rows(0).Height = 0.75 / 4
        rt.Rows(1).Height = 0.75 / 4

        rt.Cells(0, 0).Text = "Name"
        rt.Cells(0, 2).Text = MyChar.Name
        rt.Cells(0, 9).Text = "Player"
        rt.Cells(0, 11).Text = MyChar.Player
        rt.Cells(0, 15).Text = "Point Total"
        rt.Cells(0, 18).Text = MyChar.TotalCost

        rt.Cells(0, 0).SpanCols = 2
        rt.Cells(0, 2).SpanCols = 7
        rt.Cells(0, 9).SpanCols = 2
        rt.Cells(0, 11).SpanCols = 4
        rt.Cells(0, 15).SpanCols = 3
        rt.Cells(0, 18).SpanCols = 2

        rt.Cells(0, 2).Style.Font = MyOptions.Value("UserFont")
        rt.Cells(0, 11).Style.Font = MyOptions.Value("UserFont")
        rt.Cells(0, 18).Style.Font = MyOptions.Value("UserFont")
        rt.Cells(0, 2).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cells(0, 11).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cells(0, 18).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cells(0, 15).Style.TextAlignHorz = AlignHorzEnum.Right
        rt.Cells(0, 18).Style.TextAlignHorz = AlignHorzEnum.Right

        rt.Cells(1, 0).Text = "Ht"
        rt.Cells(1, 1).Text = MyChar.Height
        rt.Cells(1, 3).Text = "Wt"
        rt.Cells(1, 4).Text = MyChar.Weight
        rt.Cells(1, 6).Text = "Size Modifier"
        stat = "Size Modifier"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        If i > 0 Then
            rt.Cells(1, 9).Text = MyChar.Items.Item(i).Score
        Else
            rt.Cells(1, 9).Text = "0"
        End If
        rt.Cells(1, 11).Text = "Age"
        rt.Cells(1, 12).Text = MyChar.Age
        rt.Cells(1, 15).Text = "Unspent Pts"
        rt.Cells(1, 18).Text = MyChar.UnspentPoints

        rt.Cells(1, 1).SpanCols = 2
        rt.Cells(1, 4).SpanCols = 2
        rt.Cells(1, 6).SpanCols = 3
        rt.Cells(1, 12).SpanCols = 3
        rt.Cells(1, 15).SpanCols = 3
        rt.Cells(1, 18).SpanCols = 2

        rt.Cells(1, 1).Style.Font = MyOptions.Value("UserFont")
        rt.Cells(1, 4).Style.Font = MyOptions.Value("UserFont")
        rt.Cells(1, 9).Style.Font = MyOptions.Value("UserFont")
        rt.Cells(1, 12).Style.Font = MyOptions.Value("UserFont")
        rt.Cells(1, 18).Style.Font = MyOptions.Value("UserFont")
        rt.Cells(1, 1).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cells(1, 4).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cells(1, 9).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cells(1, 12).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cells(1, 18).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cells(1, 15).Style.TextAlignHorz = AlignHorzEnum.Right
        rt.Cells(1, 18).Style.TextAlignHorz = AlignHorzEnum.Right

        rt.Cells(2, 0).Text = "Appearance"
        rt.Cells(2, 3).Text = MyChar.Appearance
        rt.Cells(2, 3).Style.Font = MyOptions.Value("UserFont")
        rt.Cells(2, 3).Style.TextColor = MyOptions.Value("UserTextColor")

        rt.Cells(2, 0).SpanCols = 3
        rt.Cells(2, 3).SpanCols = 17

        'rt.Style.Borders.All = New C1.C1Preview.LineDef("0.2mm", Color.Green)

        MyPrintDoc.RenderDirect(blockLeft, boxTop, rt, blockWidth, 3 / 4) '9 / 16)

    End Sub
    Private Sub PrintSectionPage2Header(MaxHeight As Double)
        '*
        '* Print Graphic Header Block
        '*
        Dim boxLeft, boxTop As Double
        Dim stat As String = ""
        Dim i As Integer = 0
        Dim tmp As String = ""

        Dim ia As New C1.C1Preview.ImageAlign(C1.C1Preview.ImageAlignHorzEnum.Center, C1.C1Preview.ImageAlignVertEnum.Top, True, True, True, False, False)
        Dim bm As New Bitmap(EngineConfig.BaseSystemDataFolder & "graphics\gurps logo.bmp")

        boxLeft = MarginLeft
        boxTop = MarginTop

        MyPrintDoc.RenderDirectImage(boxLeft, boxTop, bm, 1.625, 0.5, ia)
        boxTop = boxTop + 0.5
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, "CHARACTER SHEET", 1.625, New Font(MyPrintDoc.Style.Font.Name, 12), MyOptions.Value("SheetTextColor"), C1.C1Preview.AlignHorzEnum.Center)

        boxTop = boxTop + 0.1875
        'MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Name, 1.625, 1.0625, MyStyles("UserTextCenterStyle"))
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Name, 1.625, (MaxHeight+Margintop - boxTop), MyStyles("UserTextCenterStyle"))
    End Sub
    Private Function PrintSectionBasics(TopSide As Double) As Double
        '*
        '* Print Basics
        '*

        '* Life and Damage
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double

        Dim rt As RenderTable
        Dim curitem As GCATrait

        boxLeft = MarginLeft
        boxWidth = 3 + 9 / 16
        boxTop = TopSide
        boxHeight = 3 / 16 '5 / 32 '

        rt = New RenderTable(1, 8)
        rt.Width = Unit.Auto
        rt.Cols(0).CellStyle.Padding.Left = 1 / 32

        rt.Cols(1).Width = 1 / 2

        rt.Cols(2).Width = 5 / 16
        rt.Cols(4).Width = 1 / 4
        rt.Cols(6).Width = 1 / 4

        rt.Cols(5).Width = 5 / 16
        rt.Cols(7).Width = 5 / 16

        rt.Cols(2).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(5).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(7).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(2).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(5).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(7).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(2).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(4).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(5).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(7).Style.TextAlignHorz = AlignHorzEnum.Center

        rt.Cells(0, 0).Text = "BASIC LIFT"
        rt.Cells(0, 1).Text = "(ST×ST)/5"
        rt.Cells(0, 1).Style.FontSize = MyStyles("SheetTinyTextStyle").fontsize
        rt.Cells(0, 1).Style.TextAlignVert = AlignVertEnum.Center

        curitem = MyChar.ItemByNameAndExt("Basic Lift")
        rt.Cells(0, 2).Text = curitem.Score

        rt.Cells(0, 3).Text = "DAMAGE"
        rt.Cols(3).CellStyle.Padding.Right = 1 / 32
        rt.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Right

        rt.Cells(0, 4).Text = "Thr"
        rt.Cells(0, 5).Text = MyChar.BaseTH

        rt.Cells(0, 6).Text = "Sw"
        rt.Cells(0, 7).Text = MyChar.BaseSW

        'render the table
        'rt.Style.GridLines.All = New LineDef(0.01, Color.Green, System.Drawing.Drawing2D.DashStyle.Dot)
        rt.Style.Borders.Top = New LineDef(0.01, MyOptions.Value("SheetTextColor"))
        MyPrintDoc.RenderDirect(boxLeft, boxTop, rt, boxWidth, boxHeight)


        '* Speed and Move
        boxTop = boxTop + boxHeight

        rt = New RenderTable(1, 10)
        rt.Width = Unit.Auto
        rt.Cols(0).CellStyle.Padding.Left = 1 / 32
        rt.Cols(5).CellStyle.Padding.Left = 1 / 32

        rt.Cols(1).Width = 3 / 8
        rt.Cols(6).Width = 3 / 8

        rt.Cols(2).Width = 1 / 16
        rt.Cols(4).Width = 1 / 16
        rt.Cols(7).Width = 1 / 16
        rt.Cols(9).Width = 1 / 16

        rt.Cols(3).Width = 1 / 4
        rt.Cols(8).Width = 1 / 4

        rt.Cols(1).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(3).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(6).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(8).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(1).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(3).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(6).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(8).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(2).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(4).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(6).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(7).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(8).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(9).Style.TextAlignHorz = AlignHorzEnum.Center

        'rt.Cols(5).Style.TextAlignHorz = AlignHorzEnum.Right


        curitem = MyChar.ItemByNameAndExt("Basic Speed")
        rt.Cells(0, 0).Text = "BASIC SPEED"
        rt.Cells(0, 1).Text = curitem.Score
        rt.Cells(0, 2).Text = "["
        rt.Cells(0, 3).Text = curitem.Points
        rt.Cells(0, 4).Text = "]"

        curitem = MyChar.ItemByNameAndExt("Basic Move")
        rt.Cells(0, 5).Text = "BASIC MOVE"
        rt.Cells(0, 6).Text = curitem.Score
        rt.Cells(0, 7).Text = "["
        rt.Cells(0, 8).Text = curitem.Points
        rt.Cells(0, 9).Text = "]"


        'render the table
        'rt.Style.GridLines.All = New LineDef(0.01, Color.Green, System.Drawing.Drawing2D.DashStyle.Dot)
        'rt.Style.Borders.All = New LineDef(0.01, Color.Green)
        MyPrintDoc.RenderDirect(boxLeft, boxTop, rt, boxWidth, boxHeight)

        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionMove(TopSide As Double) As Double
        '*
        '* Print Move Block
        '*
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim stat As String = ""
        Dim i As Integer = 0

        Dim curitem As GCATrait

        boxLeft = MarginLeft
        boxWidth = 3 + 9 / 16
        boxTop = TopSide
        boxHeight = 1 + 1 / 8


        Dim rt As New RenderTable(6, 6)
        rt.Width = Unit.Auto
        'rt.Style.TextColor = Color.Black
        rt.Cols(0).Width = 1.375
        rt.Cols(1).Width = 0.375
        rt.Cols(3).Width = 3 / 16
        rt.Cols(5).Width = 3 / 16

        rt.Cells(0, 0).Text = "ENCUMBRANCE"
        rt.Cells(0, 2).Text = "MOVE"
        rt.Cells(0, 4).Text = "DODGE"

        rt.Cells(0, 0).SpanCols = 2
        rt.Cells(0, 2).SpanCols = 2
        rt.Cells(0, 4).SpanCols = 2

        rt.Rows(0).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Rows(0).Style.FontBold = True

        rt.Cols(0).CellStyle.Padding.Left = 1 / 32
        rt.Cols(2).CellStyle.Padding.Left = 1 / 16
        rt.Cols(4).CellStyle.Padding.Left = 1 / 16
        rt.Cols(2).Style.Borders.Left = New LineDef(0.01, MyOptions.Value("SheetTextColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rt.Cols(4).Style.Borders.Left = New LineDef(0.01, MyOptions.Value("SheetTextColor"), System.Drawing.Drawing2D.DashStyle.Dot)

        rt.Cols(1).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(3).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(5).Style.Font = MyOptions.Value("UserFont")
        rt.Cols(1).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(3).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(5).Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Cols(1).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(5).Style.TextAlignHorz = AlignHorzEnum.Center

        If MyChar.EncumbranceLevel < 5 Then
            rt.Rows(MyChar.EncumbranceLevel + 1).Style.BackColor = Color.LightGray
        End If
        rt.Cells(1, 0).Text = "None (0) = BL"
        rt.Cells(2, 0).Text = "Light (1) = 2 × BL"
        rt.Cells(3, 0).Text = "Medium (2) = 3 × BL"
        rt.Cells(4, 0).Text = "Heavy (3) = 6 × BL"
        rt.Cells(5, 0).Text = "X-Heavy (4) = 10 × BL"

        curitem = MyChar.ItemByNameAndExt("No Encumbrance")
        rt.Cells(1, 1).Text = curitem.Score
        curitem = MyChar.ItemByNameAndExt("Light Encumbrance")
        rt.Cells(2, 1).Text = curitem.Score
        curitem = MyChar.ItemByNameAndExt("Medium Encumbrance")
        rt.Cells(3, 1).Text = curitem.Score
        curitem = MyChar.ItemByNameAndExt("Heavy Encumbrance")
        rt.Cells(4, 1).Text = curitem.Score
        curitem = MyChar.ItemByNameAndExt("X-Heavy Encumbrance")
        rt.Cells(5, 1).Text = curitem.Score

        rt.Cells(1, 2).Text = "BM × 1"
        rt.Cells(2, 2).Text = "BM × 0.8"
        rt.Cells(3, 2).Text = "BM × 0.6"
        rt.Cells(4, 2).Text = "BM × 0.4"
        rt.Cells(5, 2).Text = "BM × 0.2"

        curitem = MyChar.ItemByNameAndExt("No Encumbrance Move")
        rt.Cells(1, 3).Text = curitem.Score
        curitem = MyChar.ItemByNameAndExt("Light Encumbrance Move")
        rt.Cells(2, 3).Text = curitem.Score
        curitem = MyChar.ItemByNameAndExt("Medium Encumbrance Move")
        rt.Cells(3, 3).Text = curitem.Score
        curitem = MyChar.ItemByNameAndExt("Heavy Encumbrance Move")
        rt.Cells(4, 3).Text = curitem.Score
        curitem = MyChar.ItemByNameAndExt("X-Heavy Encumbrance Move")
        rt.Cells(5, 3).Text = curitem.Score

        rt.Cells(1, 4).Text = "Dodge"
        rt.Cells(2, 4).Text = "Dodge -1"
        rt.Cells(3, 4).Text = "Dodge -2"
        rt.Cells(4, 4).Text = "Dodge -3"
        rt.Cells(5, 4).Text = "Dodge -4"

        curitem = MyChar.ItemByNameAndExt("Dodge")
        rt.Cells(1, 5).Text = MaxInt(curitem.Score, 0)
        rt.Cells(2, 5).Text = MaxInt(curitem.Score - 1, 0)
        rt.Cells(3, 5).Text = MaxInt(curitem.Score - 2, 0)
        rt.Cells(4, 5).Text = MaxInt(curitem.Score - 3, 0)
        rt.Cells(5, 5).Text = MaxInt(curitem.Score - 4, 0)


        'render the table
        rt.Style.GridLines.Horz = New LineDef(0.01, MyOptions.Value("SheetTextColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rt.Style.Borders.All = New LineDef(0.01, MyOptions.Value("SheetTextColor"))
        MyPrintDoc.RenderDirect(boxLeft, boxTop, rt, boxWidth, boxHeight)

        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionCharacterNotes(TopSide As Double, MaxHeight As Double) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim inboxWidth, inboxHeight, inboxLeft, inboxTop As Double
        Dim rowHeight As Double

        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        Dim rtf As String

        boxLeft = MarginLeft
        boxWidth = 3 + 1 / 16
        boxHeight = MaxHeight
        boxTop = TopSide

        Dim tmpFont As New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold)
        MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, boxTop, "CHARACTER NOTES", boxWidth, tmpFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        rowHeight = GetHeight("X", tmpFont)

        inboxLeft = boxLeft + 1 / 32
        inboxWidth = boxWidth - 2 / 32
        inboxHeight = boxHeight - rowHeight - 2 / 32
        inboxTop = boxTop + rowHeight

        rtf = MyChar.Notes
        MyPrintDoc.RenderDirectRichText(inboxLeft, inboxTop, rtf, inboxWidth, inboxHeight, MyPrintDoc.Style)

        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        Return 0
    End Function
    Private Function PrintSectionPointsSummary() As Double
        Dim i As Integer
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim rowHeight As Double = GetHeight("X")

        boxLeft = MarginLeft
        boxWidth = 3 + 1 / 16
        boxHeight = PointsBoxHeight
        boxTop = PageHeight - MarginBottom - FooterHeight - boxHeight

        Dim rt As New RenderTable(6, 4)

        rt.Rows(0).Style.FontBold = True
        rt.Rows(2).Height = rowHeight * 2

        rt.Cols(0).CellStyle.Padding.Left = 1 / 32
        rt.Cols(2).Style.TextAlignHorz = AlignHorzEnum.Center
        rt.Cols(2).Style.TextColor = MyOptions.Value("UserTextColor")

        rt.Cols(1).Width = 1 / 16
        rt.Cols(2).Width = 1 / 4
        rt.Cols(3).Width = 1 / 16

        rt.Cells(0, 0).Text = "POINTS SUMMARY"
        rt.Cells(1, 0).Text = "Attributes/Secondary Characteristics"
        rt.Cells(2, 0).Text = "Advantages/Perks/TL/Languages/Cultural Familiarity"
        rt.Cells(3, 0).Text = "Disadvantages/Quirks"
        rt.Cells(4, 0).Text = "Skills/Techniques/Spells"
        rt.Cells(5, 0).Text = "Other"

        For i = 1 To 5
            rt.Cells(i, 1).Text = "["
            rt.Cells(i, 3).Text = "]"
        Next

        rt.Cells(1, 2).Text = MyChar.Cost(Stats)
        rt.Cells(2, 2).Text = MyChar.Cost(Ads) + MyChar.Cost(Perks) + MyChar.Cost(Templates)
        rt.Cells(3, 2).Text = MyChar.Cost(Disads) + MyChar.Cost(Quirks)
        rt.Cells(4, 2).Text = MyChar.Cost(Skills) + MyChar.Cost(Spells)
        rt.Cells(5, 2).Text = ""

        'render the table
        rt.Style.GridLines.Horz = New LineDef(0.01, MyOptions.Value("SheetTextColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rt.Style.Borders.All = New LineDef(0.01, MyOptions.Value("SheetTextColor"))
        MyPrintDoc.RenderDirect(boxLeft, boxTop, rt, boxWidth, boxHeight)

        Return boxTop
    End Function
    Private Function PrintSectionLanguages(LeftSide As Double) As Double
        '*
        '* Print Languages
        '*
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim stat As String = ""
        Dim i As Integer = 0

        Dim curitem As GCATrait

        boxLeft = LeftSide
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = MarginTop + 14 / 16 - 1 / 8
        boxHeight = 1 + 1 / 8


        Dim rt As New RenderTable(6, 6)
        'rt.Width = Unit.Auto 
        'rt.Height = boxHeight
        rt.Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Style.Font = MyOptions.Value("UserFont")

        rt.Cols(0).Width = boxWidth / 3 + 1 / 2
        rt.Cols(1).Width = boxWidth / 3 - 7 / 16
        rt.Cols(2).Width = boxWidth / 3 - 7 / 16
        rt.Cols(3).Width = 1 / 16
        rt.Cols(4).Width = 1 / 4
        rt.Cols(5).Width = 1 / 16

        rt.Cells(0, 0).Text = "Languages"
        rt.Cells(0, 1).Text = "Spoken"
        rt.Cells(0, 2).Text = "Written"

        rt.Cols(0).CellStyle.Padding.Left = 1 / 32
        rt.Cols(4).Style.TextAlignHorz = AlignHorzEnum.Center

        rt.Rows(0).Style.FontBold = True
        rt.Rows(0).Style.Font = MyOptions.Value("SheetFont")
        rt.Cols(3).Style.Font = MyOptions.Value("SheetFont")
        rt.Cols(5).Style.Font = MyOptions.Value("SheetFont")
        rt.Rows(0).Style.TextColor = MyOptions.Value("SheetTextColor")
        rt.Cols(3).Style.TextColor = MyOptions.Value("SheetTextColor")
        rt.Cols(5).Style.TextColor = MyOptions.Value("SheetTextColor")

        For i = 1 To 5
            rt.Cells(i, 3).Text = "["
            rt.Cells(i, 5).Text = "]"
        Next

        Dim upto As Integer = MyChar.ItemsByType(Languages).count
        If upto > 5 Then
            upto = 5
            OverflowFrom(Languages) = upto
            HasOverflow = True
        End If
        For i = 1 To upto
            curitem = MyChar.ItemsByType(Languages).Item(i)
            rt.Cells(i, 0).Text = curitem.FullNameTL
            If curitem.NameExt.ToLower = "spoken" Then
                rt.Cells(i, 1).Text = curitem.LevelName
            ElseIf curitem.NameExt.ToLower = "written" Then
                rt.Cells(i, 2).Text = curitem.LevelName
            Else
                rt.Cells(i, 1).Text = curitem.LevelName
                rt.Cells(i, 2).Text = curitem.LevelName
            End If
            rt.Cells(i, 4).Text = curitem.Points
        Next

        'render the table
        'rt.Style.GridLines.Horz = New LineDef(0.01, Color.Black, System.Drawing.Drawing2D.DashStyle.Dot)
        'rt.Style.Borders.All = New LineDef(0.01, Color.Black)
        rt.Style.GridLines.Horz = New LineDef(0.01, MyOptions.Value("SheetTextColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rt.Style.Borders.All = New LineDef(0.01, MyOptions.Value("SheetTextColor"))
        MyPrintDoc.RenderDirect(boxLeft, boxTop, rt, boxWidth, boxHeight)


        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionDR(LeftSide As Double, TopSide As Double) As Double
        '*
        '* Print DR
        '*
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim stat As String = ""
        Dim i As Integer = 0

        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        boxLeft = LeftSide
        boxWidth = 11 / 16
        boxTop = TopSide
        boxHeight = 15 / 16

        stat = "DR"
        i = MyChar.ItemPositionByNameAndExt(stat, Stats)
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.Items.Item(i).Score, boxWidth, boxHeight, MyStyles("AttributeCenterStyle"))
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, stat, boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)

        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionCultures(LeftSide As Double, TopSide As Double) As Double
        '*
        '* Print TL & Cultures
        '*
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim stat As String = ""
        Dim i As Integer = 0

        Dim curitem As GCATrait

        boxLeft = LeftSide
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = TopSide
        boxHeight = 15 / 16


        Dim rt As New RenderTable(5, 5)
        rt.Style.TextColor = MyOptions.Value("UserTextColor")
        rt.Style.Font = MyOptions.Value("UserFont")

        rt.Cols(0).Width = 5 / 16
        rt.Cols(1).Width = boxWidth - 11 / 16
        rt.Cols(2).Width = 1 / 16
        rt.Cols(3).Width = 4 / 16
        rt.Cols(4).Width = 1 / 16

        For i = 1 To 4
            rt.Cells(i, 0).SpanCols = 2
        Next

        i = MyChar.ItemPositionByNameAndExt("Tech Level")
        curitem = MyChar.Items.Item(i)
        rt.Cells(0, 0).Text = "TL:"
        rt.Cells(0, 1).Text = curitem.Score
        rt.Cells(0, 2).Text = "["
        rt.Cells(0, 3).Text = curitem.Points
        rt.Cells(0, 4).Text = "]"

        rt.Cells(0, 0).Style.FontBold = True
        rt.Cells(0, 0).Style.Font = MyOptions.Value("SheetFont")
        rt.Cells(0, 0).Style.TextColor = MyOptions.Value("SheetTextColor")
        rt.Cells(0, 2).Style.Font = MyOptions.Value("SheetFont")
        rt.Cells(0, 2).Style.TextColor = MyOptions.Value("SheetTextColor")
        rt.Cells(0, 4).Style.Font = MyOptions.Value("SheetFont")
        rt.Cells(0, 4).Style.TextColor = MyOptions.Value("SheetTextColor")


        rt.Cells(1, 0).Text = "Cultural Familiarities"

        rt.Cols(0).CellStyle.Padding.Left = 1 / 32
        rt.Cols(3).Style.TextAlignHorz = AlignHorzEnum.Center

        rt.Rows(1).Style.FontBold = True
        rt.Rows(1).Style.Font = MyOptions.Value("SheetFont")
        rt.Rows(1).Style.TextColor = MyOptions.Value("SheetTextColor")

        rt.Cols(2).Style.Font = MyOptions.Value("SheetFont")
        rt.Cols(4).Style.Font = MyOptions.Value("SheetFont")
        rt.Cols(2).Style.TextColor = MyOptions.Value("SheetTextColor")
        rt.Cols(4).Style.TextColor = MyOptions.Value("SheetTextColor")


        For i = 2 To 4
            rt.Cells(i, 2).Text = "["
            rt.Cells(i, 4).Text = "]"
        Next

        Dim upto As Integer = MyChar.ItemsByType(Cultures).count
        If upto > 3 Then
            upto = 3
            OverflowFrom(Cultures) = upto
            HasOverflow = True
        End If
        For i = 1 To upto
            curitem = MyChar.ItemsByType(Cultures).Item(i)
            rt.Cells(1 + i, 0).Text = curitem.FullNameTL
            rt.Cells(1 + i, 3).Text = curitem.Points
        Next

        'render the table
        'rt.Style.GridLines.Horz = New LineDef(0.01, Color.Black, System.Drawing.Drawing2D.DashStyle.Dot)
        'rt.Style.Borders.All = New LineDef(0.01, Color.Black)
        rt.Style.GridLines.Horz = New LineDef(0.01, MyOptions.Value("SheetTextColor"), System.Drawing.Drawing2D.DashStyle.Dot)
        rt.Style.Borders.All = New LineDef(0.01, MyOptions.Value("SheetTextColor"))
        MyPrintDoc.RenderDirect(boxLeft, boxTop, rt, boxWidth, boxHeight)

        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionReactions(LeftSide As Double, TopSide As Double) As Double
        '*
        '* Print Reaction Modifiers
        '*
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim intBoxWidth, intBoxHeight, intBoxLeft, intBoxTop As Double
        Dim stat As String
        Dim tmpLine As String
        Dim curVal As Integer
        Dim curItem As GCATrait

        'Dim ld As New C1.C1Preview.LineDef(0.01, Color.Black, Drawing2D.DashStyle.Solid)
        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        boxLeft = LeftSide
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = TopSide
        boxHeight = 50 / 32

        intBoxLeft = boxLeft + 1 / 32
        intBoxWidth = boxWidth - 2 / 32
        intBoxTop = boxTop + 1 / 32
        intBoxHeight = boxHeight - 2 / 32


        Dim rtf As String = ""
        Dim output As String = ""
        Dim tmp As String = ""
        Dim tmpBo1 As String = ""
        Dim tmpBo2 As String = ""

        output = output & "\qc\b Reaction Modifiers \par\ql "

        'Appearance
        stat = "Unappealing"
        curItem = MyChar.ItemByNameAndExt(stat, Stats)

        tmpLine = ""
        If curItem.TagItem("bonuslist") <> "" Then
            curVal = curItem.TagItem("syslevels")
            If curVal >= 0 Then
                tmpLine = tmpLine & "+" & curVal
            Else
                tmpLine = tmpLine & curVal
            End If

            tmp = curItem.TagItem("bonuslist")
            If tmp <> "" Then
                tmpBo1 = "{\i Unappealing Includes: }" & tmp & ". "
            Else
                tmpBo1 = ""
            End If

            tmp = curItem.TagItem("conditionallist")
            If tmp <> "" Then
                If tmpBo1 <> "" Then
                    tmpBo1 = tmpBo1 & "{\i Conditional: }" & tmp & ". "
                Else
                    tmpBo1 = "{\i Unappealing Conditional: }" & tmp & ". "
                End If
            End If
        End If

        stat = "Appealing"
        curItem = MyChar.ItemByNameAndExt(stat, Stats)
        If curItem.TagItem("bonuslist") <> "" Then
            curVal = curItem.TagItem("syslevels")
            If curVal >= 0 Then
                tmpLine = tmpLine & "/+" & curVal
            Else
                tmpLine = tmpLine & "/" & curVal
            End If

            tmp = curItem.TagItem("bonuslist")
            If tmp <> "" Then
                tmpBo2 = "{\i Appealing Includes: }" & tmp & ". "
            Else
                tmpBo2 = ""
            End If

            tmp = curItem.TagItem("conditionallist")
            If tmp <> "" Then
                If tmpBo2 <> "" Then
                    tmpBo2 = tmpBo2 & "{\i Conditional: }" & tmp & ". "
                Else
                    tmpBo2 = "{\i Appealing Conditional: }" & tmp & ". "
                End If
            End If
        End If

        tmp = ""
        If tmpBo1 <> "" Then
            tmp = tmp & tmpBo1
        End If
        If tmpBo2 <> "" Then
            tmp = tmp & tmpBo2
        End If

        output = output & "\b Appearance: \b0\cf2\f1 " & tmpLine & "\cf0\f0"
        If tmp <> "" Then
            output = output & " \line "
            output = output & "{\fs12 " & tmp & "}" & " \line " '" \par "
        Else
            output = output & " \line " '" \par "
        End If

        'Status
        stat = "Status"
        curItem = MyChar.ItemByNameAndExt(stat, Stats)
        If Val(curItem.TagItem("score")) >= 0 Then
            tmpLine = "+" & curItem.TagItem("score")
        Else
            tmpLine = curItem.TagItem("score")
        End If

        tmp = curItem.TagItem("bonuslist")
        If tmp <> "" Then
            tmpBo1 = "{\i Includes: }" & tmp & ". "
        Else
            tmpBo1 = ""
        End If

        tmp = curItem.TagItem("conditionallist")
        If tmp <> "" Then
            If tmpBo1 <> "" Then
                tmpBo1 = tmpBo1 & "{\i Conditional: }" & tmp & ". "
            Else
                tmpBo1 = "{\i Conditional: }" & tmp & ". "
            End If
        End If

        output = output & "\b Status: \b0\cf2\f1 " & tmpLine & "\cf0\f0"
        If tmpBo1 <> "" Then
            output = output & " \line "
            output = output & "{\fs12 " & tmpBo1 & "}" & " \line " '" \par "
        Else
            output = output & " \line " '" \par "
        End If


        'Other
        stat = "Reaction"
        curItem = MyChar.ItemByNameAndExt(stat, Stats)
        If Val(curItem.TagItem("score")) >= 0 Then
            tmpLine = "+" & curItem.TagItem("score")
        Else
            tmpLine = curItem.TagItem("score")
        End If

        tmp = curItem.TagItem("bonuslist")
        If tmp <> "" Then
            tmpBo1 = "{\i Includes: }" & tmp & ". "
        Else
            tmpBo1 = ""
        End If

        tmp = curItem.TagItem("conditionallist")
        If tmp <> "" Then
            If tmpBo1 <> "" Then
                tmpBo1 = tmpBo1 & "{\i Conditional: }" & tmp & ". "
            Else
                tmpBo1 = "{\i Conditional: }" & tmp & ". "
            End If
        End If

        output = output & "\b Other: \b0\cf2\f1 " & tmpLine & "\cf0\f0"
        If tmpBo1 <> "" Then
            output = output & " \line "
            output = output & "{\fs12 " & tmpBo1 & "}" & " \line " '" \par "
        Else
            output = output & " \line " '" \par "
        End If


        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        'rtf = ReturnRTFHeader() & output & "}"
        rtf = ReturnRTFHeaderUserOptions() & output & "}"
        MyPrintDoc.RenderDirectRichText(intBoxLeft, intBoxTop, rtf, intBoxWidth, intBoxHeight, MyPrintDoc.Style)


        Return boxTop + boxHeight
    End Function
    Private Function PrintSectionParryBlock(LeftSide As Double, TopSide As Double) As Double
        '*
        '* Print Parry
        '*
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim stat As String = ""
        Dim i As Integer = 0

        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        boxLeft = LeftSide
        boxWidth = 11 / 16
        boxTop = TopSide
        boxHeight = 25 / 32

        stat = "Parry"
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.ParryScore, boxWidth, boxHeight, MyStyles("AttributeCenterStyle"))
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, stat, boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.ParryUsing, boxWidth, boxHeight, MyStyles("SheetTinyTextAtBottomStyle"))

        '*
        '* Print Block
        '*
        boxTop = boxTop + boxHeight

        stat = "Block"
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.BlockScore, boxWidth, boxHeight, MyStyles("AttributeCenterStyle"))
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, stat, boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        MyPrintDoc.RenderDirectText(boxLeft, boxTop, MyChar.BlockUsing, boxWidth, boxHeight, MyStyles("SheetTinyTextAtBottomStyle"))

        Return boxTop + boxHeight
    End Function

    Private Function PrintOverflowTraits(ItemType As Integer, LeftSide As Double, TopSide As Double, ColWidth As Double, MaxHeight As Double, StartFrom As Integer) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim Title As String = ReturnLongListName(ItemType).ToUpper
        If ItemType = Equipment Then
            Title = "Armor & Possessions".ToUpper
        End If

        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        '* Generate traits
        boxLeft = LeftSide
        boxWidth = ColWidth
        boxTop = TopSide
        boxHeight = MaxHeight

        MaxWidth = boxWidth

        SetTraitColumns(ItemType, boxLeft)
        CreateFields(ItemType)
        SetFieldSizes()

        'It is possible to get to this point without actually having a true Overflow,
        'such as when printing Advantages and running out of room before getting to the
        'Perks at all. In which case, Perks Overflow will be at 0, even though there may
        'not be any valid Perks to print.
        '
        'So: make sure, before printing anything.
        '
        'We have a field count now, so we can verify that we have something
        'to print before doing so.

        If Fields.Count <= StartFrom Then
            'Nothing to print
            OverflowFrom(ItemType) = -1
            Return boxTop - 1 / 16
        End If

        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft, curTop, Title, boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        curTop = curTop + 0.1875
        MaxHeight = boxHeight - (curTop - boxTop)


        'print sub-heads
        Select Case ItemType
            Case Ads, Disads, Quirks, Perks, Cultures, Languages
            Case Else
                curTop = PrintSubHeads(boxLeft, curTop)
                MaxHeight = boxHeight - (curTop - boxTop)
        End Select

        'and now print the fields
        curTop = PrintFields(ItemType, boxLeft, curTop, MaxHeight, StartFrom)


        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, curTop - boxTop, ld)

        Return curTop
    End Function
    Private Function PrintOverflowWeapons(WeaponType As Integer, TopSide As Double, MaxHeight As Double, StartFrom As Integer) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim Title As String

        'Dim ld As New C1.C1Preview.LineDef(0.01, Color.Green, Drawing2D.DashStyle.Solid)
        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        '* Generate Weapons
        boxLeft = MarginLeft
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = TopSide
        boxHeight = MaxHeight

        MaxWidth = boxWidth

        Select Case WeaponType
            Case 1
                Title = "HAND WEAPONS"
            Case Else
                Title = "RANGED WEAPONS"
        End Select

        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, curTop, Title, boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        curTop = curTop + 0.1875
        MaxHeight = boxHeight - (curTop - boxTop)

        SetWeaponColumns(WeaponType, boxLeft)
        CreateWeaponFields(WeaponType)
        SetFieldSizes()

        'print sub-heads
        curTop = PrintSubHeads(boxLeft, curTop)
        MaxHeight = boxHeight - (curTop - boxTop)

        'and now print the fields
        curTop = PrintFields(Equipment, boxLeft, curTop, MaxHeight, StartFrom, True, WeaponType)
        MaxHeight = boxHeight - (curTop - boxTop)

        '*
        '* Draw the broken boxes around the info
        '*
        Dim mainboxWidth As Double = ColLefts(BrokenEffectCol) - boxLeft
        Dim subboxLeft As Double = ColLefts(BrokenEffectCol)
        Dim subboxRight As Double = ColLefts(numCols) + ColWidths(numCols)
        Dim curHeight As Double = curTop - boxTop

        'draw the right box
        MyPrintDoc.RenderDirectRectangle(subboxLeft, boxTop, subboxRight - subboxLeft, curHeight, ld)

        'clear the white area
        Dim ldWhite As New C1.C1Preview.LineDef(0.01, Color.White, Drawing2D.DashStyle.Solid)
        MyPrintDoc.RenderDirectRectangle(ColLefts(BrokenEffectCol), boxTop - 1 / 32, ColWidths(BrokenEffectCol), curHeight + 1 / 16, ldWhite, Color.White)

        'draw the left box
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, mainboxWidth, curHeight, ld)


        Return curTop
    End Function
    Private Function PrintOverflowGrimoire(TopSide As Double, MaxHeight As Double, StartFrom As Integer) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim subHeadHeight As Double
        Dim Title As String
        Dim TitleHeight As Double

        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        '* Generate Weapons
        boxLeft = MarginLeft
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = TopSide
        boxHeight = MaxHeight


        MaxWidth = boxWidth

        Title = "Grimoire of " & MyChar.Name
        Dim tmpSize As Single = MyPrintDoc.Style.FontSize * 1.5
        Dim tmpFont As New Font(MyPrintDoc.Style.FontName, tmpSize, FontStyle.Bold)

        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, curTop, Title, boxWidth, tmpFont, MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        TitleHeight = GetHeight("X", tmpFont)

        curTop = curTop + TitleHeight '0.1875
        MaxHeight = boxHeight - (curTop - boxTop)

        boxHeight = boxHeight - TitleHeight
        boxTop = curTop

        SetGrimoireColumns(boxLeft)
        CreateGrimoireFields()
        SetFieldSizes(True)

        'print sub-heads
        subHeadHeight = GetSubHeadHeight()
        curTop = PrintSubHeads(boxLeft, curTop, subHeadHeight)
        MaxHeight = boxHeight - (curTop - boxTop)

        'and now print the fields
        curTop = PrintGrimoireFields(Equipment, boxLeft, curTop, MaxHeight, StartFrom)
        MaxHeight = boxHeight - (curTop - boxTop)


        'draw the box
        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, curTop - boxTop, ld)
        'MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        Return curTop
    End Function
    Private Function PrintSectionAdsDisads(TopSide As Double) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim maxHeight As Double
        Dim LocalOverflow As Boolean = False

        Dim BottomOffset As Double = 1 'always leave room for some Disads

        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        '* Generate Advantages
        boxLeft = MarginLeft
        boxWidth = 3 + 9 / 16
        boxTop = TopSide + 1 / 16
        boxHeight = PageHeight - MarginBottom - boxTop - FooterHeight

        MaxWidth = boxWidth
        maxHeight = boxHeight

        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft, curTop, "ADVANTAGES AND PERKS", boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        curTop = curTop + 0.1875
        maxHeight = boxHeight - (curTop - boxTop)

        '***** Advantages
        'set the column widths we'll need
        SetTraitColumns(Ads, boxLeft)

        'set the list of advantages we'll be printing
        CreateFields(Ads)

        'get the sizes of all our data lines
        SetFieldSizes()

        'and now print the fields
        curTop = PrintFields(Ads, boxLeft, curTop, maxHeight - BottomOffset)
        maxHeight = boxHeight - (curTop - boxTop)
        If OverflowFrom(Ads) >= 0 Then LocalOverflow = True

        '***** Perks
        If Not LocalOverflow Then
            'no point in printing if we didn't even get through the advantages

            '* Generate Perks
            'uses same columns as Ads
            CreateFields(Perks)
            SetFieldSizes()

            curTop = PrintFields(Perks, boxLeft, curTop, maxHeight - BottomOffset)
            maxHeight = boxHeight - (curTop - boxTop)
            If OverflowFrom(Perks) >= 0 Then LocalOverflow = True
        Else
            'didn't get to print any, so automatically overflowed
            OverflowFrom(Perks) = 0
        End If

        '***** Disadvantages
        LocalOverflow = False 'always print some Disads
        'If Not LocalOverflow Then
        'no point in printing if we didn't even get through the above
        MyPrintDoc.RenderDirectText(boxLeft, curTop, "DISADVANTAGES AND QUIRKS", boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        curTop = curTop + 0.1875
        maxHeight = boxHeight - (curTop - boxTop)

        '* Generate Disadvantages
        'uses same columns as Ads
        CreateFields(Disads)
        SetFieldSizes()

        curTop = PrintFields(Disads, boxLeft, curTop, maxHeight)
        maxHeight = boxHeight - (curTop - boxTop)
        If OverflowFrom(Disads) >= 0 Then LocalOverflow = True
        'End If

        '***** Quirks
        If Not LocalOverflow Then
            'no point in printing if we didn't even get through the above

            '* Generate Quirks
            'uses same columns as Ads
            CreateFields(Quirks)
            SetFieldSizes()

            curTop = PrintFields(Quirks, boxLeft, curTop, maxHeight)
            If OverflowFrom(Quirks) >= 0 Then LocalOverflow = True
        Else
            'didn't get to print any, so automatically overflowed
            OverflowFrom(Quirks) = 0
        End If

        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        Return curTop
    End Function

    Private Function PrintSectionSkills(LeftSide As Double, TopSide As Double) As Double
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim curTop As Double
        Dim maxHeight As Double

        'Dim ld As New C1.C1Preview.LineDef(0.01, Color.Green, Drawing2D.DashStyle.Solid)
        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        '* Generate Advantages
        boxLeft = LeftSide
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = TopSide + 1 / 16
        boxHeight = PageHeight - MarginBottom - boxTop - FooterHeight

        maxWidth = boxWidth
        maxHeight = boxHeight

        'MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld, Brushes.Orange)

        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft, curTop, "SKILLS", boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Center)
        curTop = curTop + 0.1875
        maxHeight = boxHeight - (curTop - boxTop)

        '***** Skills
        'set the column widths we'll need
        SetTraitColumns(Skills, boxLeft)

        'set the list of advantages we'll be printing
        CreateFields(Skills)

        'get the sizes of all our data lines
        SetFieldSizes()

        'print sub-heads
        curTop = PrintSubHeads(boxLeft, curTop)
        maxHeight = boxHeight - (curTop - boxTop)

        'and now print the fields
        curTop = PrintFields(Skills, boxLeft, curTop, maxHeight)
        maxHeight = boxHeight - (curTop - boxTop)


        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        Return curTop
    End Function
    Private Function GetHeight(Text As String) As Double
        Dim sizD As SizeD
        sizD = GetSize(Text)

        Return sizD.Height
    End Function
    Private Function GetHeight(Text As String, WithFont As Font) As Double
        Dim sizD As SizeD

        Dim rt As New RenderText(Text, WithFont)

        sizD = GetSize(rt)

        Return sizD.Height
    End Function
    Private Function GetHeight(Text As String, WithStyle As Style) As Double
        Dim sizD As SizeD

        Dim rt As New RenderText(Text, WithStyle)

        sizD = GetSize(rt)

        Return sizD.Height
    End Function
    Private Function GetSize(UnAttachedRenderObject As RenderObject) As SizeD
        'Gets size based on a given Render object, 
        'that is NOT YET attached to the document,
        'with no additional given size restrictions

        Dim szD As SizeD

        'get size
        MyPrintDoc.Body.Children.Add(UnAttachedRenderObject)
        'measure
        szD = UnAttachedRenderObject.CalcSize(Unit.Auto, Unit.Auto)
        MyPrintDoc.Body.Children.Remove(UnAttachedRenderObject)

        Return szD
    End Function
    Private Function GetSize(UnAttachedRenderText As RenderText) As SizeD
        'Gets size based on a given RenderText object, 
        'that is NOT YET attached to the document,
        'with no additional given size restrictions

        Dim szD As SizeD

        'get size
        MyPrintDoc.Body.Children.Add(UnAttachedRenderText)
        'measure
        szD = UnAttachedRenderText.CalcSize(Unit.Auto, Unit.Auto)
        MyPrintDoc.Body.Children.Remove(UnAttachedRenderText)

        Return szD
    End Function
    Private Function GetSize(UnAttachedRenderText As RenderText, AvailableWidth As Double) As SizeD
        'Gets size based on a given RenderText object, 
        'that is NOT YET attached to the document,
        'given a fixed available width

        Dim szD As SizeD

        'get size
        MyPrintDoc.Body.Children.Add(UnAttachedRenderText)
        'measure
        szD = UnAttachedRenderText.CalcSize(AvailableWidth, Unit.Auto)
        MyPrintDoc.Body.Children.Remove(UnAttachedRenderText)

        Return szD
    End Function
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
    Private Function GetSize(Text As String, AvailableWidth As Double, WithStyle As Style) As SizeD
        'Gets size based on given style, 
        'given a fixed available width

        Dim szD As SizeD
        Dim rText As RenderText

        rText = New RenderText(Text, WithStyle)

        'get size
        MyPrintDoc.Body.Children.Add(rText)
        'measure
        szD = rText.CalcSize(AvailableWidth, Unit.Auto)
        MyPrintDoc.Body.Children.Remove(rText)

        Return szD
    End Function
    Private Function GetSize(Text As String, AvailableWidth As Double) As SizeD
        'Gets size based on normal text, 
        'given a fixed available width

        Return GetSize(Text, AvailableWidth, MyPrintDoc.Style)
    End Function

    Private Sub CreateFooter()
        Dim footer As New C1.C1Preview.RenderArea
        Dim table As New C1.C1Preview.RenderTable(MyPrintDoc)
        Dim font As New Font("Times New Roman", 6)

        'Setup the table's style
        table.Style.Font = font
        table.Style.WordWrap = False
        table.Cols(0).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Left
        table.Cols(1).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Center
        table.Cols(2).Style.TextAlignHorz = C1.C1Preview.AlignHorzEnum.Right

        'Add our footer content to the table
        table.Cells(0, 0).Text = MyChar.Name '"GURPS Character Assistant 5"
        table.Cells(0, 1).Text = "Page [[PageNo]] of [[PageCount]]."
        table.Cells(0, 2).Text = "http://www.sjgames.com/gurps/characterassistant/"

        'Add the table and some padding to the footer object
        footer.Children.Add(New RenderEmpty(0.05))
        footer.Children.Add(table)

        'Calculate the size of the footer object and save it
        'n.b. Until the table is rendered its contents aren't finalized and thus don't
        '     have a calculable size.  So, we calculate the footer object size, for the
        '     non-zero size of the table itself and the padding object, and then add the
        '     calculated size of a text blob which is equivalent to the table's payload.
        '     Thanks ever so much, ComponentOne!
        FooterHeight = GetSize(footer).Height + GetSize(New RenderText("DX", font)).Height

        'Add the footer to the page layout
        MyPrintDoc.PageLayout.PageFooter = footer

    End Sub
    Private Sub CreateStyles()
        Dim tmpSheetFont As Font = MyOptions.Value("SheetFont")
        Dim tmpUserFont As Font = MyOptions.Value("UserFont")

        '*
        '* Form Text
        '*
        Dim SheetTextLeftStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetTextLeftStyle.Font = MyOptions.Value("SheetFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        SheetTextLeftStyle.TextAlignVert = AlignVertEnum.Top
        SheetTextLeftStyle.TextAlignHorz = AlignHorzEnum.Left
        SheetTextLeftStyle.TextColor = MyOptions.Value("SheetTextColor") 'Color.Black
        MyStyles.Add(SheetTextLeftStyle, "SheetTextLeftStyle")

        Dim SheetTextRightStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetTextRightStyle.Font = MyOptions.Value("SheetFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        SheetTextRightStyle.TextAlignVert = AlignVertEnum.Top
        SheetTextRightStyle.TextAlignHorz = AlignHorzEnum.Right
        SheetTextRightStyle.TextColor = MyOptions.Value("SheetTextColor") 'Color.Black
        MyStyles.Add(SheetTextRightStyle, "SheetTextRightStyle")

        Dim SheetTextCenterStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetTextCenterStyle.Font = MyOptions.Value("SheetFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        SheetTextCenterStyle.TextAlignVert = AlignVertEnum.Top
        SheetTextCenterStyle.TextAlignHorz = AlignHorzEnum.Center
        SheetTextCenterStyle.TextColor = MyOptions.Value("SheetTextColor") 'Color.Black
        MyStyles.Add(SheetTextCenterStyle, "SheetTextCenterStyle")

        Dim SheetTextLeftBoldStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetTextLeftBoldStyle.Font = MyOptions.Value("SheetFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        SheetTextLeftBoldStyle.TextAlignVert = AlignVertEnum.Top
        SheetTextLeftBoldStyle.TextAlignHorz = AlignHorzEnum.Left
        SheetTextLeftBoldStyle.TextColor = MyOptions.Value("SheetTextColor") 'Color.Black
        SheetTextLeftBoldStyle.FontBold = True
        MyStyles.Add(SheetTextLeftBoldStyle, "SheetTextLeftBoldStyle")

        Dim SheetTextRightBoldStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetTextRightBoldStyle.Font = MyOptions.Value("SheetFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        SheetTextRightBoldStyle.TextAlignVert = AlignVertEnum.Top
        SheetTextRightBoldStyle.TextAlignHorz = AlignHorzEnum.Right
        SheetTextRightBoldStyle.TextColor = MyOptions.Value("SheetTextColor") 'Color.Black
        SheetTextRightBoldStyle.FontBold = True
        MyStyles.Add(SheetTextRightBoldStyle, "SheetTextRightBoldStyle")

        Dim SheetTextCenterBoldStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetTextCenterBoldStyle.Font = MyOptions.Value("SheetFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        SheetTextCenterBoldStyle.TextAlignVert = AlignVertEnum.Top
        SheetTextCenterBoldStyle.TextAlignHorz = AlignHorzEnum.Center
        SheetTextCenterBoldStyle.TextColor = MyOptions.Value("SheetTextColor") 'Color.Black
        SheetTextCenterBoldStyle.FontBold = True
        MyStyles.Add(SheetTextCenterBoldStyle, "SheetTextCenterBoldStyle")




        Dim SheetTinyTextStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetTinyTextStyle.Font = New Font(tmpSheetFont.Name, 6)
        SheetTinyTextStyle.TextAlignVert = AlignVertEnum.Center
        SheetTinyTextStyle.TextAlignHorz = AlignHorzEnum.Center
        MyStyles.Add(SheetTinyTextStyle, "SheetTinyTextStyle")

        Dim SheetTinyTextAtBottomStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetTinyTextAtBottomStyle.Font = New Font(tmpSheetFont.Name, 6)
        SheetTinyTextAtBottomStyle.TextAlignVert = AlignVertEnum.Bottom
        SheetTinyTextAtBottomStyle.TextAlignHorz = AlignHorzEnum.Center
        MyStyles.Add(SheetTinyTextAtBottomStyle, "SheetTinyTextAtBottomStyle")

        Dim SheetAttributeLeftStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetAttributeLeftStyle.Font = New Font(tmpSheetFont.Name, 18, FontStyle.Bold)
        SheetAttributeLeftStyle.TextAlignVert = AlignVertEnum.Center
        SheetAttributeLeftStyle.TextAlignHorz = AlignHorzEnum.Left
        MyStyles.Add(SheetAttributeLeftStyle, "SheetAttributeLeftStyle")

        Dim SheetAttributeRightStyle As Style = MyPrintDoc.Style.Children.Add()
        SheetAttributeRightStyle.Font = New Font(tmpSheetFont.Name, 18, FontStyle.Bold)
        SheetAttributeRightStyle.TextAlignVert = AlignVertEnum.Center
        SheetAttributeRightStyle.TextAlignHorz = AlignHorzEnum.Right
        MyStyles.Add(SheetAttributeRightStyle, "SheetAttributeRightStyle")


        '*
        '* User-Entered Text
        '*
        Dim UserTextLeftStyle As Style = MyPrintDoc.Style.Children.Add()
        UserTextLeftStyle.Font = MyOptions.Value("UserFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        UserTextLeftStyle.TextAlignVert = AlignVertEnum.Top
        UserTextLeftStyle.TextAlignHorz = AlignHorzEnum.Left
        UserTextLeftStyle.TextColor = MyOptions.Value("UserTextColor")
        MyStyles.Add(UserTextLeftStyle, "UserTextLeftStyle")

        Dim UserTextRightStyle As Style = MyPrintDoc.Style.Children.Add()
        UserTextRightStyle.Font = MyOptions.Value("UserFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        UserTextRightStyle.TextAlignVert = AlignVertEnum.Top
        UserTextRightStyle.TextAlignHorz = AlignHorzEnum.Right
        UserTextRightStyle.TextColor = MyOptions.Value("UserTextColor")
        MyStyles.Add(UserTextRightStyle, "UserTextRightStyle")

        Dim UserTextCenterStyle As Style = MyPrintDoc.Style.Children.Add()
        UserTextCenterStyle.Font = MyOptions.Value("UserFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        UserTextCenterStyle.TextAlignVert = AlignVertEnum.Top
        UserTextCenterStyle.TextAlignHorz = AlignHorzEnum.Center
        UserTextCenterStyle.TextColor = MyOptions.Value("UserTextColor")
        MyStyles.Add(UserTextCenterStyle, "UserTextCenterStyle")

        Dim UserTextCenterCenterStyle As Style = MyPrintDoc.Style.Children.Add()
        UserTextCenterCenterStyle.Font = MyOptions.Value("UserFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        UserTextCenterCenterStyle.TextAlignVert = AlignVertEnum.Top
        UserTextCenterCenterStyle.TextAlignHorz = AlignHorzEnum.Center
        UserTextCenterCenterStyle.TextColor = MyOptions.Value("UserTextColor")
        MyStyles.Add(UserTextCenterCenterStyle, "UserTextCenterCenterStyle")

        'NEW 2015 05 25
        'For in-line bonuses and conditionals
        Dim UserSmallTextStyle As Style = MyPrintDoc.Style.Children.Add()
        UserSmallTextStyle.Font = MyOptions.Value("UserBonusesFont")
        UserSmallTextStyle.TextAlignVert = AlignVertEnum.Top
        UserSmallTextStyle.TextAlignHorz = AlignHorzEnum.Left
        UserSmallTextStyle.TextColor = MyOptions.Value("UserBonusesColor")
        MyStyles.Add(UserSmallTextStyle, "UserSmallTextStyle")
        'END NEW

        'NEW 2015 09 08
        'for grouped trait headings
        Dim UserTextLeftStyleBold As Style = MyPrintDoc.Style.Children.Add()
        UserTextLeftStyleBold.Font = MyOptions.Value("UserFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        UserTextLeftStyleBold.TextAlignVert = AlignVertEnum.Top
        UserTextLeftStyleBold.TextAlignHorz = AlignHorzEnum.Left
        UserTextLeftStyleBold.TextColor = MyOptions.Value("UserTextColor")
        UserTextLeftStyleBold.FontBold = True
        MyStyles.Add(UserTextLeftStyleBold, "UserTextLeftStyleBold")

        Dim UserTextRightStyleBold As Style = MyPrintDoc.Style.Children.Add()
        UserTextRightStyleBold.Font = MyOptions.Value("UserFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        UserTextRightStyleBold.TextAlignVert = AlignVertEnum.Top
        UserTextRightStyleBold.TextAlignHorz = AlignHorzEnum.Right
        UserTextRightStyleBold.TextColor = MyOptions.Value("UserTextColor")
        UserTextRightStyleBold.FontBold = True
        MyStyles.Add(UserTextRightStyleBold, "UserTextRightStyleBold")

        Dim UserTextCenterStyleBold As Style = MyPrintDoc.Style.Children.Add()
        UserTextCenterStyleBold.Font = MyOptions.Value("UserFont") 'New Font(MyPrintDoc.Style.Font.Name, 10)
        UserTextCenterStyleBold.TextAlignVert = AlignVertEnum.Top
        UserTextCenterStyleBold.TextAlignHorz = AlignHorzEnum.Center
        UserTextCenterStyleBold.TextColor = MyOptions.Value("UserTextColor")
        UserTextCenterStyleBold.FontBold = True
        MyStyles.Add(UserTextCenterStyleBold, "UserTextCenterStyleBold")
        'END NEW



        Dim AttributeLeftStyle As Style = MyPrintDoc.Style.Children.Add()
        AttributeLeftStyle.Font = New Font(tmpUserFont.Name, 18, FontStyle.Bold)
        AttributeLeftStyle.TextAlignVert = AlignVertEnum.Center
        AttributeLeftStyle.TextAlignHorz = AlignHorzEnum.Left
        MyStyles.Add(AttributeLeftStyle, "AttributeLeftStyle")

        Dim AttributeCenterStyle As Style = MyPrintDoc.Style.Children.Add()
        AttributeCenterStyle.Font = New Font(tmpUserFont.Name, 18, FontStyle.Bold)
        AttributeCenterStyle.TextAlignVert = AlignVertEnum.Center
        AttributeCenterStyle.TextAlignHorz = AlignHorzEnum.Center
        AttributeCenterStyle.TextColor = MyOptions.Value("UserTextColor")
        MyStyles.Add(AttributeCenterStyle, "AttributeCenterStyle")

        Dim AttributeRightStyle As Style = MyPrintDoc.Style.Children.Add()
        AttributeRightStyle.Font = New Font(tmpUserFont.Name, 18, FontStyle.Bold)
        AttributeRightStyle.TextAlignVert = AlignVertEnum.Center
        AttributeRightStyle.TextAlignHorz = AlignHorzEnum.Right
        MyStyles.Add(AttributeRightStyle, "AttributeRightStyle")

        Dim AttributePointsStyle As Style = MyPrintDoc.Style.Children.Add()
        AttributePointsStyle.Font = New Font(tmpUserFont.Name, 10, FontStyle.Bold)
        AttributePointsStyle.TextAlignVert = AlignVertEnum.Center
        AttributePointsStyle.TextAlignHorz = AlignHorzEnum.Center
        AttributePointsStyle.TextColor = MyOptions.Value("UserTextColor")
        MyStyles.Add(AttributePointsStyle, "AttributePointsStyle")


    End Sub

    Public Function ReturnRTFHeaderUserOptions() As String
        'Remember: this header + RTF text + "}"
        Dim cS As Color = MyOptions.Value("SheetTextColor")
        Dim cU As Color = MyOptions.Value("UserTextColor")

        Dim f As Font
        f = MyOptions.Value("SheetFont")
        Dim fS As String = f.Name
        f = MyOptions.Value("UserFont")
        Dim fU As String = f.Name

        Dim RTFHeader As String
        RTFHeader = "{\rtf1\ansi\ansicpg1252\deff0\deflang1033"
        RTFHeader = RTFHeader & "{\fonttbl{\f0\fnil\fcharset0 " & fS & ";}"
        RTFHeader = RTFHeader & "{\f1\fnil\fcharset0 " & fU & ";}"
        RTFHeader = RTFHeader & "}"
        'RTFHeader = RTFHeader & "{\colortbl ;\red255\green0\blue0;\red0\green77\blue187;\red0\green176\blue80;\red156\green133\blue192;}"
        RTFHeader = RTFHeader & "{\colortbl ;\red" & cS.R.ToString & "\green" & cS.G.ToString & "\blue" & cS.B.ToString & ";"
        RTFHeader = RTFHeader & "\red" & cU.R.ToString & "\green" & cU.G.ToString & "\blue" & cU.B.ToString & ";\red255\green0\blue0;}"
        RTFHeader = RTFHeader & "\viewkind4\uc1\pard\sa200\sl276\slmult1\lang9\fs20 "

        Return RTFHeader
    End Function
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
            Case Equipment
                'weapons are printed elsewhere, so don't include them here
                If curItem.DamageModeTagItemCount("charreach") > 0 Then
                    Return False
                End If
                If curItem.DamageModeTagItemCount("charrangemax") > 0 Then
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

        'Dim MyTraits As SortedTraitCollection = MyChar.ItemsByType(ItemType)
        Fields = New ArrayList '0 based



        '*****
        'NEW TESTING 2015 09 08
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
                curItem = MyTraits.Item(curTrait)

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




        'For curTrait = 1 To MyTraits.Count
        '    curItem = MyTraits.Item(curTrait)

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
    Private Sub CreateGrimoireFields()
        'This routine creates all the fields we'll have in the list,
        'for all the valid traits, and if ShowComponents, then
        'for all children of valid traits, too.
        Dim curItem As GCATrait
        Dim tmp As String = ""

        Dim curTrait As Integer = 0

        Fields = New ArrayList '0 based

        For curTrait = 1 To MyChar.ItemsByType(Spells).Count
            curItem = MyChar.ItemsByType(Spells).Item(curTrait)

            'If ValidGrimoireTrait(curItem) Then
            'This trait is valid, so include it.
            AddGrimoireField(curItem)
            'End If
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

        curField.Values = New ArrayList
        curField.Values.Add("") 'icon col
        Select Case ItemType
            Case Languages
                '4 columns: icon & name & level & points
                curField.Values.Add(curItem.FullNameTL)
                curField.Values.Add(curItem.LevelName)
                curField.Values.Add(curItem.Points)

            Case Stats
                '4 columns: icon & name & score & points
                curField.Values.Add(curItem.DisplayName)
                curField.Values.Add(curItem.Score)
                curField.Values.Add(curItem.Points)

            Case Skills
                '5 columns: icon & name & level & rel level & points
                curField.Values.Add(curItem.DisplayName)
                If curField.IsCombo Then
                    curField.Values.Add(curItem.TagItem("combolevel"))
                    curField.Values.Add(curItem.RelativeLevel)
                Else
                    curField.Values.Add(curItem.Level)
                    curField.Values.Add(curItem.RelativeLevel)
                End If
                curField.Values.Add(curItem.Points)

            Case Spells
                '4 columns: icon & name & level & points
                curField.Values.Add(curItem.DisplayName)
                curField.Values.Add(curItem.Level)
                curField.Values.Add(curItem.Points)

            Case Equipment
                '7 columns: icon & qty & name & location & split & cost & weight
                curField.Values.Add(curItem.TagItem("count"))
                curField.Values.Add(curItem.DisplayName)
                curField.Values.Add(curItem.TagItem("location"))
                curField.Values.Add("") 'for the broken box column
                curField.Values.Add(curItem.TagItem("cost"))
                curField.Values.Add(curItem.TagItem("weight"))

            Case Else
                '3 columns: icon & name & points
                curField.Values.Add(curItem.DisplayName)
                curField.Values.Add(curItem.Points)
        End Select
        curField.Values.Add("") 'icon col

        'NEW TESTING 2015 09 08
        If curItem.IDKey = 0 Then
            'PLACEHOLDER
            'only print name of the item, but leave the blanks for the other columns
            If ItemType = Equipment Then
                For curCol = 3 To curField.Values.Count - 1
                    curField.Values(curCol) = ""
                Next
            Else
                For curCol = 2 To curField.Values.Count - 1
                    curField.Values(curCol) = ""
                Next
            End If
        End If
        'END NEW


        'NEW 2015 05 25
        'Allow for increasing the height to print included bonuses.
        Dim myValue As String

        myValue = ""
        If curItem.BonusListItemsCount > 0 Then
            myValue = "Includes: " & curItem.TagItem("bonuslist")
        End If
        If curItem.ConditionalListItemsCount > 0 Then
            If myValue <> "" Then myValue = myValue + vbCrLf
            myValue = myValue & "Conditional: " & curItem.TagItem("conditionallist")
        End If

        If myValue <> "" Then
            curField.BonusText = myValue
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
                    curField.Values.Add(curItem.TagItem("count"))

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

                    curField.Values.Add("") 'for the broken box column
                    curField.Values.Add(curItem.TagItem("cost"))
                    curField.Values.Add(curItem.TagItem("weight"))
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
                    curField.Values.Add(curItem.TagItem("count"))

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
                        curField.Values.Add("")
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

                        curField.Values.Add("") 'for the broken box column
                        curField.Values.Add("")
                        curField.Values.Add("")
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
                    curField.Values.Add(curItem.TagItem("count"))

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

                    curField.Values.Add("") 'for the broken box column
                    curField.Values.Add(curItem.TagItem("cost"))
                    curField.Values.Add(curItem.TagItem("weight"))
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
                    curField.Values.Add(curItem.TagItem("count"))

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

                    curField.Values.Add("") 'for the broken box column
                    curField.Values.Add(curItem.TagItem("cost"))
                    curField.Values.Add(curItem.TagItem("weight"))
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
                        curField.Values.Add("")
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

                        curField.Values.Add("") 'for the broken box column
                        curField.Values.Add("")
                        curField.Values.Add("")
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
    Private Sub AddGrimoireField(curItem As GCATrait)
        Dim curField As clsDisplayField

        'create the field
        curField = New clsDisplayField

        'add the trait
        curField.Data = curItem
        curField.LinePen = New Pen(Brushes.LightGray) 'New Pen(MyColorProfile.TraitTextColor)

        Fields.Add(curField)

        curField.Values = New ArrayList
        curField.Values.Add("") 'icon col

        curField.Values.Add(curItem.FullNameTL)
        curField.Values.Add(curItem.TagItem("class"))
        curField.Values.Add(curItem.Level)
        curField.Values.Add(curItem.TagItem("time"))
        curField.Values.Add(curItem.TagItem("duration"))
        curField.Values.Add(curItem.TagItem("castingcost"))
        curField.Values.Add(curItem.TagItem("cat"))
        curField.Values.Add(curItem.TagItem("page"))

        curField.Values.Add("") 'icon col

    End Sub
    Private Sub SetFieldSizes(Optional ByVal GrimoireStyle As Boolean = False)
        '*****
        '* Size and store all our trait value lines
        '*
        Dim curCol As Integer = 0
        Dim curTop As Double = 0
        Dim AltLine As Boolean = True
        Dim tmpLevel As String = ""
        Dim sz As SizeF
        Dim szD As SizeD

        Dim curField As clsDisplayField
        Dim curItem As GCATrait

        Dim myLeft As Double = 0
        Dim myWidth As Double = 0
        Dim myValue As String = ""
        Dim bonusLeft As Double = 0 'NEW 2015 05 25
        Dim FieldIndex As Integer

        For FieldIndex = 0 To Fields.Count - 1
            'Set our Field values
            curField = Fields(FieldIndex)
            curItem = curField.Data

            sz = New SizeF(MaxWidth, 0)

            For curCol = 1 To numCols - 1
                'myValue = Values(curCol)
                myValue = curField.Values(curCol) ' Values(curCol)

                If curField.IsCombo Then
                    myLeft = AltColLefts(curCol)
                    myWidth = AltColWidths(curCol)
                Else
                    myLeft = ColLefts(curCol)
                    myWidth = ColWidths(curCol)
                End If

                If curCol = ComponentDrawCol AndAlso curField.IsComponent > 0 Then
                    myLeft = myLeft + (IndentStepSize * curField.IsComponent)
                    myWidth = myWidth - (IndentStepSize * curField.IsComponent) 'NEW 2015 05 25, bug fix
                End If
                'NEW 2015 05 25
                'Allow for increasing the size to print included bonuses.
                If curCol = ComponentDrawCol Then
                    bonusLeft = myLeft
                End If
                'END NEW

                'Size our text, getting the greatest height of all cols
                'to use for our field height
                If GrimoireStyle Then
                    szD = GetSize(myValue, myWidth, MyStyles("SheetTextLeftStyle"))
                Else
                    szD = GetSize(myValue, myWidth, MyStyles("UserTextLeftStyle"))
                End If
                If szD.Height > sz.Height Then
                    sz.Height = szD.Height
                End If
            Next

            'NEW 2015 05 25
            'Allow for increasing the size to print included bonuses.
            If curField.BonusText <> "" Then
                curField.BonusTop = sz.Height
                curField.BonusLeft = bonusLeft
                curField.BonusWidth = MaxWidth - bonusLeft + ColLefts(0)

                szD = GetSize(curField.BonusText, curField.BonusWidth, MyStyles("UserSmallTextStyle"))

                sz.Height = sz.Height + szD.Height
            End If
            'END NEW

            curField.Point = New PointF(0, 0)
            curField.Size = sz
            curField.Alt = AltLine

            'Adjust for next lines
            AltLine = Not AltLine
        Next

    End Sub
    Private Function GetSubHeadHeight() As Double
        Dim curCol As Integer = 0
        Dim sz As SizeF
        Dim szD As SizeD

        Dim myWidth As Double = 0
        Dim myValue As String = ""

        sz = New SizeF(MaxWidth, 0)

        For curCol = 1 To numCols - 1
            myValue = SubHeads(curCol)
            myWidth = ColWidths(curCol)

            'Size our text, getting the greatest height of all cols
            'to use for our field height
            szD = GetSize(myValue, myWidth, MyStyles("SheetTextLeftStyle"))
            If szD.Height > sz.Height Then
                sz.Height = szD.Height
            End If
        Next

        Return sz.Height
    End Function
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

        ComponentDrawCol = 2

        '0 based columns
        Select Case WeaponType
            Case Hand
                'If Options.Value("ShowMinST") Then
                '    numCols = 11
                'else
                numCols = 10
                'End If
                ReDim ColLefts(numCols)
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                SubHeads(curCol) = "Qty"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = miniColWidth

                curCol += 1
                SubHeads(curCol) = "Weapon"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = medColWidth

                curCol += 1
                SubHeads(curCol) = "Damage"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = wideColWidth

                curCol += 1
                SubHeads(curCol) = "Reach"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = minColWidth

                curCol += 1
                'If Options.Value("ShowSkillParry") Then
                'SubHeads(curCol) = "Lvl(Pry)"
                'Else
                SubHeads(curCol) = "Parry"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = minColWidth
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
                ColWidths(curCol) = medColWidth

                curCol += 1
                'for the broken box effect
                BrokenEffectCol = curCol
                SubHeads(curCol) = ""
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = 1 / 16

                curCol += 1
                SubHeads(curCol) = "Cost"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = medColWidth

                curCol += 1
                SubHeads(curCol) = "Weight"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = medColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth


                'build the Lefts
                ColLefts(0) = IconColLeft
                ColLefts(1) = ColLefts(0) + ColWidths(0)

                ColLefts(numCols) = LeftOffset + MaxWidth - ColWidths(numCols)

                For curCol = numCols - 1 To 3 Step -1
                    ColLefts(curCol) = ColLefts(curCol + 1) - ColWidths(curCol)
                Next

                'let the Name col use whatever space is available
                ColLefts(2) = ColLefts(1) + ColWidths(1)
                ColWidths(2) = ColLefts(3) - ColLefts(2) 'col to the right - my starting left

            Case Ranged
                'If Options.Value("ShowRangedLevel") Then
                '    numCols = 17
                'else
                numCols = 16
                'End If

                ReDim ColLefts(numCols)
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                SubHeads(curCol) = "Qty"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = miniColWidth

                curCol += 1
                SubHeads(curCol) = "Weapon"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = minColWidth

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
                ColWidths(curCol) = miniColWidth

                curCol += 1
                SubHeads(curCol) = "LC"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = miniColWidth

                curCol += 1
                SubHeads(curCol) = "Notes"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = medColWidth

                curCol += 1
                'for the broken box effect
                BrokenEffectCol = curCol
                SubHeads(curCol) = ""
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = 1 / 16

                curCol += 1
                SubHeads(curCol) = "Cost"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = medColWidth

                curCol += 1
                SubHeads(curCol) = "Weight"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = medColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth


                'build the Lefts
                ColLefts(0) = IconColLeft
                ColLefts(1) = ColLefts(0) + ColWidths(0)

                ColLefts(numCols) = LeftOffset + MaxWidth - ColWidths(numCols)

                For curCol = numCols - 1 To 3 Step -1
                    ColLefts(curCol) = ColLefts(curCol + 1) - ColWidths(curCol)
                Next

                'let the Name col use whatever space is available
                ColLefts(2) = ColLefts(1) + ColWidths(1)
                ColWidths(2) = ColLefts(3) - ColLefts(2) 'col to the right - my starting left

        End Select

    End Sub
    Private Sub SetGrimoireColumns(LeftOffset As Double)
        Dim miniColWidth As Double = 0
        Dim minColWidth As Double = 0
        Dim medColWidth As Double = 0
        Dim medplusColWidth As Double = 0
        Dim medwideColWidth As Double = 0
        Dim wideColWidth As Double = 0
        Dim wideplusColWidth As Double = 0
        Dim extrawideColWidth As Double = 0
        Dim curCol As Integer

        PointsCol = 0

        miniColWidth = 0.25
        minColWidth = 0.375
        medColWidth = 0.5
        medplusColWidth = 0.625
        medwideColWidth = 0.75 '0.625
        wideColWidth = 0.875
        wideplusColWidth = 1
        extrawideColWidth = 1.25

        IconColLeft = LeftOffset
        IconColWidth = 1 / 32 ' 0

        ComponentDrawCol = 1

        numCols = 9

        ReDim ColLefts(numCols)
        ReDim ColAligns(numCols)
        ReDim ColWidths(numCols)
        ReDim SubHeads(numCols)

        curCol = 0
        ColWidths(curCol) = IconColWidth
        'icon col

        curCol += 1
        SubHeads(curCol) = "Spell Name"
        ColAligns(curCol) = AlignHorzEnum.Left
        ColWidths(curCol) = miniColWidth

        curCol += 1
        SubHeads(curCol) = "Class"
        ColAligns(curCol) = AlignHorzEnum.Left
        ColWidths(curCol) = wideplusColWidth

        curCol += 1
        SubHeads(curCol) = "Skill Level"
        ColAligns(curCol) = AlignHorzEnum.Left
        ColWidths(curCol) = medColWidth

        curCol += 1
        SubHeads(curCol) = "Time to Cast"
        ColAligns(curCol) = AlignHorzEnum.Left
        ColWidths(curCol) = medplusColWidth

        curCol += 1
        SubHeads(curCol) = "Duration"
        ColAligns(curCol) = AlignHorzEnum.Left
        ColWidths(curCol) = medplusColWidth

        curCol += 1
        SubHeads(curCol) = "Cost to Cast/Maintain"
        ColAligns(curCol) = AlignHorzEnum.Left
        ColWidths(curCol) = wideColWidth

        curCol += 1
        SubHeads(curCol) = "College"
        ColAligns(curCol) = AlignHorzEnum.Left
        ColWidths(curCol) = extrawideColWidth

        curCol += 1
        SubHeads(curCol) = "Page No."
        ColAligns(curCol) = AlignHorzEnum.Left
        ColWidths(curCol) = medwideColWidth

        curCol += 1
        'icon col
        ColWidths(curCol) = IconColWidth


        'build the Lefts
        ColLefts(0) = IconColLeft

        ColLefts(numCols) = LeftOffset + MaxWidth - ColWidths(numCols)

        For curCol = numCols - 1 To 2 Step -1
            ColLefts(curCol) = ColLefts(curCol + 1) - ColWidths(curCol)
        Next

        'let the Name col use whatever space is available
        ColLefts(1) = ColLefts(0) + ColWidths(0)
        ColWidths(1) = ColLefts(2) - ColLefts(1) 'col to the right - my starting left

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

        'Dim szD As SizeD
        'szD = GetSize("[-99.9]")
        'ptsColWidth = szD.Width
        ptsColWidth = 0.375
        minColWidth = ptsColWidth
        miniColWidth = 0.25
        medColWidth = 0.5
        medplusColWidth = 0.75 '0.625
        wideColWidth = 0.875

        IconColLeft = LeftOffset
        IconColWidth = 1 / 32 ' 0

        ComponentDrawCol = 1

        '0 based columns 
        Select Case ItemType
            Case Languages
                '5 columns: icon & name & level & points & icon
                numCols = 4
                ReDim ColLefts(numCols)
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                SubHeads(1) = "Language"
                SubHeads(2) = "Level"
                SubHeads(3) = "Points"
                ColAligns(1) = AlignHorzEnum.Left
                ColAligns(2) = AlignHorzEnum.Left
                ColAligns(3) = AlignHorzEnum.Center
                PointsCol = 3

                'set icon col width
                ColLefts(0) = IconColLeft
                ColWidths(0) = IconColWidth

                'set rightmost col width and left side - Other Icon col
                ColLefts(4) = LeftOffset + MaxWidth - IconColWidth
                ColWidths(4) = IconColWidth

                'set next col to the left col width and left side - Points col
                ColLefts(3) = ColLefts(4) - ptsColWidth
                ColWidths(3) = ptsColWidth

                'set next col to the left col width and left side - Level col
                ColLefts(2) = ColLefts(3) - wideColWidth
                ColWidths(2) = wideColWidth

                'get name col width and left side - Name col
                ColLefts(1) = ColLefts(0) + ColWidths(0)
                ColWidths(1) = ColLefts(2) - ColLefts(1) 'col to the right - my starting left


            Case Stats
                '5 columns: icon & name & score & points & icon
                numCols = 4
                ReDim ColLefts(numCols)
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                SubHeads(1) = "Attribute"
                SubHeads(2) = "Score"
                SubHeads(3) = "Points"
                ColAligns(1) = AlignHorzEnum.Left
                ColAligns(2) = AlignHorzEnum.Center
                ColAligns(3) = AlignHorzEnum.Center
                PointsCol = 3

                'set icon col width
                ColLefts(0) = IconColLeft
                ColWidths(0) = IconColWidth

                'set rightmost col width and left side - Other Icon col
                ColLefts(4) = LeftOffset + MaxWidth - IconColWidth
                ColWidths(4) = IconColWidth

                'set next col to the left col width and left side - Points col
                ColLefts(3) = ColLefts(4) - ptsColWidth
                ColWidths(3) = ptsColWidth

                'set next col to the left col width and left side - Score col
                ColLefts(2) = ColLefts(3) - ptsColWidth
                ColWidths(2) = ptsColWidth

                'get name col width and left side - Name col
                ColLefts(1) = ColLefts(0) + ColWidths(0)
                ColWidths(1) = ColLefts(2) - ColLefts(1) 'col to the right - my starting left

            Case Skills
                'combos can have pretty wide levels, so we need to allow more room for them
                'szD = GetSize("-12+-12+-12")
                'altColWidth = szD.Width
                altColWidth = 0.75

                numCols = 5
                ReDim ColLefts(numCols)
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                SubHeads(1) = "Skill"
                SubHeads(2) = "Level"
                SubHeads(3) = "Relative"
                SubHeads(4) = "Points"
                ColAligns(1) = AlignHorzEnum.Left
                ColAligns(2) = AlignHorzEnum.Right
                ColAligns(3) = AlignHorzEnum.Center
                ColAligns(4) = AlignHorzEnum.Center
                PointsCol = 4

                'set icon col width
                ColLefts(0) = IconColLeft
                ColWidths(0) = IconColWidth

                'set rightmost col width and left side - Other Icon col
                ColLefts(5) = LeftOffset + MaxWidth - IconColWidth
                ColWidths(5) = IconColWidth

                'set next col to the left col width and left side - Points col
                ColLefts(4) = ColLefts(5) - ptsColWidth
                ColWidths(4) = ptsColWidth

                'set next col to the left col width and left side - Rel Level col
                ColLefts(3) = ColLefts(4) - medplusColWidth
                ColWidths(3) = medplusColWidth

                'set next col to the left col width and left side - Level col
                ColLefts(2) = ColLefts(3) - minColWidth
                ColWidths(2) = minColWidth

                'get name col width and left side - Name col
                ColLefts(1) = ColLefts(0) + ColWidths(0)
                ColWidths(1) = ColLefts(2) - ColLefts(1) 'col to the right - my starting left

                'ALTS
                ReDim AltColLefts(numCols) 'for combos
                ReDim AltColWidths(numCols) 'for combos

                'set icon col width
                AltColLefts(0) = IconColLeft
                AltColWidths(0) = IconColWidth

                'set rightmost col width and left side - Other Icon col
                AltColLefts(5) = LeftOffset + MaxWidth - IconColWidth
                AltColWidths(5) = IconColWidth

                'set next col to the left col width and left side - Points col
                AltColLefts(4) = AltColLefts(5) - ptsColWidth
                AltColWidths(4) = ptsColWidth

                'set next col to the left col width and left side - Rel Level col
                AltColLefts(3) = AltColLefts(4) - medplusColWidth
                AltColWidths(3) = medplusColWidth

                'set next col to the left col width and left side - Level col
                AltColLefts(2) = AltColLefts(3) - altColWidth
                AltColWidths(2) = altColWidth

                'get name col width and left side - Name col
                AltColLefts(1) = AltColLefts(0) + AltColWidths(0)
                AltColWidths(1) = AltColLefts(2) - AltColLefts(1) 'col to the right - my starting left

            Case Spells
                numCols = 4
                ReDim ColLefts(numCols)
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                SubHeads(1) = "Spell"
                SubHeads(2) = "Level"
                SubHeads(3) = "Points"
                ColAligns(1) = AlignHorzEnum.Left
                ColAligns(2) = AlignHorzEnum.Right
                ColAligns(3) = AlignHorzEnum.Center
                PointsCol = 3

                'set icon col width
                ColLefts(0) = IconColLeft
                ColWidths(0) = IconColWidth

                'set rightmost col width and left side - Other Icon col
                ColLefts(4) = LeftOffset + MaxWidth - IconColWidth
                ColWidths(4) = IconColWidth

                'set next col to the left col width and left side - Points col
                ColLefts(3) = ColLefts(4) - ptsColWidth
                ColWidths(3) = ptsColWidth

                'set next col to the left col width and left side - Level col
                ColLefts(2) = ColLefts(3) - minColWidth
                ColWidths(2) = minColWidth

                'get name col width and left side - Name col
                ColLefts(1) = ColLefts(0) + ColWidths(0)
                ColWidths(1) = ColLefts(2) - ColLefts(1) 'col to the right - my starting left

            Case Equipment
                '8 columns: icon & qty & name & location & split & cost & weight & icon
                Dim curCol As Integer

                numCols = 7
                ReDim ColLefts(numCols)
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)
                PointsCol = 0
                ComponentDrawCol = 2

                curCol = 0
                ColWidths(curCol) = IconColWidth
                'icon col

                curCol += 1
                SubHeads(curCol) = "Qty"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = miniColWidth

                curCol += 1
                SubHeads(curCol) = "Item"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = minColWidth

                curCol += 1
                SubHeads(curCol) = "Location"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = medplusColWidth

                curCol += 1
                'this col is for the broken box effect
                BrokenEffectCol = curCol
                SubHeads(curCol) = ""
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = 1 / 16

                curCol += 1
                SubHeads(curCol) = "Cost"
                ColAligns(curCol) = AlignHorzEnum.Left
                ColWidths(curCol) = medColWidth

                curCol += 1
                SubHeads(curCol) = "Weight"
                ColAligns(curCol) = AlignHorzEnum.Right
                ColWidths(curCol) = medColWidth

                curCol += 1
                'icon col
                ColWidths(curCol) = IconColWidth


                'build the Lefts
                ColLefts(0) = IconColLeft
                ColLefts(1) = ColLefts(0) + ColWidths(0)

                ColLefts(numCols) = LeftOffset + MaxWidth - ColWidths(numCols)

                For curCol = numCols - 1 To 3 Step -1
                    ColLefts(curCol) = ColLefts(curCol + 1) - ColWidths(curCol)
                Next

                'let the Name col use whatever space is available
                ColLefts(2) = ColLefts(1) + ColWidths(1)
                ColWidths(2) = ColLefts(3) - ColLefts(2) 'col to the right - my starting left


            Case Else
                '4 columns: icon & name & points & icon
                numCols = 3
                ReDim ColLefts(numCols)
                ReDim ColAligns(numCols)
                ReDim ColWidths(numCols)
                ReDim SubHeads(numCols)

                Select Case ItemType
                    Case Perks
                        SubHeads(1) = "Perk"
                    Case Disads
                        SubHeads(1) = "Disadvantage"
                    Case Quirks
                        SubHeads(1) = "Quirk"
                    Case Templates
                        SubHeads(1) = "Template"
                    Case Cultures
                        SubHeads(1) = "Culture"

                    Case Else
                        SubHeads(1) = "Advantage"
                End Select
                SubHeads(2) = "Points"
                ColAligns(1) = AlignHorzEnum.Left
                ColAligns(2) = AlignHorzEnum.Center
                PointsCol = 2

                'set icon col width
                ColLefts(0) = IconColLeft
                ColWidths(0) = IconColWidth

                'set rightmost col width and left side - Other Icon col
                ColLefts(3) = LeftOffset + MaxWidth - IconColWidth
                ColWidths(3) = IconColWidth

                'set next col to the left col width and left side - Points col
                ColLefts(2) = ColLefts(3) - ptsColWidth
                ColWidths(2) = ptsColWidth

                'get name col width and left side - Name col
                ColLefts(1) = ColLefts(0) + ColWidths(0)
                ColWidths(1) = ColLefts(2) - ColLefts(1) 'col to the right - my starting left
        End Select

    End Sub
    Private Function PrintSubHeads(LeftSide As Double, TopSide As Double, Optional ByVal LineHeight As Double = 0.1875) As Double
        Dim curCol As Integer = 0
        Dim curTop As Double = 0

        Dim myLeft As Double = 0
        Dim myWidth As Double = 0
        Dim myValue As String = ""
        'Dim myHeight As Double = 0.1875
        Dim myAlign As AlignHorzEnum

        Dim tmpStyle As Style
        Dim rText As RenderText

        Dim ld As New C1.C1Preview.LineDef(0.01, Color.LightGray, Drawing2D.DashStyle.Solid)
        Dim ldLine As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Dot)

        curTop = TopSide

        'Print the column values 
        For curCol = 0 To numCols

            myValue = SubHeads(curCol)
            myAlign = ColAligns(curCol)
            myLeft = ColLefts(curCol)
            myWidth = ColWidths(curCol)

            'MyPrintDoc.RenderDirectLine(boxLeft, curTop, boxLeft + boxWidth, curTop, ld)
            Select Case myAlign
                Case AlignHorzEnum.Center
                    tmpStyle = MyStyles("SheetTextCenterBoldStyle")
                Case AlignHorzEnum.Right
                    tmpStyle = MyStyles("SheetTextRightBoldStyle")
                Case Else
                    tmpStyle = MyStyles("SheetTextLeftBoldStyle")
            End Select

            rText = New RenderText(myValue, tmpStyle)
            MyPrintDoc.RenderDirect(myLeft, curTop, rText, myWidth, LineHeight)
        Next

        curTop = curTop + LineHeight
        '*****

        Return curTop
    End Function
    Private Function PrintFields(ItemType As Integer, LeftSide As Double, TopSide As Double, MaxHeight As Double, Optional ByVal StartFrom As Integer = 0, Optional ByVal PrintingWeapons As Boolean = False, Optional ByVal WeaponType As Integer = 0) As Double
        '*****
        '* Draw all our traits
        '*
        Dim i, FieldIndex As Integer

        Dim curCol As Integer = 0
        Dim curTop As Double = 0

        Dim pv1, pv2, ph1, ph2 As PointD
        Dim IndentStepSize As Double = 0.125
        Dim BottomEdge As Double = TopSide + MaxHeight

        Dim IsPlaceholder As Boolean

        Dim curField As clsDisplayField
        Dim curItem As GCATrait

        Dim myLeft As Double = 0
        Dim myWidth As Double = 0
        Dim myValue As String = ""
        Dim myHeight As Double = 0
        Dim myAlign As AlignHorzEnum

        Dim Values(numCols) As String
        Dim tmpStyle As Style
        Dim rText As RenderText

        Dim ld As New C1.C1Preview.LineDef(0.01, Color.LightGray, Drawing2D.DashStyle.Solid)
        Dim ldLine As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Dot)

        'ComponentDrawCol
        'Dim ComponentLeft As Double = LeftSide + IconColWidth
        Dim ComponentLeft = ColLefts(ComponentDrawCol - 1) + ColWidths(ComponentDrawCol - 1)
        Dim ldComponentLine As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        curTop = TopSide

        For FieldIndex = StartFrom To Fields.Count - 1
            curField = Fields(FieldIndex)
            curItem = curField.Data

            myHeight = curField.Height

            'NEW TESTING 2015 09 08
            IsPlaceholder = False
            If curItem.ItemType = TraitTypes.None Then IsPlaceholder = True
            'END NEW

            If curTop + myHeight > BottomEdge Then
                'won't fit, get out

                HasOverflow = True
                If PrintingWeapons Then
                    HasOverflowWeapons = True
                    WeaponOverflowFrom(WeaponType) = FieldIndex
                Else
                    OverflowFrom(ItemType) = FieldIndex
                End If

                Return curTop
            End If

            'Adjust text brush for special cases.
            'Fields for items that are normally hidden
            'or such, and are shown due to options settings,
            'should have a property for same added to Fields
            'so we can adjust their text brush here.
            'If Fields(curTrait).Hidden Then
            '    curTextBrush = Brushes.LightSlateGray
            '    If MyTraits.Item(curTrait).BonusListItemsCount > 0 Then
            '        curTextBrush = Brushes.LightBlue
            '    End If
            'Else
            'If curItem.BonusListItemsCount > 0 Then
            '    curTextBrush = Brushes.Blue
            'End If
            'End If


            'Shading if AltLine = True 
            If curItem.TagItem("highlight") <> "" Then
                MyPrintDoc.RenderDirectRectangle(LeftSide, curTop, MaxWidth, myHeight, ld, Brushes.Yellow)
            Else
                If curField.Alt And ShadeAltLines Then
                    MyPrintDoc.RenderDirectRectangle(LeftSide, curTop, MaxWidth, myHeight, ld, Brushes.LightGray)
                End If
            End If
            If LineBefore Then
                MyPrintDoc.RenderDirectLine(LeftSide, curTop, LeftSide + MaxWidth, curTop, ldLine)
            End If

            'Draw Component Indent Lines
            If curField.IsComponent > 0 Then
                'horizontal line, from middle of column to right edge.
                ph1 = New PointF(ComponentLeft + (IndentStepSize * curField.IsComponent) - IndentStepSize / 2, curTop + myHeight / 2)
                ph2 = New PointF(ComponentLeft + (IndentStepSize * curField.IsComponent), ph1.Y)

                'vertical line, from top of my field to bottom, or half-way to bottom for last component
                pv1 = New PointD(ph1.X, curTop)
                If curField.IsLastComponent Then
                    pv2 = New PointD(ph1.X, ph1.Y)
                Else
                    pv2 = New PointD(ph1.X, curTop + myHeight)
                End If

                'horizontal line
                MyPrintDoc.RenderDirectLine(ph1.X, ph1.Y, ph2.X, ph2.Y, ldComponentLine)
                'vertical line
                MyPrintDoc.RenderDirectLine(pv1.X, pv1.Y, pv2.X, pv2.Y, ldComponentLine)

                'draw outside vertical connector lines
                For i = 1 To curField.IsComponent - 1 'we already draw the .IsComponent bit above
                    If curField.OutsideComponentLinesToDraw And (2 ^ i) Then
                        'vertical line, from top to bottom of my field
                        pv1 = New PointD(ComponentLeft + (IndentStepSize * i) - IndentStepSize / 2, curTop)
                        pv2 = New PointD(pv1.X, curTop + myHeight)
                        MyPrintDoc.RenderDirectLine(pv1.X, pv1.Y, pv2.X, pv2.Y, ldComponentLine)
                    End If
                Next
            End If


            'Print the column values 
            'skipping col 0 (and last col) because it's the Icons, drawn below
            For curCol = 0 To numCols
                'set alignment
                'Select Case curCol
                '    Case 1
                '        tf.Alignment = StringAlignment.Near
                '    Case Else
                '        tf.Alignment = StringAlignment.Far
                'End Select

                myValue = curField.Values(curCol)
                myAlign = ColAligns(curCol)

                If curField.IsCombo Then
                    myLeft = AltColLefts(curCol)
                    myWidth = AltColWidths(curCol)
                Else
                    myLeft = ColLefts(curCol)
                    myWidth = ColWidths(curCol)
                End If

                If curCol = ComponentDrawCol AndAlso curField.IsComponent > 0 Then
                    myLeft = myLeft + (IndentStepSize * curField.IsComponent)
                    myWidth = myWidth - (IndentStepSize * curField.IsComponent)
                End If

                If curCol = PointsCol AndAlso PointsCol <> 0 Then
                    If Not IsPlaceholder Then 'NEW TESTING 2015 09 08 just the enclosing IF
                        MyPrintDoc.RenderDirectText(myLeft, curTop, "[", myWidth, MyStyles("SheetTextLeftStyle").Font, MyOptions.Value("SheetTextColor"), MyStyles("SheetTextLeftStyle").TextAlignHorz)
                        MyPrintDoc.RenderDirectText(myLeft, curTop, "]", myWidth, MyStyles("SheetTextRightStyle").Font, MyOptions.Value("SheetTextColor"), MyStyles("SheetTextRightStyle").TextAlignHorz)
                    End If
                End If

                'MyPrintDoc.RenderDirectLine(boxLeft, curTop, boxLeft + boxWidth, curTop, ld)
                Select Case myAlign
                    Case AlignHorzEnum.Center
                        tmpStyle = MyStyles("UserTextCenterStyle")
                    Case AlignHorzEnum.Right
                        tmpStyle = MyStyles("UserTextRightStyle")
                    Case Else
                        tmpStyle = MyStyles("UserTextLeftStyle")
                End Select
                'NEW TESTING 2015 09 08
                If IsPlaceholder Then
                    Select Case myAlign
                        Case AlignHorzEnum.Center
                            tmpStyle = MyStyles("UserTextCenterStyleBold")
                        Case AlignHorzEnum.Right
                            tmpStyle = MyStyles("UserTextRightStyleBold")
                        Case Else
                            tmpStyle = MyStyles("UserTextLeftStyleBold")
                    End Select
                End If
                'END NEW

                rText = New RenderText(myValue, tmpStyle)
                MyPrintDoc.RenderDirect(myLeft, curTop, rText, myWidth, myHeight)
            Next


            'NEW 2015 05 25
            'Allow for printing included bonuses.
            If curField.BonusText <> "" Then
                tmpStyle = MyStyles("UserSmallTextStyle")

                myValue = curField.BonusText
                myWidth = curField.BonusWidth
                myLeft = curField.BonusLeft

                System.Diagnostics.Debug.Print("myWidth = " & myWidth & " myLeft = " & myLeft)

                rText = New RenderText(myValue, tmpStyle)
                MyPrintDoc.RenderDirect(myLeft, curTop + curField.BonusTop, rText, myWidth, myHeight) ' - curField.BonusTop)
            End If
            'END NEW


            curTop = curTop + myHeight
        Next
        '*****

        'everything printed fine
        If PrintingWeapons Then
            WeaponOverflowFrom(WeaponType) = -1
        Else
            OverflowFrom(ItemType) = -1
        End If

        Return curTop
    End Function
    Private Function PrintGrimoireFields(ItemType As Integer, LeftSide As Double, TopSide As Double, MaxHeight As Double, Optional ByVal StartFrom As Integer = 0) As Double
        '*****
        '* Draw all our traits
        '*
        Dim FieldIndex As Integer

        Dim curCol As Integer = 0
        Dim curTop As Double = 0

        Dim IndentStepSize As Double = 0.125
        Dim BottomEdge As Double = TopSide + MaxHeight


        Dim curField As clsDisplayField
        Dim curItem As GCATrait

        Dim myLeft As Double = 0
        Dim myWidth As Double = 0
        Dim myValue As String = ""
        Dim myHeight As Double = 0
        Dim myAlign As AlignHorzEnum

        Dim Values(numCols) As String
        Dim tmpStyle As Style
        Dim rText As RenderText

        Dim ld As New C1.C1Preview.LineDef(0.01, Color.LightGray, Drawing2D.DashStyle.Solid)
        Dim ldLine As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Dot)

        'ComponentDrawCol
        Dim ComponentLeft = ColLefts(ComponentDrawCol - 1) + ColWidths(ComponentDrawCol - 1)
        Dim ldComponentLine As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        curTop = TopSide

        For FieldIndex = StartFrom To Fields.Count - 1
            curField = Fields(FieldIndex)
            curItem = curField.Data

            myHeight = curField.Height

            If curTop + myHeight > BottomEdge Then
                'won't fit, get out
                GrimoireOverflowFrom = FieldIndex

                Return curTop
            End If

            'Shading if AltLine = True 
            If curItem.TagItem("highlight") <> "" Then
                MyPrintDoc.RenderDirectRectangle(LeftSide, curTop, MaxWidth, myHeight, ld, Brushes.Yellow)
            Else
                If curField.Alt And ShadeAltLines Then
                    MyPrintDoc.RenderDirectRectangle(LeftSide, curTop, MaxWidth, myHeight, ld, Brushes.LightGray)
                End If
            End If
            If LineBefore Then
                MyPrintDoc.RenderDirectLine(LeftSide, curTop, LeftSide + MaxWidth, curTop, ldLine)
            End If

            'Print the column values 
            'skipping col 0 (and last col) because it's the Icons, drawn below
            For curCol = 0 To numCols
                myValue = curField.Values(curCol)
                myAlign = ColAligns(curCol)

                If curField.IsCombo Then
                    myLeft = AltColLefts(curCol)
                    myWidth = AltColWidths(curCol)
                Else
                    myLeft = ColLefts(curCol)
                    myWidth = ColWidths(curCol)
                End If

                If curCol = PointsCol AndAlso PointsCol <> 0 Then
                    MyPrintDoc.RenderDirectText(myLeft, curTop, "[", myWidth, MyStyles("SheetTextLeftStyle").Font, MyOptions.Value("SheetTextColor"), MyStyles("SheetTextLeftStyle").TextAlignHorz)
                    MyPrintDoc.RenderDirectText(myLeft, curTop, "]", myWidth, MyStyles("SheetTextRightStyle").Font, MyOptions.Value("SheetTextColor"), MyStyles("SheetTextRightStyle").TextAlignHorz)
                End If

                Select Case myAlign
                    Case AlignHorzEnum.Center
                        tmpStyle = MyStyles("SheetTextCenterStyle")
                    Case AlignHorzEnum.Right
                        tmpStyle = MyStyles("SheetTextRightStyle")
                    Case Else
                        tmpStyle = MyStyles("SheetTextLeftStyle")
                End Select

                rText = New RenderText(myValue, tmpStyle)
                MyPrintDoc.RenderDirect(myLeft, curTop, rText, myWidth, myHeight)
            Next

            curTop = curTop + myHeight
        Next
        '*****

        'everything printed fine
        GrimoireOverflowFrom = -1

        Return curTop
    End Function



    Private Function PrintSectionProtection(LeftSide As Double, TopSide As Double) As Double
        Dim curTop As Double = 0
        Dim boxWidth, boxHeight, boxLeft, boxTop As Double
        Dim maxHeight As Double
        Dim ld As New C1.C1Preview.LineDef(0.01, MyOptions.Value("SheetTextColor"), Drawing2D.DashStyle.Solid)

        boxLeft = LeftSide
        boxWidth = PageWidth - MarginRight - boxLeft
        boxTop = TopSide
        boxHeight = PageHeight - MarginBottom - boxTop - FooterHeight


        curTop = boxTop
        MyPrintDoc.RenderDirectText(boxLeft + 1 / 32, curTop, "PROTECTION", boxWidth, New Font(MyPrintDoc.Style.FontName, MyPrintDoc.Style.FontSize, FontStyle.Bold), MyOptions.Value("SheetTextColor"), AlignHorzEnum.Left)
        curTop = curTop + 0.1875
        maxHeight = boxHeight - (curTop - boxTop)

        'print contents
        Dim ri As New RenderImage

        Dim Settings As New ProtectionPaperDollSettings
        Settings.Character = MyChar
        Settings.TextFont = New Font(MyPrintDoc.Style.Font.Name, 14)
        Settings.TextColor = Color.Black
        Settings.TextColorAlt = Color.Black
        Settings.ShadeColor = Color.LightGray
        Settings.BackColor = Color.White
        Settings.BorderColor = Color.Gray

        'ri.Image = GetProtectionPaperDoll(Settings)
        Dim image As Bitmap = GetProtectionPaperDoll(Settings)

        Dim ia As ImageAlign
        ia.AlignHorz = ImageAlignHorzEnum.Center
        ia.AlignVert = ImageAlignVertEnum.Center
        ia.BestFit = True
        ia.KeepAspectRatio = True

        MyPrintDoc.RenderDirectImage(boxLeft, curTop, image, boxWidth, maxHeight, ia)

        MyPrintDoc.RenderDirectRectangle(boxLeft, boxTop, boxWidth, boxHeight, ld)

        Return curTop + boxHeight
    End Function




    'This class is for tracking where in a display area a field of information is,
    'or for allowing us to turn a mouse click into the item clicked upon
    Public Class clsDisplayField
        '*
        '* Variables
        '*
        Public FieldObjects As Collection
        Public Values As ArrayList 'the value for each column of data in the display area

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

        'NEW 2015 05 25
        'To allow for bonus text inline with items.
        Public Property BonusTop As Double = 0
        Public Property BonusLeft As Double = 0
        Public Property BonusWidth As Double = 0
        Public Property BonusText As String = ""
        'END NEW

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
