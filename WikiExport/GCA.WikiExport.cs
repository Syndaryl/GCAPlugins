namespace GCA.TextExport
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Linq;
    using System.Reflection;
    using System.Xml.Serialization;
    using GCA5.Interfaces;
    using GCA5Engine;
    using System.Windows.Forms;

    public sealed class WikiExporter : GCA5.Interfaces.IExportSheet
    {
        #region CTORS
        public WikiExporter()
        {
            MyOptions = new SheetOptionsManager(Name);
        }

        const string assemblyName = "GCA.WikiExport";
        const string settingsFileName = @"wikiexportoptions.xml";
        #endregion CTORS

        #region Enums
        enum CharacterSeparators
        {
            DoNothing,
            PrintStars,
            PrintEquals,
            PrintHyphens,
            PrintHTMLHR
        }

        enum NameStyle
        {
            DoNothing,
            MakeBold,
            MakeItalic
        }
        #endregion

        #region fields
        static readonly string OwnedItemText = "* = item is owned by another, its point value is included in the other item.";
        #endregion fields

        #region InterfaceImplementation
        #region InterfaceProperties
        public string Name
        {
            get
            {
                return "MediaWiki Export";
            }
        }

        public string Description
        {
            get
            {
                return "Exports a text file suitable for pasting into the source of a MediaWiki page";
            }
        }

        public string Version
        {
            get
            {
                return "0.1.0";
            }
        }

        public SheetOptionsManager MyOptions { get; private set; }
        #endregion InterfaceProperties

        /// <summary>
        /// This is the routine where all the Options we want to use are created,
        /// and where the UI for the Preferences dialog is filled out.
        /// 
        /// This is equivalent to CharacterSheetOptions from previous implementations
        /// </summary>
        /// <param name="Options"></param>
        void IBasicPlugin.CreateOptions(SheetOptionsManager Options)
        {
            //string optionsPath = @"D:\Documents\Visual Studio 2015\Projects\GCAPlugins\wikiexportoptions.xml";
            try
            {
                var domain = AppDomain.CurrentDomain;
                domain.Load(assemblyName);
                var listOfAssemblies = new List<Assembly>(domain.GetAssemblies());

#if false
                string list = "";
                foreach (var item in listOfAssemblies)
                {
                    list += item.FullName + "\n";
                }
                System.Windows.Forms.MessageBox.Show(list, assemblyName); 
#endif

                var optionsPath = Path.GetDirectoryName(listOfAssemblies.FirstOrDefault(x => x.FullName.StartsWith(assemblyName, StringComparison.Ordinal)).Location);
                optionsPath = optionsPath + Path.DirectorySeparatorChar + settingsFileName;
                Console.WriteLine(String.Format("Options Path is {0}", optionsPath));
                extractOptionsFromFile(optionsPath, Options);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString() + "\n" + e.InnerException.ToString(), assemblyName);
            }
        }

        void extractOptionsFromFile(string optionsPath, SheetOptionsManager options)
        {
            var optionsSerializer = new XmlSerializer(typeof(SheetOption[]));

            try
            {
                var inFile = new FileStream(optionsPath, FileMode.Open);
                var optionsArray = new List<SheetOption>((SheetOption[])optionsSerializer.Deserialize(inFile));

                var descFormatHeader = new SheetOptionDisplayFormat();
                descFormatHeader.BackColor = SystemColors.Info;
                descFormatHeader.CaptionLocalBackColor = SystemColors.Info;

                var descFormatBody = new SheetOptionDisplayFormat();
                //descFormatBody.BackColor = SystemColors.;
                //descFormatBody.CaptionLocalBackColor = SystemColors.Info;

                foreach (var opt in optionsArray)
                {
                    opt.DisplayFormat = opt.Type == OptionType.Header ? descFormatHeader : descFormatBody;
                    opt.UserPrompt = opt.UserPrompt.Replace(@"$Name", Name).Replace(@"$Version", Version);
                    options.AddOption(opt);
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.ToString() + "\n" + e.InnerException.ToString(), assemblyName);
            }
            #region disabledcode
            //var inFile = new FileStream(optionsPath, FileMode.Open);
            //var optionsArray = (SheetOption[])tilesetSerializer.Deserialize(inFile);
            //* Description block at top *
            //ok = insertOption(Options, descFormat, "Header_Description", Name + " " + this.Version, GCA5Engine.OptionType.Header);
            //var outFile = new FileStream(optionsPath + ".updated", FileMode.Create);
            //optionsSerializer.Serialize(outFile, optionsArray.ToArray());

            //bool ok;

            //var descFormat = new SheetOptionDisplayFormat();
            //descFormat.BackColor = SystemColors.Info;
            //descFormat.CaptionLocalBackColor = SystemColors.Info;
            //var newOption = new GCA5Engine.SheetOption();
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "Header_Description";
            //newOption.Type = GCA5Engine.OptionType.Header;
            //newOption.UserPrompt = Name + " " + this.Version;
            //newOption.DisplayFormat = descFormat;
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            //ok = insertOption(Options, descFormat, "Description", this.Description);
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "Description";
            //newOption.Type = GCA5Engine.OptionType.Caption;
            //newOption.UserPrompt = this.Description;
            //newOption.DisplayFormat = descFormat;
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            ////******************************
            ////* Characters 
            ////******************************
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "Header_Characters";
            //newOption.Type = GCA5Engine.OptionType.Header;
            //newOption.UserPrompt = "Printing Characters";
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            ////NOTE: Because List is now a 0-based Array, the number of the 
            ////DefaultValue and the selected Value is 0-based!
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "OutputCharacters";
            //newOption.Type = GCA5Engine.OptionType.ListNumber;
            //newOption.UserPrompt = "When exporting, how do you want to handle exporting when multiple characters are loaded?";
            //newOption.DefaultValue = 0;// 'first item;
            //newOption.List = new string[] { "Export just the current character", "Export all the characters to the file", "Always ask me what to do" };
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);
            ////AlwaysAskMe = 2;

            ////NOTE: Because List is now a 0-based Array, the number of the 
            ////DefaultValue and the selected Value is 0-based!
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "CharacterSeparator";
            //newOption.Type = GCA5Engine.OptionType.ListNumber;
            //newOption.UserPrompt = "Please select how you'd like to mark the break between characters when printing multiple characters to the file.";
            //newOption.DefaultValue = 1;// 'second item;
            //newOption.List = new string[] { "Do nothing", "Print a line of *", "Print a line of =", "Print a line of -", "Use HTML to indicate a horizontal rule" };
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);


            ////******************************
            ////* Included Sections 
            ////******************************
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "Header_TextBlocks";
            //newOption.Type = GCA5Engine.OptionType.Header;
            //newOption.UserPrompt = "Sections to Include";
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "ShowMovementBlock";
            //newOption.Type = GCA5Engine.OptionType.YesNo;
            //newOption.UserPrompt = "Include a block showing all the Environment Move rates?";
            //newOption.DefaultValue = true;
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "ShowMovementZero";
            //newOption.Type = GCA5Engine.OptionType.YesNo;
            //newOption.UserPrompt = "When printing the block above, include movement rates even when they're at zero?";
            //newOption.DefaultValue = false;
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "ShowAllAdditionalStats";
            //newOption.Type = GCA5Engine.OptionType.YesNo;
            //newOption.UserPrompt = "Include a block showing all additional attributes that aren't hidden?";
            //newOption.DefaultValue = false;
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "ShowAdditionalStatsAtZero";
            //newOption.Type = GCA5Engine.OptionType.YesNo;
            //newOption.UserPrompt = "When printing the block above, include attributes even when they're at zero?";
            //newOption.DefaultValue = false;
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            ////******************************
            ////* Other Stuff 
            ////******************************
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "Header_Other";
            //newOption.Type = GCA5Engine.OptionType.Header;
            //newOption.UserPrompt = "Other Options";
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            ////NOTE: Because List is now a 0-based Array, the number of the 
            ////DefaultValue and the selected Value is 0-based!
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "HeadingStyle";
            //newOption.Type = GCA5Engine.OptionType.ListNumber;
            //newOption.UserPrompt = "Please select the way you'd like to differentiate section headers from their various items.";
            //newOption.DefaultValue = 1; //second item;
            //newOption.List = new string[] { "Do nothing", "Use a row of dashes under the header", "Use BBCode to mark the header as bold", "Use HTML to mark the header as bold" };
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            ////NOTE: Because List is now a 0-based Array, the number of the 
            ////DefaultValue and the selected Value is 0-based!
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "BonusLineStyle";
            //newOption.Type = GCA5Engine.OptionType.ListNumber;
            //newOption.UserPrompt = "Please select the way you'd like to differentiate bonus lines ('Includes: +X from Z') from their related items.";
            //newOption.DefaultValue = 1;// 'second item;
            //newOption.List = new string[] { "Do nothing", "Use a tab character preceding them", "Use BBCode to mark them in italic", "Use HTML to mark them in italic", "Tab character and BBCode", "Tab character and HTML" };
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);

            ////NOTE: Because List is now a 0-based Array, the number of the 
            ////DefaultValue and the selected Value is 0-based!
            //newOption = new GCA5Engine.SheetOption();
            //newOption.Name = "PointsLineStyle";
            //newOption.Type = GCA5Engine.OptionType.ListNumber;
            //newOption.UserPrompt = "Please select the way you'd like to differentiate the Point Summary line from the surrounding text.";
            //newOption.DefaultValue = 1;// 'second item;
            //newOption.List = new string[] { "Do nothing", "Print it as multiple lines with a header", "Use BBCode to mark them in bold", "Use HTML to mark them in bold", "Use BBCode to mark them in italic", "Use HTML to mark them in italic", "Use BBCode to mark them in bold and italic", "Use HTML to mark them in bold and italic" };
            //ok = Options.AddOption(newOption);
            //optionsArray.Add(newOption);
            #endregion disabledcode
        }

        /// <summary>
        /// Wrapper for creating a new SheetOption and adding it to the SheetOptionsManager.
        /// </summary>
        /// <param name="options"></param>
        /// <param name="descFormat"></param>
        /// <param name="name"></param>
        /// <param name="userPrompt"></param>
        /// <param name="type"></param>
        /// <param name="arguments"></param>
        /// <returns></returns>
        bool insertOption(SheetOptionsManager options, SheetOptionDisplayFormat descFormat, string name, string userPrompt,
            OptionType type = GCA5Engine.OptionType.Caption, params object[] arguments)
        {
            bool ok;
            GCA5Engine.SheetOption newOption;
            newOption = new GCA5Engine.SheetOption();
            newOption.Name = name;
            newOption.Type = GCA5Engine.OptionType.Header;
            newOption.UserPrompt = userPrompt;
            newOption.DisplayFormat = descFormat;
            ok = options.AddOption(newOption);
            return ok;
        }

        /// <summary>
        ///         'This is called only when a particular plug-in is loaded the first time,
        ///and before SetOptions.
        /// I don't do anything with this.
        /// </summary>
        /// <param name="Options"></param>
        void IBasicPlugin.UpgradeOptions(SheetOptionsManager Options)
        {
            // no op
        }

        public event IExportSheet.RequestRunSpecificOptionsEventHandler RequestRunSpecificOptions;

        string IExportSheet.SupportedFileTypeFilter()
        {
            return "Text files (*.txt)|*.txt";
        }

        /// <summary>
        /// This creates the export file on disk.
        /// </summary>
        /// <param name="Party"></param>
        /// <param name="TargetFilename"></param>
        /// <param name="Options"></param>
        /// <returns></returns>
        bool IExportSheet.GenerateExport(Party Party, string TargetFilename, SheetOptionsManager Options)
        {
            MyOptions = Options;

            bool printMults = (int)MyOptions.get_Value("OutputCharacters") == 1;
            try
            {
                using (var fw = new StreamWriter(TargetFilename, true))
                {
                    if (Party.Characters.Count > 1 && printMults)
                        foreach (GCACharacter pc in Party.Characters)
                        {
                            ExportCharacter(pc, fw);
                            DoCharacterSeparator((int)MyOptions.get_Value("CharacterSeparator"), fw);
                        }
                    else
                        ExportCharacter(Party.Current, fw);
                }
            }
            catch (Exception ex)
            {
                modHelperFunctions.Notify(Name + ": failed on export. " + ex.Message + "\n" + "Stack Trace: " + "\n" + ex.StackTrace, Priority.Red);
            }

            return true;
        }

        void DoCharacterSeparator(int separatorOptionChoice, StreamWriter fw)
        {
            string separator = " ";
            switch ((CharacterSeparators)separatorOptionChoice)
            {
                case CharacterSeparators.DoNothing:
                    break;
                case CharacterSeparators.PrintStars:
                    separator = "*";
                    break;
                case CharacterSeparators.PrintEquals:
                    separator = "=";
                    break;
                case CharacterSeparators.PrintHyphens:
                    separator = "-";
                    break;
                case CharacterSeparators.PrintHTMLHR:
                    separator = "<hr/>";
                    break;
            }
            if (!Equals(separator, " ") && !Equals(separator, "<hr/>"))
            {
                fw.WriteLine(new String(separator.ToCharArray()[0], 60));
            }
            else if (Equals(separator, "<hr/>"))
            {
                fw.WriteLine(separator);
            }
            fw.WriteLine();
        }

        void ExportCharacter(GCACharacter pc, StreamWriter fw)
        {
            ExportBiography(pc, fw);
            ExportAttributes(pc, fw);
            ExportCultural(pc, fw);
            ExportReaction(pc, fw);
            ExportAdvantages(pc, fw);
            ExportDisadvantages(pc, fw);
            ExportSkills(pc, fw);
            ExportSpells(pc, fw);
            ExportMeta(pc, fw);
            ExportPointsSummary(pc, fw);
            ExportEquiment(pc, fw);
            ExportCombat(pc, fw);
            ExportLoadout(pc, fw);
        }

        void ExportLoadout(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportCombat(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportEquiment(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportPointsSummary(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportMeta(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportSpells(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportSkills(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportDisadvantages(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportAdvantages(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportReaction(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportCultural(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportAttributes(GCACharacter pc, StreamWriter fw)
        {

        }

        void ExportBiography(GCACharacter CurChar, StreamWriter fw)
        {
            fw.WriteLine(Write("Name:", CurChar.Name));
            fw.WriteLine("Name:" + CurChar.Name);
            fw.WriteLine("Player:" + CurChar.Player);
            fw.WriteLine("Race:" + CurChar.Race);
            if (CurChar.Appearance != "")
                fw.WriteLine("Appearance: " + CurChar.Appearance);

            if (CurChar.Height != "")
                fw.WriteLine("Height:" + CurChar.Height);

            if (CurChar.Weight != "")
                fw.WriteLine("Weight:" + CurChar.Weight);

            if (CurChar.Age != "")
                fw.WriteLine("Age:" + CurChar.Age);

            fw.WriteLine("");
        }

        string Write(string Label, string Value)
        {

            try
            {

                switch ((NameStyle)MyOptions.get_Value("NameStyle"))
                {
                    case NameStyle.DoNothing:
                         return String.Format("{0} {1}<br/>", Label, Value);
                    case NameStyle.MakeBold:
                        return String.Format("'''{0}''' {1}<br/>", Label, Value);
                    case NameStyle.MakeItalic:
                        return String.Format("''{0}'' {1}<br/>", Label, Value);
                    default:
                        return String.Format("{0} {1}<br/>", Label, Value);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show( String.Format("Write({0}, {1})\nMessage: {2}\nStacktrace: {3}\nInner: {4}", Label, Value, e.Message , e.StackTrace, e.InnerException), "Error in Write()");
                return String.Format("Write({0}, {1}) error", Label, Value);
            }
        }


        #endregion InterfaceImplementation
    }

}
