/*
The MIT License(MIT)
Copyright(c) 2015 Emily Smirle
Permission is hereby granted, free of charge, to any person obtaining a copy of
this software and associated documentation files (the "Software"), to deal in 
the Software without restriction, including without limitation the rights to 
use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies 
of the Software, and to permit persons to whom the Software is furnished to do 
so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all 
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR 
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING 
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER 
DEALINGS IN THE SOFTWARE.
*/

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
    using System.Text;

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

        enum SkillType
        {
            Skill,
            Technique,
            Combo
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
        public List<GCATrait> Traits { get; private set; }
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
                using (var fw = new GCAWriter(TargetFilename, false, MyOptions))
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


        public int PreferredFilterIndex()
        {
            return 0;
        }

        public bool PreviewOptions(SheetOptionsManager Options)
        {
            return true;
        }
        #endregion InterfaceImplementation


        #region Exporters
        void ExportCharacter(GCACharacter pc, GCAWriter fw)
        {
            this.Traits = new List<GCATrait>();
            foreach (GCATrait item in pc.Items)
            {
                this.Traits.Add(item);
            }

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

        void ExportLoadout(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteLine();
        }

        void ExportCombat(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteLine();
        }

        void ExportEquiment(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteLine();
        }

        void ExportPointsSummary(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteLine();
        }

        void ExportMeta(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Metatraits [" + pc.get_Cost(modConstants.Templates) + "]");
            fw.WriteLine();
        }

        void ExportSpells(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Spells [" + pc.get_Cost(modConstants.Spells) + "]");
            fw.WriteLine();
        }

        void ExportSkills(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Skills [" + pc.get_Cost(modConstants.Skills) + "]");
            foreach (var item in ComplexListTrait(TraitTypes.Skills, fw))
            {
                fw.WriteLine(item);
            }
            fw.WriteLine();
        }

        void ExportDisadvantages(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Disadvantages [" + pc.get_Cost(modConstants.Disadvantages) + "]");
            foreach (var item in ComplexListTrait(TraitTypes.Disadvantages, fw))
            {
                fw.WriteLine(item);
            }
            fw.WriteLine();
        }

        void ExportAdvantages(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Advantages [" + pc.get_Cost(modConstants.Advantages) + "]");
            foreach (var item in ComplexListTrait(TraitTypes.Advantages, fw))
            {
                fw.WriteLine(item);
            }
            fw.WriteLine();
        }

        void ExportReaction(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Reaction Modifiers");
            fw.WriteLine();
        }

        void ExportCultural(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Cultural Background");
            var label = "TL:";
            var curItem = pc.ItemByNameAndExt("Tech Level", modConstants.Stats);
            if (curItem != null)
            {
                var buffer = curItem.DisplayScore;
                if (curItem.Points != 0)
                {
                    buffer = string.Format("{0} [{1}]", buffer, curItem.Points);
                }
                fw.WriteTrait(label, buffer);
            }
            fw.WriteLine();

            if (pc.get_Count(modConstants.Cultures) > 0)
            {
                label = "Cultures: ";
                var buffer = SimpleStringTrait(TraitTypes.Cultures);
                fw.WriteTrait(label, buffer);
            }
            if (pc.get_Count(modConstants.Languages) > 0)
            {
                label = "Languages: ";
                var buffer = SimpleStringTrait(TraitTypes.Languages);
                fw.WriteTrait(label, buffer);
            }
            fw.WriteLine();
        }

        void ExportAttributes(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Attributes [" + pc.get_Cost(modConstants.Stats) + "]");

            var StatNames = new List<string> {
                "ST",
                "DX",
                "IQ",
                "HT"};
            foreach (var item in ComplexListAttributes(StatNames, fw))
            {
                fw.WriteLine(item);
            }
            fw.WriteLine();
            StatNames.Clear();

            StatNames.AddRange(new string[] { 
                "Hit Points",
                "Will",
                "Perception",
                "Fatigue Points"}
            );
            foreach (var item in ComplexListAttributes(StatNames, fw))
            {
                fw.WriteLine(item);
            }
            fw.WriteLine();
        }

        void ExportBiography(GCACharacter CurChar, GCAWriter fw)
        {
            fw.WriteTrait("Name:", CurChar.Name);
            fw.WriteTrait("Player:", CurChar.Player);
            fw.WriteTrait("Race:", CurChar.Race);
            if (CurChar.Appearance != "")
                fw.WriteTrait("Appearance: ", CurChar.Appearance);

            if (CurChar.Height != "")
                fw.WriteTrait("Height:", CurChar.Height);

            if (CurChar.Weight != "")
                fw.WriteTrait("Weight:", CurChar.Weight);

            if (CurChar.Age != "")
                fw.WriteTrait("Age:", CurChar.Age);

            fw.WriteLine("");
        }
        #endregion Exporters

        #region Formatters
        void DoCharacterSeparator(int separatorOptionChoice, GCAWriter fw)
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


        /// <summary>
        /// Generates a string listing of the specified trait type, visible only, by name and cost only. 0 cost traits do not have a cost listed.
        /// </summary>
        /// <param name="traitType">Type of trait to harvest.</param>
        /// <returns>A comma separated, period-terminated list.</returns>
        string SimpleStringTrait(TraitTypes traitType)
        {
            var result = from trait in Traits
                         where trait.ItemType == traitType
                         where trait.get_TagItem("hide").Equals("")
                         select trait.Points != 0 ? String.Format("{0} [{1}]", trait.Name, trait.Points) : trait.Name;

            return String.Join(", ", result) + ".";
        }

        /// <summary>
        /// Generates a SJG-formatted list of the specified trait type, visible traits only.
        /// This includes level, name extensions, and modifiers, but not "effective skill".
        /// </summary>
        /// <param name="traitType">Type of trait to harvest.</param>
        /// <returns>A semicolon separated, period-terminated list.</returns>
        string ComplexStringAdsDisads(TraitTypes traitType, GCAWriter fw)
        {
            var result = from trait in Traits
                         where trait.ItemType == traitType
                         where trait.get_TagItem("hide").Equals("")
                         select trait;

            return String.Join("; ", result.Select(x => AdvantageFormatter(x, fw))) + ".";
        }

        IEnumerable<string> ComplexListTrait(TraitTypes traitType, GCAWriter fw)
        {
            var result = from trait in Traits
                         where trait.ItemType == traitType
                         where trait.get_TagItem("hide").Equals("")
                         select FormatTrait(trait, fw);
            return result;
        }


        IEnumerable<string> ComplexListAttributes(List<string> attributeList, GCAWriter fw)
        {
            var result = from trait in Traits
                         where trait.ItemType == TraitTypes.Attributes
                         where !trait.get_TagItem("mainwin").Equals("")
                         where trait.get_TagItem("display").Equals("")
                         where trait.get_TagItem("hide").Equals("")
                         orderby trait.get_TagItem("mainwin")
                         select FormatTrait(trait, fw);
            return result;
        }
        delegate string TraitFormatter(GCATrait trait, GCAWriter fw);

        string FormatTrait(GCATrait trait, GCAWriter fw)
        {
            TraitFormatter formatter = null;
            switch (trait.ItemType)
            {
                case TraitTypes.Attributes:
                    formatter = AttributeFormatter;
                    break;
                case TraitTypes.Languages:
                    formatter = AdvantageFormatter;
                    break;
                case TraitTypes.Cultures:
                    formatter = AdvantageFormatter;
                    break;
                case TraitTypes.Advantages:
                    formatter = AdvantageFormatter;
                    break;
                case TraitTypes.Perks:
                    formatter = AdvantageFormatter;
                    break;
                case TraitTypes.Disadvantages:
                    formatter = AdvantageFormatter;
                    break;
                case TraitTypes.Quirks:
                    formatter = AdvantageFormatter;
                    break;
                case TraitTypes.Skills:
                    formatter = SkillFormatter;
                    break;
                case TraitTypes.Spells:
                    formatter = SkillFormatter;
                    break;
                case TraitTypes.Equipment:
                    break;
                case TraitTypes.Templates:
                    break;
            }
            return formatter != null ? formatter(trait, fw) : "";
        }

        string AttributeFormatter(GCATrait trait, GCAWriter fw)
        {
            var builder = new StringBuilder();
            builder.Append(trait.DisplayName);

            var label = builder.ToString();

            builder.Clear();
            builder.Append(trait.DisplayScore);

            builder.AppendFormat(" [{0}]", trait.Points);
            return fw.FormatTrait(label, builder.ToString());
        }

        string AdvantageFormatter(GCATrait trait, GCAWriter fw)
        {
            var builder = new StringBuilder();
            builder.Append(trait.Name);
            if (!trait.get_TagItem("level").Equals("1") || !trait.get_TagItem("upto").Equals("") || !trait.LevelName.Equals("")) // has more than one level
            {
                builder.AppendFormat(" {0}", trait.LevelName.Equals("") ? trait.get_TagItem("level") : trait.LevelName);
            }

            var label = builder.ToString();

            builder.Clear();
            if (!trait.NameExt.Equals("") || trait.Mods.Count() > 0)
                builder.Append(" (");
            if (!trait.NameExt.Equals(""))
                builder.Append(trait.NameExt);
            if (!trait.NameExt.Equals("") && trait.Mods.Count() > 0)
                builder.Append("; ");
            if (trait.Mods.Count() > 0)
            {
                var mods = new List<GCAModifier>();
                foreach (GCAModifier item in trait.Mods)
                {
                    mods.Add(item);
                }
                builder.Append(String.Join("; ", mods.Select(x => ModifierFormatter(x, fw))));
            }
            if (!trait.NameExt.Equals("") || trait.Mods.Count() > 0)
                builder.Append(")");

            builder.AppendFormat(" [{0}]", trait.Points);

            return fw.FormatTrait(label, builder.ToString());
        }

        string ModifierFormatter(GCAModifier trait, GCAWriter fw)
        {
            var builder = new StringBuilder();
            builder.Append(trait.DisplayName.Trim());
            if (!trait.get_TagItem("level").Equals("1") || !trait.get_TagItem("upto").Equals("") || !trait.LevelName().Equals("")) // has more than one level or has named levels
            {
                builder.AppendFormat(" {0}", trait.LevelName().Equals("") ? trait.get_TagItem("level") : trait.LevelName());
            }
            builder.AppendFormat(", {0}", trait.get_TagItem("value"));
            return builder.ToString();
        }

        /// <summary>
        /// Generates a SJG-formatted list of the specified trait type, visible traits only.
        /// This includes "effective skill", name extensions, but not modifiers or "level" in the advantage sense.
        /// </summary>
        /// <param name="traitType">Type of trait to harvest.</param>
        /// <returns>A comma separated, period-terminated list.</returns>
        string ComplexStringSkillsSpells(TraitTypes traitType)
        {
            var result = from trait in Traits
                         where trait.ItemType == traitType
                         where trait.get_TagItem("hide").Equals("")
                         select trait;

            return String.Join(", ", result) + ".";
        }

        string SkillFormatter(GCATrait trait, GCAWriter fw)
        {
            var builder = new StringBuilder();
            builder.AppendFormat("{0} {1}:", trait.DisplayName, trait.get_TagItem("type"));
            var label = builder.ToString();

            builder.Clear();
            switch ((SkillType)int.Parse(trait.get_TagItem("sd")))
            {
                case SkillType.Skill:
                    string stepOff = trait.get_TagItem("stepoff");
                    if (!stepOff.Equals(string.Empty))
                    {
                        builder.Append(stepOff);
                        string step = trait.get_TagItem("step");
                        if (!step.Equals(string.Empty))
                            builder.Append(step);
                        else
                            builder.Append("?");
                    }
                    else
                    {
                        builder.Append("?+?");
                    }
                    builder.AppendFormat(" - {0}", trait.get_TagItem("level"));
                    break;
                case SkillType.Technique:
                    builder.AppendFormat(" - {0}", trait.get_TagItem("level"));
                    break;
                case SkillType.Combo:
                    builder.AppendFormat(" - {0}", trait.get_TagItem("combolevel"));
                    break;
                default:
                    break;
            }
            builder.AppendFormat(" [{0}]", trait.Points);

            return fw.FormatTrait(label, builder.ToString());
        }
        #endregion Formatters
    }

    class GCAWriter : StreamWriter
    {

        enum GeneralStyles
        {
            DoNothing,
            MakeBold,
            MakeItalic,
            MakeBoldItalic
        }

        enum HeaderStyles
        {
            DoNothing,
            MinorWikiHeader,
            Bold
        }

        public GCAWriter(string path, bool append, SheetOptionsManager options) : base(path, append)
        {
            MyOptions = options;
        }

        public SheetOptionsManager MyOptions { get; private set; }

        public void WriteTrait(string label, string value)
        {
            WriteLine(FormatTrait(label, value));
        }
        public void WriteHeader(string Header)
        {
            WriteLine(FormatHeader(Header));
        }

        internal string FormatHeader(string Header)
        {
            try
            {
                switch ((HeaderStyles)MyOptions.get_Value("HeadingStyle"))
                {
                    case HeaderStyles.DoNothing:
                        return string.Format("{0}<br/>", Header);
                    case HeaderStyles.MinorWikiHeader:
                        return string.Format("==={0}===", Header);
                    case HeaderStyles.Bold:
                        return string.Format("'''{0}'''<br/>", Header);
                    default:
                        return string.Format("{0}<br/>", Header);
                }
            }
            catch (Exception)
            {
                return string.Format("{0}<br/>", Header);
            }
        }

        internal string FormatTrait(string label, string value)
        {
            try
            {
                switch ((GeneralStyles)MyOptions.get_Value("TraitNameStyle"))
                {
                    case GeneralStyles.DoNothing:
                        return String.Format("{0} {1}<br/>", label, value);
                    case GeneralStyles.MakeBold:
                        return String.Format("'''{0}''' {1}<br/>", label, value);
                    case GeneralStyles.MakeItalic:
                        return String.Format("''{0}'' {1}<br/>", label, value);
                    case GeneralStyles.MakeBoldItalic:
                        return String.Format("'''''{0}''''' {1}<br/>", label, value);
                    default:
                        return String.Format("{0} {1}<br/>", label, value);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(String.Format("FormatTrait({0}, {1})\nMessage: {2}\nStacktrace: {3}\nInner: {4}", label, value, e.Message, e.StackTrace, e.InnerException), "Error in FormatTrait()");
                return String.Format("FormatTrait({0}, {1}) error", label, value);
            }
        }
    }
}
