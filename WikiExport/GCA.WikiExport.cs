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

namespace GCA.WikiExport
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
    using GCA.Export;
    using System.Globalization;
    using System.Diagnostics.Contracts;

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

        // Suppressed "not used" warning as this is something I'm going to be implementing in the future, and by then I'll totally forget the standard wording :P
        //  Amusingly, not code left over from the past, but code left over from the future.
#pragma warning disable 0169
        static readonly string ownedItemText = "* = item is owned by another, its point value is included in the other item.";
#pragma warning restore 0169
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
                return "1.2.0";
            }
        }

        public SheetOptionsManager MyOptions { get; private set; }
        public List<GCATrait> Traits { get; private set; }

        public static string OwnedItemText
        {
            get
            {
                return ownedItemText;
            }
        }

        public bool Metric { get; private set; }
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

                var optionsPath = Path.GetDirectoryName(listOfAssemblies.FirstOrDefault(x => x.FullName.StartsWith(assemblyName, StringComparison.Ordinal)).Location);
                optionsPath = optionsPath + Path.DirectorySeparatorChar + settingsFileName;
                Console.WriteLine(string.Format("Options Path is {0}", optionsPath));
                extractOptionsFromFile(optionsPath, Options);
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e + "\n" + e.InnerException, assemblyName);
            }
        }

        void extractOptionsFromFile(string optionsPath, SheetOptionsManager options)
        {
            var optionsSerializer = new XmlSerializer(typeof(SheetOption[]));

            try
            {
                var optionsArray = new List<SheetOption>();
                using (var inFile = new FileStream(optionsPath, FileMode.Open))
                {
                    optionsArray.AddRange((SheetOption[])optionsSerializer.Deserialize(inFile));
                }

                var descFormatHeader = new SheetOptionDisplayFormat();
                descFormatHeader.BackColor = SystemColors.Info;
                descFormatHeader.CaptionLocalBackColor = SystemColors.Info;

                var descFormatBody = new SheetOptionDisplayFormat();

                foreach (var opt in optionsArray)
                {
                    opt.DisplayFormat = opt.Type == OptionType.Header ? descFormatHeader : descFormatBody;
                    opt.UserPrompt = opt.UserPrompt.Replace(@"$Name", Name).Replace(@"$Version", Version);
                    options.AddOption(opt);
                }
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e + "\n" + e.InnerException, assemblyName);
            }
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

#pragma warning disable 0169
        public event IExportSheet.RequestRunSpecificOptionsEventHandler RequestRunSpecificOptions;
#pragma warning restore 0169

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

            Metric = pc.IsMetric();

            fw.Write(Encoding.UTF8.GetPreamble());
            ExportBiography(pc, fw);
            ExportAttributes(pc, fw);
            ExportMovement(pc, fw);
            ExportCultural(pc, fw);
            ExportReaction(pc, fw);
            ExportAdvantages(pc, fw);
            ExportPerks(pc, fw);
            ExportDisadvantages(pc, fw);
            ExportQuirks(pc, fw);
            ExportSkills(pc, fw);
            ExportSpells(pc, fw);
            ExportMeta(pc, fw);
            ExportPointsSummary(pc, fw);
            ExportEquiment(pc, fw);
            ExportCombat(pc, fw);
            ExportLoadouts(pc, fw);
            ExportLongDescription(pc, fw);
            ExportNotes(pc, fw);
            ExportCampaignLog(pc, fw);
        }

        void ExportCampaignLog(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader(string.Format("Campaign Log"));
            fw.WriteCampaign(pc.Campaign);
            fw.WriteLine();
        }

        void ExportNotes(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader(string.Format("Notes"));
            fw.WriteRTF(pc.Notes);
            fw.WriteLine();
        }

        void ExportLongDescription(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader(string.Format("Description"));
            fw.WriteRTF(pc.Description);
            fw.WriteLine();
        }

        void ExportMovement(GCACharacter pc, GCAWriter fw)
        {
            var NeededStats = new List<string>();
            NeededStats.Add("Basic Lift");
            NeededStats.Add("Super-Effort Basic Lift");
            NeededStats.Add("Air Move");
            NeededStats.Add("Brachiation Move");
            NeededStats.Add("Ground Move");
            NeededStats.Add("Space Move");
            NeededStats.Add("TK Move");
            NeededStats.Add("Tunneling Move");
            NeededStats.Add("Water Move");
            NeededStats.Add("Dodge");

            var result = from moveTrait in Traits
                         where moveTrait.ItemType == TraitTypes.Stats
                         where NeededStats.Contains(moveTrait.Name)
                         where string.IsNullOrEmpty(moveTrait.get_TagItem("hide"))
                         orderby moveTrait.get_TagItem("mainwin")
                         select moveTrait;

            if (result.Count() > 0)
            {
                //var sb = new StringBuilder();
                var gridAttributes = new List<string> {
                    "Name", "Unenc.", "Light", "Medium", "Heavy", "Xtra Heavy"
                };
                fw.WriteLine("{|");
                fw.WriteLine("|+ Encumbrance");
                foreach (var label in gridAttributes)
                {
                    fw.WriteLine("! {0}", label);
                }

                foreach (var moveTrait in result)
                {
                    fw.WriteLine("|-");
                    if (moveTrait.LCaseName.Equals("basic lift") || moveTrait.LCaseName.Equals("super-effort basic lift"))
                    {
                        var unit = new UsCustomaryUnitScale(moveTrait.Score);

                        gridAttributes = new List<string> {
                            moveTrait.LCaseName.Equals("basic lift")? "Carry ("+ unit.Units +")" : "Super-Effort Carry ("+unit.Units+")"
                        };
                        foreach (var mult in new int[] { 1, 2, 3, 6, 10 })
                        {
                            gridAttributes.Add(Math.Round(moveTrait.Score * mult / unit.UnitMultiplier, 2).ToString(unit.FormatString));
                        }
                    }
                    else if (moveTrait.LCaseName.Equals("dodge"))
                    {
                        gridAttributes = new List<string> {
                            moveTrait.Name
                        };
                        // , moveTrait.Score.ToString(), (moveTrait.Score-1).ToString(),  (moveTrait.Score-2).ToString(),  (moveTrait.Score-3).ToString(),  (moveTrait.Score-4).ToString()
                        foreach (var penalty in new int[] { 0, 1, 2, 3, 4 })
                        {
                            gridAttributes.Add(Math.Floor(moveTrait.Score - penalty).ToString());
                        }
                    }
                    else
                    {
                        gridAttributes = new List<string> {
                            moveTrait.Name
                        };
                        // Math.Floor(Math.Max( moveTrait.Score, 1)).ToString(), Math.Floor(Math.Max( moveTrait.Score*0.8, 1)).ToString(),  Math.Floor(Math.Max( moveTrait.Score*0.6, 1)).ToString(),  Math.Floor(Math.Max( moveTrait.Score*0.4, 1)).ToString(),  Math.Floor(Math.Max( moveTrait.Score*0.2, 1)).ToString()

                        foreach (var mult in new double[] { 1, 0.8, 0.6, 0.4, 0.2 })
                        {
                            gridAttributes.Add(Math.Floor(Math.Max(moveTrait.Score * mult, 1)).ToString("#,##0"));
                        }
                    }

                    fw.WriteLine("| {0}", gridAttributes.First());
                    gridAttributes.RemoveAt(0);

                    foreach (var content in gridAttributes)
                    {
                        fw.WriteLine("| style=\"width:6em; text-align:center;\" | {0}", content);
                    }

                }
                fw.WriteLine("|}");
            }
            fw.WriteLine();
        }

        public class UsCustomaryUnitScale
        {
            double value;
            string units;
            double unitMultiplier;
            string formatString;
            string unitsLong;

            public string Units
            {
                get
                {
                    return units;
                }

                set
                {
                    units = value;
                }
            }

            public double UnitMultiplier
            {
                get
                {
                    return unitMultiplier;
                }

                set
                {
                    unitMultiplier = value;
                }
            }

            public string FormatString
            {
                get
                {
                    return formatString;
                }

                set
                {
                    formatString = value;
                }
            }

            public string UnitsLong
            {
                get
                {
                    return unitsLong;
                }

                set
                {
                    unitsLong = value;
                }
            }

            public double Value
            {
                get
                {
                    return value;
                }

                set
                {
                    this.value = value;
                }
            }

            public UsCustomaryUnitScale(double Value)
            {
                this.Value = Value;
                FindWeightScale(Value, out units, out unitsLong, out unitMultiplier, out formatString);
            }
            private static void FindWeightScale(Double Value, out string unitsShort, out string unitsLong, out double unitScale, out string formatString)
            {
                const double dram = 0.00390626; // pounds

                if (Value > 1000000) // kilotons - I know this isn't American Customary, but come on - these are big numbers.
                {
                    unitsShort = "1000 tons";
                    unitsLong = unitsShort;
                    unitScale = 2000000;
                    formatString = "#,##0.00";

                }
                else if (Value > 1000) // start using tons at the half-ton mark.
                {
                    unitsShort = "tons";
                    unitsLong = "short tons";
                    unitScale = 2000;
                    formatString = "#,##0.00";
                }
                else if (Value < (dram * 12)) // 16 drams in an ounce - use ounces if there's 12 or more drams because we'll rapidly exit "Dram country".
                {
                    unitsShort = "dr";
                    unitsLong = "drams";
                    unitScale = 1 / 256;
                    formatString = "#,##0";
                }
                else if (Value < 0.75) // start using ounces at the point of less than 12 ounces.
                {
                    unitsShort = "oz";
                    unitsLong = "ounces";
                    unitScale = 1 / 16;
                    formatString = "#,##0";
                }
                else // pounds
                {
                    unitsShort = "lbs";
                    unitsLong = "pounds";
                    unitScale = 1;
                    formatString = "#,##0";
                }
            }

        }

        void ExportLoadouts(GCACharacter pc, GCAWriter fw)
        {
            //var loadout = pc.CurrentLoadout;
            //if (loadout == null)
            //{
            //    loadout = "";
            //}
            string loadout = "";
            try
            {
                //MessageBox.Show("Primary Loadout");
                ExportLoadout(pc, fw, loadout);
                //MessageBox.Show("Secondary Loadouts");
                loadout = pc.CurrentLoadout;
                ExportLoadout(pc, fw, loadout);
                foreach (LoadOut npLoadout in pc.LoadOuts)
                {
                    if ((!string.Empty.Equals(loadout)) && (!npLoadout.Name.Equals(loadout)))
                    {
                        //MessageBox.Show("ExportLoadouts(" + npLoadout.Name + ")");
                        fw.WriteLine();
                        ExportLoadout(pc, fw, npLoadout.Name);
                    }
                }
                fw.WriteLine();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
        }

        void ExportLoadout(GCACharacter pc, GCAWriter fw, string loadout)
        {
            // MessageBox.Show("ExportLoadout(" + loadout + ") entry");
            if (pc.LoadOuts.Contains(loadout))
            {
                // MessageBox.Show("ExportLoadout(" +loadout+") passed test");
                LoadOut curLoadout = pc.LoadOuts[loadout];
                fw.WriteLine();
                fw.WriteHeader(string.Format("Loadout: {0}", curLoadout.Name));
                foreach (GCATrait item in curLoadout.Items)
                {
                    fw.WriteLine(EquipmentFormatter(item, fw));
                }

                ExportDR(pc, fw, curLoadout.Body);
                fw.WriteLine();
                fw.WriteLine("Total Weight: {0} lbs.", GetLoadoutWeight(curLoadout.Items));
            }
        }

        void ExportCombat(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteLine();
            fw.WriteHeader(string.Format("Combat"));
            ExportMeleeAttacks(pc, fw);
            ExportRangedAttacks(pc, fw);
        }

        void ExportDR(GCACharacter pc, GCAWriter fw, GCA5Engine.Body HitLocations)
        {
            var HitLoc = new List<BodyItem>();
            try
            {
                //fw.WriteLine(string.Format("Adding {0} BodyItems", pc.Body.Count()));
                for (int x = 1; x < HitLocations.Count(); x++)
                {
                    //fw.WriteLine(string.Format("  {0}", pc.Body.Item(x).Name));
                    if (HitLocations.Item(x).Display)
                        HitLoc.Add(HitLocations.Item(x));
                }
                fw.WriteLine();
                fw.WriteHeader("Hit Locations");
                if (null != HitLoc && HitLoc.Count > 0)
                {
                    fw.WriteLine("{|");
                    fw.WriteLine("! Location");
                    fw.WriteLine("! DR");
                    try
                    {
                        var builder = new StringBuilder();
                        foreach (BodyItem item in HitLoc)
                        {
                            builder.AppendLine("|-");
                            builder.AppendFormat("| {0}{1}", item.Name, Environment.NewLine);
                            builder.AppendFormat("| {0}", item.DR);
                            fw.WriteLine(builder.ToString());
                            builder.Clear();
                        }

                    }
                    catch (Exception e)
                    {
                        fw.WriteLine(e);
                    }
                    fw.WriteLine("|}");

                }
                else
                {
                    fw.WriteLine("HitLoc is trash?");
                }
            }
            catch (Exception e)
            {
                fw.WriteLine(e);
            }
        }

        void ExportRangedAttacks(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteSubHeader(string.Format("Ranged Attacks"));
            fw.WriteLine("{|");
            fw.WriteLine("! Name");
            fw.WriteLine("! Damage");
            fw.WriteLine("! Acc");
            fw.WriteLine("! ½ Dam");
            fw.WriteLine("! Max");
            fw.WriteLine("! ROF");
            fw.WriteLine("! Shots");
            //fw.WriteLine("! Skill");
            fw.WriteLine("! Skillscore");
            fw.WriteLine("! MinST");
            fw.WriteLine("! Bulk");
            fw.WriteLine("! Rcl");
            fw.WriteLine("! LC");
            foreach (var curItem in Traits
                .Where(trait => ((string.IsNullOrEmpty(trait.get_TagItem("hide"))) || (trait.get_TagItem("owned").Equals("yes"))))
                .Where(trait => trait.DamageModeTagItemCount("charrangemax") > 0)
                )
            {
                var DamageText = new StringBuilder();
                var currentModeIndex = curItem.DamageModeTagItemAt("charrangemax");
                var modesCount = curItem.DamageModeTagItemCount("charrangemax");
                DamageText.AppendLine("|-");
                if (modesCount > 1 | curItem.DisplayName.Length > 50)
                    DamageText.Append("| colspan=12 | ");
                else
                    DamageText.Append("| ");
                DamageText.AppendFormat("'''{0}'''{1}", curItem.DisplayName, Environment.NewLine);

                do // Iterate over damage modes
                {
                    //this mode is hand!
                    if (modesCount > 1 | curItem.DisplayName.Length > 50)
                    {
                        //we're doing separate lines for each mode
                        DamageText.AppendLine("|-");
                        //print the mode name
                        DamageText.Append("| ");
                        DamageText.AppendLine(curItem.DamageModeName(currentModeIndex));
                    }

                    DamageText.Append("| ");
                    DamageText.AppendFormat(" {0}", curItem.DamageModeTagItem(currentModeIndex, "chardamage"));
                    if (!string.IsNullOrEmpty(curItem.DamageModeTagItem(currentModeIndex, "chararmordivisor")))
                    {
                        DamageText.AppendFormat(" ({0})", FormatArmorDivisor(curItem, currentModeIndex));
                    }
                    DamageText.AppendFormat(" {0}", curItem.DamageModeTagItem(currentModeIndex, "chardamtype"));

                    if (!string.IsNullOrEmpty(curItem.DamageModeTagItem(currentModeIndex, "charradius")))
                    {
                        DamageText.AppendFormat(" ({0})", curItem.DamageModeTagItem(currentModeIndex, "charradius"));
                    }
                    DamageText.AppendLine();

                    DamageText.Append("| ");
                    //print the level
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "characc"));

                    DamageText.Append("| ");
                    //print the reach
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charrangehalfdam"));
                    DamageText.Append("| ");
                    //print the reach
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charrangemax"));

                    DamageText.Append("| ");
                    //print the reach
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charrof"));
                    DamageText.Append("| ");
                    //print the reach
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charshots"));

                    DamageText.Append("| ");
                    //print the level
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charskillscore"));


                    DamageText.Append("| ");
                    //print the ST
                    DamageText.AppendLine(!string.IsNullOrEmpty(curItem.DamageModeTagItem(currentModeIndex, "charminst")) ?
                            curItem.DamageModeTagItem(currentModeIndex, "charminst")
                            : "N/A"
                            );

                    DamageText.Append("| ");
                    //print the lc
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charbulk"));

                    DamageText.Append("| ");
                    //print the lc
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charrcl"));

                    DamageText.Append("| ");
                    //print the lc
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "lc"));

                    //print the notes
                    var NotesText = curItem.DamageModeTagItem(currentModeIndex, "notes");
                    if (!string.IsNullOrEmpty(NotesText))
                    {
                        DamageText.AppendLine("|- ");
                        DamageText.AppendLine("| ");
                        DamageText.Append("| colspan=11 | ");
                        DamageText.AppendLine(NotesText);
                    }

                    if (!string.IsNullOrEmpty(curItem.Modes.Mode(currentModeIndex).ItemNotesText()))
                    {
                        DamageText.AppendLine("|- ");
                        DamageText.AppendLine("| ");
                        DamageText.Append("| colspan=11 | ");
                        DamageText.AppendLine(curItem.Modes.Mode(currentModeIndex).ItemNotesText());
                    }

                    currentModeIndex = curItem.DamageModeTagItemAt("charreach", currentModeIndex + 1);
                } while (currentModeIndex > 0);

                fw.WriteLine(DamageText);
                DamageText.Clear();
            }

            fw.WriteLine();
            fw.WriteLine("|}");

            fw.WriteLine();
        }
        void ExportMeleeAttacks(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteSubHeader(string.Format("Melee Attacks"));
            fw.WriteLine("{|");
            fw.WriteLine("! Name");
            fw.WriteLine("! Damage");
            fw.WriteLine("! Reach");
            //fw.WriteLine("! Skill");
            fw.WriteLine("! Skillscore");
            fw.WriteLine("! Parry");
            fw.WriteLine("! MinST");
            fw.WriteLine("! LC");
            foreach (var curItem in Traits
                .Where(trait => ((string.IsNullOrEmpty(trait.get_TagItem("hide"))) || (trait.get_TagItem("owned").Equals("yes"))))
                .Where(trait => trait.DamageModeTagItemCount("charreach") > 0)
                )
            {
                var DamageText = new StringBuilder();
                var currentModeIndex = curItem.DamageModeTagItemAt("charreach");
                var modesCount = curItem.DamageModeTagItemCount("charreach");
                DamageText.AppendLine("|-");
                if (modesCount > 1)
                    DamageText.Append("| colspan=7 | ");
                else
                    DamageText.Append("| ");
                DamageText.AppendFormat("'''{0}'''{1}", curItem.DisplayName, Environment.NewLine);

                do // Iterate over damage modes
                {
                    //this mode is hand!
                    if (modesCount > 1)
                    {
                        //we're doing separate lines for each mode
                        DamageText.AppendLine("|-");
                        //print the mode name
                        DamageText.Append("| ");
                        DamageText.AppendLine(curItem.DamageModeName(currentModeIndex));
                    }

                    DamageText.Append("| ");
                    DamageText.AppendFormat(" {0}", curItem.DamageModeTagItem(currentModeIndex, "chardamage"));
                    if (!string.IsNullOrEmpty(curItem.DamageModeTagItem(currentModeIndex, "chararmordivisor")))
                    {
                        DamageText.AppendFormat(" ({0})", FormatArmorDivisor(curItem, currentModeIndex));
                    }
                    DamageText.AppendFormat(" {0}", curItem.DamageModeTagItem(currentModeIndex, "chardamtype"));

                    if (!string.IsNullOrEmpty(curItem.DamageModeTagItem(currentModeIndex, "charradius")))
                    {
                        DamageText.AppendFormat(" ({0})", curItem.DamageModeTagItem(currentModeIndex, "charradius"));
                    }
                    DamageText.AppendLine();

                    DamageText.Append("| ");
                    //print the reach
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charreach"));

                    //DamageText.Append("| ");
                    //print the skill
                    //DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charskillused"));

                    DamageText.Append("| ");
                    //print the level
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "charskillscore"));

                    DamageText.Append("| ");
                    //print the parry
                    DamageText.AppendLine(
                        !string.IsNullOrEmpty(curItem.DamageModeTagItem(currentModeIndex, "charparryscore")) ?
                        curItem.DamageModeTagItem(currentModeIndex, "charparryscore")
                        : "N/A"
                        );

                    DamageText.Append("| ");
                    //print the ST
                    DamageText.AppendLine(!string.IsNullOrEmpty(curItem.DamageModeTagItem(currentModeIndex, "charminst")) ?
                            curItem.DamageModeTagItem(currentModeIndex, "charminst")
                            : "N/A"
                            );

                    DamageText.Append("| ");
                    //print the lc
                    DamageText.AppendLine(curItem.DamageModeTagItem(currentModeIndex, "lc"));

                    //print the notes
                    var NotesText = curItem.DamageModeTagItem(currentModeIndex, "notes");
                    if (!string.IsNullOrEmpty(NotesText))
                    {
                        DamageText.AppendLine("|- ");
                        DamageText.AppendLine("| ");
                        DamageText.Append("| colspan=6 | ");
                        DamageText.AppendLine(NotesText);
                    }

                    if (!string.IsNullOrEmpty(curItem.Modes.Mode(currentModeIndex).ItemNotesText()))
                    {
                        DamageText.AppendLine("|- ");
                        DamageText.AppendLine("| ");
                        DamageText.Append("| colspan=6 | ");
                        DamageText.AppendLine(curItem.Modes.Mode(currentModeIndex).ItemNotesText());
                    }

                    currentModeIndex = curItem.DamageModeTagItemAt("charreach", currentModeIndex + 1);
                } while (currentModeIndex > 0);

                fw.WriteLine(DamageText);
                DamageText.Clear();
            }
            fw.WriteLine();
            fw.WriteLine("|}");
            fw.WriteLine();
        }

        private static string FormatArmorDivisor(GCATrait curItem, int currentModeIndex)
        {
            var armorDivisor = curItem.DamageModeTagItem(currentModeIndex, "chararmordivisor");
            if (armorDivisor == "!")
                armorDivisor = Convert.ToChar(8734).ToString();
            return armorDivisor;
        }

        void ExportEquiment(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader(string.Format("Equipment [{0:C}]", pc.get_Cost(modConstants.Equipment)));
            foreach (var item in ComplexListTrait(TraitTypes.Equipment, fw).Where(x => string.IsNullOrEmpty(x) != true))
            {
                fw.Write(item);
            }
            fw.WriteLine();
            fw.WriteLine("Total Weight: {0} lbs.", GetEquipmentWeight());
            fw.WriteLine();
        }

        void ExportPointsSummary(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteLine();
        }

        void ExportMeta(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Metatraits [" + pc.get_Cost(modConstants.Templates) + "]");
            foreach (var item in ComplexListTrait(TraitTypes.Templates, fw).Where(x => string.IsNullOrEmpty(x) != true))
            {
                fw.Write(item);
            }
            fw.WriteLine();
        }

        void ExportSpells(GCACharacter pc, GCAWriter fw)
        {
            var spells = ComplexListTrait(TraitTypes.Spells, fw).Where(x => string.IsNullOrEmpty(x) != true);
            if (spells.Count() > 0)
            {
                fw.WriteHeader("Spells [" + pc.get_Cost(modConstants.Spells) + "]");
                fw.WriteLine("{|");
                fw.WriteLine("! Name");
                fw.WriteLine("! Type");
                fw.WriteLine("! Relative");
                fw.WriteLine("! Level");
                fw.WriteLine("! Points");
                foreach (var item in spells)
                {
                    fw.Write(item);
                }
                fw.WriteLine("}");
            }
        }

        void ExportSkills(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Skills [" + pc.get_Cost(modConstants.Skills) + "]");
            fw.WriteLine("{|");
            fw.WriteLine("! Name");
            fw.WriteLine("! Type");
            fw.WriteLine("! Relative");
            fw.WriteLine("! Level");
            fw.WriteLine("! Points");
            foreach (var item in ComplexListTrait(TraitTypes.Skills, fw).Where(x => string.IsNullOrEmpty(x) != true))
            {
                fw.Write(item);
            }
            fw.WriteLine("|}");
        }

        void ExportQuirks(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Quirks [" + pc.get_Cost(modConstants.Quirks) + "]");
            foreach (var item in ComplexListTrait(TraitTypes.Quirks, fw).Where(x => string.IsNullOrEmpty(x) != true))
            {
                fw.Write(item);
            }
            fw.WriteLine();
        }

        void ExportDisadvantages(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Disadvantages [" + pc.get_Cost(modConstants.Disadvantages) + "]");
            foreach (var item in ComplexListTrait(TraitTypes.Disadvantages, fw).Where(x => string.IsNullOrEmpty(x) != true))
            {
                fw.Write(item);
            }
            fw.WriteLine();
        }

        void ExportPerks(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Perks [" + pc.get_Cost(modConstants.Perks) + "]");
            foreach (var item in ComplexListTrait(TraitTypes.Perks, fw).Where(x => string.IsNullOrEmpty(x) != true))
            {
                fw.Write(item);
            }
            fw.WriteLine();
        }

        void ExportAdvantages(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Advantages [" + pc.get_Cost(modConstants.Advantages) + "]");
            foreach (var item in ComplexListTrait(TraitTypes.Advantages, fw).Where(x => string.IsNullOrEmpty(x) != true))
            {
                fw.Write(item);
            }
            fw.WriteLine();
        }

        void ExportReaction(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Reaction Modifiers");
            fw.WriteLine(
                string.Format("Reaction: {0}/{1}",
                    Traits.Find(x => x.Name.Equals("Appealing")).Score,
                    Traits.Find(x => x.Name.Equals("Unappealing")).Score
                )
            );
            try
            {
                var reaction = Traits.Find(x => x.Name.Equals("Reaction"));
                fw.WriteLine("{0}{1}.", "Conditionals: ", reaction.get_TagItem("conditionallist"));
            }
            catch (ArgumentNullException e)
            {
                modHelperFunctions.Notify(e.ToString());
            }


            //var StatNames = new List<string> {
            //    "Reaction",
            //    "Appealing",
            //    "Unappealing"};
            //foreach (var item in ComplexListAttributes(StatNames, fw))
            //{
            //    fw.Write(item);
            //}
            //fw.WriteLine(fw.FormatTrait("thr", pc.BaseTH));
            //fw.WriteLine(fw.FormatTrait("sw", pc.BaseSW));
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

            fw.WriteLine("{|");
            var coreStats = new List<string> {
                    "ST",
                    "DX",
                    "IQ",
                    "HT"
                 };
            var seconaryStats = new List<string> {
                    "Hit Points",
                    "Will",
                    "Perception",
                    "Fatigue Points"
                };
            Table tableData = new Table {
                coreStats,
                new List<string>(),
                seconaryStats,
                new List<string>()
            };
            int insertRow = 1;
            insertStatsRow(coreStats, tableData, insertRow);

            fw.WriteLine("|}");
        }

        private void insertStatsRow(List<string> statsNames, Table targetTable, int destinationRow)
        {
            IOrderedEnumerable<GCATrait> filteredVisibleAttributes = null;

            filteredVisibleAttributes = getAttributeObjects(statsNames);
            List<string> values = new List<string>();

            foreach (var trait in filteredVisibleAttributes)
            {
                values.Add(trait.DisplayName);

                values.Add(formatAttributeLevel(trait));

                values.Add(string.Format(" [{0}]", trait.Points));
                values.Add(formatBonuses(trait));
                values.Add(formatConditionals(trait));
            }
            targetTable[destinationRow] = new List<string>(values);
        }

        void ExportAttributesText(GCACharacter pc, GCAWriter fw)
        {
            fw.WriteHeader("Attributes [" + pc.get_Cost(modConstants.Stats) + "]");

            var table = new Table();

            //ExportAttributesListed(fw, new List<string> {
            //    "ST",
            //    "DX",
            //    "IQ",
            //    "HT"});

            table.Add(ComplexListAttributes(new List<string> {
                "ST",
                "DX",
                "IQ",
                "HT"}, fw).ToList());

            //foreach (var item in ComplexListAttributes(new List<string> {
            //    "Hit Points",
            //    "Will",
            //    "Perception",
            //    "Fatigue Points"}, fw))
            //{
            //    //fw.Write(item);
            //}
            //fw.WriteLine();

            table.Add(ComplexListAttributes(new List<string> {
                "Hit Points",
                "Will",
                "Perception",
                "Fatigue Points"}, fw).ToList());

            fw.Write(WikiFormatTable(table, true));

            table.Clear();

            table.Add(ComplexListAttributes(new List<string> {
                "Basic Speed",
                "Basic Move" }, fw).ToList());

            table.Add(new List<string> {
                fw.FormatTrait("thr", pc.BaseTH),
                fw.FormatTrait("sw", pc.BaseSW)
            });

            var SuperST = from attribute in Traits
                          where attribute.ItemType == TraitTypes.Attributes
                          where attribute.Name.Equals("Super-Effort Striking ST")
                          where string.IsNullOrEmpty(attribute.get_TagItem("hide"))
                          select attribute;

            if (SuperST.Count() > 0)
            {
                table.Add(new List<string> {
                    fw.FormatTrait("Super thr", pc.ReturnThrFromScore(SuperST.First().Score)),
                    fw.FormatTrait("Super sw", pc.ReturnSwFromScore(SuperST.First().Score))
                });
            }

            table.Add(ComplexListAttributes(new List<string> {
                "Freakishness",
                "RP",
                "Knockback Value"}, fw).ToList());

            table.Add(ComplexListAttributes(new List<string> {
                "Consciousness Check",
                "Death Check"}, fw).ToList());

            table.Add(ComplexListAttributes(new List<string> {
                "High Jump",
                "Broad Jump"}, fw).ToList());

            fw.Write(WikiFormatTable(table));

            table.Clear();
        }

        private void ExportAttributesListed(GCAWriter fw, List<string> StatNames)
        {
            foreach (var item in ComplexListAttributes(StatNames, fw))
            {
                fw.Write(item);
            }
            fw.WriteLine();
        }

        void ExportBiography(GCACharacter CurChar, GCAWriter fw)
        {
            fw.WriteTrait("Name:", CurChar.Name);
            fw.WriteTrait("Player:", CurChar.Player);
            fw.WriteTrait("Race:", CurChar.Race);
            if (!string.IsNullOrEmpty(CurChar.Appearance))
                fw.WriteTrait("Appearance: ", CurChar.Appearance);

            if (!string.IsNullOrEmpty(CurChar.Height))
                fw.WriteTrait("Height:", CurChar.Height);

            if (!string.IsNullOrEmpty(CurChar.Weight))
                fw.WriteTrait("Weight:", CurChar.Weight);

            if (!string.IsNullOrEmpty(CurChar.Age))
                fw.WriteTrait("Age:", CurChar.Age);

            fw.WriteLine("");
        }
        #endregion Exporters

        double GetLoadoutWeight(SortedTraitCollection LoadoutItems)
        {
            double sum = 0;
            foreach (GCATrait trait in LoadoutItems)
            {
                sum += Convert.ToDouble(trait.get_TagItem("weight"));
            }
            return sum;
        }

        double GetEquipmentWeight()
        {
            var result = from trait in Traits
                         where trait.ItemType == TraitTypes.Equipment
                         where string.IsNullOrEmpty(trait.get_TagItem("hide"))
                         where string.IsNullOrEmpty(trait.ParentKey)
                         select Convert.ToDouble(trait.get_TagItem("weight"));
            return result.Sum();
        }

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
                fw.WriteLine(new string(separator[0], 60));
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
                         where string.IsNullOrEmpty(trait.get_TagItem("hide"))
                         select trait.Points != 0 ? string.Format("{0} [{1}]", trait.Name, trait.Points) : trait.Name;

            return string.Join(", ", result) + ".";
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
                         where string.IsNullOrEmpty(trait.get_TagItem("hide"))
                         select trait;

            return string.Join("; ", result.Select(x => AdvantageFormatter(x, fw))) + ".";
        }

        IEnumerable<string> ComplexListTrait(TraitTypes traitType, GCAWriter fw)
        {
            var result = from trait in Traits
                         where trait.ItemType == traitType
                         where string.IsNullOrEmpty(trait.get_TagItem("hide"))
                         select FormatTrait(trait, fw);
            return result;
        }


        IEnumerable<string> ComplexListAttributes(List<string> attributeList, GCAWriter fw)
        {
            var result = from trait in getAttributeObjects(attributeList)
                         select FormatTrait(trait, fw);
            return result;

        }

        private IOrderedEnumerable<GCATrait> getAttributeObjects(List<string> attributeList)
        {
            var filteredVisibleAttributes = from trait in Traits
                                            where trait.ItemType == TraitTypes.Attributes
                                            where attributeList.Contains(trait.Name)
                                            //where !string.IsNullOrEmpty(trait.get_TagItem("mainwin"))
                                            //where !string.IsNullOrEmpty(trait.get_TagItem("display"))
                                            //where string.IsNullOrEmpty(trait.get_TagItem("hide"))
                                            orderby trait.get_TagItem("mainwin")
                                            select trait;
            return filteredVisibleAttributes;
        }

        private Table AttributeRowInTableFormatter(GCATrait trait, GCAWriter fw)
        {
            var row = new Table {
                new List<string>{
                    trait.DisplayName,
                    formatAttributeLevel(trait),
                    string.Format(" [{0}]", trait.Points)
                },
                new List<string> {
                    formatConditionals(trait)

                },
                new List<string> {
                    formatConditionals(trait)
                }
            };
            //var builder = new StringBuilder();
            //WikiFormatRows(row, builder);

            return row;
        }

        delegate string TraitFormatter(GCATrait trait, GCAWriter fw);

        string FormatTrait(GCATrait trait, GCAWriter fw, int childDepth = 0)
        {
            //fw.WriteLine("DEBUG TRAIT {0} DEPTH {1}", trait.Name, childDepth);
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
                    formatter = EquipmentFormatter;
                    break;
                case TraitTypes.Templates:
                    formatter = TemplateFormatter;
                    break;
            }
            var builder = new StringBuilder();

            // Handle nested traits recursively
            if (string.IsNullOrEmpty(trait.get_TagItem("parentkey")) || childDepth > 0)
            {
                for (int i = 0; i < childDepth; i++)
                {
                    builder.Append(":");
                }
                var traitStr = formatter != null ? formatter(trait, fw) : string.Empty;
                if (!string.IsNullOrEmpty(trait.get_ChildKeyList()))
                {
                    traitStr = traitStr.Replace("<br/>", "");
                    builder.Append(traitStr);
                    builder.Append(Environment.NewLine);
                    var keys = trait.get_ChildKeyList().Split(',');
                    //fw.WriteLine("DEBUG TRAIT '{0}' CHILDREN '{1}'", trait.Name, string.Join(",", keys));
                    foreach (var key in keys)
                    {
                        var cleanKey = key.Trim().Substring(1);
                        var child = Traits.FirstOrDefault(x => x.IDKey.Equals(Convert.ToInt32(cleanKey)));
                        //fw.WriteLine("DEBUG TRAIT '{0}' KEY '{1}' CHILD '{2}'", trait.Name, cleanKey, child);
                        if (child != null)
                        {
                            //fw.WriteLine("DEBUG FORMATTING CHILD '{2}'", trait.Name, cleanKey, child);
                            builder.Append(FormatTrait(child, fw, childDepth + 1).Replace("<br/>", ""));
                        }
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(trait.get_TagItem("parentkey")))
                    {
                        traitStr = traitStr.Replace("<br/>", "");
                    }
                    builder.Append(traitStr);
                    builder.Append(Environment.NewLine);
                }
            }
            return builder.ToString();
        }

        string TemplateFormatter(GCATrait trait, GCAWriter fw)
        {
            return AdvantageFormatter(trait, fw);
        }

        string EquipmentFormatter(GCATrait trait, GCAWriter fw)
        {
            var builder = new StringBuilder();
            builder.Append(trait.Name);

            var label = builder.ToString();

            builder.Clear();
            if (!string.IsNullOrEmpty(trait.NameExt) || trait.Mods.Count() > 0)
                builder.Append("(");
            if (!string.IsNullOrEmpty(trait.NameExt))
                builder.Append(trait.NameExt);
            if (!string.IsNullOrEmpty(trait.NameExt) && trait.Mods.Count() > 0)
                builder.Append("; ");
            if (trait.Mods.Count() > 0)
            {
                var mods = new List<GCAModifier>();
                foreach (GCAModifier item in trait.Mods)
                {
                    mods.Add(item);
                }
                builder.Append(string.Join("; ", mods.Select(x => ModifierFormatter(x, fw))));
            }
            if (!string.IsNullOrEmpty(trait.NameExt) || trait.Mods.Count() > 0)
                builder.Append(")");

            if (Convert.ToInt32(trait.get_TagItem("count")) > 1)
            {
                builder.AppendFormat(" {0} lbs, {1:C} ×{2} = {3} lbs {4:C}",
                    Convert.ToDouble(trait.get_TagItem("precountweight")),
                    Convert.ToDouble(trait.get_TagItem("precountcost")),
                    Convert.ToDouble(trait.get_TagItem("count")),
                    Convert.ToDouble(trait.get_TagItem("weight")),
                    Convert.ToDouble(trait.get_TagItem("cost"))
                    );
            }
            else
            {
                builder.AppendFormat(" {0} lbs {1:C}", Convert.ToDouble(trait.get_TagItem("weight")), Convert.ToDouble(trait.get_TagItem("cost")));
            }
            return fw.FormatTrait(label, builder.ToString());
        }

        string AttributeFormatter(GCATrait trait, GCAWriter fw)
        {
            var builder = new StringBuilder();
            builder.Append(trait.DisplayName);

            var label = builder.ToString();

            builder.Clear();
            formatAttributeLevel(trait, builder);

            builder.AppendFormat(" [{0}]", trait.Points);
            formatBonuses(trait, builder);
            formatConditionals(trait, builder);
            return fw.FormatTrait(label, builder.ToString());
        }

        private string formatConditionals(GCATrait trait, StringBuilder builder = null)
        {
            if (builder == null)
            {
                builder = new StringBuilder();
            }
            if (!string.Empty.Equals(trait.get_TagItem("conditionallist")))
            {
                builder.AppendFormat(" {0}{1}.", "Conditionals: ", trait.get_TagItem("conditionallist"));
            }
            return builder.ToString();
        }

        private string formatBonuses(GCATrait trait, StringBuilder builder = null)
        {
            if (builder == null)
            {
                builder = new StringBuilder();
            }
            if (!string.Empty.Equals(trait.get_TagItem("bonuslist")))
            {
                builder.AppendFormat(" {0}{1};", "Bonuses: ", trait.get_TagItem("bonuslist"));
            }
            return builder.ToString();
        }

        private string formatAttributeLevel(GCATrait trait, StringBuilder builder = null)
        {
            if (builder == null)
            {
                builder = new StringBuilder();
            }
            builder.Append(trait.DisplayScore);
            if (string.Compare(trait.get_TagItem("units"), "", StringComparison.Ordinal) != 0)
            {
                var units = trait.get_TagItem("units").Split('|');
                var unit = (Metric && units.Length < 1) ? units[1] : units[0];
                builder.Append(unit);
            }
            return builder.ToString();
        }

        string AdvantageFormatter(GCATrait trait, GCAWriter fw)
        {
            var builder = new StringBuilder();
            builder.Append(trait.Name);
            if (!trait.get_TagItem("level").Equals("1") || !string.IsNullOrEmpty(trait.get_TagItem("upto")) || !string.IsNullOrEmpty(trait.LevelName)) // has more than one level
            {
                builder.AppendFormat(" {0}", string.IsNullOrEmpty(trait.LevelName) ? trait.get_TagItem("level") : trait.LevelName);
            }

            var label = builder.ToString();

            builder.Clear();

            BuildAdvantageNameExtension(trait, fw, builder);

            builder.AppendFormat(" [{0}]", trait.Points);
            if (!string.Empty.Equals(trait.get_TagItem("bonuslist")))
            {
                builder.AppendFormat(" {0}{1};", "Bonuses: ", trait.get_TagItem("bonuslist"));
            }
            if (!string.Empty.Equals(trait.get_TagItem("conditionallist")))
            {
                builder.AppendFormat(" {0}{1}.", "Conditionals: ", trait.get_TagItem("conditionallist"));
            }

            return fw.FormatTrait(label, builder.ToString());
        }

        void BuildAdvantageNameExtension(GCATrait trait, GCAWriter fw, StringBuilder builder)
        {
            var hasNameExt = !string.IsNullOrEmpty(trait.NameExt);
            var hasMods = trait.Mods.Count() > 0;
            //if (!trait.get_TagItem("level").Equals("1") || !string.IsNullOrEmpty(trait.get_TagItem("upto")) || !string.IsNullOrEmpty(trait.LevelName)) // has more than one level
            var hasLevel = !string.IsNullOrEmpty(trait.LevelName);

            if (hasNameExt || hasMods || hasLevel)
            {
                builder.Append(" (");
                if (hasNameExt)
                {
                    builder.Append(trait.NameExt);
                    if (hasMods || hasLevel)
                        builder.Append("; ");
                }
                if (hasLevel)
                {
                    builder.AppendFormat("{0}", string.IsNullOrEmpty(trait.LevelName) ? trait.get_TagItem("level") : trait.LevelName);
                    if (hasMods)
                        builder.Append("; ");
                }
                if (hasMods)
                {
                    var mods = new List<GCAModifier>();
                    foreach (GCAModifier item in trait.Mods)
                    {
                        mods.Add(item);
                    }
                    builder.Append(string.Join("; ", mods.Select(x => ModifierFormatter(x, fw))));
                }

                builder.Append(")");
            }
        }

        string ModifierFormatter(GCAModifier trait, GCAWriter fw)
        {
            var builder = new StringBuilder();
            builder.Append(trait.DisplayName.Trim());
            if (!trait.get_TagItem("level").Equals("1") || !string.IsNullOrEmpty(trait.get_TagItem("upto")) || !string.IsNullOrEmpty(trait.LevelName())) // has more than one level or has named levels
            {
                builder.AppendFormat(" {0}", string.IsNullOrEmpty(trait.LevelName()) ? trait.get_TagItem("level") : trait.LevelName());
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
                         where string.IsNullOrEmpty(trait.get_TagItem("hide"))
                         select trait;

            return string.Join(", ", result) + ".";
        }

        private string MeleeAttackFormatter(GCATrait trait, GCAWriter fw, int mode)
        {
            var builder = new StringBuilder();
            builder.AppendLine("|-");
            builder.AppendFormat("| {0}{1}", trait.DisplayName, Environment.NewLine);
            return builder.ToString();
        }

        string SkillFormatter(GCATrait trait, GCAWriter fw)
        {
            var builder = new StringBuilder();
            builder.AppendLine("|-");
            builder.AppendFormat("| {0}{1}", trait.DisplayName, Environment.NewLine);
            builder.AppendFormat("| {0}{1}", trait.get_TagItem("type"), Environment.NewLine);
            //var label = builder.ToString();

            var skillTypeRaw = (int)SkillType.Skill;
            var skillTypeParseSuccess = int.TryParse(trait.get_TagItem("sd"), out skillTypeRaw);
            //builder.Clear();
            switch (skillTypeParseSuccess ? (SkillType)skillTypeRaw : SkillType.Skill)
            {
                case SkillType.Combo:
                    insertStepoff(trait, builder);
                    builder.AppendFormat("| {0}{1}", trait.get_TagItem("combolevel"), Environment.NewLine);
                    //builder.AppendFormat(" - {0}", trait.get_TagItem("combolevel"));
                    break;
                default: // SkillType.Skill, SkillType.Technique
                    insertStepoff(trait, builder);
                    builder.AppendFormat("| {0}{1}", trait.get_TagItem("level"), Environment.NewLine);
                    //builder.AppendFormat(" - {0}", trait.get_TagItem("level"));
                    break;
            }
            builder.AppendFormat("| {0}", trait.Points);
            //builder.AppendFormat(" [{0}]", trait.Points);
            if (!string.Empty.Equals(trait.get_TagItem("bonuslist")))
            {
                builder.AppendLine();
                builder.AppendLine("|-");
                builder.AppendFormat("| colspan=\"5\" | {0}{1}", "Bonuses: ", trait.get_TagItem("bonuslist"));
            }
            if (!string.Empty.Equals(trait.get_TagItem("conditionallist")))
            {
                builder.AppendLine();
                builder.AppendLine("|-");
                builder.AppendFormat("| colspan=\"5\" | {0}{1}", "Conditionals: ", trait.get_TagItem("conditionallist"));
            }
            //return fw.FormatTrait(label, builder.ToString());
            return builder.ToString();
        }

        static void insertStepoff(GCATrait trait, StringBuilder builder)
        {
            var stepOff = trait.get_TagItem("stepoff");
            if (!stepOff.Equals(string.Empty))
            {
                builder.AppendFormat("| {0}", stepOff);
                var step = trait.get_TagItem("step");
                if (!step.Equals(string.Empty))
                    builder.AppendFormat("{0}", step);
                else
                    builder.AppendFormat("?");
            }
            else
            {
                builder.AppendFormat("| ?+?");
            }
            builder.AppendLine();
        }

        string WikiFormatTable(Table table, bool transpose = false)
        {
            if (transpose)
            {
                table = table.Pivot();
            }

            var result = new StringBuilder("{|");

            WikiFormatRows(table, result);

            result.AppendLine("|}");
            return result.ToString();
        }

        private static void WikiFormatRows(Table table, StringBuilder result)
        {
            foreach (var row in table)
            {
                result.AppendLine("|-");
                WikiFormatCells(row, result);
            }
        }

        private static void WikiFormatCells(List<string> cellList, StringBuilder result)
        {
            foreach (var cell in cellList)
            {
                result.Append("| ");
                result.AppendLine(cell);
            }
        }

        internal static Table PivotTable(Table source)
        {
            return (Table)PivotNestedLists(source);
        }

        internal static List<List<T>> PivotNestedLists<T>(List<List<T>> source)
        {
            var numRows = source.Max(a => a.Count);

            var items = new List<List<T>>();
            for (int row = 0; row < source.Count; ++row)
            {
                for (int col = 0; col < source[row].Count; ++col)
                {
                    //Get the current "row" for this column, if any
                    if (items.Count <= col)
                        items.Add(new List<T>());

                    var current = items[col];

                    //Insert the value into the row
                    current.Add(source[row][col]);
                }
            }
            return items;
        }

        static T[][] PivotArrayToJagged<T>(T[][] source)
        {
            var numRows = source.Max(a => a.Length);

            //Will be adjusting multiple "rows" at the same time so need to use a more flexible collection
            var items = new List<List<T>>();
            for (int row = 0; row < source.Length; ++row)
            {
                for (int col = 0; col < source[row].Length; ++col)
                {
                    //Get the current "row" for this column, if any
                    if (items.Count <= col)
                        items.Add(new List<T>());

                    var current = items[col];

                    //Insert the value into the row
                    current.Add(source[row][col]);
                }
            }

            //Convert the nested lists back into a jagged array
            return (from i in items
                    select i.ToArray()
                    ).ToArray();
        }
        #endregion Formatters
    }


    public class Table : List<List<string>>
    {
        internal Table Pivot()
        {
            Contract.Ensures(Contract.Result<Table>() != null);
            return (Table) WikiExporter.PivotNestedLists(this);
        }
    }
}
