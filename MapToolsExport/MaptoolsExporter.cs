namespace GCA.MapToolsExport
{
    using System;
    using System.Collections.Generic;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using Export;
    using GCA5.Interfaces;
    using GCA5Engine;
    using System.Xml;
    using System.IO;
    using System.Xml.Serialization;
    using System.Runtime.CompilerServices;

    public sealed class MaptoolsExporter : IExportSheet
    {
        const string assemblyName = "GCA.MapToolsExport";
        const string settingsFileName = @"MaptoolsExportOptions.xml";

        #region CTORS
        public MaptoolsExporter()
        {
            MyOptions = new SheetOptionsManager(Name);
        }
        #endregion CTORS

        #region InterfaceImplementation
        #region InterfaceProperties
        public SheetOptionsManager MyOptions { get; private set; }
        public List<GCATrait> Traits { get; private set; }
        
        public string Name
        {
            get
            {
                return "Maptools Self-contained Export";
            }
        }

        public string Description
        {
            get
            {
                return "Exports a MapTools token with self-contained automation features.";
            }
        }

        public string Version
        {
            get
            {
                return "1.0.0";
            }
        }
        
        public event IExportSheet.RequestRunSpecificOptionsEventHandler RequestRunSpecificOptions;
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
            NoteMessage(String.Format("Creating Options"), Priority.Green);
            // take a look at myself
            var domain = AppDomain.CurrentDomain;
            NoteMessage(domain.ToString());
            domain.Load(assemblyName);
            NoteMessage(String.Format("Loaded {0} into domain", assemblyName));
            var listOfAssemblies = new List<System.Reflection.Assembly>(domain.GetAssemblies());
            NoteMessage(String.Format("Found {0} assemblies", listOfAssemblies.Count));
            try
            {
                // Get where the .dll is stored, this is where we will find our options XML
                var optionsPath = Path.GetDirectoryName(listOfAssemblies.FirstOrDefault(x => x.FullName.StartsWith(assemblyName, StringComparison.Ordinal)).Location);
                optionsPath = optionsPath + Path.DirectorySeparatorChar + settingsFileName;
                NoteMessage(String.Format("Options Path is {0}", optionsPath), Priority.Green);
                ExtractOptionsFromFile(optionsPath, Options);
            }
            catch (Exception ex)
            {
                NoteMessage(Name + ": failed to create options objects. " + ex.Message + "\n" + "Stack Trace: " + "\n" + ex.StackTrace + "\n" + ex.InnerException.ToString(), Priority.Red);
                //System.Windows.Forms.MessageBox.Show(ex.ToString() + "\n" + ex.InnerException.ToString(), assemblyName);
            }
        }
        /// <summary>
        /// This is where we do the heavy lifting of pulling the options out and setting the options object up.
        /// </summary>
        /// <param name="optionsPath"></param>
        /// <param name="options"></param>
        void ExtractOptionsFromFile(string optionsPath, SheetOptionsManager options)
        {
            var optionsSerializer = new XmlSerializer(typeof(SheetOption[]));

            try
            {
                var optionsArray = new List<SheetOption>();
                try {
                    using (var inFile = new FileStream(optionsPath, FileMode.Open))
                    {
                        optionsArray.AddRange((SheetOption[])optionsSerializer.Deserialize(inFile));
                    }
                }
                catch(FileNotFoundException ex)
                {
                    NoteMessage(Name + ": failed to create options. " + ex.Message + "\n" + "Stack Trace: " + "\n" + ex.StackTrace + "\n" + ex.InnerException.ToString(), Priority.Red);
                    return;
                }
                NoteMessage(String.Format( "{0}: found {1} options", Name,  optionsArray.Count), Priority.Green);
                var descFormatHeader = new SheetOptionDisplayFormat();
                descFormatHeader.BackColor = System.Drawing.SystemColors.Info;
                descFormatHeader.CaptionLocalBackColor = System.Drawing.SystemColors.Info;

                var descFormatBody = new SheetOptionDisplayFormat();

                foreach (var opt in optionsArray)
                {
                    opt.DisplayFormat = opt.Type == OptionType.Header ? descFormatHeader : descFormatBody;
                    opt.UserPrompt = opt.UserPrompt.Replace(@"$Name", Name).Replace(@"$Version", Version);
                    options.AddOption(opt);
                }
            }
            catch (Exception ex)
            {
                NoteMessage(Name + ": failed to create options outer error. " + ex.Message + "\n" + "Stack Trace: " + "\n" + ex.StackTrace + "\n" + ex.InnerException.ToString(), Priority.Red);
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
        /// This creates the export file on disk.
        /// </summary>
        /// <param name="Party"></param>
        /// <param name="TargetFilename"></param>
        /// <param name="Options"></param>
        /// <returns></returns>
        bool IExportSheet.GenerateExport(Party Party, string TargetFilename, SheetOptionsManager Options)
        {
            bool printMults;
            try
            {
                MyOptions = Options;

                printMults = (int)MyOptions.get_Value("OutputCharacters") == 1;
            }
            catch (NullReferenceException ex)
            {
                NoteMessage(Name + ": failed on loading options from a null reference. " + ex.Message + "\n" + "Stack Trace: " + "\n" + ex.StackTrace, Priority.Red);
                return false;
            }
            catch (Exception ex)
            {
                NoteMessage(Name + ": failed on failed on loading options for some reason. " + ex.Message + "\n" + "Stack Trace: " + "\n" + ex.StackTrace, Priority.Red);
                return false;
            }
            try
            {
                if (Party.Characters.Count > 1 && printMults)
                {
                    var count = 1;
                    foreach (GCACharacter pc in Party.Characters)
                    {
                        var TempXMLLocation = System.IO.Path.GetTempFileName();
                        NoteMessage(String.Format("Temp xml file: {0}", TempXMLLocation));
                        using (var fw = new GCA.Export.GCAWriter(TempXMLLocation, false, MyOptions))
                        {
                            ExportCharacter(pc, fw);
                        }
                        MakeTokenZip(TempXMLLocation, TargetFilename + count.ToString());
                        count++;
                    }
                }
                else
                {
                    var TempXMLLocation = System.IO.Path.GetTempFileName();
                    NoteMessage(String.Format("Temp xml file: {0}", TempXMLLocation));
                    using (var fw = new GCA.Export.GCAWriter(TempXMLLocation, false, MyOptions))
                    {
                        ExportCharacter(Party.Current, fw);
                    }
                    MakeTokenZip(TempXMLLocation, TargetFilename);
                }
            }
            catch (Exception ex)
            {
                NoteMessage(String.Format("{0}: failed on export. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace), Priority.Red);
            }

            return true;
        }

        public int PreferredFilterIndex()
        {
            return 0;
        }

        bool IExportSheet.PreviewOptions(SheetOptionsManager Options)
        {
            return true;
        }

        string IExportSheet.SupportedFileTypeFilter()
        {
            return "Maptools Tokens (*.rptok)|*.rptok|Zip file (*.zip)|*.zip";
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
        #endregion InterfaceImplementation

        public string GetTempDirectory()
        {

            string path = Path.Combine(Path.GetTempPath(), Path.GetRandomFileName());
            Directory.CreateDirectory(path);
            return path;

        }
        /// <summary>
        /// Takes a Maptools token XML file and creates a proper token file in the desired location.
        /// </summary>
        /// <param name="XMLPath"></param>
        /// <param name="TargetZip"></param>
        void MakeTokenZip(string XMLPath, string TargetZip)
        {
            var temp = GetTempDirectory();
            NoteMessage(String.Format("XML file is {0}, target is {1}, tempdir is {2}", XMLPath, TargetZip, temp));
            var filename = Path.GetFileName(XMLPath);
            var newfilename = Path.Combine(temp, "content.xml");

            try
            {
                File.Move(XMLPath, newfilename);
            }
            catch (DirectoryNotFoundException ex)
            {
                NoteMessage(String.Format("{0}: directory {4} not found. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, temp), Priority.Red);
                return;
            }
            catch (FileNotFoundException ex)
            {
                NoteMessage(String.Format("{0}: file {4} not found. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, XMLPath), Priority.Red);
                return;
            }
            catch (Exception ex)
            {
                NoteMessage(String.Format("{0}: move XML File {4} to {5} failed. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, XMLPath, newfilename), Priority.Red);
                return;
            }

            using (var properties = new StreamWriter(Path.Combine(temp, "properties.xml")))
            {
                properties.Write(@"<map><entry><string>version</string><string>1.3.b91</string></entry></map>");
            }

            var tempZip = Path.GetTempFileName();
            try
            {
                if (File.Exists(tempZip))
                {
                    File.Delete(tempZip);
                }
                NoteMessage(String.Format("Zipping source directory {0} to target zip {1}", temp, tempZip));
                System.IO.Compression.ZipFile.CreateFromDirectory(temp, tempZip);
            }
            catch (ArgumentNullException ex)
            {
                NoteMessage(String.Format("{0}: create zip file {4} failed. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, TargetZip), Priority.Red);
                return;
            }
            catch (DirectoryNotFoundException ex)
            {
                NoteMessage(String.Format("{0}: create zip file {4} failed. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, TargetZip), Priority.Red);
                return;
            }
            catch (PathTooLongException ex)
            {
                NoteMessage(String.Format("{0}: create zip file {4} failed. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, TargetZip), Priority.Red);
                return;
            }
            catch (ArgumentException ex)
            {
                NoteMessage(String.Format("{0}: create zip file {4} failed. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, TargetZip), Priority.Red);
                return;
            }
            catch (NotSupportedException ex)
            {
                NoteMessage(String.Format("{0}: create zip file {4} failed. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, TargetZip), Priority.Red);
                return;
            }
            catch (UnauthorizedAccessException ex)
            {
                NoteMessage(String.Format("{0}: create zip file {4} failed. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, TargetZip), Priority.Red);
                return;
            }
            catch (IOException ex)
            {
                NoteMessage(String.Format("{0}: create zip file {4} failed. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, TargetZip), Priority.Red);
                return;
            }
            catch (Exception ex)
            {
                NoteMessage(String.Format("{0}: create zip file {4} failed. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace, TargetZip), Priority.Red);
                return;
            }

            if (File.Exists(TargetZip))
            {
                File.Delete(TargetZip);
            }
            File.Move(tempZip, TargetZip);
        }
        /// <summary>
        /// Creates the XML file needed to build the MapTools token
        /// </summary>
        /// <param name="pc"></param>
        /// <param name="fw"></param>
        void ExportCharacter(GCACharacter pc, StreamWriter fw)
        {
            var token = new netrptoolsmaptoolmodelToken
            {
                beingImpersonated = false,

                anchorX = 0,
                anchorY = 0,
                x = 485,
                y = 675,
                z = 338,
                scaleX = 1,
                scaleY = 1,
                sizeScale = 1.0M,
                lastX = 0,
                lastY = 0,
                width = 64,
                height = 64,
                snapToGrid = true,
                snapToScale = true,
                isVisible = true,
                visibleOnlyToOwner = true,
                name = pc.Name,
                ownerType = "0",
                tokenShape = "CIRCLE",
                tokenType = "PC",
                layer = "TOKEN",
                propertyType = "GURPS-GCA5",
                facing = "0",
                isFlippedX = false,
                isFlippedY = false,
                sightType = "Normal",
                hasSight = true,
                label = "",
                notes = pc.Notes,
                gmNotes = ""
            };

            initializeTokenChildren(token);

            try
            {
                var serializer = new XmlSerializer(typeof(netrptoolsmaptoolmodelToken));
                serializer.Serialize(fw, token);
            }
            catch (Exception ex)
            {
                NoteMessage(String.Format("{0}: failed on XML serialization. {1}\n{2}\n{3}", Name, ex.Message, ex.InnerException, ex.StackTrace), Priority.Red);
                throw;
            }
        }

        private static void initializeTokenChildren(netrptoolsmaptoolmodelToken token)
        {
            token.id = new netrptoolsmaptoolmodelTokenID[1];
            token.id[0] = new netrptoolsmaptoolmodelTokenID();
            token.id[0].baGUID = "";

            token.exposedAreaGUID = new netrptoolsmaptoolmodelTokenExposedAreaGUID[1];
            token.exposedAreaGUID[0] = new netrptoolsmaptoolmodelTokenExposedAreaGUID();
            token.exposedAreaGUID[0].baGUID = "";

            token.macroPropertiesMap = new entry[1];
            token.macroPropertiesMap[0] = new entry();

            token.imageAssetMap = new entry[0];

            token.sizeMap = new entry[2];

            token.sizeMap[0] = new entry();
            token.sizeMap[0].javaclass = "net.rptools.maptool.model.HexGridHorizontal";
            token.sizeMap[0].netrptoolsmaptoolmodelGUID = new entryNetrptoolsmaptoolmodelGUID[1];
            token.sizeMap[0].netrptoolsmaptoolmodelGUID[0] = new entryNetrptoolsmaptoolmodelGUID();
            token.sizeMap[0].netrptoolsmaptoolmodelGUID[0].baGUID = "fwABAQllXDgBAAAAOAABAQ==";

            token.sizeMap[1] = new entry();
            token.sizeMap[1].javaclass = "net.rptools.maptool.model.HexGridVertical";
            token.sizeMap[1].netrptoolsmaptoolmodelGUID = new entryNetrptoolsmaptoolmodelGUID[1];
            token.sizeMap[1].netrptoolsmaptoolmodelGUID[0] = new entryNetrptoolsmaptoolmodelGUID();
            token.sizeMap[1].netrptoolsmaptoolmodelGUID[0].baGUID = "fwABAQllXDgBAAAAOAABAQ==";
        }

        static void NoteMessage(string message,
        Priority priority = Priority.None,
        [CallerLineNumber] int lineNumber = 0,
        [CallerMemberName] string caller = null)
        {
            modHelperFunctions.Notify("Line " + lineNumber + " (" + caller + "):\n"+message, priority);
        }
    }
}
