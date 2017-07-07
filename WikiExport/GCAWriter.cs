

namespace GCA.Export
{
    using GCA5Engine;
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.IO;
    using System.Linq;
    using System.Text;
    using System.Threading.Tasks;
    using System.Windows.Forms;

    public class GCAWriter : StreamWriter
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

        enum TraitSeparatorStyle
        {
            SoftBreak,
            List,
            Paragraph,
            Semicolon
        }

        public static string Indent {
            get
            {
                return " &nbsp; ";
            }
        }

        public GCAWriter(string path, bool append, SheetOptionsManager options) : base(path, append)
        {
            MyOptions = options;
        }

        public SheetOptionsManager MyOptions { get; private set; }

        //override public void WriteLine(string format, params object[] stuff)
        //{
        //    format = format + @"</br>";
        //    base.WriteLine(string.Format(format, stuff));
        //}

        public void WriteTrait(string label, string value)
        {
            WriteLine(FormatTrait(label, value));
        }
        public void WriteHeader(string Header)
        {
            WriteLine(FormatHeader(Header));
        }
        internal void WriteSubHeader(string SubHeader)
        {
            WriteLine(FormatSubHeader(SubHeader));
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
        internal string FormatSubHeader(string Header)
        {
            try
            {
                switch ((HeaderStyles)MyOptions.get_Value("HeadingStyle"))
                {
                    case HeaderStyles.DoNothing:
                        return string.Format("{0}<br/>", Header);
                    case HeaderStyles.MinorWikiHeader:
                        return string.Format("===={0}====", Header);
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
            string prefix = "";
            string suffix = "";
            SetPrefixSuffix(ref prefix, ref suffix);
            try
            {
                switch ((GeneralStyles)MyOptions.get_Value("TraitNameStyle"))
                {
                    case GeneralStyles.DoNothing:
                        return string.Format("{2}{0} {1}{3}", label, value, prefix, suffix);
                    case GeneralStyles.MakeBold:
                        return string.Format("{2}'''{0}''' {1}{3}", label, value, prefix, suffix);
                    case GeneralStyles.MakeItalic:
                        return string.Format("{2}''{0}'' {1}{3}", label, value, prefix, suffix);
                    case GeneralStyles.MakeBoldItalic:
                        return string.Format("{2}'''''{0}''''' {1}{3}", label, value, prefix, suffix);
                    default:
                        return string.Format("{2}{0} {1}{3}", label, value, prefix, suffix);
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(string.Format("FormatTrait({0}, {1})\nMessage: {2}\nStacktrace: {3}\nInner: {4}", label, value, e.Message, e.StackTrace, e.InnerException), "Error in FormatTrait()");
                return string.Format("FormatTrait({0}, {1}) error", label, value);
            }
        }

        private void SetPrefixSuffix(ref string prefix, ref string suffix)
        {
            switch ((TraitSeparatorStyle)MyOptions.get_Value("TraitSeparation"))
            {
                case TraitSeparatorStyle.SoftBreak:
                    prefix = "";
                    suffix = "<br/>";
                    break;
                case TraitSeparatorStyle.List:
                    prefix = ":";
                    suffix = "";
                    break;
                case TraitSeparatorStyle.Paragraph:
                    prefix = " ";
                    suffix = "";
                    break;
                case TraitSeparatorStyle.Semicolon:
                    prefix = "";
                    suffix = "; ";
                    break;
                default:
                    break;
            }
        }

        internal void WriteRTF(string rtfString)
        {
            if (rtfString.StartsWith(@"{\rtf", StringComparison.Ordinal))
            {
                //var rtfConverter = new RichTextBox();
                //rtfConverter.Rtf = rtfString;
                //Write(rtfConverter.Text);
                var converter = new SautinSoft.RtfToHtml();
                var htmlString = converter.ConvertString(rtfString);
                Write(htmlString);
                //Write(Pandoc(htmlString));
            }
            else
                Write(rtfString);
        }

        internal void WriteCampaign(CampaignInfo campaign)
        {
            WriteLine(FormatSubHeader(campaign.Name));
            var log = new Dictionary<DateTime, LogEntry>() { };

            foreach (LogEntry entry in campaign.Log)
            {
                AddEntriesByDate(log, entry);
            }
            WriteLine("{|");
            WriteLine("! Play Date");
            WriteLine("! Session Name");
            WriteLine("! Campaign Date");
            WriteLine("! Points");
            foreach (var entry in log.OrderBy(x => x.Key).Select(x => x.Value))
            {
                WriteLogEntry(entry);
            }
            WriteLine("|}");
        }

        private static void AddEntriesByDate(Dictionary<DateTime, LogEntry> log, LogEntry entry)
        {
            DateTime d = new DateTime();
            var culture = new System.Globalization.CultureInfo("en-CA");

            //try
            //{
            //    d = DateTime.ParseExact(entry.EntryDate, "M/d/yyyy", CultureInfo.InvariantCulture);
            //}
            //catch (Exception)
            //{
            //    d = DateTime.ParseExact(entry.EntryDate, "yyyy/M/d", CultureInfo.InvariantCulture);
            //}
            try
            {
                if (DateTime.TryParseExact(entry.EntryDate,
                                "d/M/yyyy",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                out d))
                {
                    log.Add(d, entry);
                }
                else if (DateTime.TryParseExact(entry.EntryDate,
                                "mm/dd/yyyy",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                out d))
                {
                    log.Add(d, entry);
                }
                else if (DateTime.TryParseExact(entry.EntryDate,
                                "yyyy/M/d",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                out d))
                {
                    log.Add(d, entry);
                }
                else if (DateTime.TryParseExact(entry.EntryDate,
                                "yyyy-M-d",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                out d))
                {
                    log.Add(d, entry);
                }
                else if (DateTime.TryParseExact(entry.EntryDate,
                                "mm-dd-yyyy",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                out d))
                {
                    log.Add(d, entry);
                }
                else if (DateTime.TryParseExact(entry.EntryDate,
                                "yyyy-mm-dd",
                                CultureInfo.InvariantCulture,
                                DateTimeStyles.None,
                out d))
                {
                    log.Add(d, entry);
                }
                else
                {
                    var substD = new DateTime(0);
                    entry.Notes = string.Format("DateTime Error {0};{2}{1}{2}", entry.EntryDate,  Environment.NewLine, entry.Notes);
                    Console.WriteLine("DateTime Error {0}", entry.EntryDate);
                    log.Add(d, entry);
                }
            }
            catch (Exception e)
            {
                var substD = new DateTime(0);
                entry.Notes = string.Format("DateTime Exception {0};{2}{1}{2}{3}", entry.EntryDate, e.ToString(), Environment.NewLine, entry.Notes);
                Console.WriteLine("DateTime Error Exception {0}", entry.EntryDate);
                log.Add(d, entry);
            }
        }

        private void WriteLogEntry(LogEntry entry)
        {
            WriteLine("|-");
            WriteLine(string.Format("| {0}", FormatDate( entry.EntryDate)));
            WriteLine(string.Format("| {0}", entry.Caption));
            WriteLine(string.Format("| {0}", entry.CampaignDate));
            WriteLine(string.Format("| {0}", entry.CharPoints));
            WriteLine("|-");
            if (!string.Empty.Equals(entry.Notes))
            {
                Write("| colspan=\"4\" | ");
                WriteRTF(entry.Notes);
                WriteLine();
            }
        }

        private string FormatDate(string date)
        {
            string formattedDate = "";
            DateTime d;
            if (DateTime.TryParseExact(date,
                "d/M/yyyy",
                CultureInfo.CurrentCulture,
                DateTimeStyles.None, 
                out d))
            {
                formattedDate = d.Date.ToShortDateString();
            }
            else if (DateTime.TryParseExact(date,
                            "mm/dd/yyyy",
                            CultureInfo.CurrentCulture,
                            DateTimeStyles.None,
            out d))
            {
                formattedDate = d.Date.ToShortDateString();
            }
            else if (DateTime.TryParseExact(date,
                            "yyyy/M/d",
                            CultureInfo.CurrentCulture,
                            DateTimeStyles.None,
            out d))
            {
                formattedDate = d.Date.ToShortDateString();
            }
            else if (DateTime.TryParseExact(date,
                            "yyyy-M-d",
                            CultureInfo.CurrentCulture,
                            DateTimeStyles.None,
            out d))
            {
                formattedDate = d.Date.ToShortDateString();
            }
            else if (DateTime.TryParseExact(date,
                            "mm-dd-yyyy",
                            CultureInfo.CurrentCulture,
                            DateTimeStyles.None,
            out d))
            {
                formattedDate = d.Date.ToShortDateString();
            }
            else if (DateTime.TryParseExact(date,
                            "yyyy-mm-dd",
                            CultureInfo.CurrentCulture,
                            DateTimeStyles.None,
            out d))
            {
                formattedDate = d.Date.ToShortDateString();
            }
            else
            {
                return date;
            }

            return formattedDate;
        }

        internal string Pandoc(string source)
        {
            string processName = (string) MyOptions.get_Value("PandocLocation");
            string args = String.Format(@"-r html -t mediawiki");

            var psi = new System.Diagnostics.ProcessStartInfo(processName, args);

            psi.RedirectStandardOutput = true;
            psi.RedirectStandardInput = true;

            var p = new System.Diagnostics.Process();
            p.StartInfo = psi;
            psi.UseShellExecute = false;
            p.Start();

            string outputString = "";
            byte[] inputBuffer = Encoding.UTF8.GetBytes(source);
            p.StandardInput.BaseStream.Write(inputBuffer, 0, inputBuffer.Length);
            p.StandardInput.Close();

            p.WaitForExit(2000);
            using (System.IO.StreamReader sr = new System.IO.StreamReader(
                                                   p.StandardOutput.BaseStream))
            {

                outputString = sr.ReadToEnd();
            }

            return outputString;
        }
    }
}
