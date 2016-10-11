using log4net;
using SchneiderElectric.UnityWrapper;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Security.Cryptography;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;

namespace SchneiderElectric.UnityComments
{
    [System.Serializable()]
    [XmlType(TypeName = "UnityApplicationComments")]
    public class UnityApplicationComments : IUnityApplicationComments, IDisposable
    {
        #region fields
        private static ILog _log;
        private IUnityEngine _engine = new UnityEngine();
        //xef tags to translate
        private const string _comment = "comment";
        private const string _textbox = "textBox";
        private const string _stSource = "STSource";
        private const string _IlSource = "ILSource";
        //textual comments extraction rule
        const string pattern = @"((\(\*)(?<" + _comment + @">(?<!\*\)).*?)(\*\)))*";
        private static Regex _reg = new Regex(pattern, RegexOptions.Compiled | RegexOptions.Multiline);
        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="UnityApplicationComments"/> class.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="target">The target.</param>
        public UnityApplicationComments(string source)
        {
            ReadSource(source);
        }
        /// <summary>
        /// Initializes a new instance of the <see cref="UnityApplicationComments"/> class.
        /// default for serialisation
        /// </summary>
        public UnityApplicationComments()
        {

        }

        #region properties
        [XmlIgnore]
        public static ILog Log
        {
            get
            {
                return _log;
            }
            set
            {
                //propagate
                UnityEngine.Log = MicrosoftTranslator.Log = _log = value;

            }
        }

        /// <summary>
        /// Gets the translator.
        /// </summary>
        /// <value>
        /// The translator.
        /// </value>
        [XmlIgnore]
        public static ITranslator Translator { get; private set; } = new MicrosoftTranslator();

        /// <summary>
        /// Gets a value indicating whether source is valid.
        /// </summary>
        /// <value>
        ///   <c>true</c> source valid; otherwise, <c>false</c>.
        /// </value>
        [XmlIgnore]
        public bool IsSourceValid
        {
            get
            {
                if (!string.IsNullOrEmpty(Source) && File.Exists(Source))
                {
                    return !string.IsNullOrEmpty(SourceCRC);
                }
                return false;
            }
        }
        /// <summary>
        /// Gets the target.
        /// </summary>
        /// <value>
        /// The target.
        /// </value>
        [XmlAttribute("FileSource")]
        public string Source { get; set; }


        /// <summary>
        /// Gets the Working Directory.
        /// </summary>
        /// <value>
        /// The Working Directory.
        /// </value>
        [XmlIgnore]
        public string WorkingDirectory { get { return !string.IsNullOrEmpty(SourceCRC) ? Path.Combine(Path.GetTempPath(), "Translation." + SourceCRC) : string.Empty; } }

        /// <summary>
        /// Gets the xef temporary path.
        /// </summary>
        /// <value>
        /// The xef temporary path.
        /// </value>
        [XmlIgnore]
        protected string XEFTempPath => Directory.Exists(WorkingDirectory) ? Path.Combine(WorkingDirectory, "unitpro.xef") : string.Empty;


        /// <summary>
        /// Gets the xef temporary path.
        /// </summary>
        /// <value>
        /// The xef temporary path.
        /// </value>
        [XmlIgnore]
        protected string TranslateTempPath => Directory.Exists(WorkingDirectory) ? Path.Combine(WorkingDirectory, "unitpro.xef.translate") : string.Empty;


        /// <summary>
        /// Gets the zef temporary path.
        /// </summary>
        /// <value>
        /// The zef temppath.
        /// </value>
        [XmlIgnore]
        protected string ZEFTempPath => Directory.Exists(WorkingDirectory) ? WorkingDirectory + ".zef" : string.Empty;


        /// <summary>
        /// Gets or sets the source CRC.
        /// </summary>
        /// <value>
        /// The source CRC.
        /// </value>
        [XmlAttribute("CRC")]
        public string SourceCRC { get; set; }

        /// <summary>
        /// Gets the map.
        /// </summary>
        /// <value>
        /// The map.
        /// </value>
        [XmlElement("Comments", Form = System.Xml.Schema.XmlSchemaForm.Unqualified)]
        public List<Comment> Comments { get; set; } = new List<Comment>();
        #endregion

        #region methods
        /// <summary>
        /// Reads the comments.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <returns></returns>
        public bool ReadSource(string source)
        {
            bool read = false;
            SetSource(source);
            if (File.Exists(Source))
            {
                if (CopySource())
                {
                    //ZEF ready to use.
                    //unzip ZEF
                    if (Unzip(ZEFTempPath, WorkingDirectory))
                    {

                        //Load XEF
                        read = LoadAllComments();
                    }
                    else
                    {
                        //pb unzip
                        Log?.Error($"error opening source : {Source}");
                    }

                }
                else
                {
                    //pb copy
                    Log?.Error($"error preparing workspace for source : {Source}");
                }
            }
            else
            {
                //file not found
                Log?.Error($"source : {Source} does not exist");
            }
            return read;
        }

        /// <summary>
        /// Saves the result.
        /// </summary>
        /// <returns></returns>
        public bool WriteTarget(string target)
        {
            UnityfileKind extension = _engine.KindOfFile(target);
            if (extension != UnityfileKind.unknown)
            {
                if (WriteComments())
                {
                    if (extension != UnityfileKind.export)
                    {
                        //build the app
                        if (_engine.Open(ZEFTempPath))
                        {
                            Backup(target, true);
                            _engine.SaveProjectAs(Path.ChangeExtension(target, ".stu"));
                            return true;
                        }
                    }
                    else
                    {
                        //make a copy
                        Backup(target, false);
                        File.Copy(ZEFTempPath, target);
                        return true;
                    }
                }

            }
            return false;
        }

        /// <summary>
        /// Loads the XML.
        /// </summary>
        /// <param name="xmlPath">The XML path.</param>
        /// <returns> UnityApplicationComments instance or nul</returns>
        public static IUnityApplicationComments LoadXml(string xmlPath)
        {
            UnityApplicationComments cmts = XmlHelper.DeserialiseFromFile<UnityApplicationComments>(xmlPath);
            if (cmts.IsSourceValid)
            {
                Update(cmts);
            }
            else
            {
                Log?.Warn($"{cmts.Source} not valid");
            }
            return cmts;
        }

        /// <summary>
        /// Saves the XML.
        /// </summary>
        /// <param name="xmlpath">The xml path.</param>
        /// <returns></returns>
        public void SaveXml(string xmlpath)
        {
            XmlHelper.SerialiseToFile<UnityApplicationComments>(xmlpath, this);
        }

        /// <summary>
        /// Detects the language.
        /// </summary>
        /// <returns></returns>
        public string DetectLanguage()
        {

            var lang = "en";
            int count = Comments == null ? 0 : Comments.Count;

            if (count > 0)
            {
                Task.Run(async () =>
                {
                    //check with 250 char max
                    var txt = (from c in Comments.Take(Math.Min(40, count))
                               select c.Source).Aggregate((current, next) => current + ", " + next);
                    const int maxLength = 250;
                    txt = txt.Length <= maxLength ? txt : txt.Substring(0, maxLength);
                    lang = await Translator.DetectSourceLanguage(txt);
                }).Wait();
            }
            return lang;
        }

        /// <summary>
        /// Translates the comments.
        /// </summary>
        /// <param name="destlang">The destlang.</param>
        /// <param name="sourceLang">The source language.</param>
        /// <param name="eraseExisting">if set to <c>true</c> [erase existing].</param>
        /// <returns></returns>
        public async Task TranslateComments(string destlang = "fr", string sourceLang = "en", bool eraseExisting = false)
        {
            //avoid calling multiple time the translator for the same...
            var list = from x in (eraseExisting ? Comments : Comments.FindAll(x => string.IsNullOrEmpty(x.Translation)))
                       group x by x.Source into same
                       select same;
            try
            {
                const int interval = 100;
                int CurrentCount = 0;
                foreach (var group in list)
                {
                    if (!string.IsNullOrEmpty(group.Key))
                    {
                        CurrentCount++;
                        string translation = await Translator.Translate(group.Key, sourceLang, destlang);
                        foreach (var comment in group)
                        {
                            comment.Translation = translation;
                            //put a minimum delay between request 
                            //to avoid microsoft translator
                            //to stop responding (DDOS defense) 
                            if (CurrentCount % interval == 0)
                            { 
                                Log?.Info($"{CurrentCount} translations done");
                                await Task.Delay(1500);
                            }
                        }
                    }
                }
            }
            catch(Exception e)
            {
                Log?.Error(e.Message);
            }

        }

        #endregion

        #region Zip
        /// <summary>
        /// Unzips the file into dest directory
        /// </summary>
        public static bool Unzip(string file, string dest)
        {
            if (File.Exists(file))
            {
                try
                {
                    CleanDirectory(dest);
                    ZipFile.ExtractToDirectory(file, dest);
                }
                catch
                {
                    //log issue when unzipping
                    return false;
                }
                return Directory.Exists(dest);
            }
            return false;
        }

        /// <summary>
        /// Cleans the directory.
        /// </summary>
        /// <param name="dir">The dir.</param>
        public static bool CleanDirectory(string directory, bool delete = false)
        {
            bool cleaned = true;
            string[] files = Directory.GetFiles(directory);
            string[] dirs = Directory.GetDirectories(directory);

            foreach (string dir in dirs)
            {
                //recurse, force delete
                cleaned &= CleanDirectory(dir, true);
            }
            if (cleaned)
            {
                foreach (string file in files)
                {
                    File.SetAttributes(file, FileAttributes.Normal);
                    try { File.Delete(file); } catch { cleaned = false; }
                }
            }
            if (delete && cleaned)
            {
                try { Directory.Delete(directory); } catch { cleaned = false; }
            }
            return cleaned;
        }

        /// <summary>
        /// Zips the specified dir.
        /// </summary>
        /// <param name="dir">The source directory.</param>
        /// <param name="file">The target file name.</param>
        /// <returns></returns>
        public static bool Zip(string dir, string file)
        {
            if (Directory.Exists(dir))
            {
                try
                {
                    Backup(file, false);
                    ZipFile.CreateFromDirectory(dir, file);
                }
                catch
                {
                    //log issue when zipping
                    return false;
                }
                return File.Exists(file);
            }
            return false;
        }

        #endregion

        #region dispose
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        /// <summary>
        /// Releases unmanaged and - optionally - managed resources.
        /// </summary>
        /// <param name="disposing"><c>true</c> to release both managed and unmanaged resources; <c>false</c> to release only unmanaged resources.</param>
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                try
                {
                    (_engine as IDisposable)?.Dispose();

                }
                finally
                {
                    _engine = null;
                }
            }
        }
    
    #endregion

    #region private

    /// <summary>
    /// Writes the comments.
    /// </summary>
    /// <returns></returns>
    private bool WriteComments()
        {
            bool writen = false;
            string file = TranslateTempPath;
            if (File.Exists(file) && PatchXef(TranslateTempPath))
            {
                //ZEF ready to use.
                writen = Zip(WorkingDirectory, ZEFTempPath);
            }
            return writen;
        }

        /// <summary>
        /// Loads the specified file.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <returns></returns>
        private bool LoadAllComments()
        {
            bool loaded = false;
            string file = XEFTempPath;
            if (File.Exists(file))
            {
                XDocument xefXml = XDocument.Load(file);
                Comments = new List<Comment>();
                ExtractComments(_comment, xefXml);
                ExtractComments(_textbox, xefXml);
                ExtractSourceComments(_stSource, xefXml);
                ExtractSourceComments(_IlSource, xefXml);
                xefXml.Save(TranslateTempPath);
                loaded = true;
            }
            return loaded;
        }

        /// <summary>
        /// Backups the specified file.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <param name="forceStu">if set to <c>true</c> [force stu].</param>
        private static void Backup(string file, bool forceStu)
        {
            string f = forceStu ? Path.ChangeExtension(file, ".stu") : file;
            if (File.Exists(f))
            {
                string backup = Path.ChangeExtension(f, ".Save");
                if (File.Exists(backup))
                {
                    File.Delete(backup);
                }
                File.Move(f, backup);
            }
        }


        /// <summary>
        /// Patch the xef.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <returns></returns>
        private bool PatchXef(string file)
        {
            bool writen = false;
            if (File.Exists(file))
            {
                try
                {
                    string fileContents = File.ReadAllText(file);
                    foreach (Comment c in Comments)
                    {
                        if (fileContents.IndexOf(c.Key,StringComparison.InvariantCulture) >= 0)
                        {
                            fileContents = fileContents.Replace(c.Key, c.Translation == null ? c.Source : c.Translation);
                        }
                        else
                        {
                            Log?.Warn($"{c.Key} not found");
                        }
                    }
                    File.WriteAllText(XEFTempPath, fileContents);
                    writen = true;
                }
                catch (Exception e)
                {
                    Log?.Error(e.Message);
                }
            }
            return writen;
        }

        /// <summary>
        /// get the Non empty comments containers.
        /// </summary>
        /// <param name="tag">The container tag.</param>
        /// <param name="xml">The document.</param>
        /// <returns></returns>
        private static IEnumerable<XElement> NonEmptyCommentsContainer(string tag, XDocument xml)
        {
            return xml.Descendants(tag).Where(x => !string.IsNullOrEmpty(x.Value));
        }

        /// <summary>
        /// Extracts the comments.
        /// </summary>
        private void ExtractComments(string tag, XDocument xml)
        {
            foreach (XElement comment in NonEmptyCommentsContainer(tag, xml))
            {
                string context = Context(comment);
                var textNode = comment.Nodes().OfType<XText>().FirstOrDefault();
                if (textNode != null)
                {
                    textNode.Value = AddComment(textNode.Value, context);
                }
            }
        }

        /// <summary>
        /// Context of the specified comment.
        /// </summary>
        /// <param name="comment">The comment.</param>
        /// <returns></returns>
        private static string Context(XElement comment)
        {
            string context = comment.Name.ToString();
            var ancester = comment;
            while (ancester != null)
            {
                XAttribute nameAtt = null;
                while (nameAtt == null && ancester != null)
                {
                    ancester = ancester.Parent;
                    if (ancester != null)
                    {
                        nameAtt = GetAttribute(ancester);
                    }
                }
                if (ancester != null)
                {
                    context = $"({ancester.Name.ToString()})\\{context}";
                }
                if (nameAtt != null)
                {
                    context = $"{nameAtt.Value}{context}";
                }

            }

            return context;
        }

        /// <summary>
        /// Gets the attribute.
        /// </summary>
        /// <param name="element">The element.</param>
        /// <returns></returns>
        private static XAttribute GetAttribute(XElement element)
        {
            string[] remaquableAttributes = { "name", "DDTName", "nameOfFBType", "task", "partNumber" };
            XAttribute att = null;
            foreach (string a in remaquableAttributes)
            {
                att = element?.Attribute(a);
                if (att != null)
                {
                    break;
                }
            }
            return att;
        }

        /// <summary>
        /// Extracts the source comments.
        /// </summary>
        /// <param name="comments">The comments.</param>
        /// <param name="map">The map.</param>
        private void ExtractSourceComments(string tag, XDocument xml)
        {

            var comments = NonEmptyCommentsContainer(tag, xml);
            int count = comments.Count();

            foreach (XElement element in comments)
            {
                //because of regex matching issue, need to remove the \r and \n
                //from original string, these new line chars will be lost
                string text = element.Value.Replace('\r', ' ').Replace('\n', ' ');
                string output = text;
                string context = Context(element);
                MatchCollection matchCollection = _reg.Matches(text);
                var i = 0;
                foreach (Match match in matchCollection)
                {
                    var group = match.Groups[_comment];
                    if (group != null)
                    {
                        foreach (Capture c in group.Captures)
                        {
                            i++;
                            output = output.Replace(c.Value, AddComment(c.Value, $"{context} code comment [{i}]"));
                        }
                    }
                }
                if (output != text)
                {
                    element.Value = output;
                }
            }
        }


        /// <summary>
        /// Adds the comment.
        /// </summary>
        /// <param name="value">The value.</param>
        /// <param name="context">The context.</param>
        /// <returns></returns>
        private string AddComment(string value, string context)
        {
            Comment c = new Comment(value, context);
            Comments.Add(c);
            return c.Key;
        }

        /// <summary>
        /// Copies the source.
        /// </summary>
        /// <returns></returns>
        private bool CopySource()
        {
            InitWorkSpace();
            UnityfileKind extension = _engine.KindOfFile(Source);
            if (extension != UnityfileKind.unknown)
            {
                string zef = ZEFTempPath;
                if (extension != UnityfileKind.export)
                {
                    //build the zef
                    try
                    {
                        if (_engine.Open(Source, true))
                        {
                            _engine.ExportProjectAs(zef);
                            return true;
                        }
                    }
                    catch(Exception e)
                    {
                        Log.Error(e.Message);
                    }
                }
                else
                {
                    try
                    {
                        //make a copy
                        Backup(zef, false);
                        File.Copy(Source, zef);
                        return true;
                    }
                    catch (Exception e)
                    {
                        Log.Error(e.Message);
                    }
                }
            }
            return false;
        }

        /// <summary>
        /// Initializes the work space.
        /// </summary>
        private void InitWorkSpace()
        {
            string wd = WorkingDirectory;
            if (!Directory.Exists(wd) && !string.IsNullOrEmpty(wd))
            {
                Directory.CreateDirectory(wd);
            }
        }

        /// <summary>
        /// Gets the checksum.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <returns></returns>
        private static string GetChecksum(string file)
        {
            using (FileStream stream = File.OpenRead(file))
            {
                var sha = new SHA256Managed();
                byte[] checksum = sha.ComputeHash(stream);
                return BitConverter.ToString(checksum).Replace("-", String.Empty);
            }
        }

        /// <summary>
        /// Sets the source.
        /// </summary>
        /// <param name="source">The source.</param>
        private void SetSource(string source)
        {
            if (source != null && File.Exists(source))
            {
                Source = source;
                SourceCRC = GetChecksum(Source);
            }
        }

        /// <summary>
        /// Updates the comments from source 
        /// called when the working directory has beedn recreated from source 
        /// or the checksum of the source is differnet
        /// </summary>
        /// <param name="cmts">The CMTS.</param>
        private static void Update(UnityApplicationComments cmts)
        {
            List<Comment> savecomments = cmts.Comments;
            cmts.Comments = null;
            var originalXmlCRC = cmts.SourceCRC;
            cmts.ReadSource(cmts.Source);
            string curCks = cmts.SourceCRC;//modified during readsource
            if (originalXmlCRC == curCks)
            {
                //unmodified source
                //the working dir has been recovered, id of the commennts have changed.. need to re update them with the translation saved in xml
                Log?.Debug($"Update comments mapping for {cmts.Source}");
                if (cmts.Comments.Count == savecomments.Count)
                {
                    int i = 0;
                    foreach (Comment c in cmts.Comments)
                    {
                        c.Translation = savecomments[i].Translation;
                        i++;
                    }
                }
            }
            else
            {
                int cmtN = cmts.Comments.Count;
                Log?.Warn($"the source file {cmts.Source} has been modified since last tranlation. update... translations");
                int mapped = 0;
                foreach (Comment c in cmts.Comments)
                {
                    var oldcmt = savecomments.FirstOrDefault(x => x.Source == c.Source);
                    if (oldcmt != null)
                    {
                        c.Translation = oldcmt.Translation;
                        mapped++;
                    }
                }
                if (mapped != cmtN)
                {
                    Log?.Warn($"translation partialy matched {mapped} / {cmtN} ({mapped * 100 / cmtN}%)");
                }
            }
        }
        #endregion
    }
}
