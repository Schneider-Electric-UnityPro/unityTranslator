// Copyright (c) 2016 Schneider-Electric
using log4net;
using log4net.Config;
using SchneiderElectric.UnityComments;
using System;
using System.Diagnostics;
using System.IO;

namespace SchneiderElectric.UnityTranslator
{
    class Program
    {
        static Stopwatch _watch;
        private static ILog _Log = LogManager.GetLogger("UnityTranslator");
        static void Main(string[] args)
        {

            XmlConfigurator.Configure();
            UnityApplicationComments.Log = _Log;


            _Log.Info($"Starts UnityTranslator");
            int l = args.Length;

            if (l > 2)
            {
                string command = args[0]?.ToLowerInvariant();

                if (command == "extract")
                {
                    string appli = l > 1 ? Absolute(args[1]) : null;
                    string file = l > 2 ? Absolute(args[2]) : null;
                    monitore(()=>Extract(appli, file));
                }
                else if (command == "apply")
                {
                    string file  = l > 1 ? Absolute(args[1]) : null;
                    string appli = l > 2 ? Absolute(args[2]) : null;
                    monitore(() => Apply(file, appli));
                }
                else if (command == "translate")
                {
                    string file = l > 1 ? Absolute(args[1]) : null;
                    string lang = l > 2 ? (args[2]) : null;
                    monitore(() => Translate(file, lang));
                }
                else
                {
                    _Log.Error($"Error : unknown command ({command})");
                    help();
                }

            }
            else
            {
                _Log.Error($"Error :missing arguments:  command +  2 parameters  expected");
                help();
            }
            _watch?.Stop();
        }


        /// <summary>
        /// Helps this instance.
        /// </summary>
        private static void help()
        {
            //help
            Console.WriteLine("commands accepted:");
            Console.WriteLine("unityTranlator.exe Extract {source application Path} {xml file path}: read the comments of unity {source application} and write them into {xml file}");
            Console.WriteLine("unityTranlator.exe Apply  {xml file path} {target application Path}: write the comments translation from {xml file } into unity {target path");
            Console.WriteLine("unityTranlator.exe Translate {xml file path} {lang}: read the comments xml file {xml file path} translate it into {lang} and write it application back as {xml file path}.Lang.zef");
        }

        /// <summary>
        /// Extracts comments from the specified source file.
        /// </summary>
        /// <param name="sourcefile">The unity application file.</param>
        /// <param name="outputfile">The xml output file.</param>
        private static void Extract(string sourcefile, string outputfile)
        {
            UnityApplicationComments model ;
            if (CheckFile(sourcefile))
            {
                model = new UnityApplicationComments(sourcefile);
                _Log.Info($"{sourcefile} : Loaded : {model.Comments.Count} comments found");
                Backup(outputfile);
                _Log.Info($"Save results in {outputfile}");
                model.SaveXml(outputfile);
            }
        }

        /// <summary>
        /// Translates the specified file.
        /// </summary>
        /// <param name="file">The file.</param>
        /// <param name="lang">The destination language.</param>
        private static void Translate(string file, string lang)
        {

            if (CheckFile(file) && CheckLang(lang))
            {
                IUnityApplicationComments model =  UnityApplicationComments.LoadXml(file);
                if (model != null)
                {
                    string l = model.DetectLanguage();
                    if (l != lang)
                    {
                        _Log.Info($"Translate {file} : from: {l} to {lang}");
                        model.TranslateComments(lang, l, true).Wait();
                        Backup(file);
                        model.SaveXml(file);
                    }
                }
            }
        }


 

        /// <summary>
        /// Updates unity application from file.
        /// </summary>
        /// <param name="param">The file.</param>
        private static void Apply(string file, string target)
        {
            if (CheckFile(file))
            {
                IUnityApplicationComments model =  UnityApplicationComments.LoadXml(file);           
                if (model != null)
                {
                    _Log.Info($"{file} : Loaded : {model.Comments?.Count} comments found");
                    try
                    {
                        _Log.Info($"generate unity application {target}:");
                        Backup(target);
                        model.WriteTarget(target);
                    }
                    catch (Exception e)
                    {
                        _Log.Error(e);
                    }
                }
                else
                {
                    _Log.Error("No application generated");
                }
            }

        }


  
        #region utilities

        /// <summary>
        /// Backups the specified file.
        /// </summary>
        /// <param name="file">The file.</param>
        private static bool Backup(string file)
        {
            if (File.Exists(file))
            {
                string backup = Path.ChangeExtension(file, ".back");
                try
                {
                    _Log.Debug($"back up {file} => {backup} ");
                    if (File.Exists(backup))
                    {
                        File.Delete(backup);
                    }
                    File.Move(file, backup);
                    return File.Exists(backup);
                }
                catch (Exception)
                {
                    _Log.Error($"fail to back up {file} into {backup}");
                }
            }
            return false;
        }


        /// <summary>
        /// Checks the file path.
        /// </summary>
        /// <param name="param">The parameter.</param>
        /// <param name="mustExist">if set to <c>true</c> when file [must exist].</param>
        /// <returns></returns>
        private static bool CheckFile(string param, bool mustExist = true)
        {
            bool res = true;
            if (string.IsNullOrEmpty(param))
            {
                _Log.Error($"Invalid parameters");
                help();
                res = false;
            }
            else if (mustExist && !File.Exists(param))
            {
                _Log.Error($"File {param} not found");
                res = false;
            }
            return res;
        }


        /// <summary>
        /// Formats the time.
        /// </summary>
        /// <param name="watch">The watch.</param>
        /// <param name="reset">if set to <c>true</c> [reset].</param>
        /// <returns></returns>
        private static string FormatTime(Stopwatch watch, bool reset = false)
        {
            watch.Stop();
            string message = "....";
            if (_watch.Elapsed.Hours > 0)
            {
                message += string.Format(":{3} hour {0} min {1} sec {2} ms", watch.Elapsed.Minutes, _watch.Elapsed.Seconds, _watch.Elapsed.Milliseconds, _watch.Elapsed.Hours);
            }
            else if (_watch.Elapsed.Minutes > 0)
            {
                message += string.Format(":{0} min {1} sec {2} ms", watch.Elapsed.Minutes, _watch.Elapsed.Seconds, _watch.Elapsed.Milliseconds);
            }
            else
            {
                message += string.Format(":{0} sec {1} ms", watch.Elapsed.Seconds, _watch.Elapsed.Milliseconds);

            }
            if (reset)
            {
                _watch.Reset();
            }
            _watch.Start();
            return message;
        }

        /// <summary>
        /// Monitores the specified action.
        /// </summary>
        /// <param name="action">The action.</param>
        private static void monitore(Action action)
        {
            _watch = Stopwatch.StartNew();
            action.Invoke();
            Console.WriteLine($"execution took: {FormatTime(_watch)}");
        }

        /// <summary>
        /// Absolute path
        /// </summary>
        /// <param name="file">The file.</param>
        /// <returns></returns>
        private static string Absolute(string file)
        {
            return string.IsNullOrEmpty(file) ? file : Path.GetFullPath(file);
        }

        /// <summary>
        /// Checks the language.
        /// </summary>
        /// <param name="lang">The language.</param>
        /// <returns></returns>
        private static bool CheckLang(string lang)
        {
            var codes = UnityApplicationComments.Translator.LanguagesCodes;
            bool exist = (codes != null && codes.Contains(lang)) ;
            if (!exist)
            {
                _Log.Info($"{lang} : language code not found");
                if (codes != null && codes.Count>0)
                {
                    _Log.Debug("supported languages:");
                    foreach (var l in codes)
                    {
                        _Log.Debug($"{l} : {UnityApplicationComments.Translator.LanguageNameFromCode(l) }");
                    }
                }
            }
            return exist;
        }


        #endregion
    }
}
