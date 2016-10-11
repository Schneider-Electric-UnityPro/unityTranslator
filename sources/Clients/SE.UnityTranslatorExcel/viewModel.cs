// Copyright (c) 2016 Schneider-Electric
using log4net;
using SchneiderElectric.UnityComments;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace SE.UnityCommentsExcel
{
    /// <summary>
    /// viewmodel for excel addins
    /// </summary>
    public class ExcelViewModel : IDisposable
    {
        #region fields
        private CommentTable _table;
        private IUnityApplicationComments _model;
        private object _locker = new object();
        private static ILog _log;
        private List<string> _lang;
        private static string _targetLang ;
        #endregion

        #region properties

        /// <summary>
        /// Gets or sets a value indicating whether [translate selection only].
        /// </summary>
        /// <value>
        /// <c>true</c> if [translate selection only]; otherwise, <c>false</c>.
        /// </value>
        public static bool TranslateSelectionOnly { get; set; } = false;
        /// <summary>
        /// Gets or sets a value indicating whether [erase existing translations].
        /// </summary>
        /// <value>
        ///   <c>true</c> if [erase existing translations]; otherwise, <c>false</c>.
        /// </value>
        public static bool EraseExistingTranslations { get; set; } = false;

        /// <summary>
        /// Gets or sets the logger.
        /// </summary>
        /// <value>
        /// The log.
        /// </value>
        public static ILog Log
        {
            get
            {
                return _log;
            }
            set
            {
                //propagate
                UnityApplicationComments.Log = _log = value;
                CommentTable.Log = value;
                CommentRow.Log = value;
            }
        }
        /// <summary>
        /// Gets the model.
        /// </summary>
        /// <value>
        /// The model.
        /// </value>
        public IUnityApplicationComments Model
        {
            get
            {
                lock (_locker)
                {
                    return _model;
                }
            }
            private set
            {
                lock (_locker)
                {
                    _model = value;
                }
            }
        }

        /// <summary>
        /// Gets the table.
        /// </summary>
        /// <value>
        /// The table.
        /// </value>
        public CommentTable Table
        {
            get
            {
                lock (_locker)
                {
                    return _table;
                }
            }

            set
            {
                lock (_locker)
                {
                    if (_table != null)
                    {
                        _table.Dispose();
                    }
                    _table = value;
                }
            }
        }

        /// <summary>
        /// Gets the languages.
        /// </summary>
        /// <value>
        /// The languages.
        /// </value>
        public List<string> Languages
        {
            get
            {
                if(_lang == null)  
                {
                    _lang = UnityApplicationComments.Translator.LanguagesNames;
                }
                return _lang;
            }
        }

        
        #endregion

        /// <summary>
        /// Initializes a new instance of the <see cref="ExcelViewModel"/> class.
        /// </summary>
        internal ExcelViewModel()
        {
            _table = new CommentTable();
            Model = new UnityApplicationComments();
        }
        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            Table = null;
            Model = null;
        }

        #region public methods


        /// <summary>
        /// Sets the target language.
        /// </summary>
        /// <param name="text">The text.</param>
        public static void SetTargetLang(string text)
        {
            if(UnityApplicationComments.Translator.LanguagesNames.Contains(text))
            {
                _targetLang = UnityApplicationComments.Translator.LanguageCodeFromName(text);
            }
        }


        /// <summary>
        /// Translates the comments to the targeted lang.
        /// </summary>
        /// <param name="ids">The ids.</param>
        /// <returns></returns>
        public async Task Translate(List<string> ids = null)
        {
            if (_targetLang != null)
            {
                var detected = Model.DetectLanguage();
                IUnityApplicationComments toTranslate = Model;
                bool filterselection = (ids != null && ids.Count > 0 && TranslateSelectionOnly);
                if (filterselection)
                {
                    //extract selected comments
                    toTranslate = new UnityApplicationComments();
                    toTranslate.Comments.AddRange (from comment in Model.Comments
                                         where ids.Contains(comment.Key)
                                         select comment);
                }
                await toTranslate.TranslateComments(_targetLang, detected,EraseExistingTranslations);
                if (filterselection)
                {
                    // re-inject translations 
                    foreach (var c in toTranslate.Comments)
                    {
                        var m = Model.Comments.FirstOrDefault(x => x.Key == c.Key);
                        if (m != null)
                        {
                            m.Translation = c.Translation;
                        }
                    }
                }
                Table = await TablefromModel(Model);
            }
        }

        /// <summary>
        /// Partials the translate.
        /// </summary>
        /// <param name="ids">The ids.</param>
        /// <returns></returns>
        public async Task<List<string>> PartialTranslate(List<string> ids)
        {
            List<string> translations = null;
            if (_targetLang != null)
            {
                bool filterselection = (ids != null && ids.Count > 0);
                if (filterselection)
                {
                    //extract selected comments
                    var toTranslate = new UnityApplicationComments();
                    toTranslate.Comments.AddRange(from comment in Model.Comments
                                                  where ids.Contains(comment.Key)
                                                  select comment);

                    await toTranslate.TranslateComments(_targetLang, toTranslate.DetectLanguage(), true);
                    translations =  (from x in toTranslate.Comments select x.Translation) .ToList();
                }
            }
            return translations;
        }

        /// <summary>
        /// Open unity application and loads the comments
        /// </summary>
        public async Task<bool> Open(string filename)
        {
            if (!string.IsNullOrEmpty(filename))
            {
                Model = new UnityApplicationComments(filename);
                Table = await TablefromModel(Model);
                return Table != null;
            }
            return false;
        }

        /// <summary>
        /// loads xml
        /// </summary>
        public async Task<bool> Import(string filename)
        {
            if (!string.IsNullOrEmpty(filename))
            {
                Model =  UnityApplicationComments.LoadXml(filename);
                Table = await TablefromModel(Model);
                return Table != null;
            }
            return false;
        }

        /// <summary>
        /// saves xml file from table.
        /// </summary>
        public async Task<bool> Export(string filename)
        {
            if (Table != null && !string.IsNullOrEmpty(filename))
            {
                Model = await ModelFromTable(Table);
                Model.SaveXml(filename);
                Log.Info($"Scan model saved in {filename}");
                return true;
            }
            return false;
        }



        /// <summary>
        /// saves xml file from table.
        /// </summary>
        public async Task<bool> Save(string filename)
        {
            if (Table != null && !string.IsNullOrEmpty(filename))
            {
                Model = await ModelFromTable(Table);
                Model.WriteTarget(filename);
                Log.Info($"Appication updated {filename}");
                return true;
            }
            return false;
        }




        /// <summary>
        /// create Model from table.
        /// </summary>
        /// <returns></returns>
        public async Task<IUnityApplicationComments> ModelFromTable(CommentTable table)
        {
            UnityApplicationComments modelfromTable = NewModel();
            await Task.Run(() =>
            {
                lock (_locker)
                {
                    foreach (CommentRow row in table.Rows)
                    {
                        var c = new Comment();
                        c.Source = row.Comment;
                        c.Context = row.Context;
                        c.Key = row.ID;
                        c.Translation = row.Translation;
                        modelfromTable.Comments.Add(c);
                    }
                }
            });
            return modelfromTable;
        }

       

        /// <summary>
        /// Builds the table.
        /// </summary>
        public async Task<CommentTable> TablefromModel(IUnityApplicationComments model)
        {
            CommentTable table = new CommentTable();
            await Task.Run(() =>
                {
                    lock (_locker)
                    {
                        if (model != null)
                        {
                            foreach (Comment c in model.Comments)
                            {
                                var row = table.CreateNewRow();
                                row.ID = c.Key;
                                row.Comment = c.Source;
                                row.Context = c.Context;
                                row.Translation = c.Translation;
                                table.Rows.Add(row);
                            }
                        }
                    }
                });
            return table;
        }

        #endregion

        #region private methods
        /// <summary>
        /// News the model.
        /// </summary>
        /// <returns></returns>
        private UnityApplicationComments NewModel()
        {
            UnityApplicationComments modelfromTable = new UnityApplicationComments();
            modelfromTable.Source = Model?.Source;
            modelfromTable.SourceCRC = ((UnityApplicationComments)Model)?.SourceCRC;
            return modelfromTable;
        }
        #endregion
    }
}
