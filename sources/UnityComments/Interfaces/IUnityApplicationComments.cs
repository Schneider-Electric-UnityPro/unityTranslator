// Copyright (c) 2016 Schneider-Electric
using System.Collections.Generic;
using System.Threading.Tasks;

namespace SchneiderElectric.UnityComments
{
    public interface IUnityApplicationComments
    {
        #region properties
        /// <summary>
        /// Gets the comments.
        /// </summary>
        /// <value>
        /// The comments.
        /// </value>
        List<Comment> Comments { get; }
        /// <summary>
        /// Gets the source.
        /// </summary>
        /// <value>
        /// The source.
        /// </value>
        string Source { get; }
        #endregion

        #region method
        /// <summary>
        /// Writes the target.
        /// </summary>
        /// <param name="targetPath">The target path.</param>
        /// <returns></returns>
        bool WriteTarget(string targetPath);
        /// <summary>
        /// Reads the source.
        /// </summary>
        /// <param name="sourcePath">The source path.</param>
        /// <returns></returns>
        bool ReadSource(string sourcePath );
        /// <summary>
        /// Saves the XML.
        /// </summary>
        /// <param name="path">The path.</param>
        void SaveXml(string path);
        #endregion

        #region automatic translation
        /// <summary>
        /// Detects the language.
        /// </summary>
        /// <returns></returns>
        string DetectLanguage();
        /// <summary>
        /// Translates the comments.
        /// </summary>
        /// <param name="destlang">The destlang.</param>
        /// <param name="sourceLang">The source language.</param>
        /// <param name="eraseExisting">if set to <c>true</c> [erase existing].</param>
        /// <returns></returns>
        Task TranslateComments(string destlang = "fr", string sourceLang = "en", bool eraseExisting = false); 
        #endregion
    }

}