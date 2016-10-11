// Copyright (c) 2016 Schneider-Electric

using SchneiderElectric.UnityWrapper;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System;
using SchneiderElectric.UnityComments;
using System.Threading.Tasks;
using Xunit;


namespace UnitTestScanner
{
    /// <summary>
    /// 
    /// </summary>
    public class TestXmlTranslator
    {

      
        /// <summary>
        /// Automatic translation.
        /// </summary>
        [Fact]
        public async Task AutoTranslate()
        {
            var code = UnityApplicationComments.Translator.LanguagesCodes;
            var lang = UnityApplicationComments.Translator.LanguagesNames;
            Assert.True(code.Count == lang.Count);
            Assert.True(lang.Any());
            int codeLang = PickLang(code);

            await Translate(Paths.TestFilePath("base.zef"),
                                Paths.TempFilePath(lang[codeLang] + "base.zef"), code[codeLang]);

            await Translate(Paths.TestFilePath("unityv11-comment.zef"),
                                 Paths.TempFilePath(lang[codeLang]+"unityv11-comment.zef"), code[codeLang]);
        }

        /// <summary>
        /// write comments in XMLformat
        /// </summary>
        [Fact]
        public async Task DetectLanguage()
        {
            MicrosoftTranslator mt= new MicrosoftTranslator();
            string detection = await mt.DetectSourceLanguage("il etait une fois un ...");
            Assert.True(detection == "fr");
            UnityApplicationComments CommentManager = new UnityApplicationComments(Paths.TestFilePath("base.zef"));
            int count = CommentManager.Comments == null? 0: CommentManager.Comments.Count;
            Assert.True(count > 0);
            detection = CommentManager.DetectLanguage();
            Assert.True(detection == "en");
            CommentManager.Dispose();

        }

        [Fact]
        public void Languages()
        {
            MicrosoftTranslator mt = new MicrosoftTranslator();
            string detection = "en";
            Assert.True(mt.LanguagesNames.Count == mt.LanguagesCodes.Count);
            Assert.True(mt.LanguagesNames.Count >0);
            Assert.True(mt.LanguageCodeFromName(mt.LanguageNameFromCode(detection)) == detection);

        }
        /// <summary>
        /// write comments in XMLformat
        /// </summary>
        [Fact]
        public void WriteXmlComments()
        {
            UnityApplicationComments CommentManager = new UnityApplicationComments(Paths.TestFilePath("base.zef") );
            int count = CommentManager.Comments.Count;
            Assert.True(count > 0);
            CommentManager.TranslateComments("ar").Wait();
            Assert.True(CommentManager.WriteTarget(Paths.TempFilePath("Xmlbase.zef")));
            CommentManager.SaveXml(Paths.TempFilePath("Xmlcomments.xml"));
            CommentManager.Dispose();

        }

        /// <summary>
        /// read XMLformat of the comments.
        /// </summary>
        [Fact]
        public void ReadXmlComments()
        {
            string tstfile = Paths.TestFilePath("Xmlcomments.xml");
            Assert.True(File.Exists(tstfile));
            using (UnityApplicationComments CommentManager = UnityApplicationComments.LoadXml(tstfile) as UnityApplicationComments)
            {
                int count = CommentManager.Comments == null ? 0 : CommentManager.Comments.Count;
                Assert.True(count > 0);
                Assert.True(CommentManager.WriteTarget(Paths.TempFilePath("Xmlcomments.zef")));
            }
        }

        /// <summary>
        /// translate XMLformat of the comments.
        /// </summary>
        [Fact]
        public async Task TranslateXmlComments()
        {
            string tstfile = Paths.TestFilePath("Xmlcomments.xml");
            Assert.True(File.Exists(tstfile));
            using (UnityApplicationComments CommentManager = UnityApplicationComments.LoadXml(tstfile) as UnityApplicationComments)
            {
                int count = CommentManager.Comments == null ? 0 : CommentManager.Comments.Count();
                Assert.True(count > 0);
                await CommentManager.TranslateComments("ar");
                foreach (var c in CommentManager.Comments)
                {
                    Assert.False(string.IsNullOrEmpty(c.Translation));
                }

                Assert.True(CommentManager.WriteTarget(Paths.TempFilePath("Xmlcomments.zef")));
            }
        }

        /// <summary>
        /// Valid  commment source?.
        /// </summary>
        [Fact]
        public void ValidCommmentSource()
        {
            string tstfile = Paths.TestFilePath("Xmlcomments.xml");
            Assert.True(File.Exists(tstfile));
            UnityApplicationComments CommentManager =(UnityApplicationComments) UnityApplicationComments.LoadXml(tstfile);
            Assert.True(CommentManager.IsSourceValid);
            CommentManager = new UnityApplicationComments();
            Assert.False(CommentManager.IsSourceValid);
            CommentManager = new UnityApplicationComments(Paths.TestFilePath("base.zef"));
            Assert.True(CommentManager.IsSourceValid);
            CommentManager.SourceCRC = null;
            Assert.False(CommentManager.IsSourceValid);
        }




        /// <summary>
        /// tests of xml loading
        /// </summary>
        [Fact]
        public void LoadXml()
        {
            using (UnityApplicationComments CommentManager = new UnityApplicationComments(Paths.TestFilePath("base.zef")))
            {
                Assert.True(CommentManager.IsSourceValid);

                string crc = CommentManager.SourceCRC;
                string working = CommentManager.WorkingDirectory;
                string xmlnok = Paths.TempFilePath("wrongcrcbasecomments.xml");
                string xmlok = Paths.TempFilePath("correctcrcbasecomments.xml");
                CommentManager.SourceCRC = crc + "FFF"; //make sure is <>
                CommentManager.SaveXml(xmlnok);
                CommentManager.SourceCRC = crc; //make sure is ==
                CommentManager.SaveXml(xmlok);

                //first test .... delete the temp dir and reload
                UnityApplicationComments.CleanDirectory(working, true);
                using (UnityApplicationComments CommentManager2 = (UnityApplicationComments)UnityApplicationComments.LoadXml(xmlok))
                {
                    Assert.True(CommentManager.SourceCRC == CommentManager2.SourceCRC);
                    Assert.True(ComparComments(CommentManager, CommentManager2));
                    Assert.True(CommentManager2.WriteTarget(Paths.TempFilePath("base2.zef")));
                }
                //second test with <> crc
                using (var CommentManager3 = (UnityApplicationComments)UnityApplicationComments.LoadXml(xmlnok))
                { 
                    Assert.True(CommentManager.SourceCRC == CommentManager3.SourceCRC);//has been updated on the fly
                    Assert.True(ComparComments(CommentManager, CommentManager3));
                    Assert.True(CommentManager3.WriteTarget(Paths.TempFilePath("base3.zef")));
                }
            }

        }

        private static bool ComparComments(UnityApplicationComments CommentManager, UnityApplicationComments CommentManager2)
        {
            bool res = CommentManager2.Comments.Count == CommentManager.Comments.Count;
            Assert.True(res);
            int count = CommentManager.Comments == null ? 0 : CommentManager.Comments.Count();

            for (int c = 0; c < count; c++)
            {
                var comment1 = CommentManager.Comments[c];
                var comment2 = CommentManager2.Comments[c];
                Assert.True(res&= comment1.Context == comment2.Context, $"[{comment2.Context}] <> [{comment1.Context}]");
                Assert.True(res &=comment2.Source == comment1.Source, $"[{comment2.Source}] <> [{comment1.Source}]");
                Assert.True(res&=comment1.Translation == comment2.Translation, $"[{comment2.Translation}] <> [{comment1.Translation}]");
            }
            return res;
        }


        /// <summary>
        /// Picks the language.
        /// </summary>
        /// <param name="code">The code.</param>
        /// <returns></returns>
        private static int PickLang(List<string> code)
        {
            var rnd = new Random();
            int i = rnd.Next(code.Count);
            while (code[i] == "en")
            {
                i = rnd.Next(code.Count);
            }

            return i;
        }

        /// <summary>
        /// Translates the specified source.
        /// </summary>
        /// <param name="source">The source.</param>
        /// <param name="dest">The dest.</param>
        private static async Task Translate(string source, string dest, string lang)
        {
            using (UnityApplicationComments CommentManager = new UnityApplicationComments())
            {
                Assert.True(CommentManager.ReadSource(source));
                int count = CommentManager.Comments == null ? 0 : CommentManager.Comments.Count();
                CommentManager.SaveXml(dest + ".source.xml");
                Assert.True(count > 0);
                var detected = CommentManager.DetectLanguage();
                Assert.True(detected == "en");
                await CommentManager.TranslateComments(lang, detected);
                CommentManager.SaveXml(dest + ".translated.xml");
                Assert.True(CommentManager.WriteTarget(dest));

                foreach (var c in CommentManager.Comments)
                {
                    Assert.False(string.IsNullOrEmpty(c.Translation));
                }
            }

        }



    }
}
