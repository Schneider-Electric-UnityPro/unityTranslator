using SchneiderElectric.UnityWrapper;
using System.IO;
using System.Threading.Tasks;
using Xunit;


namespace UnitTestScanner
{

    static class Paths
    {
        //th paths must be initialed on test machine
        // sources files for test
        private static string SourcePath = @"..\..\..\sandbox\testData";
        //temp test dir
        private static string TempPath = @"..\..\..\sandbox\testtmp";

        /// <summary>
        /// Test file path.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public static string TestFilePath(string name) => Path.GetFullPath(Path.Combine(SourcePath, name));
        /// <summary>
        /// Temporary file path.
        /// </summary>
        /// <param name="name">The name.</param>
        /// <returns></returns>
        public static string TempFilePath(string name) => Path.GetFullPath(Path.Combine(TempPath, name));



    }
    /// <summary>
    /// 
    /// </summary>
    public class TestUnity
    {

        [Fact]
        public void OpenCloseStu()
        {
            string target = PrepareBaseFile();
            using (UnityEngine test = new UnityEngine())
            {
                test.Open(Paths.TempFilePath(target));
                Assert.NotNull(test.UnityInstance);
                Assert.NotNull(test.Project);
                Assert.NotNull(test.DtmRoot);
                Assert.Equal(2, test.Masters.Count);
                test.Close();
                Assert.Null(test.Project);
                Assert.Null(test.DtmRoot);
            }
        } 

        [Fact]
        public void OpenCloseSta()
        {
            using (UnityEngine test = new UnityEngine())
            {
                test.Open(Paths.TestFilePath("base.sta"));
                Assert.NotNull(test.UnityInstance);
                Assert.NotNull(test.Project);
                Assert.NotNull(test.DtmRoot);
                Assert.True(test.Masters.Count >= 2);
                test.Close();
                Assert.Null(test.Project);
                Assert.Null(test.DtmRoot);
            }
        }


        [Fact]
        public void OpenCloseZef()
        {
            using (UnityEngine test = new UnityEngine())
            {
                test.Open(Paths.TestFilePath("base.zef"));
                Assert.NotNull(test.UnityInstance);
                Assert.NotNull(test.Project);
                Assert.NotNull(test.DtmRoot);
                Assert.True(test.Masters.Count == 2);
                test.Close();
                Assert.Null(test.Project);
                Assert.Null(test.DtmRoot);
            }
        }

        [Fact]
        public void Export()
        {
            string target = PrepareBaseFile();
            using (UnityEngine test = new UnityEngine())
            {
                test.Open(target);
                var val =test.Masters.Count;
                test.ExportProjectAs(Paths.TempFilePath("baseResult.zef"));
                test.Close();
                test.Open(Paths.TempFilePath("baseResult.zef"));
                Assert.NotNull(test.UnityInstance);
                Assert.NotNull(test.Project);
                Assert.NotNull(test.DtmRoot);
                Assert.True(test.Masters.Count == val);
                test.Close();
            }
        }

     


       
       

        #region private methods
        private static string GenerateStuFromXefOrSta(string baseFile)
        {
            var tempFile =Path.ChangeExtension(Paths.TempFilePath(Path.GetFileName(baseFile)), ".stu");
            if (!File.Exists(tempFile))
            {
                using (UnityEngine stuMaker = new UnityEngine())
                {
                    stuMaker.Open(Paths.TestFilePath(baseFile));
                    stuMaker.SaveProjectAs(Paths.TestFilePath(tempFile));
                    stuMaker.Close();
                }
            }
            return tempFile;
        }

        private static string PrepareBaseFile()
        {
            string target = GenerateStuFromXefOrSta(Paths.TestFilePath("base.zef"));
            Assert.True(File.Exists(target));
            return target;
        }

        #endregion
    }
}
