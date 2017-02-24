// Copyright (c) 2016 Schneider-Electric
using SchneiderElectric.UnityWrapper;
using System.IO;
using System.Threading.Tasks;
using Xunit;


namespace UnitTest
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
    
}
