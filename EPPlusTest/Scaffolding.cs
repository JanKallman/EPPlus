using System;
using System.IO;
using System.Linq;
using System.Reflection;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace EPPlusTest
{
    /// <summary>
    /// Code that runs once per test run.
    /// </summary>
    [TestClass] // This needs to be here for AssemblyInitialize and AssemblyCleanup to be respected.
    public sealed class Scaffolding
    {

        private static string ClipartPath => Path.Combine(Path.GetTempPath(), @"EPPlus clipart");
        internal static string WorksheetPath => Path.Combine(Path.GetTempPath(), @"EPPlus worksheets");

        [AssemblyInitialize]
        public static void InitBase(TestContext ctx)
        {
            if (!Directory.Exists(ClipartPath))
            {
                Directory.CreateDirectory(ClipartPath);
            }
#if (Core)
            var asm = Assembly.GetEntryAssembly();
#else
            var asm = Assembly.GetExecutingAssembly();
#endif
            var validExtensions = new[]
            {
                ".gif", ".wmf"
            };
            foreach (var name in asm.GetManifestResourceNames())
            {
                var ext = Path.GetExtension(name);
                if (validExtensions.Contains(ext, StringComparer.CurrentCultureIgnoreCase))
                {

                    var fileName = name.Replace("EPPlusTest.Resources.", "");
                    using (var stream = asm.GetManifestResourceStream(name))
                    using (var file = File.Create(Path.Combine(ClipartPath, fileName)))
                    {
                        stream.CopyTo(file);
                    }
                }
                else
                {
                    Console.Error.WriteLine(
                        $"File name {name} does not have a valid extension {string.Join(", ", validExtensions)}");
                }
            }
            if (!Directory.Exists(WorksheetPath))
            {
                Directory.CreateDirectory(WorksheetPath);
            }
        }

        [AssemblyCleanup]
        public static void AssemblyCleanup()
        {
            Directory.Delete(WorksheetPath, true);
        }
    }
}