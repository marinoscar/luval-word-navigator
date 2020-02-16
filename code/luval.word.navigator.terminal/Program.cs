using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace luval.word.navigator.terminal
{
    class Program
    {
        static void Main(string[] args)
        {
            var arguments = new ConsoleSwitches(args);
            var start = DateTime.UtcNow;
            Console.WriteLine("Started At UTC: {0}", start);
            Console.WriteLine("Reading files from {0}", arguments.DocumentDir.FullName);
            DoExecute(() =>
            {
                var files = arguments.DocumentDir.GetFiles("*.doc*", SearchOption.AllDirectories).Where(i => !i.Name.StartsWith("~")).ToList();
                var stats = new List<DocumentData>();
                foreach (var file in files)
                {
                    Console.WriteLine("File: {0} {1} of {2}", file.Name.PadRight(100), (files.IndexOf(file) + 1).ToString().PadLeft(4), files.Count.ToString().PadLeft(4));
                    var doc = new WordDocument(file.FullName);
                    DoExecute(() =>
                    {
                        stats.Add(doc.GetStats());
                    });
                }
                Console.WriteLine();
                Console.WriteLine();
                File.WriteAllText(arguments.OutputFile, JsonConvert.SerializeObject(stats));
            });

            Console.WriteLine("Completed. DURATION: {0}", DateTime.UtcNow.Subtract(start));
            Console.WriteLine("Results saved on: {0}", arguments.OutputFile);
            Console.WriteLine("Press any key to end");
            Console.ReadKey();
        }

        private static string GetRegion(string fileName)
        {
            return GetSegement(fileName, new[] { "-APAC-", "-NA-", "-GLOBAL-", "-EMEA-" });
        }

        private static string GetFunction(string fileName)
        {
            return GetSegement(fileName, new[] { "-P2P-", "-R2R-", "-O2C-" });
        }

        private static string GetSegement(string fileName, string[] options)
        {
            if (!string.IsNullOrWhiteSpace(fileName))
            {
                foreach (var opt in options)
                {
                    if (fileName.Contains(opt)) return opt.Replace("-", "");
                }
            }
            return "";
        }

        private static void DoExecute(Action action)
        {
            try
            {
                action();
            }
            catch (Exception ex)
            {
                var original = Console.ForegroundColor;
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine();
                Console.WriteLine("{0}\n{1}", ex.Message, ex.InnerException != null ? ex.InnerException.Message : "");
                Console.WriteLine();
                Console.ForegroundColor = original;
            }
        }
    }
}
