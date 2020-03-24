using Newtonsoft.Json;
using System;
using System.Collections.Concurrent;
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

            //ExecuteResultImport();

            ExecuteWordExtraction(arguments);

            Console.WriteLine("Completed. DURATION: {0}", DateTime.UtcNow.Subtract(start));
            Console.WriteLine("Results saved on: {0}", arguments.OutputFile);
            Console.WriteLine("Press any key to end");
            Console.ReadKey();
        }

        private static void ExecuteResultImport()
        {
            DoExecute(() => {
                var dir = new DirectoryInfo(@"C:\Users\ch489gt\Downloads\BPO-Docs\Results");
                var files = dir.GetFiles("*.json", SearchOption.AllDirectories);
                var items = new List<DocumentData>();
                foreach(var file in files)
                {
                    var i = JsonConvert.DeserializeObject<List<DocumentData>>(File.ReadAllText(file.FullName));
                    items.AddRange(i);
                }
                var processor = new ResultProcessor();
                processor.ImportToDb(items);
            });
        }

        private static void ExecuteWordExtraction(ConsoleSwitches arguments)
        {
            DoExecute(() =>
            {
                var files = arguments.DocumentDir.GetFiles("*.doc*", SearchOption.AllDirectories).Where(i => !i.Name.StartsWith("~")).ToList();
                var stats = new BlockingCollection<DocumentData>();
                var tasks = new List<Task>(arguments.ThreadCount);
                var tcount = 0;
                for (int f = 0; f < files.Count; f++)
                {
                    var file = files[f];
                    Console.WriteLine("File: {0} {1} of {2}", file.Name.PadRight(100), (files.IndexOf(file) + 1).ToString().PadLeft(4), files.Count.ToString().PadLeft(4));
                    if (tcount < arguments.ThreadCount)
                    {
                        var t = new Task(() =>
                        {
                            var doc = new WordDocument(file.FullName);
                            DoExecute(() =>
                            {
                                stats.Add(doc.GetStats());
                            });
                        });
                        tcount++;
                        tasks.Add(t);
                    }
                    if (tcount >= arguments.ThreadCount || f == (files.Count - 1)) /*is the last file*/
                    {
                        tasks.ForEach(t => t.Start());
                        tcount = 0;
                        Task.WaitAll(tasks.ToArray());
                        tasks.Clear();
                    }
                }
                Console.WriteLine();
                Console.WriteLine();
              File.WriteAllText(arguments.OutputFile, JsonConvert.SerializeObject(stats));
            });
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
