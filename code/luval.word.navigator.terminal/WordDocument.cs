using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;


namespace luval.word.navigator.terminal
{
    public class WordDocument
    {
        public WordDocument(string fileName)
        {
            DocumentFile = new FileInfo(fileName);
        }

        public FileInfo DocumentFile { get; private set; }

        public int GetImageCount()
        {
            var count = 0;
            ProcessDocument((doc) =>
            {
                count = GetImageCount(doc);
            });
            return count;
        }

        public DocumentData GetStats()
        {
            var res = new DocumentData();
            ProcessDocument((doc) =>
            {
                var regExRes = DoRegExSearches(doc);
                res = new DocumentData()
                {
                    Systems = GetSystems(doc),
                    FileName = DocumentFile.Name,
                    ImageCount = GetImageCount(doc),
                    Frequency = GetFrequency(doc),
                    Country = GetCountry(doc),
                    Urls = regExRes.Urls,
                    SAPTransactionCodes = regExRes.SAPTransactionCodes
                };
            });
            return res;
        }

        public void WithDocParragraph(Document doc, Paragraph paragraph, DocumentData docData)
        {

        }

        public int GetImageCount(Document doc)
        {
            var count = 0;
            foreach (var shape in doc.InlineShapes.Cast<InlineShape>())
            {
                shape.Range.Select();
                if (shape.Type == WdInlineShapeType.wdInlineShapePicture ||
                   shape.Type == WdInlineShapeType.wdInlineShapeLinkedPicture ||
                   shape.Type == WdInlineShapeType.wdInlineShapePictureHorizontalLine ||
                   shape.Type == WdInlineShapeType.wdInlineShapePictureBullet ||
                   shape.Type == WdInlineShapeType.wdInlineShapeLinkedPictureHorizontalLine)
                {
                    count++;
                }
            }
            return count;
        }

        public string GetDocumentStatus()
        {
            var res = "";
            ProcessDocument((doc) =>
            {
                var parragraphs = doc.Paragraphs.Cast<Paragraph>().Take(40);
                foreach (var pa in parragraphs)
                {
                    pa.Range.Select();
                    if (!string.IsNullOrWhiteSpace(pa.Range.Text) &&
                         pa.Range.Text.ToLowerInvariant().Contains("status"))
                    {
                        if (pa.Range.Text.ToLowerInvariant().Contains("approved"))
                            res = "Approved";
                        else if (pa.Range.Text.ToLowerInvariant().Contains("in progress"))
                        {
                            res = "In Progress";
                        }
                        else res = "Inactive";
                    }
                }
            });
            return res;
        }

        public string GetDocumentStatus(Document doc)
        {
            var res = "";
            var parragraphs = doc.Paragraphs.Cast<Paragraph>().Take(40);
            foreach (var pa in parragraphs)
            {
                pa.Range.Select();
                if (!string.IsNullOrWhiteSpace(pa.Range.Text) &&
                     pa.Range.Text.ToLowerInvariant().Contains("status"))
                {
                    if (pa.Range.Text.ToLowerInvariant().Contains("approved"))
                        res = "Approved";
                    else if (pa.Range.Text.ToLowerInvariant().Contains("in progress"))
                    {
                        res = "In Progress";
                    }
                    else res = "Inactive";
                    break;
                }
            }
            return res;
        }


        public string GetSystems(Document doc)
        {
            var apps = FindTextBetweenHeadings(doc, @"system|application", "heading 1", "heading 1");
            if (string.IsNullOrWhiteSpace(apps)) return "";
            return CleanString(apps);
        }

        public string GetCountry(Document doc)
        {
            var apps = FindTextBetweenHeadings(doc, @"country", "heading 1", "heading 1");
            if (string.IsNullOrWhiteSpace(apps)) return "";
            return CleanString(apps);
        }

        public string GetFrequency(Document doc)
        {
            var apps = FindTextBetweenHeadings(doc, @"FREQUENCY".ToLowerInvariant(), "heading 1", "heading 1");
            if (string.IsNullOrWhiteSpace(apps)) return "";
            return CleanString(apps);
        }

        private string CleanString(string value)
        {
            return value.Trim().Replace("\n\r", ";")
                .Replace("\r\n", ";")
                .Replace("\n", ";")
                .Replace("\r", ";")
                .Replace(";;", ";")
                .Replace("-", "")
                .Replace("_", "")
                .Replace("\t", "")
                .Replace("  ", " ")
                .Replace("   ", " ")
                .Replace("   ", " ")
                .Replace("   ", " ");
        }

        public DocumentData DoRegExSearches(Document doc)
        {
            var urls = new List<string>();
            var tcodes = new List<string>();
            const string tcodePattern = @"(\st*.(-|_| )*code\s)|(\stransaction*.(-|_| )*code\s)";
            const string urlPattern = @"(?:vnc|s3|ssh|scp|sftp|ftp|http|https)\:\/\/[\w\.]+(?:\:?\d{0,5})|(?:mailto|)\:[\w\.]+\@[\w\.]+";
            
            FindTCodeAndUrls(doc.Content.Text, tcodePattern, urlPattern, tcodes, urls);
            return new DocumentData() {
                SAPTransactionCodes = string.Join(";", tcodes.Distinct()),
                Urls = string.Join(";", urls.Distinct())
            };
        }

        public void FindTCodeAndUrls(string text, string tcodePattern, string urlPattern, List<string> tcodes, List<string> urls)
        {
            var tcodeResult = FindTextInParragraph(text, tcodePattern);
            if (tcodeResult != null && tcodeResult.Result.Any(i => i.Success))
            {
                tcodes.AddRange(ExtractTCodeFromText(tcodeResult));
            }
            var urlResult = FindTextInParragraph(text, urlPattern);
            if (urlResult != null && urlResult.Result.Any(i => i.Success))
            {
                urls.AddRange(urlResult.Result.Select(i => i.Value));
            }
        }

        private List<string> ExtractTCodeFromText(RegExResult result)
        {
            var res = new List<string>();
            foreach(var r in result.Result)
            {
                var sub = CleanString(result.Input.Substring(r.Index + r.Length, 75));
                res.Add(GetWords(sub).FirstOrDefault());
            }
            return res;
        }

        private string CleanText(string text)
        {
            //remove duplicate spaces
            RegexOptions options = RegexOptions.None;
            Regex regex = new Regex("[ ]{2,}", options);
            text = regex.Replace(text, " ");
            return text.Replace(";", "").Replace(":", "").Replace(".", "").Replace(",", "");
        }

        private List<string> GetWords(string text)
        {
            return Regex.Matches(text, @"\b[\w']*\b").Cast<Match>().Where(i => i.Success).Select(i => i.Value).ToList();
        }

        public RegExResult FindTextInParragraph(Paragraph paragraph, string pattern)
        {
            paragraph.Range.Select();
            var result = Regex.Matches(paragraph.Range.Text, pattern, RegexOptions.IgnoreCase);
            return new RegExResult() { Input = paragraph.Range.Text, Result = result.Cast<Match>().ToList() };
        }

        public RegExResult FindTextInParragraph(string text, string pattern)
        {
            var result = Regex.Matches(text, pattern, RegexOptions.IgnoreCase);
            return new RegExResult() { Input = text, Result = result.Cast<Match>().ToList() };
        }

        public string FindTextBetweenHeadings(Document doc, string pattern, string styleStart, string styleEnd)
        {
            var sw = new StringWriter();
            var start = false;
            foreach (var pa in doc.Paragraphs.Cast<Paragraph>().ToList())
            {
                pa.Range.Select();
                var text = pa.Range.Text.Trim().ToLowerInvariant();
                if (string.IsNullOrWhiteSpace(text)) continue;
                var style = (pa.Range.ParagraphStyle as Style).NameLocal.ToLowerInvariant();
                if (!start && !string.IsNullOrWhiteSpace(style) && style.Equals(styleStart.ToLowerInvariant()) && Regex.IsMatch(text, pattern))
                {
                    start = true;
                    continue;
                }
                if (start && !string.IsNullOrWhiteSpace(style) && style.Contains(styleEnd.ToLowerInvariant()))
                {
                    start = false;
                    break;
                }
                if (start)
                {
                    sw.WriteLine(pa.Range.Text);
                }
            }
            return sw.ToString();
        }

        public List<SearchResult> FindInParragraph(Document doc, List<Func<Paragraph, SearchResult>> searchFuncs)
        {
            var res = new List<SearchResult>();
            var completed = new List<Func<Paragraph, SearchResult>>();
            foreach (var pa in doc.Paragraphs.Cast<Paragraph>().ToList())
            {
                foreach (var call in searchFuncs)
                {
                    if (!completed.Contains(call))
                    {
                        var result = call(pa);
                        if (result.Completed)
                        {
                            res.Add(result);
                            completed.Add(call);
                        }
                    }
                    if (completed.Count == searchFuncs.Count) break;
                }
            }
            return res;
        }

        public void ProcessDocument(Action<Document> action)
        {
            var wordApp = new Word.Application();
            wordApp.Documents.Open(DocumentFile.FullName);
            try
            {
                action(wordApp.ActiveDocument);
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to process the file", ex);
            }
            finally
            {
                wordApp.ActiveDocument.Close();
                wordApp.Quit();
            }
        }
    }
}
