using NUnit.Framework;
using NetOffice.WordApi;
using NUnit.Framework.Internal;
using System.IO;
using System;

namespace ProtectedRangeSearch
{
    public class Tests
    {
        Application app;
        Document doc;
        const string docName = "Naive Bayes classifier.docx";

        [Test]
        [TestCase("assume that the value")]
        [TestCase("based on a common")]
        [TestCase("diameter features")]
        public void FindTextTest(string searchText)
        {
            using Range docRange = doc.Content.Duplicate;

            foreach (var paragraph in docRange.Paragraphs)
            {
                using Range paragraphRange = paragraph.Range;
                var text = paragraphRange.Text;
                var startParagraph = paragraphRange.Start;
                var endParagraph = paragraphRange.End;

                var startIndex = text.IndexOf(searchText);
                if (startIndex >= 0)
                {
                    text = GetParagraphTextWithHiddenSymbols(paragraphRange, text);
                    startIndex = text.IndexOf(searchText);
                    var startFoundRange = startParagraph + startIndex;
                    var end = startFoundRange + searchText.Length;

                    paragraphRange.SetRange(startFoundRange, end);

                    var foundText = paragraphRange.Text;
                    Assert.AreEqual(searchText, foundText);
                }
            }
        }

        private static string GetParagraphTextWithHiddenSymbols(Range paragraphRange, string initialText)
        {
            var text = initialText;
            foreach (Field field in paragraphRange.Fields)
            {
                int index = text.IndexOf(field.Result.Text);
                if (index >= 0)
                {
                    text = text.Replace(field.Result.Text, $"{{{field.Code.Text}}} {field.Result.Text}{(char)21}");
                }
            }
            return text;
        }

        [SetUp]
        public void Open()
        {
            app = new Application();
            app.Visible = true;
            doc = app.Documents.Open(Path.Combine(Environment.CurrentDirectory,"Resources", docName));
        }

        [TearDown]
        public void Close()
        {
            doc.Close();
            doc.Dispose();
            app.Quit();
            app.Dispose();
        }
    }
}