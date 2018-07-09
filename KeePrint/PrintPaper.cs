using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace KeePrint
{
    class PrintPaper
    {
        private string _username;
        private string _password;
        private string _notes;
        private string _docpath;

        public PrintPaper(string username, string password, string notes, string path)
        {
            _username = username;
            _password = password;
            _notes = notes;
            _docpath = path;
        }

        public void PrintDoc()
        {
            try
            {
                Word.Application word = new Word.Application();
                Word.Document document = new Word.Document();
                document = word.Documents.Open(_docpath, ReadOnly: false, Visible: false);
                document.Activate();

                Word.Find find = word.Selection.Find;
                find.ClearFormatting();
                find.Replacement.ClearFormatting();
                object replaceAll = Word.WdReplace.wdReplaceAll;
                object missing = System.Type.Missing;

                BookmarkSubroutine(document, "username", _username, 0);
                BookmarkSubroutine(document, "password", _password, 0);
                BookmarkSubroutine(document, "notes", _notes, 0);

                document.PrintOut();
                document.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
                word.Application.Quit();
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message);
            }
        }

        public static void BookmarkSubroutine(Word.Document Doc, String vBookmark, String vText, System.Int32 vBold)
        {
            Dictionary<string, int> names = new Dictionary<string, int>();


            // find bookmark
            var bm = Doc.Bookmarks[vBookmark];

            //// get COM object reference
            object rng = bm.Range;

            //// get original bookmark start position and name
            var bmStart = bm.Range.Start;
            var bookmarkName = bm.Name;

            // replace bookmark with new text
            bm.Range.Text = vText;

            // calculate bookmark range (use original bookmark start position, add length of new contents 
            ((Microsoft.Office.Interop.Word.Range)rng).Start = bmStart;
            if (vBookmark != "fAbladestelleE" && vText.Length > 1 && vBookmark != "fzHd")
            {
                ((Microsoft.Office.Interop.Word.Range)rng).End = bmStart + vText.Length;
            }
            else if ((vBookmark == "fNrE" || vBookmark == "fNr" || vBookmark == "fTel" || vBookmark == "fTel2") && vText.Length == 1)
            {
                ((Microsoft.Office.Interop.Word.Range)rng).End = bmStart + vText.Length;
            }

            else
            {
                ((Microsoft.Office.Interop.Word.Range)rng).End = bmStart + vText.Length - 1;
            }
            //// re-add bookmark with updated range
            Doc.Bookmarks.Add(bookmarkName, ref rng);

            // replace bookmark font with new font
            bm.Range.Font.Bold = vBold;


            //highlight text if it is Still not filled out
            string stringToCheck = vText;
            string[] stringArray = { "UNTERNEHMEN", "STRASSE", "STRASSEN NR", "NR","PLZ",
                "STADT", "LAND", "NACHNAME", "VORNAME", "VORNAME NACHNAME", "STADT PLZ","ANLAGE", "FUNKTION","POSITION","BETREFF" };
            if (stringArray.Any(stringToCheck.Contains))
            {
                bm.Range.HighlightColorIndex = Word.WdColorIndex.wdYellow;
            }
            else
            {
                bm.Range.HighlightColorIndex = Word.WdColorIndex.wdNoHighlight;
            }
        }
    }
}
