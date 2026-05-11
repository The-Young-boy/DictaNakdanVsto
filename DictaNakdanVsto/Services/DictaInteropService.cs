using System;
using Word = Microsoft.Office.Interop.Word;

namespace DictaNakdanVsto.Services
{
    public class DictaInteropService
    {
        private Word.Application _app => Globals.ThisAddIn.Application;

        public void StartUndoRecord() => _app.UndoRecord.StartCustomRecord("ניקוד דיקטה");
        public void EndUndoRecord() => _app.UndoRecord.EndCustomRecord();

        public Word.Range GetInitialSearchRange()
        {
            Word.Selection sel = _app.Selection;
            if (sel.Start != sel.End) return sel.Range; // טקסט מסומן
            return sel.Paragraphs[1].Range; // פסקה נוכחית
        }

        // מוצא ומסמן את המילה בוורד, ומחזיר את הטווח שלה
        public Word.Range FindAndSelectWord(string wordText, Word.Range bounds)
        {
            Word.Range searchRange = bounds.Duplicate;
            Word.Find find = searchRange.Find;
            find.ClearFormatting();
            find.Text = wordText;
            find.MatchWholeWord = true;
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;

            if (find.Execute())
            {
                searchRange.Select(); // קריטי: מסמן את המילה למשתמש
                return searchRange;
            }
            return null;
        }

        // מחליף את המילה ומצמצם את טווח החיפוש הנותר
        public void ReplaceWordAndUpdateBounds(Word.Range wordRange, string newText, ref Word.Range bounds)
        {
            wordRange.Text = newText;
            bounds.Start = wordRange.End; // מקדם את החיפוש הבא לאחר המילה שהוחלפה
        }
    }
}