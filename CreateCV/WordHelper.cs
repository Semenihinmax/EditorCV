using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Word.Application;

namespace CreateCV
{
    class WordHelper
    {
        private FileInfo fileInfo;

        public WordHelper(string fileName)
        {
            if (File.Exists(fileName))
            {
                fileInfo = new FileInfo(fileName);
            }
            else
            {
                throw new ArgumentException("File not found");
            }
        }

        internal bool Process(Dictionary<string, string> items, string imageAdress)
        {
            Application application = null;

            try
            {
                application = new Application();
                object file = fileInfo.FullName;

                object missing = Type.Missing;

                Document document = application.Documents.Open(file);

                foreach (var item in items)
                {
                    Find find = application.Selection.Find;
                    find.Text = item.Key;
                    find.Replacement.Text = item.Value;

                    Object wrap = WdFindWrap.wdFindContinue;
                    Object replace = WdReplace.wdReplaceAll;

                    find.Execute(FindText: Type.Missing,
                        MatchCase: false,
                        MatchWholeWord: false,
                        MatchWildcards: false,
                        MatchSoundsLike: missing,
                        MatchAllWordForms: false,
                        Forward: true,
                        Wrap: wrap,
                        Format: false,
                        ReplaceWith: missing, Replace: replace);
                }

                Object newFileName = Path.Combine(fileInfo.DirectoryName, items["<FIO>"] + "CV");
                application.ActiveDocument.SaveAs2(newFileName);
                application.ActiveDocument.Close();

                if (imageAdress != null)
                {
                    addImageInWord(newFileName.ToString(), imageAdress, "Контакты");
                }

                return true;
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if(application != null)
                {
                    application.Quit();
                }
            }

            return false;
        }

        internal List<string> ReadText()
        {
            Application application = null;

            try
            {
                application = new Application();
                object file = fileInfo.FullName;

                object missing = Type.Missing;

                Document doc = application.Documents.Open(file);
                var list = new List<string>();

                foreach (Paragraph paragraph in doc.Paragraphs)
                {
                    list.Add(paragraph.Range.Text);
                }
               
                application.ActiveDocument.Close();

                return list;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (application != null)
                {
                    application.Quit();
                }
            }

            return null;
        }

        internal Image ReadImage()
        {
            Application application = null;
            int i = 0;

            try
            {
                application = new Application()
                {
                    //Ignore Word notifications (read only)
                    Visible = false,
                    AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
                };

                object file = fileInfo.FullName;

                Document doc = application.Documents.Open(file);
                Image image = null;

                foreach (InlineShape shape in doc.InlineShapes)
                {
                    shape.Range.Select();
                    if (shape.Type == WdInlineShapeType.wdInlineShapePicture)
                    {
                        doc.ActiveWindow.Selection.Range.CopyAsPicture();
                        var ImageData = Clipboard.GetDataObject();
                        image = (Image)ImageData.GetData(DataFormats.Bitmap);
                    }
                    // return last image
                    return image;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (application != null)
                {
                    application.Quit();
                }
            }

            return null;
        }

        private void addImageInWord(string fileName, string imageAdress, string tag)
        {
            Application application = new Application();
            Document document = application.Documents.Open(fileName);

            Range rng = document.Content;
            Find wdFind = rng.Find;

            wdFind.Text = tag;
            bool found = wdFind.Execute();

            if (found)
            {
                InlineShape ils = rng.InlineShapes.AddPicture(imageAdress, false, true, rng);
            }

            application.ActiveDocument.SaveAs2(fileName);
            application.ActiveDocument.Close();
            application.Quit();
        }
    }
}
