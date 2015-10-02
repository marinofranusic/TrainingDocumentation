using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Core;
using System.IO;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using System.Xml;
using System.Windows.Forms;
using System.ComponentModel;

namespace TrainingDocumentation
{
    class HandbookDoc
    {
        public void CreateDocument(string pptName, BackgroundWorker bckgWorker, string saveName)
        {
            try
            {
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                object missing = System.Reflection.Missing.Value;
                object oHeadingStyle1 = WdBuiltinStyle.wdStyleHeading1;
                object oHeadingStyle2 = WdBuiltinStyle.wdStyleHeading2;
                object oHeadingStyle3 = WdBuiltinStyle.wdStyleHeading3;
                object oBodyTextStyle = WdBuiltinStyle.wdStyleBodyText;

                string templateFile = Environment.CurrentDirectory + "\\Templates\\StudentHandbookTemplate.docx";

                Document document = winword.Documents.Open(templateFile);
                object oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */


                string savePath = Path.GetDirectoryName(pptName) + "\\" + Path.GetFileNameWithoutExtension(pptName);

                string[] fileEntries = Directory.GetFiles(savePath);

                var app = new Microsoft.Office.Interop.PowerPoint.Application();
                var pres = app.Presentations;
                var pptfile = pres.Open(pptName, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse);

                int i = 1;
                int counter = 0;
                foreach (string fileName in fileEntries)
                {
                    counter++;
                    float progress = (((float)counter / fileEntries.Length) / 2) * 100;
                    bckgWorker.ReportProgress(50 + (int)progress);
                    int a = fileName.IndexOf("Slide");
                    string tempString = fileName.Substring(0, a) + "Slide" + i.ToString() + ".png";

                    if (pptfile.Slides[i].SlideShowTransition.Hidden != MsoTriState.msoTrue)
                    {
                        object what = WdGoToItem.wdGoToPercent;
                        object which = WdGoToDirection.wdGoToLast;
                        document.Application.Selection.GoTo(ref what, ref which, ref missing, ref missing);
                        document.Application.Selection.TypeBackspace();
                        document.Application.Selection.InsertBreak(WdBreakType.wdPageBreak);

                        var pPic = document.Paragraphs.Add();
                        pPic.Format.SpaceAfter = 6f;

                        InlineShape shpPic = document.InlineShapes.AddPicture(tempString, Range: pPic.Range);
                        shpPic.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleSingle;


                        shpPic.LockAspectRatio = MsoTriState.msoCTrue;
                        shpPic.Height = 370f;

                        object style_name = "Heading 1";
                        //notes.Range.set_Style(ref oBodyTextStyle);
                        HandbookPPT hbp = new HandbookPPT();
                        string strNotes = hbp.GetNotes(pptfile.Slides[i]);
                        string h1 = GetHeading1(strNotes);
                        if (h1 != "")
                        {
                            var h1p = document.Paragraphs.Add();

                            h1p.Range.Text = h1;
                            h1p.set_Style(ref style_name);

                            h1p.Range.InsertParagraphAfter();
                            h1p.Outdent();
                            strNotes = RemoveHeading1(strNotes);

                        }

                        string h2 = GetHeading2(strNotes);
                        if (h2 != "")
                        {
                            var h2p = document.Paragraphs.Add();

                            h2p.Range.Text = h2;
                            style_name = "Heading 2";
                            h2p.set_Style(ref style_name);

                            h2p.Range.InsertParagraphAfter();
                            h2p.Outdent();
                            strNotes = RemoveHeading2(strNotes);

                        }

                        string h3 = GetHeading3(strNotes);
                        if (h3 != "")
                        {
                            var h3p = document.Paragraphs.Add();

                            h3p.Range.Text = h3;
                            style_name = "Heading 3";
                            h3p.set_Style(ref style_name);

                            h3p.Range.InsertParagraphAfter();
                            h3p.Outdent();
                            strNotes = RemoveHeading3(strNotes);

                        }

                        var notes = document.Paragraphs.Add();

                        notes.Range.Text = strNotes;
                        style_name = "Normal";
                        notes.set_Style(ref style_name);
                        notes.Outdent();
                        notes.Range.InsertParagraphAfter();
                    }

                    i++;
                }

                //Save the document
                string savePathDocFile = saveName;
                
                object filename = savePathDocFile;
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;

                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }


        

        private string GetHeading1(string inputString)
        {
            string testString = "";
            int a = inputString.IndexOf("*h1*");
            if (a != -1)
            {
                int b = inputString.Substring(a + 2).IndexOf("*h1*");
                if (b + a > a)
                {
                    testString = inputString.Substring(a + 4, b - 2);
                }
            }
            return testString;
        }
        private string RemoveHeading1(string inputString)
        {
            string testString = inputString;
            int a = 0;
            int b = 0;

            a = testString.IndexOf("*h1*");
            b = testString.Substring(a + 2).IndexOf("*h1*");
            if (b + a > a)
            {
                testString = testString.Remove(a, b + 6);
            }

            return testString;
        }

        private string GetHeading2(string inputString)
        {
            string testString = "";
            int a = inputString.IndexOf("*h2*");
            if (a != -1)
            {
                int b = inputString.Substring(a + 2).IndexOf("*h2*");
                if (b + a > a)
                {
                    testString = inputString.Substring(a + 4, b - 2);
                }
            }
            return testString;
        }
        private string RemoveHeading2(string inputString)
        {
            string testString = inputString;
            int a = 0;
            int b = 0;

            a = testString.IndexOf("*h2*");
            b = testString.Substring(a + 2).IndexOf("*h2*");
            if (b + a > a)
            {
                testString = testString.Remove(a, b + 6);
            }

            return testString;
        }

        private string GetHeading3(string inputString)
        {
            string testString = "";
            int a = inputString.IndexOf("*h3*");
            if (a != -1)
            {
                int b = inputString.Substring(a + 2).IndexOf("*h3*");
                if (b + a > a)
                {
                    testString = inputString.Substring(a + 4, b - 2);
                }
            }
            return testString;
        }
        private string RemoveHeading3(string inputString)
        {
            string testString = inputString;
            int a = 0;
            int b = 0;

            a = testString.IndexOf("*h3*");
            b = testString.Substring(a + 2).IndexOf("*h3*");
            if (b + a > a)
            {
                testString = testString.Remove(a, b + 6);
            }

            return testString;
        }

        public int CheckIfTemplateExists()
        {
            int retVal = -1;

            string templateFile = Environment.CurrentDirectory + "\\Templates\\StudentHandbookTemplate.docx";
            if(File.Exists(templateFile))
            {
                retVal = 1;
            }
            return retVal;

        }

    }
}
