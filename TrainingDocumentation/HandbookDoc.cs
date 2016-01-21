using System;
using System.Drawing;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Core;
using System.IO;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Packaging;
using P = DocumentFormat.OpenXml.Presentation;
using D = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml;
using System.Xml;
using System.Windows.Forms;
using System.ComponentModel;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;
using WordDoc = DocumentFormat.OpenXml.Wordprocessing;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;

namespace TrainingDocumentation
{
    class HandbookDoc
    {
        public void CreateDocument(string pptName, BackgroundWorker bckgWorker, string saveName, bool InstructorGuide)
        {
            PresentationDocument presentationDocument = null;
            WordprocessingDocument wordprocessingDocument = null;
            try
            {
                string templateFile = Environment.CurrentDirectory + "\\Templates\\StudentHandbookTemplate.docx";
                if(InstructorGuide)
                {
                    templateFile = Environment.CurrentDirectory + "\\Templates\\InstructorGuideTemplate.docx";
                }
                File.Copy(templateFile, saveName, true);
                string savePath = Path.GetDirectoryName(pptName) + "\\" + Path.GetFileNameWithoutExtension(pptName);

                presentationDocument = PresentationDocument.Open(pptName, true);
                wordprocessingDocument = WordprocessingDocument.Open(saveName, true);

                string[] fileEntries = Directory.GetFiles(savePath);
                HandbookPPT hbp = new HandbookPPT();
                int i = 1;
                int counter = 0;
                foreach (string fileName in fileEntries)
                {
                    counter++;
                    float progress = (((float)counter / fileEntries.Length) / 2) * 100;
                    bckgWorker.ReportProgress(33 + (int)progress);
                    int a = fileName.IndexOf("Slide");
                    string tempString = fileName.Substring(0, a) + "Slide" + i.ToString() + ".png";
                    int hbSlideYesNo = GetHBSlideStatus(presentationDocument, i - 1);
                    if (((hbp.SlideVisibile(presentationDocument, i-1) == true && hbSlideYesNo==0) || hbSlideYesNo==1) && (hbSlideYesNo!=2))
                    {
                        //if (i == 2)
                        //    AddSectionBreak(true, wordprocessingDocument);
                        InsertAPicture(wordprocessingDocument, tempString);
                        NotesInSlideToWord(presentationDocument, i-1, wordprocessingDocument, InstructorGuide);
                        //if (i == 2)
                        //    AddSectionBreak(false, wordprocessingDocument);
                        if (i < fileEntries.Length)
                        {
                            InsertPageBreak(wordprocessingDocument); 
                        }
                    }

                    i++;
                }
                bckgWorker.ReportProgress(85);

            }
            catch (Exception ex)
            {
                AddToLog("Problem with creating handbook document. " + ex.Message);
                MessageBox.Show("Problem with creating handbook document. " + ex.Message);
            }
            finally
            {
                presentationDocument.Close();
                wordprocessingDocument.Close();
                bckgWorker.ReportProgress(95);
            }
        }

        public void AddSectionBreak(bool landscape, WordprocessingDocument doc)
        {
            //NOTE: does not work
            if (landscape)
            {
                doc.MainDocumentPart.Document.Body.Append(
                  new WordDoc.Paragraph(
                      new WordDoc.ParagraphProperties(
                          new WordDoc.SectionProperties(
                              new WordDoc.PageSize()
                              {
                                  Width = (UInt32Value)16840U,
                                  Height = (UInt32Value)11907U//,
                                  //Orient = WordDoc.PageOrientationValues.Landscape,
                                  //Code = (UInt16Value)9
                              },
                              new WordDoc.PageMargin()
                              {
                                  Top = 851,
                                  Right = 567,
                                  Bottom = 1344,
                                  Left = 981,
                                  Header = (UInt32Value)567U,
                                  Footer = (UInt32Value)567U,
                                  Gutter = (UInt32Value)0U
                              },
                              new WordDoc.SectionType()
                              {
                                  Val = WordDoc.SectionMarkValues.NextPage
                              }
                              ))));
            }
            else
            {
                doc.MainDocumentPart.Document.Body.Append(
                  new WordDoc.Paragraph(
                      new WordDoc.ParagraphProperties(
                          new WordDoc.SectionProperties(
                              new WordDoc.PageSize()
                              {
                                  Width = (UInt32Value)12240U,
                                  Height = (UInt32Value)15840U
                              },
                              new WordDoc.PageMargin()
                              {
                                  Top = 720,
                                  Right = Convert.ToUInt32(1 * 1440.0),
                                  Bottom = 360,
                                  Left = Convert.ToUInt32(1 * 1440.0),
                                  Header = (UInt32Value)450U,
                                  Footer = (UInt32Value)720U,
                                  Gutter = (UInt32Value)0U
                              }))));
            }
        }


        private int GetHBSlideStatus(PresentationDocument presentationDocument, int slideIndex)
        {
            int retVal = 0;

            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }
            if (slideIndex < 0)
            {
                throw new ArgumentOutOfRangeException("slideIndex");
            }
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            if (presentationPart != null && presentationPart.Presentation != null)
            {
                P.Presentation presentation = presentationPart.Presentation;
                if (presentation.SlideIdList != null)
                {
                    OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;
                    if (slideIndex < slideIds.Count)
                    {
                        string slidePartRelationshipId = (slideIds[slideIndex] as P.SlideId).RelationshipId;
                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);
                        NotesSlidePart notesSlidePart1 = slidePart.NotesSlidePart;
                        if (notesSlidePart1 == null) return 0;
                        P.NotesSlide notesSlide = notesSlidePart1.NotesSlide;
                        P.CommonSlideData csldData = notesSlide.CommonSlideData;
                        P.ShapeTree shpTree = csldData.ShapeTree;
                        P.Shape shp = shpTree.Elements<P.Shape>().ElementAt(1);
                        P.TextBody tb = shp.Elements<P.TextBody>().ElementAt(0);
                        Drawing.Paragraph para = tb.Elements<Drawing.Paragraph>().ElementAt(0);
                        string tempStr = "";
                        foreach (Drawing.Run r in para.Elements<Drawing.Run>())
                        {
                            tempStr += r.InnerText;
                        }
                        if(tempStr.Contains("*hbno*"))
                        {
                            retVal = 2;
                        }
                        else if(tempStr.Contains("*hbyes*"))
                        {
                            retVal = 1;
                        }
                    }
                }
            }

            return retVal;
        }


        private void NotesInSlideToWord(PresentationDocument presentationDocument, int slideIndex, WordprocessingDocument wordprocessingDocument, bool instructorGuide)
        {
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }
            if (slideIndex < 0)
            {
                throw new ArgumentOutOfRangeException("slideIndex");
            }
            PresentationPart presentationPart = presentationDocument.PresentationPart;
            if (presentationPart != null && presentationPart.Presentation != null)
            {
                P.Presentation presentation = presentationPart.Presentation;
                if (presentation.SlideIdList != null)
                {
                    OpenXmlElementList slideIds = presentation.SlideIdList.ChildElements;
                    if (slideIndex < slideIds.Count)
                    {
                        string slidePartRelationshipId = (slideIds[slideIndex] as P.SlideId).RelationshipId;
                        SlidePart slidePart = (SlidePart)presentationPart.GetPartById(slidePartRelationshipId);
                        NotesSlidePart notesSlidePart1 = slidePart.NotesSlidePart;
                        if (notesSlidePart1 == null) return;
                        P.NotesSlide notesSlide = notesSlidePart1.NotesSlide;
                        P.CommonSlideData csldData = notesSlide.CommonSlideData;
                        P.ShapeTree shpTree = csldData.ShapeTree;
                        P.Shape shp = shpTree.Elements<P.Shape>().ElementAt(1);
                        P.TextBody tb = shp.Elements<P.TextBody>().ElementAt(0);
                        string collectedText = "";
                        int numberingID = -1;
                        bool h1open = false;
                        bool h2open = false;
                        bool h3open = false;
                        bool instrNotesopen = false;
                        bool collectText = false;
                        int heading = 0;
                        string firstParaText = "";
                        string paragraphText = "";
                        foreach (Drawing.Paragraph para in tb.Elements<Drawing.Paragraph>())
                        {
                            paragraphText = "";
                            foreach (Drawing.Run r in para.Elements<Drawing.Run>())
                            {
                                firstParaText += r.InnerText;
                                paragraphText += r.InnerText;
                            }
                            if (firstParaText.Contains("*hbno*") || firstParaText.Contains("*hbyes*"))
                            {
                                firstParaText = "";
                                continue;
                            }
                            

                            bool b = false;
                            bool i = false;
                            bool u = false;
                            Drawing.ParagraphProperties pp = null;
                            foreach (OpenXmlElement el in para.Elements<Drawing.ParagraphProperties>())
                            {
                                if (el.GetType() == typeof(Drawing.ParagraphProperties))
                                {
                                    pp = (Drawing.ParagraphProperties)el;
                                }
                            }
                            Drawing.CharacterBullet cb = null;
                            Drawing.AutoNumberedBullet anb = null;
                            if (pp != null)
                            {
                                OpenXmlElementList li = pp.ChildElements;
                                foreach (OpenXmlElement el in li)
                                {
                                    if (el.GetType() == typeof(Drawing.CharacterBullet))
                                    {
                                        cb = (Drawing.CharacterBullet)el;
                                    }
                                    if (el.GetType() == typeof(Drawing.AutoNumberedBullet))
                                    {
                                        anb = (Drawing.AutoNumberedBullet)el;
                                    }
                                }
                            }
                            if (cb != null && !instrNotesopen)
                            {
                                AddParagraphToWordBullet(wordprocessingDocument, cb.Char);
                            }
                            else if (anb != null && !instrNotesopen)
                            {
                                numberingID = AddParagraphToWordNumbering(wordprocessingDocument, numberingID);
                            }
                            else if (!instrNotesopen && !paragraphText.Contains("**") && !paragraphText.Contains("*h1*") && !paragraphText.Contains("*h2*") && !paragraphText.Contains("*h3*"))
                            {
                                AddParagraphToWord(wordprocessingDocument);
                                numberingID = -1;
                            }

                            foreach (Drawing.Run r in para.Elements<Drawing.Run>())
                            {
                                b = false;
                                i = false;
                                u = false;
                                string addedText = "";
                                string tempStr = r.InnerText;
                                collectedText += r.InnerText;
                                tempStr = collectedText;
                                string textToWrite = "";
                                
                                int astLoc = tempStr.IndexOf("*");
                                if (astLoc != -1)
                                {
                                    if (astLoc > 0)
                                    {
                                        textToWrite = tempStr.Substring(0, astLoc);
                                    }
                                    if (tempStr.Length > astLoc + 1)
                                    {
                                        if (tempStr.Substring(astLoc, 2) == "**")
                                        {
                                            if (!instructorGuide)
                                            {
                                                if (!instrNotesopen)
                                                {
                                                    collectedText = tempStr.Substring(astLoc + 2);
                                                    int a = collectedText.IndexOf("**");
                                                    if (a != -1)
                                                    {
                                                        int a1 = collectedText.IndexOf("**");
                                                        if (collectedText.Length > a + 2)
                                                            addedText = collectedText.Substring(a + 2);
                                                        collectedText = collectedText.Remove(collectedText.IndexOf("**"), 2);
                                                        instrNotesopen = !instrNotesopen;
                                                        collectText = !collectText;
                                                    }

                                                    if (instrNotesopen)
                                                    {
                                                        collectedText = "";
                                                    }
                                                }
                                                else
                                                {
                                                    int a = collectedText.IndexOf("**");
                                                    if (collectedText.Length > a + 2)
                                                        addedText = collectedText.Substring(a + 2);
                                                    collectedText = collectedText.Remove(collectedText.IndexOf("**"));
                                                    collectedText = "";
                                                }
                                                instrNotesopen = !instrNotesopen;
                                                collectText = !collectText;
                                            }
                                            else
                                            {
                                                collectedText = collectedText.Replace("**", "");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        collectText = true;
                                    }
                                    if (tempStr.Length > astLoc + 3)
                                    {
                                        if (tempStr.Substring(astLoc, 4) == "*h1*")
                                        {
                                            if (!h1open)
                                            {
                                                collectedText = tempStr.Substring(astLoc + 4);
                                                int a = collectedText.IndexOf("*h1*");
                                                collectText = false;    //20.01. testing added this
                                                if (a != -1)
                                                {
                                                    int a1 = collectedText.IndexOf("*h1*");
                                                    if (collectedText.Length > a1 + 4)
                                                        addedText = collectedText.Substring(a1 + 4);
                                                    collectedText = collectedText.Remove(collectedText.IndexOf("*h1*"));
                                                    h1open = !h1open;
                                                    collectText = !collectText;
                                                    heading = 1;
                                                }
                                            }
                                            else
                                            {
                                                int a = collectedText.IndexOf("*h1*");
                                                if (collectedText.Length > a + 4)
                                                    addedText = collectedText.Substring(a + 4);
                                                collectedText = collectedText.Remove(collectedText.IndexOf("*h1*"));
                                                heading = 1;
                                            }
                                            h1open = !h1open;
                                            collectText = !collectText;
                                        }
                                        else if (tempStr.Substring(astLoc, 4) == "*h2*")
                                        {
                                            if (!h2open)
                                            {
                                                collectedText = tempStr.Substring(astLoc + 4);
                                                int a = collectedText.IndexOf("*h2*");
                                                collectText = false;    //20.01. testing added this
                                                if (a != -1)
                                                {
                                                    int a1 = collectedText.IndexOf("*h2*");
                                                    if (collectedText.Length > a1 + 4)
                                                        addedText = collectedText.Substring(a1 + 4);
                                                    collectedText = collectedText.Remove(collectedText.IndexOf("*h2*"));
                                                    h1open = !h1open;
                                                    collectText = !collectText;
                                                    heading = 2;
                                                }
                                            }
                                            else
                                            {
                                                int a = collectedText.IndexOf("*h2*");
                                                if (collectedText.Length > a + 4)
                                                    addedText = collectedText.Substring(a + 4);
                                                collectedText = collectedText.Remove(collectedText.IndexOf("*h2*"));
                                                heading = 2;
                                            }
                                            h2open = !h2open;
                                            collectText = !collectText;
                                        }
                                        else if (tempStr.Substring(astLoc, 4) == "*h3*")
                                        {
                                            if (!h3open)
                                            {
                                                collectedText = tempStr.Substring(astLoc + 4);
                                                int a = collectedText.IndexOf("*h3*");
                                                collectText = false;    //20.01. testing added this
                                                if (a != -1)
                                                {
                                                    int a1 = collectedText.IndexOf("*h3*");
                                                    if (collectedText.Length > a1 + 4)
                                                        addedText = collectedText.Substring(a1 + 4);
                                                    collectedText = collectedText.Remove(collectedText.IndexOf("*h3*"));
                                                    h1open = !h1open;
                                                    collectText = !collectText;
                                                    heading = 3;
                                                }
                                            }
                                            else
                                            {
                                                int a = collectedText.IndexOf("*h3*");
                                                if (collectedText.Length > a + 4)
                                                    addedText = collectedText.Substring(a + 4);
                                                collectedText = collectedText.Remove(collectedText.IndexOf("*h3*"));
                                                heading = 3;
                                            }
                                            h3open = !h3open;
                                            collectText = !collectText;
                                        }
                                    }
                                    //if (tempStr.Length > astLoc + 1)
                                    //{
                                    //    if (tempStr.Substring(astLoc, 2) == "**")
                                    //    {
                                    //        int a = collectedText.IndexOf("**");
                                    //        if (collectedText.Length > a + 2)
                                    //            addedText = collectedText.Substring(a + 2);
                                    //        collectedText = collectedText.Remove(collectedText.IndexOf("**"), 2);
                                    //        instrNotesopen = !instrNotesopen;
                                    //        collectText = !collectText;
                                    //        if (!instrNotesopen)
                                    //        {
                                    //            collectedText = "";
                                    //        }
                                    //    }
                                    //}
                                    //else
                                    //{
                                    //    collectText = true;
                                    //}
                                }

                                Drawing.RunProperties rp = r.RunProperties;
                                b = false;
                                i = false;
                                u = false;
                                if (rp.Bold != null && rp.Bold.ToString() != "0")
                                    b = true;
                                if (rp.Italic != null && rp.Italic.ToString() != "0")
                                    i = true;
                                if (rp.Underline != null && rp.Underline != "none")
                                    u = true;

                                if (!collectText && collectedText.Length > 0)
                                {
                                    AddTextToWord(wordprocessingDocument, collectedText, b, i, u, heading);
                                    heading = 0;
                                    collectedText = "";
                                }
                                else
                                {

                                }
                                if (addedText.Length > 0)
                                    collectedText = addedText;
                            }
                            if (!collectText && collectedText.Length > 0 && collectedText.IndexOf("*") == -1)
                            {
                                AddTextToWord(wordprocessingDocument, collectedText, b, i, u, heading);
                                heading = 0;
                                collectedText = "";
                            }
                        }
                        if (collectedText.Length > 0)
                        {
                            AddParagraphToWord(wordprocessingDocument);
                            AddTextToWord(wordprocessingDocument, collectedText, false, false, false, 0);

                            collectedText = "";
                        }
                    }
                }
            }

        }

        private int AddParagraphToWordNumbering(WordprocessingDocument doc, int numberingID)
        {

            MainDocumentPart mainPart = doc.MainDocumentPart;
            WordDoc.Body body = mainPart.Document.Body;
            WordDoc.Paragraph para = body.AppendChild(new WordDoc.Paragraph());
            NumberingDefinitionsPart numberingPart = mainPart.NumberingDefinitionsPart;
            WordDoc.Numbering nn = numberingPart.Numbering;
            int numid = 0;
            if (numberingID == -1)
            {
                foreach (WordDoc.NumberingInstance num in nn.Elements<WordDoc.NumberingInstance>())
                {
                    numid = num.NumberID;
                }
                numid++;
                WordDoc.NumberingInstance newNI =
                    new WordDoc.NumberingInstance(
                        new WordDoc.AbstractNumId() { Val = numid },
                        new WordDoc.LevelOverride(
                            new WordDoc.StartOverrideNumberingValue() { Val = 1 }
                            )
                        { LevelIndex = 0 }
                        )
                    { NumberID = numid };
                nn.Append(newNI);
                nn.Save(numberingPart);
            }
            else
            {
                numid = numberingID;
            }

            WordDoc.ParagraphProperties pp = para.AppendChild(new WordDoc.ParagraphProperties(
                new WordDoc.NumberingProperties(
                new WordDoc.NumberingLevelReference() { Val = 0 },
                new WordDoc.NumberingId() { Val = numid })));

            return numid;
        }


        public void AddParagraphToWordBullet(WordprocessingDocument doc, StringValue bulletNumChar)
        {

            MainDocumentPart mainPart = doc.MainDocumentPart;
            WordDoc.Body body = mainPart.Document.Body;
            WordDoc.Paragraph para = body.AppendChild(new WordDoc.Paragraph());
            
            WordDoc.ParagraphProperties pp = para.AppendChild(new WordDoc.ParagraphProperties(
                new WordDoc.NumberingProperties(
                new WordDoc.NumberingLevelReference() { Val = 0 },
                new WordDoc.NumberingId() { Val = 1 })));
        }

        public void AddParagraphToWord(WordprocessingDocument doc)
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            WordDoc.Body body = mainPart.Document.Body;
            WordDoc.Paragraph para = body.AppendChild(new WordDoc.Paragraph());
        }

        

        public void AddTextToWord(WordprocessingDocument doc, string text, bool bold, bool italic, bool underline, int heading)
        {
            MainDocumentPart mainPart = doc.MainDocumentPart;
            WordDoc.Body body = mainPart.Document.Body;
            WordDoc.Paragraph para = new WordDoc.Paragraph();
            foreach (WordDoc.Paragraph p in body.Elements<WordDoc.Paragraph>())
            {
                para = p;
            }
            if (heading != 0)
            {
                para = body.AppendChild(new WordDoc.Paragraph());
                para.ParagraphProperties = new WordDoc.ParagraphProperties(new WordDoc.ParagraphStyleId() { Val = "Heading" + heading.ToString() });
                bold = false;
                italic = false;
                underline = false;
            }
            else
            {
                if (para.ParagraphProperties != null)
                {
                    if (para.ParagraphProperties.ParagraphStyleId != null)
                    {
                        if (para.ParagraphProperties.ParagraphStyleId.Val == "Heading1" ||
                          para.ParagraphProperties.ParagraphStyleId.Val == "Heading2" ||
                          para.ParagraphProperties.ParagraphStyleId.Val == "Heading3")
                        {
                            para = body.AppendChild(new WordDoc.Paragraph());
                        }
                    }
                }
            }

            string[] t1 = text.Split('\n');
            for (int j = 0; j < t1.Length; j++)
            {

                WordDoc.Run run = para.AppendChild(new WordDoc.Run());
                WordDoc.RunProperties rp = run.AppendChild(new WordDoc.RunProperties());
                if (bold)
                {
                    WordDoc.Bold b = rp.AppendChild(new WordDoc.Bold());
                }
                if (italic)
                {
                    WordDoc.Italic i = rp.AppendChild(new WordDoc.Italic());
                }
                if (underline)
                {
                    WordDoc.Underline u = new WordDoc.Underline()
                    {
                        Val = WordDoc.UnderlineValues.Single
                    };
                    rp.AppendChild(u);
                }
                WordDoc.Text t = new WordDoc.Text()
                {
                    Text = t1[j],
                    Space = SpaceProcessingModeValues.Preserve
                };
                run.AppendChild(t);
                if (j != t1.Length - 1)
                    run.AppendChild(new WordDoc.Break());

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

        private void InsertPageBreak(WordprocessingDocument wordprocessingDocument)
        {
            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
            WordDoc.Body body = mainPart.Document.Body;
            WordDoc.Paragraph para = new WordDoc.Paragraph();
            para = body.AppendChild(new WordDoc.Paragraph());

            WordDoc.Run run = para.AppendChild(new WordDoc.Run(
                new WordDoc.Break() { Type = WordDoc.BreakValues.Page }
                ));
        }


        private void InsertAPicture(WordprocessingDocument wordprocessingDocument, string fileName)
        {
            
            MainDocumentPart mainPart = wordprocessingDocument.MainDocumentPart;
            ImagePart imagePart = mainPart.AddImagePart(ImagePartType.Png);
            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            Bitmap img = new Bitmap(fileName);
            int widthPx = img.Width;
            int heightPx = img.Height;
            float horzRezDpi = img.HorizontalResolution;
            float vertRezDpi = img.VerticalResolution;
            const int emusPerInch = 914400;
            const int emusPerCm = 360000;
            float maxWidthCm = Properties.Settings.Default.HANDBOOK_MAX_SLIDE_WIDTH;
            long widthEmus = (long)(widthPx / horzRezDpi * emusPerInch);
            long heightEmus = (long)(heightPx / vertRezDpi * emusPerInch);

            long maxWidthEmus = (long)(maxWidthCm * emusPerCm);
            if (widthEmus > maxWidthEmus)
            {
                var ratio = (heightEmus * 1.0m) / widthEmus;
                widthEmus = maxWidthEmus;
                heightEmus = (long)(widthEmus * ratio);
            }


            AddImageToBody(wordprocessingDocument, mainPart.GetIdOfPart(imagePart), heightEmus, widthEmus);
            
        }

        private void AddToLog(string message)
        {
            try
            {
                //string fileName = Environment.CurrentDirectory + "\\Log.txt";
                string fileName = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Log.txt";
                using (StreamWriter sw = File.AppendText(fileName))
                {
                    sw.WriteLine("{0}: {1}", DateTime.Now.ToString(), message);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to update the log file! " + ex.Message);
            }
        }

        private static void AddImageToBody(WordprocessingDocument wordDoc, string relationshipId, long h, long w)
        {
            // Define the reference of the image.
            var element =
                 new WordDoc.Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = w, Cy = h },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 19050L,
                             TopEdge = 19050L,
                             RightEdge = 24765L,
                             BottomEdge = 19050L
                         },
                         new DW.DocProperties()
                         {
                             Id = (DocumentFormat.OpenXml.UInt32Value)1U,
                             Name = "Picture 1"
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new Drawing.GraphicFrameLocks() { NoChangeAspect = true }),
                         new Drawing.Graphic(
                             new Drawing.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (DocumentFormat.OpenXml.UInt32Value)0U,
                                             Name = "New Image.png"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new Drawing.Blip(
                                             new Drawing.BlipExtensionList(
                                                 new Drawing.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             Drawing.BlipCompressionValues.Print
                                         },
                                         new Drawing.Stretch(
                                             new Drawing.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new Drawing.Transform2D(
                                             new Drawing.Offset() { X = 0L, Y = 0L },
                                             new Drawing.Extents() { Cx = w, Cy = h }),
                                         new Drawing.PresetGeometry(
                                             new Drawing.AdjustValueList()
                                         )
                                         { Preset = Drawing.ShapeTypeValues.Rectangle },
                                         new Drawing.Outline(new Drawing.SolidFill(new Drawing.SchemeColor() { Val = Drawing.SchemeColorValues.Text1 }))
                                         ))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (DocumentFormat.OpenXml.UInt32Value)0U,
                         DistanceFromBottom = (DocumentFormat.OpenXml.UInt32Value)0U,
                         DistanceFromLeft = (DocumentFormat.OpenXml.UInt32Value)0U,
                         DistanceFromRight = (DocumentFormat.OpenXml.UInt32Value)0U,
                         EditId = "50D07946"
                     });

            
            wordDoc.MainDocumentPart.Document.Body.AppendChild(new WordDoc.Paragraph(new WordDoc.Run(element)));
        }

    }
}
