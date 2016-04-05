using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Threading;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using System.IO;
using x = System.Xml;
using WordDoc = DocumentFormat.OpenXml.Wordprocessing;

namespace TrainingDocumentation
{
    class LPConvertPPT
    {

        public void CreateTemplateFromXML(string xmlFileName, PresentationDocument presentationDocument, BackgroundWorker bckgW, int subjectNumber)
        {
            try
            {
                int slideNumber = 0;
                x.XmlDocument xmlDoc = new x.XmlDocument();
                xmlDoc.Load(xmlFileName);
                x.XmlNode node = xmlDoc.DocumentElement.SelectSingleNode("/Course/Title");

                string courseTitle = node.InnerText;

                slideNumber = InsertTitleSlide(presentationDocument, slideNumber, courseTitle, Properties.Settings.Default.DEFAULT_TITLE_SLIDE_SUBTITLE, true);

                node = xmlDoc.DocumentElement.SelectSingleNode("/Course/Subjects");
                List<List<string>> objectives = new List<List<string>>();

                int totalNumOfTopics = 0;

                foreach (x.XmlNode n in node.ChildNodes)
                {
                    if (n.Name == "Subject")
                    {
                        List<string> subjObjectives = new List<string>();
                        foreach (x.XmlNode nc in n.ChildNodes)
                        {
                            if (nc.Name == "Topic")
                            {
                                totalNumOfTopics++;
                                int headingVal = Convert.ToInt32(nc.Attributes["Heading"].Value);
                                if (headingVal == 1)
                                {
                                    string objective = nc["Objective"].InnerText;
                                    subjObjectives.Add(objective);
                                }
                            }
                        }
                        objectives.Add(subjObjectives);
                    }
                }

                int subjectCounter = 0;
                int topicCounter = 0;
                foreach (x.XmlNode n in node.ChildNodes)
                {
                    if (n.Name == "Subject")
                    {
                        if ((subjectCounter == subjectNumber) || subjectNumber == -1)
                        {
                            string lessonGoal = "";
                            string subjectTitle = n.Attributes["SubjectTitle"].Value;
                            if (string.Equals(subjectTitle, Properties.Settings.Default.INTRODUCTION_SUBJECT_TITLE, StringComparison.OrdinalIgnoreCase))
                            {
                                slideNumber = InsertIntroductionSlide(presentationDocument, slideNumber);

                            }
                            else
                            {
                                slideNumber = InsertTitleSlide(presentationDocument, slideNumber, courseTitle, subjectTitle, false);

                                foreach (x.XmlNode nc in n.ChildNodes)
                                {
                                    if (nc.Name == "LessonGoal")
                                    {
                                        lessonGoal = nc.InnerText;
                                        slideNumber = InsertLessonGoalSlide(presentationDocument, slideNumber, lessonGoal, true);

                                        slideNumber = InsertLessonObjectivesSlide(presentationDocument, slideNumber, objectives[subjectCounter], true);

                                    }
                                    if (nc.Name == "Topic")
                                    {
                                        float progress = (((float)topicCounter / totalNumOfTopics) / 2) * 100;
                                        bckgW.ReportProgress(50 + (int)progress);
                                        int headingVal = Convert.ToInt32(nc.Attributes["Heading"].Value);
                                        string topicTitle = nc["Title"].InnerText;
                                        if (headingVal == 1)
                                        {
                                            List<string> topicDividerNotes = new List<string>();
                                            topicDividerNotes.Add(nc["Objective"].InnerText);
                                            topicDividerNotes.Add(nc["Methods"].InnerText);
                                            topicDividerNotes.Add(nc["Evaluation"].InnerText);
                                            topicDividerNotes.Add(nc["Time"].InnerText);
                                            topicDividerNotes.Add(nc["Materials"].InnerText);
                                            topicDividerNotes.Add(nc["Notes"].InnerText);
                                            slideNumber = InsertTopicDividerSlide(presentationDocument, slideNumber, topicTitle, topicDividerNotes);


                                        }
                                        slideNumber = InsertContentSlide(presentationDocument, slideNumber, topicTitle);
                                        topicCounter++;
                                    }
                                }

                            }

                            //finishing up the subject
                            slideNumber = InsertQuestionsSlide(presentationDocument, slideNumber);
                            slideNumber = InsertLessonObjectivesSlide(presentationDocument, slideNumber, objectives[subjectCounter], false);
                            slideNumber = InsertLessonGoalSlide(presentationDocument, slideNumber, lessonGoal, false);
                        }
                    }
                    subjectCounter++;
                }
            }
            catch(Exception ex)
            {
                AddToLog("Problem with creating presentation document. " + ex.Message);
                MessageBox.Show("Problem with creating presentation document. " + ex.Message);
            }
        }

        private void AddToLog(string message)
        {
            //string fileName = Environment.CurrentDirectory + "\\Log.txt";
            string fileName = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\Log.txt";
            using (StreamWriter sw = File.AppendText(fileName))
            {
                sw.WriteLine("{0}: {1}", DateTime.Now.ToString(), message);
            }
        }

        private int InsertQuestionsSlide(PresentationDocument presentationDocument, int slideNumber)
        {
            int thisSlideNumber = slideNumber;
            if (presentationDocument == null)
            {
                MessageBox.Show("Can't insert slide in the presentation.");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart == null)
            {
                MessageBox.Show("The presentation document is empty.");
            }

            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            string layoutName = Properties.Settings.Default.QUESTIONS_SLIDE;

            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slideLayoutPart = null;
            foreach (SlideMasterPart smp in presentationPart.SlideMasterParts)
            {
                foreach (SlideLayoutPart slp in smp.SlideLayoutParts)
                {
                    if (slp.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName, StringComparison.OrdinalIgnoreCase))
                    {
                        slideLayoutPart = slp;
                    }
                }
            }

            if (slideLayoutPart == null)
            {
                MessageBox.Show("Slide Layout not found");
                return slideNumber;
            }
            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideLayoutPart.SlideLayout.CommonSlideData.Clone();
            using (Stream stream = slideLayoutPart.GetStream())
            {
                slidePart.SlideLayoutPart.FeedData(stream);
            }

            foreach (ImagePart iPart in slideLayoutPart.ImageParts)
            {
                ImagePart newImagePart = slidePart.AddImagePart(iPart.ContentType, slideLayoutPart.GetIdOfPart(iPart));
                newImagePart.FeedData(iPart.GetStream());
            }

            NotesSlidePart notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
            NotesSlide notesSlide = CreateDeafultNotesPart(Properties.Settings.Default.FORCE_SLIDE_HANDBOOK_NO + Environment.NewLine);

            string notesText = 
                Properties.Settings.Default.HEADING2_MARKUP +
                Properties.Settings.Default.QUESTIONS_SLIDE_TITLE +
                Properties.Settings.Default.HEADING2_MARKUP +
                Environment.NewLine + Environment.NewLine;

            notesSlidePart.NotesSlide = notesSlide;
            Shape shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            Drawing.Run newRun = new Drawing.Run();
            Drawing.RunProperties runProp = new Drawing.RunProperties();
            runProp.Bold = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            Drawing.Text newText = new Drawing.Text(notesText);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            Drawing.Paragraph newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);

            shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            newRun = new Drawing.Run();
            runProp = new Drawing.RunProperties();
            runProp.Bold = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            newText = new Drawing.Text(Properties.Settings.Default.DEFAULT_NOTES_ENDING + Environment.NewLine);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);

            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                slideNumber--;
                if (slideNumber == 0)
                {
                    prevSlideId = slideId;
                }

            }
            maxSlideId++;
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();

            thisSlideNumber++;
            return thisSlideNumber;
        }


        private int InsertContentSlide(PresentationDocument presentationDocument, int slideNumber, string title)
        {
            int thisSlideNumber = slideNumber;
            if (presentationDocument == null)
            {
                MessageBox.Show("Can't insert slide in the presentation.");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart == null)
            {
                MessageBox.Show("The presentation document is empty.");
            }

            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            string layoutName = Properties.Settings.Default.CONTENT_SLIDE;

            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slideLayoutPart = null;
            foreach (SlideMasterPart smp in presentationPart.SlideMasterParts)
            {
                foreach (SlideLayoutPart slp in smp.SlideLayoutParts)
                {
                    if (slp.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName, StringComparison.OrdinalIgnoreCase))
                    {
                        slideLayoutPart = slp;
                    }
                }
            }

            if (slideLayoutPart == null)
            {
                MessageBox.Show("Slide Layout not found");
                return slideNumber;
            }
            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideLayoutPart.SlideLayout.CommonSlideData.Clone();
            using (Stream stream = slideLayoutPart.GetStream())
            {
                slidePart.SlideLayoutPart.FeedData(stream);
            }

            foreach (ImagePart iPart in slideLayoutPart.ImageParts)
            {
                ImagePart newImagePart = slidePart.AddImagePart(iPart.ContentType, slideLayoutPart.GetIdOfPart(iPart));
                newImagePart.FeedData(iPart.GetStream());
            }


            foreach (Shape shp in slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>())
            {
                if (shp.InnerText.Contains(Properties.Settings.Default.CONTENT_SLIDE_TITLE_CONTAINER_KEYWORD))
                {
                    foreach (Drawing.Paragraph p in shp.TextBody.Elements<Drawing.Paragraph>())
                    {
                        if (p.InnerText.Contains(Properties.Settings.Default.CONTENT_SLIDE_TITLE_CONTAINER_KEYWORD))
                        {
                            int runCounter = 0;
                            foreach (Drawing.Run r in p.Elements<Drawing.Run>())
                            {
                                if (runCounter == 0)
                                {
                                    r.Text.Text = title;
                                }
                                else
                                {
                                    r.Text.Text = "";
                                }
                            }

                        }
                        else
                        {
                            p.RemoveAllChildren();
                        }
                    }
                }
                if (shp.InnerText.Contains(Properties.Settings.Default.CONTENT_SLIDE_CONTENT_CONTAINER_KEYWORD))
                {
                    foreach (Drawing.Paragraph p in shp.TextBody.Elements<Drawing.Paragraph>())
                    {
                        if (p.InnerText.Contains(Properties.Settings.Default.CONTENT_SLIDE_CONTENT_CONTAINER_KEYWORD))
                        {
                            int runCounter = 0;
                            foreach (Drawing.Run r in p.Elements<Drawing.Run>())
                            {
                                if (runCounter == 0)
                                {
                                    r.Text.Text = Properties.Settings.Default.CONTENT_SLIDE_DEFAULT_CONTENT;
                                }
                                else
                                {
                                    r.Text.Text = "";
                                }
                            }

                        }
                        else
                        {
                            p.RemoveAllChildren();
                        }
                    }
                }

            }

            NotesSlidePart notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
            NotesSlide notesSlide = CreateDeafultNotesPart(string.Empty);

            string SlideNotesText = "";
            SlideNotesText = Properties.Settings.Default.HEADING2_MARKUP
                + title + 
                Properties.Settings.Default.HEADING2_MARKUP + 
                Environment.NewLine + Environment.NewLine + 
                Properties.Settings.Default.DEFAULT_NOTES_ENDING;

            notesSlidePart.NotesSlide = notesSlide;
            Shape shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            Drawing.Run newRun = new Drawing.Run();
            Drawing.RunProperties runProp = new Drawing.RunProperties();
            runProp.Bold = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            Drawing.Text newText = new Drawing.Text(SlideNotesText);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            Drawing.Paragraph newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);

            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                slideNumber--;
                if (slideNumber == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            presentationPart.Presentation.Save();
            thisSlideNumber++;
            return thisSlideNumber;
        }



        private int InsertTopicDividerSlide(PresentationDocument presentationDocument, int slideNumber, string topicTitle, List<string> topicDividerNotes)
        {
            int thisSlideNumber = slideNumber;
            if (presentationDocument == null)
            {
                MessageBox.Show("Can't insert slide in the presentation.");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart == null)
            {
                MessageBox.Show("The presentation document is empty.");
            }

            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            string layoutName = Properties.Settings.Default.TOPIC_DIVIDER_SLIDE;

            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slideLayoutPart = null;
            foreach (SlideMasterPart smp in presentationPart.SlideMasterParts)
            {
                foreach (SlideLayoutPart slp in smp.SlideLayoutParts)
                {
                    if (slp.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName, StringComparison.OrdinalIgnoreCase))
                    {
                        slideLayoutPart = slp;
                    }
                }
            }

            if (slideLayoutPart == null)
            {
                MessageBox.Show("Slide Layout not found");
                return slideNumber;
            }
            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideLayoutPart.SlideLayout.CommonSlideData.Clone();
            using (Stream stream = slideLayoutPart.GetStream())
            {
                slidePart.SlideLayoutPart.FeedData(stream);
            }

            foreach (ImagePart iPart in slideLayoutPart.ImageParts)
            {
                ImagePart newImagePart = slidePart.AddImagePart(iPart.ContentType, slideLayoutPart.GetIdOfPart(iPart));
                newImagePart.FeedData(iPart.GetStream());
            }


            foreach (Shape shp in slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>())
            {
                if (shp.InnerText.Contains(Properties.Settings.Default.TOPIC_DIVIDER_SLIDE_CONTAINER_KEYWORD))
                {
                    foreach (Drawing.Paragraph p in shp.TextBody.Elements<Drawing.Paragraph>())
                    {
                        if (p.InnerText.Contains(Properties.Settings.Default.TOPIC_DIVIDER_SLIDE_CONTAINER_KEYWORD))
                        {
                            int runCounter = 0;
                            foreach (Drawing.Run r in p.Elements<Drawing.Run>())
                            {
                                if (runCounter == 0)
                                {
                                    r.Text.Text = topicTitle;
                                }
                                else
                                {
                                    r.Text.Text = "";
                                }
                            }
                        }
                        else
                        {
                            p.RemoveAllChildren();
                        }
                    }
                }
            }

            NotesSlidePart notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
            NotesSlide notesSlide = CreateDeafultNotesPart(string.Empty);

            string SlideNotesText = "";
            SlideNotesText = Properties.Settings.Default.HEADING1_MARKUP +
                topicTitle + 
                Properties.Settings.Default.HEADING1_MARKUP + 
                Environment.NewLine + Environment.NewLine + 
                Properties.Settings.Default.DEFAULT_NOTES_ENDING;

            notesSlidePart.NotesSlide = notesSlide;
            Shape shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            Drawing.Run newRun = new Drawing.Run();
            Drawing.RunProperties runProp = new Drawing.RunProperties();
            runProp.Bold = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            Drawing.Text newText = new Drawing.Text(SlideNotesText);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            Drawing.Paragraph newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);

            SlideNotesText = Properties.Settings.Default.INSTRUCTOR_NOTES_MARKUP;
            SlideNotesText += Properties.Settings.Default.TOPIC_DIVIDER_TOPIC_OBJECTIVE + topicDividerNotes[0] + Environment.NewLine;
            SlideNotesText += Properties.Settings.Default.TOPIC_DIVIDER_TOPIC_METHODS + topicDividerNotes[1] + Environment.NewLine;
            SlideNotesText += Properties.Settings.Default.TOPIC_DIVIDER_TOPIC_EVALUATION + topicDividerNotes[2] + Environment.NewLine;
            SlideNotesText += Properties.Settings.Default.TOPIC_DIVIDER_TOPIC_TIME + topicDividerNotes[3] + Environment.NewLine;
            SlideNotesText += Properties.Settings.Default.TOPIC_DIVIDER_TOPIC_MATERIALS + topicDividerNotes[4] + Environment.NewLine;
            SlideNotesText += Properties.Settings.Default.TOPIC_DIVIDER_TOPIC_NOTES + topicDividerNotes[5] + Environment.NewLine;
            SlideNotesText += Properties.Settings.Default.INSTRUCTOR_NOTES_MARKUP;

            shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            newRun = new Drawing.Run();
            runProp = new Drawing.RunProperties();
            runProp.Italic = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            newText = new Drawing.Text(SlideNotesText);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);

            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                slideNumber--;
                if (slideNumber == 0)
                {
                    prevSlideId = slideId;
                }

            }
            maxSlideId++;

            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }

            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();

            thisSlideNumber++;
            return thisSlideNumber;
        }


        private int InsertLessonObjectivesSlide(PresentationDocument presentationDocument, int slideNumber, List<string> objectives, bool beginningOfSubject)
        {
            int thisSlideNumber = slideNumber;
            if (presentationDocument == null)
            {
                MessageBox.Show("Can't insert slide in the presentation.");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart == null)
            {
                MessageBox.Show("The presentation document is empty.");
            }

            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            string layoutName = Properties.Settings.Default.LESSON_OBJECTIVES_SLIDE;
            if (!beginningOfSubject)
                layoutName = Properties.Settings.Default.LESSON_OBJECTIVES_REVIEW_SLIDE;

            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slideLayoutPart = null;
            foreach (SlideMasterPart smp in presentationPart.SlideMasterParts)
            {
                foreach (SlideLayoutPart slp in smp.SlideLayoutParts)
                {
                    if (slp.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName, StringComparison.OrdinalIgnoreCase))
                    {
                        slideLayoutPart = slp;
                    }
                }
            }

            if (slideLayoutPart == null)
            {
                MessageBox.Show("Slide Layout not found");
                return slideNumber;
            }
            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideLayoutPart.SlideLayout.CommonSlideData.Clone();
            using (Stream stream = slideLayoutPart.GetStream())
            {
                slidePart.SlideLayoutPart.FeedData(stream);
            }

            foreach (ImagePart iPart in slideLayoutPart.ImageParts)
            {
                ImagePart newImagePart = slidePart.AddImagePart(iPart.ContentType, slideLayoutPart.GetIdOfPart(iPart));
                newImagePart.FeedData(iPart.GetStream());
            }


            foreach (Shape shp in slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>())
            {
                if (shp.InnerText.Contains(Properties.Settings.Default.LESSON_OBJECTIVES_SLIDE_CONTAINER_KEYWORD))
                {
                    foreach (Drawing.Paragraph p in shp.TextBody.Elements<Drawing.Paragraph>())
                    {
                        if (p.InnerText.Contains(Properties.Settings.Default.LESSON_OBJECTIVES_SLIDE_CONTAINER_KEYWORD))
                        {
                            int runCounter = 0;
                            foreach (Drawing.Run r in p.Elements<Drawing.Run>())
                            {
                                if (runCounter == 0)
                                {

                                    r.Text.Text = "";
                                    int objcounter = 0;
                                    foreach (string obj in objectives)
                                    {
                                        if (objcounter < objectives.Count - 1)
                                        {
                                            r.Text.Text += obj + Environment.NewLine;
                                        }
                                        else
                                        {
                                            r.Text.Text += obj;
                                        }
                                        objcounter++;
                                    }
                                }
                                else
                                {
                                    r.Text.Text = "";
                                }
                            }

                        }
                        else
                        {
                            p.RemoveAllChildren();
                        }
                    }
                }


            }

            NotesSlidePart notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
            string defaultNotes = beginningOfSubject ? string.Empty : (Properties.Settings.Default.FORCE_SLIDE_HANDBOOK_NO + Environment.NewLine);
            NotesSlide notesSlide = CreateDeafultNotesPart(defaultNotes);

            string SlideNotesText = "";
            SlideNotesText = Properties.Settings.Default.HEADING2_MARKUP+
                Properties.Settings.Default.LESSON_OBJECTIVES_SLIDE_TITLE+
                Properties.Settings.Default.HEADING2_MARKUP+
                Environment.NewLine + Environment.NewLine + 
                Properties.Settings.Default.DEFAULT_NOTES_ENDING;

            if (!beginningOfSubject)
                SlideNotesText = Properties.Settings.Default.HEADING2_MARKUP +
                Properties.Settings.Default.LESSON_OBJECTIVES_REVIEW_SLIDE_TITLE +
                Properties.Settings.Default.HEADING2_MARKUP +
                Environment.NewLine + Environment.NewLine +
                Properties.Settings.Default.DEFAULT_NOTES_ENDING;

            notesSlidePart.NotesSlide = notesSlide;
            Shape shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            Drawing.Run newRun = new Drawing.Run();
            Drawing.RunProperties runProp = new Drawing.RunProperties();
            runProp.Bold = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            Drawing.Text newText = new Drawing.Text(SlideNotesText);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            Drawing.Paragraph newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);

            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                slideNumber--;
                if (slideNumber == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }
            
            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();

            thisSlideNumber++;
            return thisSlideNumber;
        }



        private int InsertLessonGoalSlide(PresentationDocument presentationDocument, int slideNumber, string lessonGoal, bool beginningOfSubject)
        {
            int thisSlideNumber = slideNumber;
            if (presentationDocument == null)
            {
                MessageBox.Show("Can't insert slide in the presentation.");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart == null)
            {
                MessageBox.Show("The presentation document is empty.");
            }

            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            string layoutName = Properties.Settings.Default.LESSON_GOAL_SLIDE;
            if (!beginningOfSubject)
                layoutName = Properties.Settings.Default.LESSON_GOAL_OBTAINED_SLIDE;

            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slideLayoutPart = null;
            foreach (SlideMasterPart smp in presentationPart.SlideMasterParts)
            {
                foreach (SlideLayoutPart slp in smp.SlideLayoutParts)
                {
                    if (slp.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName, StringComparison.OrdinalIgnoreCase))
                    {
                        slideLayoutPart = slp;
                    }
                }
            }

            if (slideLayoutPart == null)
            {
                MessageBox.Show("Slide Layout not found");
                return slideNumber;
            }
            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideLayoutPart.SlideLayout.CommonSlideData.Clone();
            using (Stream stream = slideLayoutPart.GetStream())
            {
                slidePart.SlideLayoutPart.FeedData(stream);
            }

            foreach (ImagePart iPart in slideLayoutPart.ImageParts)
            {
                ImagePart newImagePart = slidePart.AddImagePart(iPart.ContentType, slideLayoutPart.GetIdOfPart(iPart));
                newImagePart.FeedData(iPart.GetStream());
            }


            foreach (Shape shp in slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>())
            {
                if (shp.InnerText.Contains(Properties.Settings.Default.LESSON_GOAL_SLIDE_CONTAINER_KEYWORD))
                {
                    foreach (Drawing.Paragraph p in shp.TextBody.Elements<Drawing.Paragraph>())
                    {
                        if (p.InnerText.Contains(Properties.Settings.Default.LESSON_GOAL_SLIDE_CONTAINER_KEYWORD))
                        {
                            int runCounter = 0;
                            foreach (Drawing.Run r in p.Elements<Drawing.Run>())
                            {
                                if (runCounter == 0)
                                {
                                    r.Text.Text = lessonGoal;
                                }
                                else
                                {
                                    r.Text.Text = "";
                                }
                            }
                        }
                        else
                        {
                            p.RemoveAllChildren();
                        }
                    }
                }


            }

            NotesSlidePart notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
            string defaultNotes = beginningOfSubject ? string.Empty : (Properties.Settings.Default.FORCE_SLIDE_HANDBOOK_NO + Environment.NewLine);
            NotesSlide notesSlide = CreateDeafultNotesPart(defaultNotes);

            string SlideNotesText = "";
            SlideNotesText = Properties.Settings.Default.HEADING2_MARKUP +
                Properties.Settings.Default.LESSON_GOAL_SLIDE_TITLE +
                Properties.Settings.Default.HEADING2_MARKUP +
                Environment.NewLine + Environment.NewLine +
                Properties.Settings.Default.DEFAULT_NOTES_ENDING;
            if (!beginningOfSubject)
                SlideNotesText = Properties.Settings.Default.HEADING2_MARKUP +
                Properties.Settings.Default.LESSON_GOAL_OBTAINED_SLIDE_TITLE +
                Properties.Settings.Default.HEADING2_MARKUP +
                Environment.NewLine + Environment.NewLine +
                Properties.Settings.Default.DEFAULT_NOTES_ENDING;
            notesSlidePart.NotesSlide = notesSlide;
            Shape shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            Drawing.Run newRun = new Drawing.Run();
            Drawing.RunProperties runProp = new Drawing.RunProperties();
            runProp.Bold = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            Drawing.Text newText = new Drawing.Text(SlideNotesText);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            Drawing.Paragraph newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);

            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                slideNumber--;
                if (slideNumber == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }


            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();
            thisSlideNumber++;
            return thisSlideNumber;
        }

        private int InsertIntroductionSlide(PresentationDocument presentationDocument, int slideNumber)
        {
            int thisSlideNumber = slideNumber;
            if (presentationDocument == null)
            {
                MessageBox.Show("Can't insert slide in the presentation.");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart == null)
            {
                MessageBox.Show("The presentation document is empty.");
            }

            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            string layoutName = Properties.Settings.Default.INTRODUCTION_SLIDE;

            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slideLayoutPart = null;
            foreach (SlideMasterPart smp in presentationPart.SlideMasterParts)
            {
                foreach (SlideLayoutPart slp in smp.SlideLayoutParts)
                {
                    if (slp.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName, StringComparison.OrdinalIgnoreCase))
                    {
                        slideLayoutPart = slp;
                    }
                }
            }

            if (slideLayoutPart == null)
            {
                MessageBox.Show("Slide Layout not found");
                return slideNumber;
            }
            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideLayoutPart.SlideLayout.CommonSlideData.Clone();
            using (Stream stream = slideLayoutPart.GetStream())
            {
                slidePart.SlideLayoutPart.FeedData(stream);
            }

            foreach (ImagePart iPart in slideLayoutPart.ImageParts)
            {
                ImagePart newImagePart = slidePart.AddImagePart(iPart.ContentType, slideLayoutPart.GetIdOfPart(iPart));
                newImagePart.FeedData(iPart.GetStream());
            }

           
            NotesSlidePart notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
            NotesSlide notesSlide = CreateDeafultNotesPart(Properties.Settings.Default.FORCE_SLIDE_HANDBOOK_NO + Environment.NewLine);

            string notesText = Properties.Settings.Default.HEADING1_MARKUP +
                Properties.Settings.Default.INTRODUCTION_SUBJECT_TITLE +
                Properties.Settings.Default.HEADING1_MARKUP + Environment.NewLine + Environment.NewLine;

            notesSlidePart.NotesSlide = notesSlide;
            Shape shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            Drawing.Run newRun = new Drawing.Run();
            Drawing.RunProperties runProp = new Drawing.RunProperties();
            runProp.Bold = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            Drawing.Text newText = new Drawing.Text(notesText);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            Drawing.Paragraph newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);

            shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            newRun = new Drawing.Run();
            runProp = new Drawing.RunProperties();
            runProp.Bold = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            newText = new Drawing.Text(Properties.Settings.Default.DEFAULT_NOTES_ENDING + Environment.NewLine);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);


            // Modify the slide ID list in the presentation part.
            // The slide ID list should not be null.
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                slideNumber--;
                if (slideNumber == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }


            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();

            thisSlideNumber++;
            return thisSlideNumber;
        }



        private int InsertTitleSlide(PresentationDocument presentationDocument, int slideNumber, string slideTitle, string subTitle, bool firstSlide)
        {
            int thisSlideNumber = slideNumber;
            if (presentationDocument == null)
            {
                MessageBox.Show("Can't insert slide in the presentation.");
            }

            PresentationPart presentationPart = presentationDocument.PresentationPart;

            if (presentationPart == null)
            {
                MessageBox.Show("The presentation document is empty.");
            }

            Slide slide = new Slide(new CommonSlideData(new ShapeTree()));
            SlidePart slidePart = presentationPart.AddNewPart<SlidePart>();
            slide.Save(slidePart);

            string layoutName = Properties.Settings.Default.TITLE_SLIDE;

            SlideMasterPart slideMasterPart = presentationPart.SlideMasterParts.First();
            SlideLayoutPart slideLayoutPart = null;
            foreach (SlideMasterPart smp in presentationPart.SlideMasterParts)
            {
                foreach (SlideLayoutPart slp in smp.SlideLayoutParts)
                {
                    if (slp.SlideLayout.CommonSlideData.Name.Value.Equals(layoutName, StringComparison.OrdinalIgnoreCase))
                    {
                        slideLayoutPart = slp;
                    }
                }
            }

            if (slideLayoutPart == null)
            {
                MessageBox.Show("Slide Layout not found");
                return slideNumber;
            }

            slidePart.AddPart<SlideLayoutPart>(slideLayoutPart);
            slidePart.Slide.CommonSlideData = (CommonSlideData)slideLayoutPart.SlideLayout.CommonSlideData.Clone();
            using (Stream stream = slideLayoutPart.GetStream())
            {
                slidePart.SlideLayoutPart.FeedData(stream);
            }

            foreach (ImagePart iPart in slideLayoutPart.ImageParts)
            {
                ImagePart newImagePart = slidePart.AddImagePart(iPart.ContentType, slideLayoutPart.GetIdOfPart(iPart));
                newImagePart.FeedData(iPart.GetStream());
            }
            
            foreach (Shape shp in slidePart.Slide.CommonSlideData.ShapeTree.Elements<Shape>())
            {
                if (shp.InnerText.Contains(Properties.Settings.Default.TITLE_SLIDE_MAIN_TITLE_KEYWORD))
                {
                    foreach (Drawing.Paragraph p in shp.TextBody.Elements<Drawing.Paragraph>())
                    {
                        if (p.InnerText.Contains(Properties.Settings.Default.TITLE_SLIDE_MAIN_TITLE_KEYWORD))
                        {
                            int runCounter = 0;
                            foreach (Drawing.Run r in p.Elements<Drawing.Run>())
                            {
                                if (runCounter == 0)
                                {
                                    r.Text.Text = slideTitle;
                                }
                                else
                                {
                                    r.Text.Text = "";
                                }
                            }
                        }
                        else
                        {
                            p.RemoveAllChildren();
                        }
                    }
                }
                if (shp.InnerText.Contains(Properties.Settings.Default.TITLE_SLIDE_SUBTITLE_KEYWORD))
                {
                    foreach (Drawing.Paragraph p in shp.TextBody.Elements<Drawing.Paragraph>())
                    {
                        if (p.InnerText.Contains(Properties.Settings.Default.TITLE_SLIDE_SUBTITLE_KEYWORD))
                        {
                            int runCounter = 0;
                            foreach (Drawing.Run r in p.Elements<Drawing.Run>())
                            {
                                if (runCounter == 0)
                                {
                                    r.Text.Text = subTitle;
                                }
                                else
                                {
                                    r.Text.Text = "";
                                }
                            }
                        }
                        else
                        {
                            p.RemoveAllChildren();
                        }
                    }
                }

            }

            NotesSlidePart notesSlidePart = slidePart.AddNewPart<NotesSlidePart>();
            NotesSlide notesSlide = CreateDeafultNotesPart(string.Empty);


            string SlideNotesText = "";
            if (firstSlide)
            {
                SlideNotesText = Properties.Settings.Default.DEFAULT_TITLE_SLIDE_NOTES + 
                    Environment.NewLine + Environment.NewLine + 
                    Properties.Settings.Default.DEFAULT_NOTES_ENDING;
            }
            else
            {
                SlideNotesText = Properties.Settings.Default.HEADING1_MARKUP + subTitle + Properties.Settings.Default.HEADING1_MARKUP
                    + Environment.NewLine + Environment.NewLine + 
                    Properties.Settings.Default.DEFAULT_NOTES_ENDING;
            }
            notesSlidePart.NotesSlide = notesSlide;
            Shape shpNotes = (Shape)notesSlide.CommonSlideData.ShapeTree.LastChild;
            Drawing.Run newRun = new Drawing.Run();
            Drawing.RunProperties runProp = new Drawing.RunProperties();
            runProp.Bold = DocumentFormat.OpenXml.BooleanValue.FromBoolean(true);
            Drawing.Text newText = new Drawing.Text(SlideNotesText);
            newRun.AppendChild<Drawing.Text>(newText);
            newRun.RunProperties = runProp;
            Drawing.Paragraph newP = new Drawing.Paragraph(newRun);
            shpNotes.TextBody.Append(newP);
            
            SlideIdList slideIdList = presentationPart.Presentation.SlideIdList;

            // Find the highest slide ID in the current list.
            uint maxSlideId = 1;
            SlideId prevSlideId = null;

            foreach (SlideId slideId in slideIdList.ChildElements)
            {
                if (slideId.Id > maxSlideId)
                {
                    maxSlideId = slideId.Id;
                }

                slideNumber--;
                if (slideNumber == 0)
                {
                    prevSlideId = slideId;
                }

            }

            maxSlideId++;

            // Get the ID of the previous slide.
            SlidePart lastSlidePart;

            if (prevSlideId != null)
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(prevSlideId.RelationshipId);
            }
            else
            {
                lastSlidePart = (SlidePart)presentationPart.GetPartById(((SlideId)(slideIdList.ChildElements[0])).RelationshipId);
            }


            // Insert the new slide into the slide list after the previous slide.
            SlideId newSlideId = slideIdList.InsertAfter(new SlideId(), prevSlideId);
            newSlideId.Id = maxSlideId;
            newSlideId.RelationshipId = presentationPart.GetIdOfPart(slidePart);

            // Save the modified presentation.
            presentationPart.Presentation.Save();

            thisSlideNumber++;
            return thisSlideNumber;
        }

        private NotesSlide CreateDeafultNotesPart(string force)
        {
            string defaultNotesText = force +
                Properties.Settings.Default.DEFAULT_NOTES_TEXT +
                Environment.NewLine;
            NotesSlide notesSlide = new NotesSlide(
                        new CommonSlideData(new ShapeTree(
                          new NonVisualGroupShapeProperties(
                            new NonVisualDrawingProperties() { Id = (UInt32)1U, Name = "" },
                            new NonVisualGroupShapeDrawingProperties(),
                            new ApplicationNonVisualDrawingProperties()),
                            new GroupShapeProperties(new Drawing.TransformGroup()),
                            new Shape(
                                new NonVisualShapeProperties(
                                    new NonVisualDrawingProperties() { Id = (UInt32)2U, Name = "Slide Image Placeholder 1" },
                                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true, NoRotation = true, NoChangeAspect = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.SlideImage })),
                                new ShapeProperties()),
                            new Shape(
                                new NonVisualShapeProperties(
                                    new NonVisualDrawingProperties() { Id = (UInt32)3U, Name = "Notes Placeholder 2" },
                                    new NonVisualShapeDrawingProperties(new Drawing.ShapeLocks() { NoGrouping = true }),
                                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape() { Type = PlaceholderValues.Body, Index = (UInt32)1U })),
                                new ShapeProperties(),
                                new TextBody(
                                    new Drawing.BodyProperties(),
                                    new Drawing.ListStyle(),
                                    new Drawing.Paragraph(
                                        new Drawing.Run(
                                            new Drawing.RunProperties() { Language = "en-US", Dirty = false },
                                            new Drawing.Text() { Text = defaultNotesText }),
                                        new Drawing.EndParagraphRunProperties() { Language = "en-US", Dirty = false }))
                                    ))),
                        new ColorMapOverride(new Drawing.MasterColorMapping()));
            return notesSlide;
        }

        public void CopyTemplateToPresentation(string presentationFileTemplate, string presentationFile, BackgroundWorker bckgW)
        {
            try
            {
                var app = new Microsoft.Office.Interop.PowerPoint.Application();
                var pres = app.Presentations;
                var pptfile = pres.Open(presentationFileTemplate, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse);
                pptfile.SaveCopyAs(presentationFile, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, Microsoft.Office.Core.MsoTriState.msoTrue);
                
            }
            catch (Exception ex)
            {
                AddToLog("Problem with saving PPTM to PPTX. " + ex.Message);
                MessageBox.Show("Powerpoint template problem. " + ex.Message);
            }
            bckgW.ReportProgress(40);
            //Thread.Sleep(1000);
            var file = new FileInfo(presentationFile);
            while (IsFileLocked(file))
            { }
            bckgW.ReportProgress(50);
        }

        private bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;
            try
            {
                stream = file.Open(FileMode.Open, FileAccess.ReadWrite, FileShare.None);
            }
            catch (IOException)
            {
                return true;
            }
            finally
            {
                if (stream != null)
                    stream.Close();
            }
            return false;
        }

    }
}
