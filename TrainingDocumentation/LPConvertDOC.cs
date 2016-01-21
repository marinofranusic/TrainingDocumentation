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
    class LPConvertDOC
    {

        public int CreateXMLFromLPTable(string filepath, string xmlLP, BackgroundWorker bckgW)
        {
            int retVal = 0;
            AddToLog("Start creating XML.");
            // Use the file name and path passed in as an argument to 
            // open an existing document.    
            StringBuilder sb = new StringBuilder();
            
            x.XmlWriter writer = x.XmlWriter.Create(xmlLP);
            writer.WriteStartDocument();

            bool subjectOpen = false;
            bool readSubjectTitle = false;
            int lastSubjectTitleRow = -1;
            int rowCounter = 0;
            int readTableCellFontSize = 0;
            int irrelevant = 0;

            string COURSE_TITLE_STRING = Properties.Settings.Default.COURSE_TITLE_STRING;
            List<string> SUBJECT_DIVIDERS = Properties.Settings.Default.SUBJECT_DIVIDERS.Cast<string>().ToList();
            List<string> SUBJECT_TOPIC_NAMING = Properties.Settings.Default.SUBJECT_TOPIC_NAMING.Cast<string>().ToList();
            string PART_DIVIDER = Properties.Settings.Default.PART_DIVIDER;

            int totalRowCount = 0;

            using (WordprocessingDocument doc = WordprocessingDocument.Open(filepath, true))
            {
                try
                {
                    foreach (WordDoc.Table table in doc.MainDocumentPart.Document.Body.Elements<WordDoc.Table>())
                    {
                        totalRowCount += table.Elements<WordDoc.TableRow>().Count();
                    }
                }
                catch(Exception ex)
                {
                    AddToLog("Problem with reading tables from Lesson plan document. " + ex.Message);
                    MessageBox.Show("Problem with reading tables from Lesson plan document. " + ex.Message);
                    totalRowCount = 0;
                    retVal = -1;
                }

                if (totalRowCount > 0)
                {
                    // Go through all the tables in the document.
                    foreach (WordDoc.Table table in doc.MainDocumentPart.Document.Body.Elements<WordDoc.Table>())
                    {
                        foreach (WordDoc.TableRow tr in table.Elements<WordDoc.TableRow>())
                        {
                            float progress = (((float)rowCounter / totalRowCount) / 3) * 100;
                            bckgW.ReportProgress((int)progress);
                            readTableCellFontSize = 0;
                            try
                            {
                                string readFromCell = ReadTableCell(tr, 0, out readTableCellFontSize);
                                //Course title
                                if (readFromCell.Length >= COURSE_TITLE_STRING.Length)
                                {
                                    if (string.Equals(readFromCell.Substring(0, COURSE_TITLE_STRING.Length), COURSE_TITLE_STRING, StringComparison.OrdinalIgnoreCase))
                                    {
                                        string readCourseTitle = ReadTableCell(tr, 1, out irrelevant);
                                        writer.WriteStartElement("Course");
                                        writer.WriteStartElement("Title");
                                        writer.WriteString(readCourseTitle);
                                        writer.WriteEndElement();
                                        writer.WriteStartElement("Subjects");
                                    }
                                }

                                if (SUBJECT_DIVIDERS.Contains(readFromCell))
                                {
                                    subjectOpen = true;
                                    readSubjectTitle = false;
                                }
                                else
                                {
                                    if ((lastSubjectTitleRow > 0) && (rowCounter - lastSubjectTitleRow) >= 2)
                                    {
                                        string firstWord = "";
                                        if (readFromCell.IndexOf(' ') > 0)
                                        {
                                            firstWord = readFromCell.Substring(0, readFromCell.IndexOf(' '));
                                        }
                                        if (SUBJECT_TOPIC_NAMING.Contains(firstWord))
                                            readSubjectTitle = false;
                                    }

                                    if (readTableCellFontSize == 28 || (!readSubjectTitle && subjectOpen))
                                    //if (!readSubjectTitle && subjectOpen)
                                    {
                                        if (!string.Equals(readFromCell.Substring(0, PART_DIVIDER.Length), PART_DIVIDER, StringComparison.OrdinalIgnoreCase))
                                        {
                                            subjectOpen = true;
                                            string firstWord = "";
                                            if (readFromCell.IndexOf(' ') > 0)
                                            {
                                                firstWord = readFromCell.Substring(0, readFromCell.IndexOf(' '));
                                            }
                                            string subjectTitle = "";
                                            if (SUBJECT_TOPIC_NAMING.Contains(firstWord))
                                            {
                                                int posSemicolon = readFromCell.IndexOf(':');
                                                int posMinus = readFromCell.IndexOf('-');
                                                int posMinus2 = readFromCell.IndexOf('–');
                                                if (posSemicolon > 0 && posSemicolon < 15)
                                                {
                                                    //Subject name starting like Subject 2.1.: Introduction
                                                    subjectTitle = readFromCell.Substring(posSemicolon + 2);
                                                }
                                                else if (posMinus > 0 && posMinus < 15)
                                                {
                                                    subjectTitle = readFromCell.Substring(posMinus + 2);
                                                }
                                                else if (posMinus2 > 0 && posMinus2 < 15)
                                                {
                                                    subjectTitle = readFromCell.Substring(posMinus2 + 2);
                                                }
                                            }
                                            else
                                            {
                                                //Subject name without subject naming, just Introduction
                                                subjectTitle = readFromCell;
                                            }
                                            readSubjectTitle = true;
                                            if (lastSubjectTitleRow != -1)
                                                writer.WriteEndElement();
                                            writer.WriteStartElement("Subject");
                                            retVal++;
                                            writer.WriteAttributeString("SubjectTitle", subjectTitle);
                                            string subjectGoal = ReadTableCell(tr, 1, out irrelevant);
                                            //subjectGoal.Replace(Environment.NewLine, "");
                                            int posSemicolon2 = subjectGoal.IndexOf(':');
                                            if (posSemicolon2 > 0 && posSemicolon2 < 25)
                                            {
                                                //remove "overall subject goal:" or "lesson goal:"
                                                if (subjectGoal.Length > posSemicolon2 + 3)
                                                {
                                                    subjectGoal = subjectGoal.Substring(posSemicolon2 + 2);
                                                }
                                                else
                                                {
                                                    subjectGoal = "";
                                                }
                                            }

                                            writer.WriteStartElement("LessonGoal");
                                            writer.WriteString(subjectGoal);
                                            writer.WriteEndElement();

                                            lastSubjectTitleRow = rowCounter;
                                        }

                                    }
                                    else
                                    {
                                        //Read the topics
                                        if (subjectOpen)
                                        {
                                            if ((readFromCell.Length < PART_DIVIDER.Length) || (readFromCell.Length >= PART_DIVIDER.Length && !string.Equals(readFromCell.Substring(0, PART_DIVIDER.Length), PART_DIVIDER, StringComparison.OrdinalIgnoreCase)))
                                            {
                                                string[] lines = readFromCell.Split(new string[] { Environment.NewLine }, StringSplitOptions.None);
                                                int counter = 0;
                                                foreach (string line in lines)
                                                {
                                                    writer.WriteStartElement("Topic");
                                                    if (counter == 0)
                                                    {
                                                        writer.WriteAttributeString("Heading", "1");
                                                        //for each topic i need to read the objective
                                                        string objective = ReadTableCell(tr, 1, out irrelevant);
                                                        writer.WriteStartElement("Objective");
                                                        writer.WriteString(objective);
                                                        writer.WriteEndElement();
                                                        string methodsActivities = ReadTableCell(tr, 3, out irrelevant);
                                                        writer.WriteStartElement("Methods");
                                                        writer.WriteString(methodsActivities);
                                                        writer.WriteEndElement();
                                                        string evaluation = ReadTableCell(tr, 4, out irrelevant);
                                                        writer.WriteStartElement("Evaluation");
                                                        writer.WriteString(evaluation);
                                                        writer.WriteEndElement();
                                                        string time = ReadTableCell(tr, 5, out irrelevant);
                                                        writer.WriteStartElement("Time");
                                                        writer.WriteString(time);
                                                        writer.WriteEndElement();
                                                        string materials = ReadTableCell(tr, 6, out irrelevant);
                                                        writer.WriteStartElement("Materials");
                                                        writer.WriteString(materials);
                                                        writer.WriteEndElement();
                                                        string notes = ReadTableCell(tr, 7, out irrelevant);
                                                        writer.WriteStartElement("Notes");
                                                        writer.WriteString(notes);
                                                        writer.WriteEndElement();
                                                    }
                                                    else
                                                    {
                                                        writer.WriteAttributeString("Heading", "2");
                                                    }
                                                    //writer.WriteString(line);
                                                    writer.WriteStartElement("Title");
                                                    writer.WriteString(line);
                                                    writer.WriteEndElement();
                                                    writer.WriteEndElement();
                                                    counter++;
                                                }



                                            }
                                        }
                                    }
                                }
                            }
                            catch
                            { }
                            rowCounter++;
                        }

                    }
                }
            }
            try
            {
                if (lastSubjectTitleRow != rowCounter)
                    writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndDocument();
            }
            catch
            {
                retVal = -1;
            }
            finally
            {
                writer.Close();
                AddToLog("XML created.");
            }
            return retVal;
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

        private string ReadTableCell(WordDoc.TableRow tr, int column, out int tableCellFontSize)
        {
            tableCellFontSize = 0;
            string cellFontSize = "";
            int tcellFontSize = 0;
            StringBuilder sb = new StringBuilder();
            WordDoc.TableCell tc = tr.Elements<WordDoc.TableCell>().ElementAt(column);
            if (tc.Elements<WordDoc.Paragraph>().First() != null)
            {
                foreach (WordDoc.Paragraph tcp in tc.Elements<WordDoc.Paragraph>())
                {
                    //Check for bullets
                    WordDoc.ParagraphProperties pp = tcp.Elements<WordDoc.ParagraphProperties>().First();
                    if (pp.NumberingProperties != null)
                    {
                        //LogText(pp.NumberingProperties.NumberingLevelReference.Val.ToString(), true);
                        //LogText(pp.NumberingProperties.NumberingId.Val.ToString(), true);
                    }

                    if (tcp.Elements<WordDoc.Run>().FirstOrDefault() != null)
                    {
                        foreach (WordDoc.Run tcpr in tcp.Elements<WordDoc.Run>())
                        {
                            WordDoc.RunProperties tcprp = tcpr.RunProperties;
                            cellFontSize = tcprp.FontSize.Val.ToString();
                            tcellFontSize = Convert.ToInt32(cellFontSize);
                            if (tcpr.Elements<WordDoc.Text>().FirstOrDefault() != null)
                            {
                                foreach (WordDoc.Text tcprt in tcpr.Elements<WordDoc.Text>())
                                {
                                    sb.Append(tcprt.Text);
                                }
                            }
                        }

                    }
                    sb.Append(Environment.NewLine);
                }

            }
            tableCellFontSize = tcellFontSize;
            return sb.ToString().Trim();
        }

    }
}
