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
using P=DocumentFormat.OpenXml.Presentation;

namespace TrainingDocumentation
{
    class HandbookPPT
    {
        public string ExportSlides(string fileName)
        {
            string retVal = "";
            try
            {
                var app = new Microsoft.Office.Interop.PowerPoint.Application();
                var pres = app.Presentations;
                var pptfile = pres.Open(fileName, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoTrue, Microsoft.Office.Core.MsoTriState.msoFalse);
                string savePath = Path.GetDirectoryName(fileName) + "\\" + Path.GetFileNameWithoutExtension(fileName) + ".png";
                //Exporting slides to the folder where the PPTX file is, in the subfolder named as the presentation
                pptfile.SaveCopyAs(savePath, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPNG, Microsoft.Office.Core.MsoTriState.msoTrue);
                retVal = "OK";
            }
            catch (Exception ex)
            {
                retVal = ex.Message;
            }
            return retVal;
        }

        

        public string GetNotes(Microsoft.Office.Interop.PowerPoint.Slide slide)
        {
            if (slide.HasNotesPage == MsoTriState.msoFalse)
                return string.Empty;
            string slideNodes = string.Empty;
            var notesPage = slide.NotesPage;
            int length = 0;
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in notesPage.Shapes)
            {
                if (shape.Type == MsoShapeType.msoPlaceholder)
                {
                    var tf = shape.TextFrame;
                    try
                    {
                        //Some TextFrames do not have a range
                        var range = tf.TextRange;
                        if (range.Length > length)
                        {   //Some have a digit in the text, 
                            //so find the longest text item and return that
                            slideNodes = range.Text;
                            length = range.Length;
                        }
                        Marshal.ReleaseComObject(range);
                    }
                    catch (Exception)
                    { }
                    finally
                    { //Ensure clear up
                        Marshal.ReleaseComObject(tf);
                    }
                }
                Marshal.ReleaseComObject(shape);
            }
            return RemoveInstructorNotes(slideNodes);
        }

        private string RemoveInstructorNotes(string inputString)
        {
            string testString = inputString;
            int a = testString.IndexOf("**");
            int b = 0;
            int count = testString.Length - testString.Replace("**", "").Length;
            if (count % 4 == 0)
            {
                while (testString.IndexOf("**") != -1)
                {
                    a = testString.IndexOf("**");
                    b = testString.Substring(a + 2).IndexOf("**");
                    if (b + a > a)
                    {
                        testString = testString.Remove(a, b + 4);
                    }
                }
            }
            return testString;
        }

        public bool SlideVisibile(PresentationDocument presentationDocument, int slideIndex)
        {
            bool retVal = true;

            PresentationPart pp = presentationDocument.PresentationPart;
            P.Presentation p = pp.Presentation;
            P.SlideIdList slideIDList = p.SlideIdList;
            P.SlideId slideId = slideIDList.ChildElements[slideIndex] as P.SlideId;
            string slideRelID = slideId.RelationshipId;
            SlidePart slidePart = pp.GetPartById(slideRelID) as SlidePart;
            if (slidePart.Slide.Show != null)
            {
                if (!slidePart.Slide.Show.Value)
                {
                    retVal = false;
                }
                else
                {
                    retVal = true;
                }
            }
            else
            {
                retVal = true;
            }

            return retVal;
        }


    }
}
