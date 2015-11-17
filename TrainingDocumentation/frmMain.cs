using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace TrainingDocumentation
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();

            bckgHandbook.DoWork += BckgHandbook_DoWork;
            bckgHandbook.ProgressChanged += BckgHandbook_ProgressChanged;
            bckgHandbook.WorkerReportsProgress = true;
            bckgHandbook.RunWorkerCompleted += BckgHandbook_RunWorkerCompleted;

            bckgLP.DoWork += BckgLP_DoWork;
            bckgLP.ProgressChanged += BckgLP_ProgressChanged;
            bckgLP.WorkerReportsProgress = true;
            bckgLP.RunWorkerCompleted += BckgLP_RunWorkerCompleted;

        }

        private void BckgLP_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnLPConvert.Enabled = true;
            UpdateLPLabel("Done!");
        }

        private void BckgLP_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarLP.Value = e.ProgressPercentage;
        }

        private void BckgLP_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                AddToLog("Start Lesson plan conversion.");
                this.Invoke(new MethodInvoker(delegate { UpdateLPLabel("Reading lesson plan..."); }));
                bckgLP.ReportProgress(0);
                LPConvertDOC lpDOC = new LPConvertDOC();
                string lessonPlanFileName = txtLPFile.Text;
                AddToLog("File to convert: " + txtLPFile.Text);
                string xmlFileName = Path.GetDirectoryName(lessonPlanFileName) + "\\" + Path.GetFileNameWithoutExtension(lessonPlanFileName) + ".xml";
                AddToLog("Creating XML file: " + xmlFileName);
                lpDOC.CreateXMLFromLPTable(lessonPlanFileName, xmlFileName, bckgLP);
                bckgLP.ReportProgress(33);
                this.Invoke(new MethodInvoker(delegate { UpdateLPLabel("Preparing Powerpoint template..."); }));
                string presentationFileTemplate = Environment.CurrentDirectory + "\\Templates\\Template.pptm";
                string presentationFile = txtLPSaveLocation.Text;
                LPConvertPPT lpcPPT = new LPConvertPPT();
                AddToLog("Copy PPTM to PPTX.");
                lpcPPT.CopyTemplateToPresentation(presentationFileTemplate, presentationFile, bckgLP);
                this.Invoke(new MethodInvoker(delegate { UpdateLPLabel("Filling the template..."); }));
                DocumentFormat.OpenXml.Packaging.PresentationDocument presentationDocument = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(presentationFile, true);
                AddToLog("Creating presentation from XML.");
                lpcPPT.CreateTemplateFromXML(xmlFileName, presentationDocument, bckgLP);
                presentationDocument.Close();
                if (chkLPDeleteXML.Checked)
                {
                    AddToLog("Deleting XML file.");
                    File.Delete(xmlFileName);
                }
                bckgLP.ReportProgress(100);
                AddToLog("Lesson plan conversion done.");
            }
            catch(Exception ex)
            {
                AddToLog("Problem in Lesson plan conversion: " + ex.Message);
                MessageBox.Show("There was a problem with Lesson plan conversion." + ex.Message);
            }
        }

        private void BckgHandbook_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnHandbookConvert.Enabled = true;
            UpdateHandbookLabel("Done!");
        }

        private void BckgHandbook_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            progressBarHandbook.Value = e.ProgressPercentage;
        }

        private void BckgHandbook_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                AddToLog("Start Handbook conversion.");
                this.Invoke(new MethodInvoker(delegate { UpdateHandbookLabel("Exporting slides (this may take a while)..."); }));
                bckgHandbook.ReportProgress(0);

                HandbookPPT hbPPT = new HandbookPPT();
                string pptFileName = txtHandbookFile.Text;
                AddToLog("Start exporting slides from: " + pptFileName);
                string exportSlidesRetVal = hbPPT.ExportSlides(pptFileName);
                bckgHandbook.ReportProgress(33);

                this.Invoke(new MethodInvoker(delegate { UpdateHandbookLabel("Creating the handbook..."); }));
                HandbookDoc hbDOC = new HandbookDoc();
                string docFileName = txtHandbookSaveLocation.Text;
                bool instructorGuide = chkHandbookInstructorGuide.Checked;
                AddToLog("Start creating handbook document: " + docFileName);
                hbDOC.CreateDocument(pptFileName, bckgHandbook, docFileName, instructorGuide);
                bckgHandbook.ReportProgress(100);
                AddToLog("Handbook conversion finished.");
                if (chkHandbookDeletePictures.Checked)
                {
                    string folderName = Path.GetDirectoryName(pptFileName) + "\\" + Path.GetFileNameWithoutExtension(pptFileName);
                    AddToLog("Deleting slide pictures folder: " + folderName);
                    Directory.Delete(folderName, true);
                }
            }
            catch(Exception ex)
            {
                AddToLog("Problem in handbook conversion: " + ex.Message);
                MessageBox.Show("There was a problem with handbook conversion." + ex.Message);
            }
        }

        private void UpdateHandbookLabel(string lblText)
        {
            lblHandbookStatus.Text = lblText;
            lblHandbookStatus.Left = progressBarHandbook.Left + (progressBarHandbook.Width / 2 - lblHandbookStatus.Width / 2);
        }

        private void UpdateLPLabel(string lblText)
        {
            lblLPStatus.Text = lblText;
            lblLPStatus.Left = progressBarLP.Left + (progressBarLP.Width / 2 - lblLPStatus.Width / 2);
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void chkHandbookOverrideDefaultSave_CheckedChanged(object sender, EventArgs e)
        {
            if(chkHandbookOverrideDefaultSave.Checked)
            {
                txtHandbookSaveLocation.Enabled = true;
                btnHandbookBrowseSave.Enabled = true;
            }
            else
            {
                txtHandbookSaveLocation.Enabled = false;
                btnHandbookBrowseSave.Enabled = false;
                txtHandbookSaveLocation.Text = Path.GetDirectoryName(txtHandbookFile.Text) + "\\" + Path.GetFileNameWithoutExtension(txtHandbookFile.Text) + ".docx";
            }
        }

        private void btnHandbookBrowseOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileHandbook = new OpenFileDialog();
            openFileHandbook.Filter = "Powerpoint presentations (.pptx)|*.pptx";
            openFileHandbook.Title = "Open presentation";
            openFileHandbook.Multiselect = false;
            if (txtHandbookFile.Text == "")
            {
                openFileHandbook.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileHandbook.FileName = "";
            }
            else
            {
                openFileHandbook.InitialDirectory = Path.GetDirectoryName(txtHandbookFile.Text);
                openFileHandbook.FileName = Path.GetFileName(txtHandbookFile.Text);
            }
            DialogResult dlgOK = openFileHandbook.ShowDialog();
            if(dlgOK==DialogResult.OK)
            {
                txtHandbookFile.Text = openFileHandbook.FileName;
                if(!chkHandbookOverrideDefaultSave.Checked)
                {
                    txtHandbookSaveLocation.Text = Path.GetDirectoryName(txtHandbookFile.Text) + "\\" + Path.GetFileNameWithoutExtension(txtHandbookFile.Text) + ".docx";
                }
            }
            
        }

        private void btnHandbookBrowseSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileHandbook = new SaveFileDialog();
            saveFileHandbook.Filter = "Word document (*.docx)|*.docx";
            saveFileHandbook.Title = "Save handbook file";
            saveFileHandbook.ShowDialog();
            if (saveFileHandbook.FileName != "")
            {
                txtHandbookSaveLocation.Text = saveFileHandbook.FileName;
            }
        }

        private void btnHandbookConvert_Click(object sender, EventArgs e)
        {
            if(txtHandbookFile.Text!="" && txtHandbookSaveLocation.Text!="")
            {
                btnHandbookConvert.Enabled = false;
                bckgHandbook.RunWorkerAsync();
            }
        }

        private void btnLPConvert_Click(object sender, EventArgs e)
        {
            
            btnLPConvert.Enabled = false;
            bckgLP.RunWorkerAsync();
        }

        private void btnLPBrowseOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileLP = new OpenFileDialog();
            openFileLP.Filter = "Word document (.docx)|*.docx";
            openFileLP.Title = "Open Lesson plan";
            openFileLP.Multiselect = false;
            if (txtLPFile.Text == "")
            {
                openFileLP.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                openFileLP.FileName = "";
            }
            else
            {
                openFileLP.InitialDirectory = Path.GetDirectoryName(txtLPFile.Text);
                openFileLP.FileName = Path.GetFileName(txtLPFile.Text);
            }
            DialogResult dlgOK = openFileLP.ShowDialog();
            if (dlgOK == DialogResult.OK)
            {
                txtLPFile.Text = openFileLP.FileName;
                if (!chkLPOverrideDefaultSave.Checked)
                {
                    txtLPSaveLocation.Text = Path.GetDirectoryName(txtLPFile.Text) + "\\" + Path.GetFileNameWithoutExtension(txtLPFile.Text) + ".pptx";
                }
            }
        }

        private void btnLPBrowseSave_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileLP = new SaveFileDialog();
            saveFileLP.Filter = "Powerpoint presentation (*.pptx)|*.pptx";
            saveFileLP.Title = "Save presentation template";
            saveFileLP.ShowDialog();
            if (saveFileLP.FileName != "")
            {
                txtLPSaveLocation.Text = saveFileLP.FileName;
            }
        }

        private void chkLPOverrideDefaultSave_CheckedChanged(object sender, EventArgs e)
        {
            if (chkLPOverrideDefaultSave.Checked)
            {
                txtLPSaveLocation.Enabled = true;
                btnLPBrowseSave.Enabled = true;
            }
            else
            {
                txtLPSaveLocation.Enabled = false;
                btnLPBrowseSave.Enabled = false;
                txtLPSaveLocation.Text = Path.GetDirectoryName(txtLPFile.Text) + "\\" + Path.GetFileNameWithoutExtension(txtLPFile.Text) + ".pptx";
            }
        }

        private void txtLPFile_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (s[0].Substring(s[0].Length - 4) == "docx")
                txtLPFile.Text = s[0];
        }

        private void txtLPFile_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void txtHandbookFile_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.All;
            else
                e.Effect = DragDropEffects.None;
        }

        private void txtHandbookFile_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            if (s[0].Substring(s[0].Length - 4) == "pptx")
                txtHandbookFile.Text = s[0];
        }

        private void txtHandbookFile_TextChanged(object sender, EventArgs e)
        {
            if (!chkHandbookOverrideDefaultSave.Checked)
            {
                try
                {
                    txtHandbookSaveLocation.Text = Path.GetDirectoryName(txtHandbookFile.Text) + "\\" + Path.GetFileNameWithoutExtension(txtHandbookFile.Text) + ".docx";
                }
                catch
                { }
            }
        }

        private void txtLPFile_TextChanged(object sender, EventArgs e)
        {
            if (!chkLPOverrideDefaultSave.Checked)
            {
                try
                {
                    txtLPSaveLocation.Text = Path.GetDirectoryName(txtLPFile.Text) + "\\" + Path.GetFileNameWithoutExtension(txtLPFile.Text) + ".pptx";
                }
                catch
                { }
            }
        }

        private void AddToLog(string message)
        {
            string fileName = Environment.CurrentDirectory + "\\Log.txt";
            using (StreamWriter sw = File.AppendText(fileName))
            {
                sw.WriteLine("{0}: {1}", DateTime.Now.ToString(), message);
            }
        }
    }
}
