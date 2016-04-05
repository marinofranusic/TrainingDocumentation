namespace TrainingDocumentation
{
    partial class frmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.tabContainer = new System.Windows.Forms.TabControl();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.chkHandbookInstructorGuide = new System.Windows.Forms.CheckBox();
            this.btnHandbookConvert = new System.Windows.Forms.Button();
            this.chkHandbookOverrideDefaultSave = new System.Windows.Forms.CheckBox();
            this.btnHandbookBrowseSave = new System.Windows.Forms.Button();
            this.txtHandbookSaveLocation = new System.Windows.Forms.TextBox();
            this.chkHandbookDeletePictures = new System.Windows.Forms.CheckBox();
            this.progressBarHandbook = new System.Windows.Forms.ProgressBar();
            this.lblHandbookStatus = new System.Windows.Forms.Label();
            this.lblHandbookFileTitle = new System.Windows.Forms.Label();
            this.btnHandbookBrowseOpen = new System.Windows.Forms.Button();
            this.txtHandbookFile = new System.Windows.Forms.TextBox();
            this.tabPage3 = new System.Windows.Forms.TabPage();
            this.btnBatchHandbookClearList = new System.Windows.Forms.Button();
            this.btnBatchHandbookRemove = new System.Windows.Forms.Button();
            this.lstBatchHandbookFiles = new System.Windows.Forms.ListBox();
            this.btnBatchHandbookAddToList = new System.Windows.Forms.Button();
            this.chkBatchHandbookInstructorGuide = new System.Windows.Forms.CheckBox();
            this.btnBatchHandbookConvert = new System.Windows.Forms.Button();
            this.progressBarBatchHandbook = new System.Windows.Forms.ProgressBar();
            this.lblBatchHandbookStatus = new System.Windows.Forms.Label();
            this.lblBatchHandbookTitle = new System.Windows.Forms.Label();
            this.btnBatchHandbookBrowseOpen = new System.Windows.Forms.Button();
            this.txtBatchHandbookFile = new System.Windows.Forms.TextBox();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.chkLPSeparate = new System.Windows.Forms.CheckBox();
            this.rb169 = new System.Windows.Forms.RadioButton();
            this.rb43 = new System.Windows.Forms.RadioButton();
            this.chkLPDeleteXML = new System.Windows.Forms.CheckBox();
            this.chkLPOverrideDefaultSave = new System.Windows.Forms.CheckBox();
            this.btnLPBrowseSave = new System.Windows.Forms.Button();
            this.txtLPSaveLocation = new System.Windows.Forms.TextBox();
            this.lblLPFileTitle = new System.Windows.Forms.Label();
            this.btnLPBrowseOpen = new System.Windows.Forms.Button();
            this.txtLPFile = new System.Windows.Forms.TextBox();
            this.progressBarLP = new System.Windows.Forms.ProgressBar();
            this.lblLPStatus = new System.Windows.Forms.Label();
            this.btnLPConvert = new System.Windows.Forms.Button();
            this.btnExit = new System.Windows.Forms.Button();
            this.bckgHandbook = new System.ComponentModel.BackgroundWorker();
            this.bckgLP = new System.ComponentModel.BackgroundWorker();
            this.label1 = new System.Windows.Forms.Label();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.bckgBatch = new System.ComponentModel.BackgroundWorker();
            this.chkGoalObjSamePage = new System.Windows.Forms.CheckBox();
            this.chkBatchHandbookGoalObjSamePage = new System.Windows.Forms.CheckBox();
            this.tabContainer.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.tabPage3.SuspendLayout();
            this.tabPage1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabContainer
            // 
            this.tabContainer.Controls.Add(this.tabPage2);
            this.tabContainer.Controls.Add(this.tabPage3);
            this.tabContainer.Controls.Add(this.tabPage1);
            this.tabContainer.Location = new System.Drawing.Point(12, 111);
            this.tabContainer.Name = "tabContainer";
            this.tabContainer.SelectedIndex = 0;
            this.tabContainer.Size = new System.Drawing.Size(589, 369);
            this.tabContainer.TabIndex = 0;
            // 
            // tabPage2
            // 
            this.tabPage2.Controls.Add(this.chkGoalObjSamePage);
            this.tabPage2.Controls.Add(this.chkHandbookInstructorGuide);
            this.tabPage2.Controls.Add(this.btnHandbookConvert);
            this.tabPage2.Controls.Add(this.chkHandbookOverrideDefaultSave);
            this.tabPage2.Controls.Add(this.btnHandbookBrowseSave);
            this.tabPage2.Controls.Add(this.txtHandbookSaveLocation);
            this.tabPage2.Controls.Add(this.chkHandbookDeletePictures);
            this.tabPage2.Controls.Add(this.progressBarHandbook);
            this.tabPage2.Controls.Add(this.lblHandbookStatus);
            this.tabPage2.Controls.Add(this.lblHandbookFileTitle);
            this.tabPage2.Controls.Add(this.btnHandbookBrowseOpen);
            this.tabPage2.Controls.Add(this.txtHandbookFile);
            this.tabPage2.Location = new System.Drawing.Point(4, 25);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(581, 340);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Handbook creator";
            this.tabPage2.UseVisualStyleBackColor = true;
            // 
            // chkHandbookInstructorGuide
            // 
            this.chkHandbookInstructorGuide.AutoSize = true;
            this.chkHandbookInstructorGuide.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHandbookInstructorGuide.Location = new System.Drawing.Point(20, 177);
            this.chkHandbookInstructorGuide.Name = "chkHandbookInstructorGuide";
            this.chkHandbookInstructorGuide.Size = new System.Drawing.Size(202, 24);
            this.chkHandbookInstructorGuide.TabIndex = 12;
            this.chkHandbookInstructorGuide.Text = "Create Instructor guide";
            this.chkHandbookInstructorGuide.UseVisualStyleBackColor = true;
            // 
            // btnHandbookConvert
            // 
            this.btnHandbookConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnHandbookConvert.Location = new System.Drawing.Point(233, 249);
            this.btnHandbookConvert.Name = "btnHandbookConvert";
            this.btnHandbookConvert.Size = new System.Drawing.Size(100, 35);
            this.btnHandbookConvert.TabIndex = 11;
            this.btnHandbookConvert.Text = "Convert";
            this.btnHandbookConvert.UseVisualStyleBackColor = true;
            this.btnHandbookConvert.Click += new System.EventHandler(this.btnHandbookConvert_Click);
            // 
            // chkHandbookOverrideDefaultSave
            // 
            this.chkHandbookOverrideDefaultSave.AutoSize = true;
            this.chkHandbookOverrideDefaultSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHandbookOverrideDefaultSave.Location = new System.Drawing.Point(20, 84);
            this.chkHandbookOverrideDefaultSave.Name = "chkHandbookOverrideDefaultSave";
            this.chkHandbookOverrideDefaultSave.Size = new System.Drawing.Size(331, 24);
            this.chkHandbookOverrideDefaultSave.TabIndex = 10;
            this.chkHandbookOverrideDefaultSave.Text = "Override default save location and name";
            this.chkHandbookOverrideDefaultSave.UseVisualStyleBackColor = true;
            this.chkHandbookOverrideDefaultSave.CheckedChanged += new System.EventHandler(this.chkHandbookOverrideDefaultSave_CheckedChanged);
            // 
            // btnHandbookBrowseSave
            // 
            this.btnHandbookBrowseSave.Enabled = false;
            this.btnHandbookBrowseSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnHandbookBrowseSave.Location = new System.Drawing.Point(472, 105);
            this.btnHandbookBrowseSave.Name = "btnHandbookBrowseSave";
            this.btnHandbookBrowseSave.Size = new System.Drawing.Size(100, 35);
            this.btnHandbookBrowseSave.TabIndex = 9;
            this.btnHandbookBrowseSave.Text = "Browse";
            this.btnHandbookBrowseSave.UseVisualStyleBackColor = true;
            this.btnHandbookBrowseSave.Click += new System.EventHandler(this.btnHandbookBrowseSave_Click);
            // 
            // txtHandbookSaveLocation
            // 
            this.txtHandbookSaveLocation.Enabled = false;
            this.txtHandbookSaveLocation.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtHandbookSaveLocation.Location = new System.Drawing.Point(20, 111);
            this.txtHandbookSaveLocation.Name = "txtHandbookSaveLocation";
            this.txtHandbookSaveLocation.Size = new System.Drawing.Size(427, 23);
            this.txtHandbookSaveLocation.TabIndex = 8;
            // 
            // chkHandbookDeletePictures
            // 
            this.chkHandbookDeletePictures.AutoSize = true;
            this.chkHandbookDeletePictures.Checked = true;
            this.chkHandbookDeletePictures.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkHandbookDeletePictures.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHandbookDeletePictures.Location = new System.Drawing.Point(20, 147);
            this.chkHandbookDeletePictures.Name = "chkHandbookDeletePictures";
            this.chkHandbookDeletePictures.Size = new System.Drawing.Size(338, 24);
            this.chkHandbookDeletePictures.TabIndex = 7;
            this.chkHandbookDeletePictures.Text = "Delete slide pictures folder after finishing";
            this.chkHandbookDeletePictures.UseVisualStyleBackColor = true;
            // 
            // progressBarHandbook
            // 
            this.progressBarHandbook.Location = new System.Drawing.Point(20, 314);
            this.progressBarHandbook.Name = "progressBarHandbook";
            this.progressBarHandbook.Size = new System.Drawing.Size(527, 23);
            this.progressBarHandbook.TabIndex = 6;
            // 
            // lblHandbookStatus
            // 
            this.lblHandbookStatus.AutoSize = true;
            this.lblHandbookStatus.Location = new System.Drawing.Point(257, 294);
            this.lblHandbookStatus.Name = "lblHandbookStatus";
            this.lblHandbookStatus.Size = new System.Drawing.Size(53, 17);
            this.lblHandbookStatus.TabIndex = 5;
            this.lblHandbookStatus.Text = "Ready.";
            // 
            // lblHandbookFileTitle
            // 
            this.lblHandbookFileTitle.AutoSize = true;
            this.lblHandbookFileTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblHandbookFileTitle.Location = new System.Drawing.Point(16, 16);
            this.lblHandbookFileTitle.Name = "lblHandbookFileTitle";
            this.lblHandbookFileTitle.Size = new System.Drawing.Size(282, 20);
            this.lblHandbookFileTitle.TabIndex = 4;
            this.lblHandbookFileTitle.Text = "Presentation to convert to handbook:";
            // 
            // btnHandbookBrowseOpen
            // 
            this.btnHandbookBrowseOpen.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnHandbookBrowseOpen.Location = new System.Drawing.Point(472, 42);
            this.btnHandbookBrowseOpen.Name = "btnHandbookBrowseOpen";
            this.btnHandbookBrowseOpen.Size = new System.Drawing.Size(100, 35);
            this.btnHandbookBrowseOpen.TabIndex = 3;
            this.btnHandbookBrowseOpen.Text = "Browse";
            this.btnHandbookBrowseOpen.UseVisualStyleBackColor = true;
            this.btnHandbookBrowseOpen.Click += new System.EventHandler(this.btnHandbookBrowseOpen_Click);
            // 
            // txtHandbookFile
            // 
            this.txtHandbookFile.AllowDrop = true;
            this.txtHandbookFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtHandbookFile.Location = new System.Drawing.Point(19, 48);
            this.txtHandbookFile.Name = "txtHandbookFile";
            this.txtHandbookFile.Size = new System.Drawing.Size(428, 23);
            this.txtHandbookFile.TabIndex = 2;
            this.txtHandbookFile.TextChanged += new System.EventHandler(this.txtHandbookFile_TextChanged);
            this.txtHandbookFile.DragDrop += new System.Windows.Forms.DragEventHandler(this.txtHandbookFile_DragDrop);
            this.txtHandbookFile.DragEnter += new System.Windows.Forms.DragEventHandler(this.txtHandbookFile_DragEnter);
            // 
            // tabPage3
            // 
            this.tabPage3.Controls.Add(this.chkBatchHandbookGoalObjSamePage);
            this.tabPage3.Controls.Add(this.btnBatchHandbookClearList);
            this.tabPage3.Controls.Add(this.btnBatchHandbookRemove);
            this.tabPage3.Controls.Add(this.lstBatchHandbookFiles);
            this.tabPage3.Controls.Add(this.btnBatchHandbookAddToList);
            this.tabPage3.Controls.Add(this.chkBatchHandbookInstructorGuide);
            this.tabPage3.Controls.Add(this.btnBatchHandbookConvert);
            this.tabPage3.Controls.Add(this.progressBarBatchHandbook);
            this.tabPage3.Controls.Add(this.lblBatchHandbookStatus);
            this.tabPage3.Controls.Add(this.lblBatchHandbookTitle);
            this.tabPage3.Controls.Add(this.btnBatchHandbookBrowseOpen);
            this.tabPage3.Controls.Add(this.txtBatchHandbookFile);
            this.tabPage3.Location = new System.Drawing.Point(4, 25);
            this.tabPage3.Name = "tabPage3";
            this.tabPage3.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage3.Size = new System.Drawing.Size(581, 340);
            this.tabPage3.TabIndex = 2;
            this.tabPage3.Text = "Batch Handbook creator";
            this.tabPage3.UseVisualStyleBackColor = true;
            // 
            // btnBatchHandbookClearList
            // 
            this.btnBatchHandbookClearList.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBatchHandbookClearList.Location = new System.Drawing.Point(472, 165);
            this.btnBatchHandbookClearList.Name = "btnBatchHandbookClearList";
            this.btnBatchHandbookClearList.Size = new System.Drawing.Size(100, 35);
            this.btnBatchHandbookClearList.TabIndex = 23;
            this.btnBatchHandbookClearList.Text = "Clear";
            this.btnBatchHandbookClearList.UseVisualStyleBackColor = true;
            this.btnBatchHandbookClearList.Click += new System.EventHandler(this.btnBatchHandbookClearList_Click);
            // 
            // btnBatchHandbookRemove
            // 
            this.btnBatchHandbookRemove.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBatchHandbookRemove.Location = new System.Drawing.Point(472, 124);
            this.btnBatchHandbookRemove.Name = "btnBatchHandbookRemove";
            this.btnBatchHandbookRemove.Size = new System.Drawing.Size(100, 35);
            this.btnBatchHandbookRemove.TabIndex = 22;
            this.btnBatchHandbookRemove.Text = "Remove";
            this.btnBatchHandbookRemove.UseVisualStyleBackColor = true;
            this.btnBatchHandbookRemove.Click += new System.EventHandler(this.btnBatchHandbookRemove_Click);
            // 
            // lstBatchHandbookFiles
            // 
            this.lstBatchHandbookFiles.AllowDrop = true;
            this.lstBatchHandbookFiles.FormattingEnabled = true;
            this.lstBatchHandbookFiles.HorizontalScrollbar = true;
            this.lstBatchHandbookFiles.ItemHeight = 16;
            this.lstBatchHandbookFiles.Location = new System.Drawing.Point(20, 83);
            this.lstBatchHandbookFiles.Name = "lstBatchHandbookFiles";
            this.lstBatchHandbookFiles.Size = new System.Drawing.Size(427, 116);
            this.lstBatchHandbookFiles.TabIndex = 21;
            this.lstBatchHandbookFiles.DragDrop += new System.Windows.Forms.DragEventHandler(this.lstBatchHandbookFiles_DragDrop);
            this.lstBatchHandbookFiles.DragEnter += new System.Windows.Forms.DragEventHandler(this.lstBatchHandbookFiles_DragEnter);
            // 
            // btnBatchHandbookAddToList
            // 
            this.btnBatchHandbookAddToList.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBatchHandbookAddToList.Location = new System.Drawing.Point(472, 83);
            this.btnBatchHandbookAddToList.Name = "btnBatchHandbookAddToList";
            this.btnBatchHandbookAddToList.Size = new System.Drawing.Size(100, 35);
            this.btnBatchHandbookAddToList.TabIndex = 20;
            this.btnBatchHandbookAddToList.Text = "Add to List";
            this.btnBatchHandbookAddToList.UseVisualStyleBackColor = true;
            this.btnBatchHandbookAddToList.Click += new System.EventHandler(this.btnBatchHandbookAddToList_Click);
            // 
            // chkBatchHandbookInstructorGuide
            // 
            this.chkBatchHandbookInstructorGuide.AutoSize = true;
            this.chkBatchHandbookInstructorGuide.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkBatchHandbookInstructorGuide.Location = new System.Drawing.Point(19, 205);
            this.chkBatchHandbookInstructorGuide.Name = "chkBatchHandbookInstructorGuide";
            this.chkBatchHandbookInstructorGuide.Size = new System.Drawing.Size(211, 24);
            this.chkBatchHandbookInstructorGuide.TabIndex = 19;
            this.chkBatchHandbookInstructorGuide.Text = "Create Instructor guides";
            this.chkBatchHandbookInstructorGuide.UseVisualStyleBackColor = true;
            // 
            // btnBatchHandbookConvert
            // 
            this.btnBatchHandbookConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBatchHandbookConvert.Location = new System.Drawing.Point(233, 253);
            this.btnBatchHandbookConvert.Name = "btnBatchHandbookConvert";
            this.btnBatchHandbookConvert.Size = new System.Drawing.Size(100, 35);
            this.btnBatchHandbookConvert.TabIndex = 18;
            this.btnBatchHandbookConvert.Text = "Convert";
            this.btnBatchHandbookConvert.UseVisualStyleBackColor = true;
            this.btnBatchHandbookConvert.Click += new System.EventHandler(this.btnBatchHandbookConvert_Click);
            // 
            // progressBarBatchHandbook
            // 
            this.progressBarBatchHandbook.Location = new System.Drawing.Point(20, 314);
            this.progressBarBatchHandbook.Name = "progressBarBatchHandbook";
            this.progressBarBatchHandbook.Size = new System.Drawing.Size(527, 23);
            this.progressBarBatchHandbook.TabIndex = 17;
            // 
            // lblBatchHandbookStatus
            // 
            this.lblBatchHandbookStatus.AutoSize = true;
            this.lblBatchHandbookStatus.Location = new System.Drawing.Point(257, 294);
            this.lblBatchHandbookStatus.Name = "lblBatchHandbookStatus";
            this.lblBatchHandbookStatus.Size = new System.Drawing.Size(53, 17);
            this.lblBatchHandbookStatus.TabIndex = 16;
            this.lblBatchHandbookStatus.Text = "Ready.";
            // 
            // lblBatchHandbookTitle
            // 
            this.lblBatchHandbookTitle.AutoSize = true;
            this.lblBatchHandbookTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblBatchHandbookTitle.Location = new System.Drawing.Point(16, 16);
            this.lblBatchHandbookTitle.Name = "lblBatchHandbookTitle";
            this.lblBatchHandbookTitle.Size = new System.Drawing.Size(337, 20);
            this.lblBatchHandbookTitle.TabIndex = 15;
            this.lblBatchHandbookTitle.Text = "Presentation to add to batch conversion list:";
            // 
            // btnBatchHandbookBrowseOpen
            // 
            this.btnBatchHandbookBrowseOpen.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnBatchHandbookBrowseOpen.Location = new System.Drawing.Point(472, 42);
            this.btnBatchHandbookBrowseOpen.Name = "btnBatchHandbookBrowseOpen";
            this.btnBatchHandbookBrowseOpen.Size = new System.Drawing.Size(100, 35);
            this.btnBatchHandbookBrowseOpen.TabIndex = 14;
            this.btnBatchHandbookBrowseOpen.Text = "Browse";
            this.btnBatchHandbookBrowseOpen.UseVisualStyleBackColor = true;
            // 
            // txtBatchHandbookFile
            // 
            this.txtBatchHandbookFile.AllowDrop = true;
            this.txtBatchHandbookFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtBatchHandbookFile.Location = new System.Drawing.Point(19, 48);
            this.txtBatchHandbookFile.Name = "txtBatchHandbookFile";
            this.txtBatchHandbookFile.Size = new System.Drawing.Size(428, 23);
            this.txtBatchHandbookFile.TabIndex = 13;
            this.txtBatchHandbookFile.DragDrop += new System.Windows.Forms.DragEventHandler(this.txtBatchHandbookFile_DragDrop);
            this.txtBatchHandbookFile.DragEnter += new System.Windows.Forms.DragEventHandler(this.txtBatchHandbookFile_DragEnter);
            // 
            // tabPage1
            // 
            this.tabPage1.Controls.Add(this.chkLPSeparate);
            this.tabPage1.Controls.Add(this.rb169);
            this.tabPage1.Controls.Add(this.rb43);
            this.tabPage1.Controls.Add(this.chkLPDeleteXML);
            this.tabPage1.Controls.Add(this.chkLPOverrideDefaultSave);
            this.tabPage1.Controls.Add(this.btnLPBrowseSave);
            this.tabPage1.Controls.Add(this.txtLPSaveLocation);
            this.tabPage1.Controls.Add(this.lblLPFileTitle);
            this.tabPage1.Controls.Add(this.btnLPBrowseOpen);
            this.tabPage1.Controls.Add(this.txtLPFile);
            this.tabPage1.Controls.Add(this.progressBarLP);
            this.tabPage1.Controls.Add(this.lblLPStatus);
            this.tabPage1.Controls.Add(this.btnLPConvert);
            this.tabPage1.Location = new System.Drawing.Point(4, 25);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(581, 340);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Presentation template creator";
            this.tabPage1.UseVisualStyleBackColor = true;
            // 
            // chkLPSeparate
            // 
            this.chkLPSeparate.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkLPSeparate.Location = new System.Drawing.Point(394, 150);
            this.chkLPSeparate.Name = "chkLPSeparate";
            this.chkLPSeparate.Size = new System.Drawing.Size(153, 64);
            this.chkLPSeparate.TabIndex = 25;
            this.chkLPSeparate.Text = "Separate presentation by subjects";
            this.chkLPSeparate.UseVisualStyleBackColor = true;
            // 
            // rb169
            // 
            this.rb169.AutoSize = true;
            this.rb169.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rb169.Location = new System.Drawing.Point(19, 180);
            this.rb169.Name = "rb169";
            this.rb169.Size = new System.Drawing.Size(131, 24);
            this.rb169.TabIndex = 24;
            this.rb169.Text = "16:9 template";
            this.rb169.UseVisualStyleBackColor = true;
            // 
            // rb43
            // 
            this.rb43.AutoSize = true;
            this.rb43.Checked = true;
            this.rb43.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rb43.Location = new System.Drawing.Point(19, 150);
            this.rb43.Name = "rb43";
            this.rb43.Size = new System.Drawing.Size(122, 24);
            this.rb43.TabIndex = 23;
            this.rb43.TabStop = true;
            this.rb43.Text = "4:3 template";
            this.rb43.UseVisualStyleBackColor = true;
            this.rb43.CheckedChanged += new System.EventHandler(this.radioButton1_CheckedChanged);
            // 
            // chkLPDeleteXML
            // 
            this.chkLPDeleteXML.AutoSize = true;
            this.chkLPDeleteXML.Checked = true;
            this.chkLPDeleteXML.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkLPDeleteXML.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkLPDeleteXML.Location = new System.Drawing.Point(19, 219);
            this.chkLPDeleteXML.Name = "chkLPDeleteXML";
            this.chkLPDeleteXML.Size = new System.Drawing.Size(302, 24);
            this.chkLPDeleteXML.TabIndex = 21;
            this.chkLPDeleteXML.Text = "Delete temporary files after finishing";
            this.chkLPDeleteXML.UseVisualStyleBackColor = true;
            // 
            // chkLPOverrideDefaultSave
            // 
            this.chkLPOverrideDefaultSave.AutoSize = true;
            this.chkLPOverrideDefaultSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkLPOverrideDefaultSave.Location = new System.Drawing.Point(20, 84);
            this.chkLPOverrideDefaultSave.Name = "chkLPOverrideDefaultSave";
            this.chkLPOverrideDefaultSave.Size = new System.Drawing.Size(331, 24);
            this.chkLPOverrideDefaultSave.TabIndex = 20;
            this.chkLPOverrideDefaultSave.Text = "Override default save location and name";
            this.chkLPOverrideDefaultSave.UseVisualStyleBackColor = true;
            this.chkLPOverrideDefaultSave.CheckedChanged += new System.EventHandler(this.chkLPOverrideDefaultSave_CheckedChanged);
            // 
            // btnLPBrowseSave
            // 
            this.btnLPBrowseSave.Enabled = false;
            this.btnLPBrowseSave.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLPBrowseSave.Location = new System.Drawing.Point(472, 105);
            this.btnLPBrowseSave.Name = "btnLPBrowseSave";
            this.btnLPBrowseSave.Size = new System.Drawing.Size(100, 35);
            this.btnLPBrowseSave.TabIndex = 19;
            this.btnLPBrowseSave.Text = "Browse";
            this.btnLPBrowseSave.UseVisualStyleBackColor = true;
            this.btnLPBrowseSave.Click += new System.EventHandler(this.btnLPBrowseSave_Click);
            // 
            // txtLPSaveLocation
            // 
            this.txtLPSaveLocation.Enabled = false;
            this.txtLPSaveLocation.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLPSaveLocation.Location = new System.Drawing.Point(20, 111);
            this.txtLPSaveLocation.Name = "txtLPSaveLocation";
            this.txtLPSaveLocation.Size = new System.Drawing.Size(427, 23);
            this.txtLPSaveLocation.TabIndex = 18;
            // 
            // lblLPFileTitle
            // 
            this.lblLPFileTitle.AutoSize = true;
            this.lblLPFileTitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblLPFileTitle.Location = new System.Drawing.Point(16, 16);
            this.lblLPFileTitle.Name = "lblLPFileTitle";
            this.lblLPFileTitle.Size = new System.Drawing.Size(369, 20);
            this.lblLPFileTitle.TabIndex = 17;
            this.lblLPFileTitle.Text = "Lesson plan to convert to presentation template:";
            // 
            // btnLPBrowseOpen
            // 
            this.btnLPBrowseOpen.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLPBrowseOpen.Location = new System.Drawing.Point(472, 42);
            this.btnLPBrowseOpen.Name = "btnLPBrowseOpen";
            this.btnLPBrowseOpen.Size = new System.Drawing.Size(100, 35);
            this.btnLPBrowseOpen.TabIndex = 16;
            this.btnLPBrowseOpen.Text = "Browse";
            this.btnLPBrowseOpen.UseVisualStyleBackColor = true;
            this.btnLPBrowseOpen.Click += new System.EventHandler(this.btnLPBrowseOpen_Click);
            // 
            // txtLPFile
            // 
            this.txtLPFile.AllowDrop = true;
            this.txtLPFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLPFile.Location = new System.Drawing.Point(19, 48);
            this.txtLPFile.Name = "txtLPFile";
            this.txtLPFile.Size = new System.Drawing.Size(428, 23);
            this.txtLPFile.TabIndex = 15;
            this.txtLPFile.TextChanged += new System.EventHandler(this.txtLPFile_TextChanged);
            this.txtLPFile.DragDrop += new System.Windows.Forms.DragEventHandler(this.txtLPFile_DragDrop);
            this.txtLPFile.DragEnter += new System.Windows.Forms.DragEventHandler(this.txtLPFile_DragEnter);
            // 
            // progressBarLP
            // 
            this.progressBarLP.Location = new System.Drawing.Point(20, 314);
            this.progressBarLP.Name = "progressBarLP";
            this.progressBarLP.Size = new System.Drawing.Size(527, 23);
            this.progressBarLP.TabIndex = 14;
            // 
            // lblLPStatus
            // 
            this.lblLPStatus.AutoSize = true;
            this.lblLPStatus.Location = new System.Drawing.Point(257, 294);
            this.lblLPStatus.Name = "lblLPStatus";
            this.lblLPStatus.Size = new System.Drawing.Size(53, 17);
            this.lblLPStatus.TabIndex = 13;
            this.lblLPStatus.Text = "Ready.";
            // 
            // btnLPConvert
            // 
            this.btnLPConvert.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnLPConvert.Location = new System.Drawing.Point(233, 249);
            this.btnLPConvert.Name = "btnLPConvert";
            this.btnLPConvert.Size = new System.Drawing.Size(100, 35);
            this.btnLPConvert.TabIndex = 12;
            this.btnLPConvert.Text = "Convert";
            this.btnLPConvert.UseVisualStyleBackColor = true;
            this.btnLPConvert.Click += new System.EventHandler(this.btnLPConvert_Click);
            // 
            // btnExit
            // 
            this.btnExit.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExit.Location = new System.Drawing.Point(513, 486);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 35);
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "Exit";
            this.btnExit.UseVisualStyleBackColor = true;
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(-1, 512);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(47, 17);
            this.label1.TabIndex = 3;
            this.label1.Text = "v1.3.1";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::TrainingDocumentation.Properties.Resources.LogoTraining;
            this.pictureBox1.Location = new System.Drawing.Point(16, 8);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(581, 97);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            this.pictureBox1.Click += new System.EventHandler(this.pictureBox1_Click);
            // 
            // chkGoalObjSamePage
            // 
            this.chkGoalObjSamePage.AutoSize = true;
            this.chkGoalObjSamePage.Checked = true;
            this.chkGoalObjSamePage.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkGoalObjSamePage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkGoalObjSamePage.Location = new System.Drawing.Point(20, 207);
            this.chkGoalObjSamePage.Name = "chkGoalObjSamePage";
            this.chkGoalObjSamePage.Size = new System.Drawing.Size(381, 24);
            this.chkGoalObjSamePage.TabIndex = 13;
            this.chkGoalObjSamePage.Text = "Lesson goals and objectives on the same page";
            this.chkGoalObjSamePage.UseVisualStyleBackColor = true;
            // 
            // chkBatchHandbookGoalObjSamePage
            // 
            this.chkBatchHandbookGoalObjSamePage.AutoSize = true;
            this.chkBatchHandbookGoalObjSamePage.Checked = true;
            this.chkBatchHandbookGoalObjSamePage.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkBatchHandbookGoalObjSamePage.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkBatchHandbookGoalObjSamePage.Location = new System.Drawing.Point(19, 228);
            this.chkBatchHandbookGoalObjSamePage.Name = "chkBatchHandbookGoalObjSamePage";
            this.chkBatchHandbookGoalObjSamePage.Size = new System.Drawing.Size(381, 24);
            this.chkBatchHandbookGoalObjSamePage.TabIndex = 24;
            this.chkBatchHandbookGoalObjSamePage.Text = "Lesson goals and objectives on the same page";
            this.chkBatchHandbookGoalObjSamePage.UseVisualStyleBackColor = true;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(614, 528);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.tabContainer);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Training documentation";
            this.tabContainer.ResumeLayout(false);
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.tabPage3.ResumeLayout(false);
            this.tabPage3.PerformLayout();
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TabControl tabContainer;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Button btnExit;
        private System.Windows.Forms.CheckBox chkHandbookOverrideDefaultSave;
        private System.Windows.Forms.Button btnHandbookBrowseSave;
        private System.Windows.Forms.TextBox txtHandbookSaveLocation;
        private System.Windows.Forms.CheckBox chkHandbookDeletePictures;
        private System.Windows.Forms.ProgressBar progressBarHandbook;
        private System.Windows.Forms.Label lblHandbookStatus;
        private System.Windows.Forms.Label lblHandbookFileTitle;
        private System.Windows.Forms.Button btnHandbookBrowseOpen;
        private System.Windows.Forms.TextBox txtHandbookFile;
        private System.ComponentModel.BackgroundWorker bckgHandbook;
        private System.Windows.Forms.Button btnHandbookConvert;
        private System.Windows.Forms.Button btnLPConvert;
        private System.Windows.Forms.ProgressBar progressBarLP;
        private System.Windows.Forms.Label lblLPStatus;
        private System.ComponentModel.BackgroundWorker bckgLP;
        private System.Windows.Forms.CheckBox chkLPOverrideDefaultSave;
        private System.Windows.Forms.Button btnLPBrowseSave;
        private System.Windows.Forms.TextBox txtLPSaveLocation;
        private System.Windows.Forms.Label lblLPFileTitle;
        private System.Windows.Forms.Button btnLPBrowseOpen;
        private System.Windows.Forms.TextBox txtLPFile;
        private System.Windows.Forms.CheckBox chkLPDeleteXML;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox chkHandbookInstructorGuide;
        private System.Windows.Forms.RadioButton rb43;
        private System.Windows.Forms.RadioButton rb169;
        private System.Windows.Forms.TabPage tabPage3;
        private System.Windows.Forms.Button btnBatchHandbookClearList;
        private System.Windows.Forms.Button btnBatchHandbookRemove;
        private System.Windows.Forms.ListBox lstBatchHandbookFiles;
        private System.Windows.Forms.Button btnBatchHandbookAddToList;
        private System.Windows.Forms.CheckBox chkBatchHandbookInstructorGuide;
        private System.Windows.Forms.Button btnBatchHandbookConvert;
        private System.Windows.Forms.ProgressBar progressBarBatchHandbook;
        private System.Windows.Forms.Label lblBatchHandbookStatus;
        private System.Windows.Forms.Label lblBatchHandbookTitle;
        private System.Windows.Forms.Button btnBatchHandbookBrowseOpen;
        private System.Windows.Forms.TextBox txtBatchHandbookFile;
        private System.ComponentModel.BackgroundWorker bckgBatch;
        private System.Windows.Forms.CheckBox chkLPSeparate;
        private System.Windows.Forms.CheckBox chkGoalObjSamePage;
        private System.Windows.Forms.CheckBox chkBatchHandbookGoalObjSamePage;
    }
}

