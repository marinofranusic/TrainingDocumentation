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
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.btnExit = new System.Windows.Forms.Button();
            this.btnHandbookBrowseOpen = new System.Windows.Forms.Button();
            this.txtHandbookFile = new System.Windows.Forms.TextBox();
            this.lblHandbookFileTitle = new System.Windows.Forms.Label();
            this.lblHandbookStatus = new System.Windows.Forms.Label();
            this.progressBarHandbook = new System.Windows.Forms.ProgressBar();
            this.chkHandbookDeletePictures = new System.Windows.Forms.CheckBox();
            this.txtHandbookSaveLocation = new System.Windows.Forms.TextBox();
            this.btnHandbookBrowseSave = new System.Windows.Forms.Button();
            this.chkHandbookOverrideDefaultSave = new System.Windows.Forms.CheckBox();
            this.bckgHandbook = new System.ComponentModel.BackgroundWorker();
            this.btnHandbookConvert = new System.Windows.Forms.Button();
            this.btnLPConvert = new System.Windows.Forms.Button();
            this.progressBarLP = new System.Windows.Forms.ProgressBar();
            this.lblLPStatus = new System.Windows.Forms.Label();
            this.bckgLP = new System.ComponentModel.BackgroundWorker();
            this.chkLPOverrideDefaultSave = new System.Windows.Forms.CheckBox();
            this.btnLPBrowseSave = new System.Windows.Forms.Button();
            this.txtLPSaveLocation = new System.Windows.Forms.TextBox();
            this.lblLPFileTitle = new System.Windows.Forms.Label();
            this.btnLPBrowseOpen = new System.Windows.Forms.Button();
            this.txtLPFile = new System.Windows.Forms.TextBox();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.chkLPDeleteXML = new System.Windows.Forms.CheckBox();
            this.tabContainer.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // tabContainer
            // 
            this.tabContainer.Controls.Add(this.tabPage1);
            this.tabContainer.Controls.Add(this.tabPage2);
            this.tabContainer.Location = new System.Drawing.Point(12, 111);
            this.tabContainer.Name = "tabContainer";
            this.tabContainer.SelectedIndex = 0;
            this.tabContainer.Size = new System.Drawing.Size(589, 369);
            this.tabContainer.TabIndex = 0;
            // 
            // tabPage1
            // 
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
            // tabPage2
            // 
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
            this.txtHandbookFile.Enabled = false;
            this.txtHandbookFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtHandbookFile.Location = new System.Drawing.Point(19, 48);
            this.txtHandbookFile.Name = "txtHandbookFile";
            this.txtHandbookFile.Size = new System.Drawing.Size(428, 23);
            this.txtHandbookFile.TabIndex = 2;
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
            // lblHandbookStatus
            // 
            this.lblHandbookStatus.AutoSize = true;
            this.lblHandbookStatus.Location = new System.Drawing.Point(257, 294);
            this.lblHandbookStatus.Name = "lblHandbookStatus";
            this.lblHandbookStatus.Size = new System.Drawing.Size(53, 17);
            this.lblHandbookStatus.TabIndex = 5;
            this.lblHandbookStatus.Text = "Ready.";
            // 
            // progressBarHandbook
            // 
            this.progressBarHandbook.Location = new System.Drawing.Point(20, 314);
            this.progressBarHandbook.Name = "progressBarHandbook";
            this.progressBarHandbook.Size = new System.Drawing.Size(527, 23);
            this.progressBarHandbook.TabIndex = 6;
            // 
            // chkHandbookDeletePictures
            // 
            this.chkHandbookDeletePictures.AutoSize = true;
            this.chkHandbookDeletePictures.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkHandbookDeletePictures.Location = new System.Drawing.Point(20, 180);
            this.chkHandbookDeletePictures.Name = "chkHandbookDeletePictures";
            this.chkHandbookDeletePictures.Size = new System.Drawing.Size(338, 24);
            this.chkHandbookDeletePictures.TabIndex = 7;
            this.chkHandbookDeletePictures.Text = "Delete slide pictures folder after finishing";
            this.chkHandbookDeletePictures.UseVisualStyleBackColor = true;
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
            this.txtLPFile.Enabled = false;
            this.txtLPFile.Font = new System.Drawing.Font("Microsoft Sans Serif", 8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtLPFile.Location = new System.Drawing.Point(19, 48);
            this.txtLPFile.Name = "txtLPFile";
            this.txtLPFile.Size = new System.Drawing.Size(428, 23);
            this.txtLPFile.TabIndex = 15;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::TrainingDocumentation.Properties.Resources.Logo;
            this.pictureBox1.Location = new System.Drawing.Point(16, 8);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(581, 97);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 1;
            this.pictureBox1.TabStop = false;
            // 
            // chkLPDeleteXML
            // 
            this.chkLPDeleteXML.AutoSize = true;
            this.chkLPDeleteXML.Font = new System.Drawing.Font("Microsoft Sans Serif", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.chkLPDeleteXML.Location = new System.Drawing.Point(20, 180);
            this.chkLPDeleteXML.Name = "chkLPDeleteXML";
            this.chkLPDeleteXML.Size = new System.Drawing.Size(302, 24);
            this.chkLPDeleteXML.TabIndex = 21;
            this.chkLPDeleteXML.Text = "Delete temporary files after finishing";
            this.chkLPDeleteXML.UseVisualStyleBackColor = true;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ControlLightLight;
            this.ClientSize = new System.Drawing.Size(614, 528);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.tabContainer);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Training documentation";
            this.tabContainer.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);

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
    }
}

