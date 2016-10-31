namespace TFSTestCaseExportToExcel
{
    partial class FrmExport
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmExport));
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.comBoxTestSuite = new System.Windows.Forms.ComboBox();
            this.lblTestSuite = new System.Windows.Forms.Label();
            this.comBoxTestPlan = new System.Windows.Forms.ComboBox();
            this.lblTestPlan = new System.Windows.Forms.Label();
            this.btnTeamProject = new System.Windows.Forms.Button();
            this.txtTeamProject = new System.Windows.Forms.TextBox();
            this.lblTeamProject = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.txtFileName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnFolderBrowse = new System.Windows.Forms.Button();
            this.txtSaveFolder = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.folderBrowserDialog = new System.Windows.Forms.FolderBrowserDialog();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.grpResultFilter = new System.Windows.Forms.GroupBox();
            this.cmbRunType = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.grpResultFilter.SuspendLayout();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.comBoxTestSuite);
            this.groupBox1.Controls.Add(this.lblTestSuite);
            this.groupBox1.Controls.Add(this.comBoxTestPlan);
            this.groupBox1.Controls.Add(this.lblTestPlan);
            this.groupBox1.Controls.Add(this.btnTeamProject);
            this.groupBox1.Controls.Add(this.txtTeamProject);
            this.groupBox1.Controls.Add(this.lblTeamProject);
            this.groupBox1.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(12, 190);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(586, 199);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Source";
            // 
            // comBoxTestSuite
            // 
            this.comBoxTestSuite.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBoxTestSuite.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comBoxTestSuite.FormattingEnabled = true;
            this.comBoxTestSuite.Location = new System.Drawing.Point(11, 158);
            this.comBoxTestSuite.Name = "comBoxTestSuite";
            this.comBoxTestSuite.Size = new System.Drawing.Size(452, 23);
            this.comBoxTestSuite.TabIndex = 3;
            this.comBoxTestSuite.SelectedIndexChanged += new System.EventHandler(this.comBoxTestSuite_SelectedIndexChanged);
            // 
            // lblTestSuite
            // 
            this.lblTestSuite.AutoSize = true;
            this.lblTestSuite.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTestSuite.Location = new System.Drawing.Point(6, 138);
            this.lblTestSuite.Name = "lblTestSuite";
            this.lblTestSuite.Size = new System.Drawing.Size(66, 17);
            this.lblTestSuite.TabIndex = 0;
            this.lblTestSuite.Text = "Test Suite:";
            // 
            // comBoxTestPlan
            // 
            this.comBoxTestPlan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comBoxTestPlan.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comBoxTestPlan.FormattingEnabled = true;
            this.comBoxTestPlan.Location = new System.Drawing.Point(11, 99);
            this.comBoxTestPlan.Name = "comBoxTestPlan";
            this.comBoxTestPlan.Size = new System.Drawing.Size(452, 23);
            this.comBoxTestPlan.TabIndex = 2;
            this.comBoxTestPlan.SelectedIndexChanged += new System.EventHandler(this.comBoxTestPlan_SelectedIndexChanged);
            // 
            // lblTestPlan
            // 
            this.lblTestPlan.AutoSize = true;
            this.lblTestPlan.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTestPlan.Location = new System.Drawing.Point(6, 79);
            this.lblTestPlan.Name = "lblTestPlan";
            this.lblTestPlan.Size = new System.Drawing.Size(62, 17);
            this.lblTestPlan.TabIndex = 0;
            this.lblTestPlan.Text = "Test Plan:";
            // 
            // btnTeamProject
            // 
            this.btnTeamProject.Location = new System.Drawing.Point(475, 44);
            this.btnTeamProject.Name = "btnTeamProject";
            this.btnTeamProject.Size = new System.Drawing.Size(35, 23);
            this.btnTeamProject.TabIndex = 1;
            this.btnTeamProject.Text = "...";
            this.btnTeamProject.UseVisualStyleBackColor = true;
            this.btnTeamProject.Click += new System.EventHandler(this.btnTeamProject_Click);
            // 
            // txtTeamProject
            // 
            this.txtTeamProject.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtTeamProject.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtTeamProject.ForeColor = System.Drawing.SystemColors.InactiveCaption;
            this.txtTeamProject.Location = new System.Drawing.Point(11, 43);
            this.txtTeamProject.Name = "txtTeamProject";
            this.txtTeamProject.ReadOnly = true;
            this.txtTeamProject.Size = new System.Drawing.Size(453, 24);
            this.txtTeamProject.TabIndex = 0;
            this.txtTeamProject.Text = "Select TFS Server Location";
            // 
            // lblTeamProject
            // 
            this.lblTeamProject.AutoSize = true;
            this.lblTeamProject.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblTeamProject.Location = new System.Drawing.Point(6, 21);
            this.lblTeamProject.Name = "lblTeamProject";
            this.lblTeamProject.Size = new System.Drawing.Size(87, 17);
            this.lblTeamProject.TabIndex = 0;
            this.lblTeamProject.Text = "Team-Project:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.txtFileName);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.btnFolderBrowse);
            this.groupBox2.Controls.Add(this.txtSaveFolder);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(12, 533);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(586, 142);
            this.groupBox2.TabIndex = 0;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Destination";
            // 
            // txtFileName
            // 
            this.txtFileName.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtFileName.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtFileName.Location = new System.Drawing.Point(11, 95);
            this.txtFileName.Name = "txtFileName";
            this.txtFileName.Size = new System.Drawing.Size(411, 24);
            this.txtFileName.TabIndex = 5;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(8, 75);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(156, 17);
            this.label2.TabIndex = 0;
            this.label2.Text = "Specify Name of Excel File:";
            // 
            // btnFolderBrowse
            // 
            this.btnFolderBrowse.Font = new System.Drawing.Font("Calibri", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnFolderBrowse.Location = new System.Drawing.Point(428, 42);
            this.btnFolderBrowse.Name = "btnFolderBrowse";
            this.btnFolderBrowse.Size = new System.Drawing.Size(82, 28);
            this.btnFolderBrowse.TabIndex = 4;
            this.btnFolderBrowse.Text = "Browse...";
            this.btnFolderBrowse.UseVisualStyleBackColor = true;
            this.btnFolderBrowse.Click += new System.EventHandler(this.btnFolderBrowse_Click);
            // 
            // txtSaveFolder
            // 
            this.txtSaveFolder.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.txtSaveFolder.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.txtSaveFolder.ForeColor = System.Drawing.SystemColors.InactiveCaption;
            this.txtSaveFolder.Location = new System.Drawing.Point(11, 42);
            this.txtSaveFolder.Name = "txtSaveFolder";
            this.txtSaveFolder.ReadOnly = true;
            this.txtSaveFolder.Size = new System.Drawing.Size(411, 24);
            this.txtSaveFolder.TabIndex = 7;
            this.txtSaveFolder.Text = "C:\\Users\\Psrivastav\\Downloads";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(6, 22);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(235, 17);
            this.label1.TabIndex = 0;
            this.label1.Text = "Specify Location to Save Exported Script:";
            // 
            // btnCancel
            // 
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(403, 698);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(83, 25);
            this.btnCancel.TabIndex = 7;
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnExport
            // 
            this.btnExport.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnExport.Location = new System.Drawing.Point(147, 698);
            this.btnExport.Name = "btnExport";
            this.btnExport.Size = new System.Drawing.Size(83, 25);
            this.btnExport.TabIndex = 6;
            this.btnExport.Text = "Export";
            this.btnExport.UseVisualStyleBackColor = true;
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label6);
            this.groupBox3.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(12, 12);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(586, 172);
            this.groupBox3.TabIndex = 0;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Welcome";
            // 
            // label6
            // 
            this.label6.AllowDrop = true;
            this.label6.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(10, 21);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(559, 143);
            this.label6.TabIndex = 0;
            this.label6.Text = resources.GetString("label6.Text");
            // 
            // grpResultFilter
            // 
            this.grpResultFilter.Controls.Add(this.cmbRunType);
            this.grpResultFilter.Controls.Add(this.label3);
            this.grpResultFilter.Enabled = false;
            this.grpResultFilter.Font = new System.Drawing.Font("Calibri", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.grpResultFilter.Location = new System.Drawing.Point(12, 395);
            this.grpResultFilter.Name = "grpResultFilter";
            this.grpResultFilter.Size = new System.Drawing.Size(586, 120);
            this.grpResultFilter.TabIndex = 9;
            this.grpResultFilter.TabStop = false;
            this.grpResultFilter.Text = "Filter Results";
            // 
            // cmbRunType
            // 
            this.cmbRunType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cmbRunType.FormattingEnabled = true;
            this.cmbRunType.Items.AddRange(new object[] {
            "All",
            "Automated",
            "Manual"});
            this.cmbRunType.Location = new System.Drawing.Point(13, 51);
            this.cmbRunType.Name = "cmbRunType";
            this.cmbRunType.Size = new System.Drawing.Size(452, 26);
            this.cmbRunType.TabIndex = 9;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri", 10F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(9, 29);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 17);
            this.label3.TabIndex = 8;
            this.label3.Text = "Run Type:";
            // 
            // FrmExport
            // 
            this.AcceptButton = this.btnExport;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(609, 747);
            this.Controls.Add(this.grpResultFilter);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.btnExport);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Font = new System.Drawing.Font("Calibri", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.Name = "FrmExport";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "TFS Test Case Export to Excel";
            this.Load += new System.EventHandler(this.FrmExport_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.grpResultFilter.ResumeLayout(false);
            this.grpResultFilter.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.ComboBox comBoxTestSuite;
        private System.Windows.Forms.Label lblTestSuite;
        private System.Windows.Forms.ComboBox comBoxTestPlan;
        private System.Windows.Forms.Label lblTestPlan;
        private System.Windows.Forms.Button btnTeamProject;
        private System.Windows.Forms.TextBox txtTeamProject;
        private System.Windows.Forms.Label lblTeamProject;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.TextBox txtFileName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button btnFolderBrowse;
        private System.Windows.Forms.TextBox txtSaveFolder;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.GroupBox grpResultFilter;
        private System.Windows.Forms.ComboBox cmbRunType;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
    }
}

