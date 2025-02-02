﻿namespace GraphReports
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            tableLayoutPanel1 = new TableLayoutPanel();
            textBoxTenant = new TextBox();
            label1 = new Label();
            label3 = new Label();
            textBoxClientID = new TextBox();
            buttonExport = new Button();
            dataGridView1 = new DataGridView();
            progressBar1 = new ProgressBar();
            tabPage3 = new TabPage();
            buttonGetAllDevices = new Button();
            buttonNonComplaintDevices = new Button();
            groupBox1 = new GroupBox();
            textBoxDomainName = new TextBox();
            buttonGetDomainDependency = new Button();
            buttonGetDomains = new Button();
            buttonGetSubs = new Button();
            tabPage2 = new TabPage();
            groupBox2 = new GroupBox();
            textBoxGroupName = new TextBox();
            buttonGetGroupMembers = new Button();
            buttonGetLicensedGroups = new Button();
            buttonMailEnabledSec = new Button();
            buttonGetAllSec = new Button();
            buttonDistributionGroups = new Button();
            buttonUnifiedGroups = new Button();
            buttonGetAllGroups = new Button();
            tabPage1 = new TabPage();
            buttonMFAReg = new Button();
            buttonGetAdmins = new Button();
            buttonGetUnlicensed = new Button();
            buttonGetGuests = new Button();
            buttonGetSynced = new Button();
            buttonGetAllUsers = new Button();
            tabControl1 = new TabControl();
            tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).BeginInit();
            tabPage3.SuspendLayout();
            groupBox1.SuspendLayout();
            tabPage2.SuspendLayout();
            groupBox2.SuspendLayout();
            tabPage1.SuspendLayout();
            tabControl1.SuspendLayout();
            SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            tableLayoutPanel1.ColumnCount = 2;
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 19.7329369F));
            tableLayoutPanel1.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 80.26706F));
            tableLayoutPanel1.Controls.Add(textBoxTenant, 1, 2);
            tableLayoutPanel1.Controls.Add(label1, 0, 0);
            tableLayoutPanel1.Controls.Add(label3, 0, 2);
            tableLayoutPanel1.Controls.Add(textBoxClientID, 1, 0);
            tableLayoutPanel1.Location = new Point(11, 9);
            tableLayoutPanel1.Name = "tableLayoutPanel1";
            tableLayoutPanel1.RowCount = 3;
            tableLayoutPanel1.RowStyles.Add(new RowStyle());
            tableLayoutPanel1.RowStyles.Add(new RowStyle());
            tableLayoutPanel1.RowStyles.Add(new RowStyle());
            tableLayoutPanel1.Size = new Size(674, 88);
            tableLayoutPanel1.TabIndex = 0;
            // 
            // textBoxTenant
            // 
            textBoxTenant.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            textBoxTenant.Location = new Point(135, 32);
            textBoxTenant.Name = "textBoxTenant";
            textBoxTenant.Size = new Size(536, 23);
            textBoxTenant.TabIndex = 5;
            textBoxTenant.Text = "25412cef-d489-431e-87a3-8d6aa23d0853";
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(3, 0);
            label1.Name = "label1";
            label1.Size = new Size(52, 15);
            label1.TabIndex = 0;
            label1.Text = "Client ID";
            // 
            // label3
            // 
            label3.AutoSize = true;
            label3.Location = new Point(3, 29);
            label3.Name = "label3";
            label3.Size = new Size(103, 15);
            label3.TabIndex = 2;
            label3.Text = "Tenant ID/Domain";
            // 
            // textBoxClientID
            // 
            textBoxClientID.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            textBoxClientID.Location = new Point(135, 3);
            textBoxClientID.Name = "textBoxClientID";
            textBoxClientID.Size = new Size(536, 23);
            textBoxClientID.TabIndex = 3;
            textBoxClientID.Text = "6733563a-6624-404c-aaa1-5a860b0a721a";
            // 
            // buttonExport
            // 
            buttonExport.Location = new Point(561, 107);
            buttonExport.Name = "buttonExport";
            buttonExport.Size = new Size(124, 47);
            buttonExport.TabIndex = 2;
            buttonExport.Text = "Export to CSV";
            buttonExport.UseVisualStyleBackColor = true;
            buttonExport.Click += buttonExport_Click;
            // 
            // dataGridView1
            // 
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.AllowUserToDeleteRows = false;
            dataGridView1.AllowUserToOrderColumns = true;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView1.Location = new Point(14, 317);
            dataGridView1.Name = "dataGridView1";
            dataGridView1.ReadOnly = true;
            dataGridView1.SelectionMode = DataGridViewSelectionMode.CellSelect;
            dataGridView1.ShowEditingIcon = false;
            dataGridView1.Size = new Size(671, 398);
            dataGridView1.TabIndex = 3;
            // 
            // progressBar1
            // 
            progressBar1.Location = new Point(14, 287);
            progressBar1.Name = "progressBar1";
            progressBar1.Size = new Size(140, 24);
            progressBar1.TabIndex = 4;
            progressBar1.Visible = false;
            // 
            // tabPage3
            // 
            tabPage3.Controls.Add(buttonGetAllDevices);
            tabPage3.Controls.Add(buttonNonComplaintDevices);
            tabPage3.Controls.Add(groupBox1);
            tabPage3.Controls.Add(buttonGetDomains);
            tabPage3.Controls.Add(buttonGetSubs);
            tabPage3.Location = new Point(4, 24);
            tabPage3.Name = "tabPage3";
            tabPage3.Size = new Size(676, 122);
            tabPage3.TabIndex = 2;
            tabPage3.Text = "Other Reports";
            tabPage3.UseVisualStyleBackColor = true;
            // 
            // buttonGetAllDevices
            // 
            buttonGetAllDevices.Location = new Point(3, 61);
            buttonGetAllDevices.Name = "buttonGetAllDevices";
            buttonGetAllDevices.Size = new Size(128, 52);
            buttonGetAllDevices.TabIndex = 8;
            buttonGetAllDevices.Text = "Get All Devices";
            buttonGetAllDevices.UseVisualStyleBackColor = true;
            buttonGetAllDevices.Click += buttonGetAllDevices_Click;
            // 
            // buttonNonComplaintDevices
            // 
            buttonNonComplaintDevices.Location = new Point(137, 61);
            buttonNonComplaintDevices.Name = "buttonNonComplaintDevices";
            buttonNonComplaintDevices.Size = new Size(128, 52);
            buttonNonComplaintDevices.TabIndex = 7;
            buttonNonComplaintDevices.Text = "Get Non Complaint Devices";
            buttonNonComplaintDevices.UseVisualStyleBackColor = true;
            buttonNonComplaintDevices.Click += buttonNonComplaintDevices_Click;
            // 
            // groupBox1
            // 
            groupBox1.Controls.Add(textBoxDomainName);
            groupBox1.Controls.Add(buttonGetDomainDependency);
            groupBox1.Location = new Point(271, 3);
            groupBox1.Name = "groupBox1";
            groupBox1.Size = new Size(128, 116);
            groupBox1.TabIndex = 5;
            groupBox1.TabStop = false;
            groupBox1.Text = "Get Domain Dependency";
            // 
            // textBoxDomainName
            // 
            textBoxDomainName.Location = new Point(6, 41);
            textBoxDomainName.Name = "textBoxDomainName";
            textBoxDomainName.Size = new Size(116, 23);
            textBoxDomainName.TabIndex = 5;
            // 
            // buttonGetDomainDependency
            // 
            buttonGetDomainDependency.Location = new Point(3, 70);
            buttonGetDomainDependency.Name = "buttonGetDomainDependency";
            buttonGetDomainDependency.Size = new Size(119, 40);
            buttonGetDomainDependency.TabIndex = 4;
            buttonGetDomainDependency.Text = "Get Dependency";
            buttonGetDomainDependency.UseVisualStyleBackColor = true;
            buttonGetDomainDependency.Click += buttonGetDomainDependency_Click;
            // 
            // buttonGetDomains
            // 
            buttonGetDomains.Location = new Point(137, 3);
            buttonGetDomains.Name = "buttonGetDomains";
            buttonGetDomains.Size = new Size(128, 52);
            buttonGetDomains.TabIndex = 3;
            buttonGetDomains.Text = "Get All Domains";
            buttonGetDomains.UseVisualStyleBackColor = true;
            buttonGetDomains.Click += buttonGetDomains_Click;
            // 
            // buttonGetSubs
            // 
            buttonGetSubs.Location = new Point(3, 3);
            buttonGetSubs.Name = "buttonGetSubs";
            buttonGetSubs.Size = new Size(128, 52);
            buttonGetSubs.TabIndex = 2;
            buttonGetSubs.Text = "Get All Subsciption";
            buttonGetSubs.UseVisualStyleBackColor = true;
            buttonGetSubs.Click += buttonGetSubs_Click;
            // 
            // tabPage2
            // 
            tabPage2.Controls.Add(groupBox2);
            tabPage2.Controls.Add(buttonGetLicensedGroups);
            tabPage2.Controls.Add(buttonMailEnabledSec);
            tabPage2.Controls.Add(buttonGetAllSec);
            tabPage2.Controls.Add(buttonDistributionGroups);
            tabPage2.Controls.Add(buttonUnifiedGroups);
            tabPage2.Controls.Add(buttonGetAllGroups);
            tabPage2.Location = new Point(4, 24);
            tabPage2.Name = "tabPage2";
            tabPage2.Padding = new Padding(3);
            tabPage2.Size = new Size(676, 122);
            tabPage2.TabIndex = 1;
            tabPage2.Text = "Group Reports";
            tabPage2.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            groupBox2.Controls.Add(textBoxGroupName);
            groupBox2.Controls.Add(buttonGetGroupMembers);
            groupBox2.Font = new Font("Segoe UI", 6.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
            groupBox2.Location = new Point(408, 6);
            groupBox2.Name = "groupBox2";
            groupBox2.Size = new Size(128, 116);
            groupBox2.TabIndex = 7;
            groupBox2.TabStop = false;
            groupBox2.Text = "Get Group Members(Enter Name)";
            // 
            // textBoxGroupName
            // 
            textBoxGroupName.Location = new Point(6, 41);
            textBoxGroupName.Name = "textBoxGroupName";
            textBoxGroupName.Size = new Size(116, 19);
            textBoxGroupName.TabIndex = 5;
            // 
            // buttonGetGroupMembers
            // 
            buttonGetGroupMembers.Location = new Point(3, 70);
            buttonGetGroupMembers.Name = "buttonGetGroupMembers";
            buttonGetGroupMembers.Size = new Size(119, 40);
            buttonGetGroupMembers.TabIndex = 4;
            buttonGetGroupMembers.Text = "Get Members";
            buttonGetGroupMembers.UseVisualStyleBackColor = true;
            buttonGetGroupMembers.Click += buttonGetGroupMembers_Click;
            // 
            // buttonGetLicensedGroups
            // 
            buttonGetLicensedGroups.Location = new Point(6, 64);
            buttonGetLicensedGroups.Name = "buttonGetLicensedGroups";
            buttonGetLicensedGroups.Size = new Size(128, 52);
            buttonGetLicensedGroups.TabIndex = 6;
            buttonGetLicensedGroups.Text = "Get All Licensed Groups";
            buttonGetLicensedGroups.UseVisualStyleBackColor = true;
            buttonGetLicensedGroups.Click += buttonGetLicensedGroups_Click;
            // 
            // buttonMailEnabledSec
            // 
            buttonMailEnabledSec.Location = new Point(140, 64);
            buttonMailEnabledSec.Name = "buttonMailEnabledSec";
            buttonMailEnabledSec.Size = new Size(128, 52);
            buttonMailEnabledSec.TabIndex = 5;
            buttonMailEnabledSec.Text = "Get All Mail Enabled Security Groups";
            buttonMailEnabledSec.UseVisualStyleBackColor = true;
            buttonMailEnabledSec.Click += buttonMailEnabledSec_Click;
            // 
            // buttonGetAllSec
            // 
            buttonGetAllSec.Location = new Point(274, 64);
            buttonGetAllSec.Name = "buttonGetAllSec";
            buttonGetAllSec.Size = new Size(128, 52);
            buttonGetAllSec.TabIndex = 4;
            buttonGetAllSec.Text = "Get All Security Groups";
            buttonGetAllSec.UseVisualStyleBackColor = true;
            buttonGetAllSec.Click += button1_Click;
            // 
            // buttonDistributionGroups
            // 
            buttonDistributionGroups.Location = new Point(274, 6);
            buttonDistributionGroups.Name = "buttonDistributionGroups";
            buttonDistributionGroups.Size = new Size(128, 52);
            buttonDistributionGroups.TabIndex = 3;
            buttonDistributionGroups.Text = "Get All Distribution Groups";
            buttonDistributionGroups.UseVisualStyleBackColor = true;
            buttonDistributionGroups.Click += buttonDistributionGroups_Click;
            // 
            // buttonUnifiedGroups
            // 
            buttonUnifiedGroups.Location = new Point(140, 6);
            buttonUnifiedGroups.Name = "buttonUnifiedGroups";
            buttonUnifiedGroups.Size = new Size(128, 52);
            buttonUnifiedGroups.TabIndex = 2;
            buttonUnifiedGroups.Text = "Get All Unified Groups";
            buttonUnifiedGroups.UseVisualStyleBackColor = true;
            buttonUnifiedGroups.Click += buttonUnifiedGroups_Click;
            // 
            // buttonGetAllGroups
            // 
            buttonGetAllGroups.Location = new Point(6, 6);
            buttonGetAllGroups.Name = "buttonGetAllGroups";
            buttonGetAllGroups.Size = new Size(128, 52);
            buttonGetAllGroups.TabIndex = 1;
            buttonGetAllGroups.Text = "Get All Groups";
            buttonGetAllGroups.UseVisualStyleBackColor = true;
            buttonGetAllGroups.Click += buttonGetAllGroups_Click;
            // 
            // tabPage1
            // 
            tabPage1.Controls.Add(buttonMFAReg);
            tabPage1.Controls.Add(buttonGetAdmins);
            tabPage1.Controls.Add(buttonGetUnlicensed);
            tabPage1.Controls.Add(buttonGetGuests);
            tabPage1.Controls.Add(buttonGetSynced);
            tabPage1.Controls.Add(buttonGetAllUsers);
            tabPage1.Location = new Point(4, 24);
            tabPage1.Name = "tabPage1";
            tabPage1.Padding = new Padding(3);
            tabPage1.Size = new Size(676, 122);
            tabPage1.TabIndex = 0;
            tabPage1.Text = "User Reports";
            tabPage1.UseVisualStyleBackColor = true;
            // 
            // buttonMFAReg
            // 
            buttonMFAReg.Location = new Point(8, 64);
            buttonMFAReg.Name = "buttonMFAReg";
            buttonMFAReg.Size = new Size(128, 52);
            buttonMFAReg.TabIndex = 5;
            buttonMFAReg.Text = "Get All Users Auth Registrations";
            buttonMFAReg.UseVisualStyleBackColor = true;
            buttonMFAReg.Click += buttonMFAReg_Click;
            // 
            // buttonGetAdmins
            // 
            buttonGetAdmins.Location = new Point(544, 9);
            buttonGetAdmins.Name = "buttonGetAdmins";
            buttonGetAdmins.Size = new Size(128, 52);
            buttonGetAdmins.TabIndex = 4;
            buttonGetAdmins.Text = "Get All Admins";
            buttonGetAdmins.UseVisualStyleBackColor = true;
            buttonGetAdmins.Click += buttonGetAdmins_Click;
            // 
            // buttonGetUnlicensed
            // 
            buttonGetUnlicensed.Location = new Point(410, 9);
            buttonGetUnlicensed.Name = "buttonGetUnlicensed";
            buttonGetUnlicensed.Size = new Size(128, 52);
            buttonGetUnlicensed.TabIndex = 3;
            buttonGetUnlicensed.Text = "Get All Unlicensed";
            buttonGetUnlicensed.UseVisualStyleBackColor = true;
            buttonGetUnlicensed.Click += buttonGetUnlicensed_Click;
            // 
            // buttonGetGuests
            // 
            buttonGetGuests.Location = new Point(276, 9);
            buttonGetGuests.Name = "buttonGetGuests";
            buttonGetGuests.Size = new Size(128, 52);
            buttonGetGuests.TabIndex = 2;
            buttonGetGuests.Text = "Get All Guests";
            buttonGetGuests.UseVisualStyleBackColor = true;
            buttonGetGuests.Click += buttonGetGuests_Click;
            // 
            // buttonGetSynced
            // 
            buttonGetSynced.Location = new Point(142, 9);
            buttonGetSynced.Name = "buttonGetSynced";
            buttonGetSynced.Size = new Size(128, 52);
            buttonGetSynced.TabIndex = 1;
            buttonGetSynced.Text = "Get All Synced Users";
            buttonGetSynced.UseVisualStyleBackColor = true;
            buttonGetSynced.Click += buttonGetSynced_Click;
            // 
            // buttonGetAllUsers
            // 
            buttonGetAllUsers.Location = new Point(8, 9);
            buttonGetAllUsers.Name = "buttonGetAllUsers";
            buttonGetAllUsers.Size = new Size(128, 52);
            buttonGetAllUsers.TabIndex = 0;
            buttonGetAllUsers.Text = "Get All Users";
            buttonGetAllUsers.UseVisualStyleBackColor = true;
            buttonGetAllUsers.Click += buttonGetAllUsers_Click;
            // 
            // tabControl1
            // 
            tabControl1.Controls.Add(tabPage1);
            tabControl1.Controls.Add(tabPage2);
            tabControl1.Controls.Add(tabPage3);
            tabControl1.Location = new Point(11, 135);
            tabControl1.Name = "tabControl1";
            tabControl1.SelectedIndex = 0;
            tabControl1.Size = new Size(684, 150);
            tabControl1.TabIndex = 1;
            // 
            // Form1
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(697, 727);
            Controls.Add(progressBar1);
            Controls.Add(dataGridView1);
            Controls.Add(buttonExport);
            Controls.Add(tabControl1);
            Controls.Add(tableLayoutPanel1);
            Icon = (Icon)resources.GetObject("$this.Icon");
            Name = "Form1";
            Text = "Graph Reports";
            tableLayoutPanel1.ResumeLayout(false);
            tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridView1).EndInit();
            tabPage3.ResumeLayout(false);
            groupBox1.ResumeLayout(false);
            groupBox1.PerformLayout();
            tabPage2.ResumeLayout(false);
            groupBox2.ResumeLayout(false);
            groupBox2.PerformLayout();
            tabPage1.ResumeLayout(false);
            tabControl1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private TableLayoutPanel tableLayoutPanel1;
        private Label label1;
        private Label label3;
        private TextBox textBoxClientID;
        private TextBox textBoxTenant;
        private Button buttonExport;
        private DataGridView dataGridView1;
        private ProgressBar progressBar1;
        private TabPage tabPage3;
        private Button buttonGetDomains;
        private Button buttonGetSubs;
        private TabPage tabPage2;
        private Button buttonMailEnabledSec;
        private Button buttonGetAllSec;
        private Button buttonDistributionGroups;
        private Button buttonUnifiedGroups;
        private Button buttonGetAllGroups;
        private TabPage tabPage1;
        private Button buttonGetAdmins;
        private Button buttonGetUnlicensed;
        private Button buttonGetGuests;
        private Button buttonGetSynced;
        private Button buttonGetAllUsers;
        private TabControl tabControl1;
        private Button buttonGetDomainDependency;
        private GroupBox groupBox1;
        private TextBox textBoxDomainName;
        private Button buttonGetLicensedGroups;
        private GroupBox groupBox2;
        private TextBox textBoxGroupName;
        private Button buttonGetGroupMembers;
        private Button buttonMFAReg;
        private Button buttonNonComplaintDevices;
        private Button buttonGetAllDevices;
    }
}
