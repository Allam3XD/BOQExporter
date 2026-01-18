namespace BOQExporter.UI
{
    partial class MainFrom
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainFrom));
            this.ExtractBtn = new System.Windows.Forms.Button();
            this.ExportBtn = new System.Windows.Forms.Button();
            this.CloseBtn = new System.Windows.Forms.Button();
            this.dgvItems = new System.Windows.Forms.DataGridView();
            this.ItemCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Description = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Usage = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.lblCount = new System.Windows.Forms.Label();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.clbCategories = new System.Windows.Forms.CheckedListBox();
            this.ArchCB = new System.Windows.Forms.CheckBox();
            this.MechCB = new System.Windows.Forms.CheckBox();
            this.ElecCB = new System.Windows.Forms.CheckBox();
            this.StruCB = new System.Windows.Forms.CheckBox();
            this.SelectedItems = new System.Windows.Forms.Label();
            this.metroStyleManager1 = new MetroFramework.Components.MetroStyleManager(this.components);
            this.txtProjectName = new System.Windows.Forms.TextBox();
            this.txtClientName = new System.Windows.Forms.TextBox();
            this.txtLocation = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.Location = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.btnBrowseLogo = new System.Windows.Forms.Button();
            this.picLogo = new System.Windows.Forms.PictureBox();
            this.dtpIssueDate = new System.Windows.Forms.DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this.dgvItems)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.metroStyleManager1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picLogo)).BeginInit();
            this.SuspendLayout();
            // 
            // ExtractBtn
            // 
            this.ExtractBtn.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ExtractBtn.Location = new System.Drawing.Point(611, 281);
            this.ExtractBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ExtractBtn.Name = "ExtractBtn";
            this.ExtractBtn.Size = new System.Drawing.Size(133, 39);
            this.ExtractBtn.TabIndex = 0;
            this.ExtractBtn.Text = "Extract";
            this.ExtractBtn.UseVisualStyleBackColor = false;
            this.ExtractBtn.Click += new System.EventHandler(this.ExtractBtn_Click);
            // 
            // ExportBtn
            // 
            this.ExportBtn.BackColor = System.Drawing.Color.WhiteSmoke;
            this.ExportBtn.Location = new System.Drawing.Point(611, 336);
            this.ExportBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ExportBtn.Name = "ExportBtn";
            this.ExportBtn.Size = new System.Drawing.Size(133, 39);
            this.ExportBtn.TabIndex = 1;
            this.ExportBtn.Text = "Export";
            this.ExportBtn.UseVisualStyleBackColor = false;
            this.ExportBtn.Click += new System.EventHandler(this.ExportBtn_Click);
            // 
            // CloseBtn
            // 
            this.CloseBtn.BackColor = System.Drawing.Color.WhiteSmoke;
            this.CloseBtn.Location = new System.Drawing.Point(611, 391);
            this.CloseBtn.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.CloseBtn.Name = "CloseBtn";
            this.CloseBtn.Size = new System.Drawing.Size(133, 39);
            this.CloseBtn.TabIndex = 2;
            this.CloseBtn.Text = "Close";
            this.CloseBtn.UseVisualStyleBackColor = false;
            this.CloseBtn.Click += new System.EventHandler(this.CloseBtn_Click);
            // 
            // dgvItems
            // 
            this.dgvItems.BackgroundColor = System.Drawing.Color.LightSkyBlue;
            this.dgvItems.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvItems.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.ItemCode,
            this.Description,
            this.Quantity,
            this.Usage});
            this.dgvItems.Location = new System.Drawing.Point(27, 280);
            this.dgvItems.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.dgvItems.Name = "dgvItems";
            this.dgvItems.RowHeadersWidth = 51;
            this.dgvItems.RowTemplate.Height = 24;
            this.dgvItems.Size = new System.Drawing.Size(553, 158);
            this.dgvItems.TabIndex = 3;
            // 
            // ItemCode
            // 
            this.ItemCode.HeaderText = "ItemCode";
            this.ItemCode.MinimumWidth = 6;
            this.ItemCode.Name = "ItemCode";
            this.ItemCode.Width = 125;
            // 
            // Description
            // 
            this.Description.HeaderText = "Description";
            this.Description.MinimumWidth = 6;
            this.Description.Name = "Description";
            this.Description.Width = 125;
            // 
            // Quantity
            // 
            this.Quantity.HeaderText = "Quantity";
            this.Quantity.MinimumWidth = 6;
            this.Quantity.Name = "Quantity";
            this.Quantity.Width = 125;
            // 
            // Usage
            // 
            this.Usage.HeaderText = "Usage";
            this.Usage.MinimumWidth = 6;
            this.Usage.Name = "Usage";
            this.Usage.Width = 125;
            // 
            // lblCount
            // 
            this.lblCount.AutoSize = true;
            this.lblCount.Location = new System.Drawing.Point(24, 252);
            this.lblCount.Name = "lblCount";
            this.lblCount.Size = new System.Drawing.Size(110, 16);
            this.lblCount.TabIndex = 4;
            this.lblCount.Text = "Items extracted: 0";
            // 
            // clbCategories
            // 
            this.clbCategories.FormattingEnabled = true;
            this.clbCategories.Location = new System.Drawing.Point(27, 89);
            this.clbCategories.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.clbCategories.Name = "clbCategories";
            this.clbCategories.Size = new System.Drawing.Size(196, 157);
            this.clbCategories.TabIndex = 5;
            this.clbCategories.SelectedIndexChanged += new System.EventHandler(this.clbCategories_SelectedIndexChanged);
            // 
            // ArchCB
            // 
            this.ArchCB.AutoSize = true;
            this.ArchCB.Location = new System.Drawing.Point(235, 103);
            this.ArchCB.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ArchCB.Name = "ArchCB";
            this.ArchCB.Size = new System.Drawing.Size(99, 20);
            this.ArchCB.TabIndex = 7;
            this.ArchCB.Text = "Architecture";
            this.ArchCB.UseVisualStyleBackColor = true;
            this.ArchCB.CheckedChanged += new System.EventHandler(this.ArchCB_CheckedChanged);
            // 
            // MechCB
            // 
            this.MechCB.AutoSize = true;
            this.MechCB.Location = new System.Drawing.Point(235, 182);
            this.MechCB.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.MechCB.Name = "MechCB";
            this.MechCB.Size = new System.Drawing.Size(98, 20);
            this.MechCB.TabIndex = 8;
            this.MechCB.Text = "Mechanical";
            this.MechCB.UseVisualStyleBackColor = true;
            this.MechCB.CheckedChanged += new System.EventHandler(this.MechCB_CheckedChanged);
            // 
            // ElecCB
            // 
            this.ElecCB.AutoSize = true;
            this.ElecCB.Location = new System.Drawing.Point(235, 219);
            this.ElecCB.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.ElecCB.Name = "ElecCB";
            this.ElecCB.Size = new System.Drawing.Size(84, 20);
            this.ElecCB.TabIndex = 9;
            this.ElecCB.Text = "Electrical";
            this.ElecCB.UseVisualStyleBackColor = true;
            this.ElecCB.CheckedChanged += new System.EventHandler(this.ElecCB_CheckedChanged_1);
            // 
            // StruCB
            // 
            this.StruCB.AutoSize = true;
            this.StruCB.Location = new System.Drawing.Point(235, 143);
            this.StruCB.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.StruCB.Name = "StruCB";
            this.StruCB.Size = new System.Drawing.Size(81, 20);
            this.StruCB.TabIndex = 10;
            this.StruCB.Text = "Structure";
            this.StruCB.UseVisualStyleBackColor = true;
            this.StruCB.CheckedChanged += new System.EventHandler(this.StruCB_CheckedChanged_1);
            // 
            // SelectedItems
            // 
            this.SelectedItems.AutoSize = true;
            this.SelectedItems.Location = new System.Drawing.Point(24, 64);
            this.SelectedItems.Name = "SelectedItems";
            this.SelectedItems.Size = new System.Drawing.Size(146, 16);
            this.SelectedItems.TabIndex = 11;
            this.SelectedItems.Text = "Selected Categories : 0";
            // 
            // metroStyleManager1
            // 
            this.metroStyleManager1.Owner = this;
            // 
            // txtProjectName
            // 
            this.txtProjectName.Location = new System.Drawing.Point(449, 119);
            this.txtProjectName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtProjectName.Name = "txtProjectName";
            this.txtProjectName.Size = new System.Drawing.Size(296, 22);
            this.txtProjectName.TabIndex = 12;
            // 
            // txtClientName
            // 
            this.txtClientName.Location = new System.Drawing.Point(449, 156);
            this.txtClientName.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtClientName.Name = "txtClientName";
            this.txtClientName.Size = new System.Drawing.Size(296, 22);
            this.txtClientName.TabIndex = 13;
            // 
            // txtLocation
            // 
            this.txtLocation.Location = new System.Drawing.Point(449, 191);
            this.txtLocation.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.txtLocation.Name = "txtLocation";
            this.txtLocation.Size = new System.Drawing.Size(296, 22);
            this.txtLocation.TabIndex = 14;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(353, 122);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(95, 16);
            this.label1.TabIndex = 16;
            this.label1.Text = "Project Name :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(353, 160);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(86, 16);
            this.label2.TabIndex = 17;
            this.label2.Text = "Client Name :";
            // 
            // Location
            // 
            this.Location.AutoSize = true;
            this.Location.Location = new System.Drawing.Point(353, 194);
            this.Location.Name = "Location";
            this.Location.Size = new System.Drawing.Size(64, 16);
            this.Location.TabIndex = 18;
            this.Location.Text = "Location :";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(353, 230);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 16);
            this.label3.TabIndex = 19;
            this.label3.Text = "Issue Date :";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(353, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(44, 16);
            this.label4.TabIndex = 21;
            this.label4.Text = "Logo :";
            // 
            // btnBrowseLogo
            // 
            this.btnBrowseLogo.BackColor = System.Drawing.Color.WhiteSmoke;
            this.btnBrowseLogo.Location = new System.Drawing.Point(611, 49);
            this.btnBrowseLogo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.btnBrowseLogo.Name = "btnBrowseLogo";
            this.btnBrowseLogo.Size = new System.Drawing.Size(133, 39);
            this.btnBrowseLogo.TabIndex = 22;
            this.btnBrowseLogo.Text = "Browse For Logo";
            this.btnBrowseLogo.UseVisualStyleBackColor = false;
            this.btnBrowseLogo.Click += new System.EventHandler(this.btnBrowseLogo_Click);
            // 
            // picLogo
            // 
            this.picLogo.Location = new System.Drawing.Point(449, 25);
            this.picLogo.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.picLogo.Name = "picLogo";
            this.picLogo.Size = new System.Drawing.Size(139, 87);
            this.picLogo.TabIndex = 23;
            this.picLogo.TabStop = false;
            // 
            // dtpIssueDate
            // 
            this.dtpIssueDate.CustomFormat = "\"dd/mm/yyyy\"";
            this.dtpIssueDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.dtpIssueDate.Location = new System.Drawing.Point(449, 230);
            this.dtpIssueDate.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.dtpIssueDate.Name = "dtpIssueDate";
            this.dtpIssueDate.Size = new System.Drawing.Size(293, 22);
            this.dtpIssueDate.TabIndex = 24;
            // 
            // MainFrom
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(773, 452);
            this.Controls.Add(this.dtpIssueDate);
            this.Controls.Add(this.picLogo);
            this.Controls.Add(this.btnBrowseLogo);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Location);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txtLocation);
            this.Controls.Add(this.txtClientName);
            this.Controls.Add(this.txtProjectName);
            this.Controls.Add(this.SelectedItems);
            this.Controls.Add(this.ElecCB);
            this.Controls.Add(this.StruCB);
            this.Controls.Add(this.clbCategories);
            this.Controls.Add(this.MechCB);
            this.Controls.Add(this.lblCount);
            this.Controls.Add(this.ArchCB);
            this.Controls.Add(this.dgvItems);
            this.Controls.Add(this.CloseBtn);
            this.Controls.Add(this.ExportBtn);
            this.Controls.Add(this.ExtractBtn);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Margin = new System.Windows.Forms.Padding(3, 2, 3, 2);
            this.Name = "MainFrom";
            this.Padding = new System.Windows.Forms.Padding(20, 74, 20, 20);
            this.Text = "BOQ Exporter";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainFrom_FormClosing);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MainFrom_FormClosed);
            ((System.ComponentModel.ISupportInitialize)(this.dgvItems)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.metroStyleManager1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picLogo)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button ExtractBtn;
        private System.Windows.Forms.Button ExportBtn;
        private System.Windows.Forms.Button CloseBtn;
        private System.Windows.Forms.DataGridView dgvItems;
        private System.Windows.Forms.Label lblCount;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.CheckedListBox clbCategories;
        private System.Windows.Forms.DataGridViewTextBoxColumn ItemCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn Description;
        private System.Windows.Forms.DataGridViewTextBoxColumn Quantity;
        private System.Windows.Forms.DataGridViewTextBoxColumn Usage;
        private System.Windows.Forms.CheckBox ArchCB;
        private System.Windows.Forms.CheckBox MechCB;
        private System.Windows.Forms.CheckBox ElecCB;
        private System.Windows.Forms.CheckBox StruCB;
        private System.Windows.Forms.Label SelectedItems;
        public MetroFramework.Components.MetroStyleManager metroStyleManager1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtLocation;
        private System.Windows.Forms.TextBox txtClientName;
        private System.Windows.Forms.TextBox txtProjectName;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label Location;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnBrowseLogo;
        private System.Windows.Forms.PictureBox picLogo;
        private System.Windows.Forms.DateTimePicker dtpIssueDate;
    }
}