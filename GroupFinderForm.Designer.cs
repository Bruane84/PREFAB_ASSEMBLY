namespace TeklaGroupFinder
{
    partial class GroupFinderForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.label1 = new System.Windows.Forms.Label();
            this.projectNumberTextBox = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.toleranceTextBox = new System.Windows.Forms.TextBox();
            this.createSignatureButton = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label3 = new System.Windows.Forms.Label();
            this.udaNameTextBox = new System.Windows.Forms.TextBox();
            this.findMatchesButton = new System.Windows.Forms.Button();
            this.openCatalogueButton = new System.Windows.Forms.Button();
            this.statusLabel = new System.Windows.Forms.Label();
            this.connectionStatusPanel = new System.Windows.Forms.Panel();
            this.label4 = new System.Windows.Forms.Label();
            this.processingProgressBar = new System.Windows.Forms.ProgressBar();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 26);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(83, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Project Number:";
            // 
            // projectNumberTextBox
            // 
            this.projectNumberTextBox.Location = new System.Drawing.Point(101, 23);
            this.projectNumberTextBox.Name = "projectNumberTextBox";
            this.projectNumberTextBox.Size = new System.Drawing.Size(100, 20);
            this.projectNumberTextBox.TabIndex = 1;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.toleranceTextBox);
            this.groupBox1.Controls.Add(this.createSignatureButton);
            this.groupBox1.Location = new System.Drawing.Point(20, 68);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(182, 156);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "1. Create Signature File";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(42, 32);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Tolerance (mm):";
            // 
            // toleranceTextBox
            // 
            this.toleranceTextBox.Location = new System.Drawing.Point(36, 48);
            this.toleranceTextBox.Name = "toleranceTextBox";
            this.toleranceTextBox.Size = new System.Drawing.Size(100, 20);
            this.toleranceTextBox.TabIndex = 1;
            this.toleranceTextBox.Text = ".5";
            // 
            // createSignatureButton
            // 
            this.createSignatureButton.Location = new System.Drawing.Point(15, 91);
            this.createSignatureButton.Name = "createSignatureButton";
            this.createSignatureButton.Size = new System.Drawing.Size(150, 46);
            this.createSignatureButton.TabIndex = 0;
            this.createSignatureButton.Text = "Create Signature from Selection";
            this.createSignatureButton.UseVisualStyleBackColor = true;
            this.createSignatureButton.Click += new System.EventHandler(this.createSignatureButton_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.udaNameTextBox);
            this.groupBox2.Controls.Add(this.findMatchesButton);
            this.groupBox2.Location = new System.Drawing.Point(20, 230);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(217, 191);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "2. Find & Group Matches";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(32, 94);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(104, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "UDA Name to Write:";
            // 
            // udaNameTextBox
            // 
            this.udaNameTextBox.Location = new System.Drawing.Point(15, 110);
            this.udaNameTextBox.Name = "udaNameTextBox";
            this.udaNameTextBox.Size = new System.Drawing.Size(150, 20);
            this.udaNameTextBox.TabIndex = 4;
            this.udaNameTextBox.Text = "SUB_ASSEMBLY_ID";
            // 
            // findMatchesButton
            // 
            this.findMatchesButton.Location = new System.Drawing.Point(15, 31);
            this.findMatchesButton.Name = "findMatchesButton";
            this.findMatchesButton.Size = new System.Drawing.Size(150, 46);
            this.findMatchesButton.TabIndex = 0;
            this.findMatchesButton.Text = "Find Matches Using Signature";
            this.findMatchesButton.UseVisualStyleBackColor = true;
            this.findMatchesButton.Click += new System.EventHandler(this.findMatchesButton_Click);
            // 
            // openCatalogueButton
            // 
            this.openCatalogueButton.Location = new System.Drawing.Point(20, 437);
            this.openCatalogueButton.Name = "openCatalogueButton";
            this.openCatalogueButton.Size = new System.Drawing.Size(182, 51);
            this.openCatalogueButton.TabIndex = 4;
            this.openCatalogueButton.Text = "Open Signature Catalogue";
            this.openCatalogueButton.UseVisualStyleBackColor = true;
            this.openCatalogueButton.Click += new System.EventHandler(this.openCatalogueButton_Click);
            // 
            // statusLabel
            // 
            this.statusLabel.AutoSize = true;
            this.statusLabel.Location = new System.Drawing.Point(2, 513);
            this.statusLabel.Name = "statusLabel";
            this.statusLabel.Size = new System.Drawing.Size(41, 13);
            this.statusLabel.TabIndex = 5;
            this.statusLabel.Text = "Ready.";
            // 
            // connectionStatusPanel
            // 
            this.connectionStatusPanel.BackColor = System.Drawing.Color.Red;
            this.connectionStatusPanel.Location = new System.Drawing.Point(233, 491);
            this.connectionStatusPanel.Name = "connectionStatusPanel";
            this.connectionStatusPanel.Size = new System.Drawing.Size(16, 16);
            this.connectionStatusPanel.TabIndex = 6;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(2, 491);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(97, 13);
            this.label4.TabIndex = 7;
            this.label4.Text = "Connection Status:";
            // 
            // processingProgressBar
            // 
            this.processingProgressBar.Location = new System.Drawing.Point(20, 497);
            this.processingProgressBar.Name = "processingProgressBar";
            this.processingProgressBar.Size = new System.Drawing.Size(221, 10);
            this.processingProgressBar.TabIndex = 8;
            this.processingProgressBar.Visible = false;
            // 
            // GroupFinderForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(261, 532);
            this.Controls.Add(this.processingProgressBar);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.connectionStatusPanel);
            this.Controls.Add(this.statusLabel);
            this.Controls.Add(this.openCatalogueButton);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.projectNumberTextBox);
            this.Controls.Add(this.label1);
            this.Name = "GroupFinderForm";
            this.ShowInTaskbar = true;
            this.Text = "MegaPanelPost Grouping";
            this.TopMost = true;
            //this.Load += new System.EventHandler(this.GroupFinderForm_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox projectNumberTextBox;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.Button createSignatureButton;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox toleranceTextBox;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button findMatchesButton;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox udaNameTextBox;
        private System.Windows.Forms.Button openCatalogueButton;
        private System.Windows.Forms.Label statusLabel;
        private System.Windows.Forms.Panel connectionStatusPanel;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ProgressBar processingProgressBar;
    }
}