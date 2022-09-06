namespace RevitAddinAcademy
{
    partial class FrmTestForm
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
            this.lblLabel1 = new System.Windows.Forms.Label();
            this.lbxText = new System.Windows.Forms.ListBox();
            this.btnButton1 = new System.Windows.Forms.Button();
            this.btnButton2 = new System.Windows.Forms.Button();
            this.btnButton3 = new System.Windows.Forms.Button();
            this.tbxTextBox = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // lblLabel1
            // 
            this.lblLabel1.AutoSize = true;
            this.lblLabel1.Location = new System.Drawing.Point(12, 9);
            this.lblLabel1.Name = "lblLabel1";
            this.lblLabel1.Size = new System.Drawing.Size(131, 25);
            this.lblLabel1.TabIndex = 0;
            this.lblLabel1.Text = "This is a label";
            // 
            // lbxText
            // 
            this.lbxText.FormattingEnabled = true;
            this.lbxText.ItemHeight = 24;
            this.lbxText.Location = new System.Drawing.Point(23, 46);
            this.lbxText.Name = "lbxText";
            this.lbxText.Size = new System.Drawing.Size(376, 388);
            this.lbxText.TabIndex = 1;
            this.lbxText.DoubleClick += new System.EventHandler(this.lbxText_DoubleClick);
            // 
            // btnButton1
            // 
            this.btnButton1.Location = new System.Drawing.Point(446, 46);
            this.btnButton1.Name = "btnButton1";
            this.btnButton1.Size = new System.Drawing.Size(169, 50);
            this.btnButton1.TabIndex = 2;
            this.btnButton1.Text = "Do Something";
            this.btnButton1.UseVisualStyleBackColor = true;
            this.btnButton1.Click += new System.EventHandler(this.btnButton1_Click);
            // 
            // btnButton2
            // 
            this.btnButton2.Location = new System.Drawing.Point(446, 141);
            this.btnButton2.Name = "btnButton2";
            this.btnButton2.Size = new System.Drawing.Size(169, 50);
            this.btnButton2.TabIndex = 3;
            this.btnButton2.Text = "Do Something 2";
            this.btnButton2.UseVisualStyleBackColor = true;
            this.btnButton2.Click += new System.EventHandler(this.btnButton2_Click);
            // 
            // btnButton3
            // 
            this.btnButton3.Location = new System.Drawing.Point(446, 247);
            this.btnButton3.Name = "btnButton3";
            this.btnButton3.Size = new System.Drawing.Size(169, 50);
            this.btnButton3.TabIndex = 4;
            this.btnButton3.Text = "Do Something 3";
            this.btnButton3.UseVisualStyleBackColor = true;
            this.btnButton3.Click += new System.EventHandler(this.btnButton3_Click);
            // 
            // tbxTextBox
            // 
            this.tbxTextBox.Location = new System.Drawing.Point(429, 341);
            this.tbxTextBox.Name = "tbxTextBox";
            this.tbxTextBox.Size = new System.Drawing.Size(323, 29);
            this.tbxTextBox.TabIndex = 5;
            this.tbxTextBox.Text = "This is default text";
            // 
            // FrmTestForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(11F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(811, 463);
            this.Controls.Add(this.tbxTextBox);
            this.Controls.Add(this.btnButton3);
            this.Controls.Add(this.btnButton2);
            this.Controls.Add(this.btnButton1);
            this.Controls.Add(this.lbxText);
            this.Controls.Add(this.lblLabel1);
            this.Name = "FrmTestForm";
            this.Text = "This is a test form";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lblLabel1;
        private System.Windows.Forms.ListBox lbxText;
        private System.Windows.Forms.Button btnButton1;
        private System.Windows.Forms.Button btnButton2;
        private System.Windows.Forms.Button btnButton3;
        private System.Windows.Forms.TextBox tbxTextBox;
    }
}