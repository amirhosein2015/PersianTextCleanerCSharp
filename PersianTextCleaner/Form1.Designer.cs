
namespace PersianTextCleaner
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.btnOpen = new System.Windows.Forms.Button();
            this.textBoxOriginal = new System.Windows.Forms.TextBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.textBoxCleaned = new System.Windows.Forms.TextBox();
            this.btnCopy = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.btnSaveAs = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.btnClear = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnOpen
            // 
            this.btnOpen.BackColor = System.Drawing.Color.Khaki;
            this.btnOpen.Font = new System.Drawing.Font("Trebuchet MS", 8.1F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnOpen.Location = new System.Drawing.Point(157, 31);
            this.btnOpen.Name = "btnOpen";
            this.btnOpen.Size = new System.Drawing.Size(148, 83);
            this.btnOpen.TabIndex = 0;
            this.btnOpen.Text = "Open";
            this.btnOpen.UseVisualStyleBackColor = false;
            this.btnOpen.Click += new System.EventHandler(this.btnOpen_Click);
            // 
            // textBoxOriginal
            // 
            this.textBoxOriginal.Font = new System.Drawing.Font("Mikhak Medium", 8.1F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBoxOriginal.Location = new System.Drawing.Point(157, 138);
            this.textBoxOriginal.Multiline = true;
            this.textBoxOriginal.Name = "textBoxOriginal";
            this.textBoxOriginal.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBoxOriginal.Size = new System.Drawing.Size(1818, 588);
            this.textBoxOriginal.TabIndex = 1;
            // 
            // btnRun
            // 
            this.btnRun.BackColor = System.Drawing.Color.Khaki;
            this.btnRun.Font = new System.Drawing.Font("Trebuchet MS", 9.900001F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnRun.Location = new System.Drawing.Point(2025, 369);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(183, 132);
            this.btnRun.TabIndex = 2;
            this.btnRun.Text = "Run";
            this.btnRun.UseVisualStyleBackColor = false;
            this.btnRun.Click += new System.EventHandler(this.buttonRun_Click);
            // 
            // textBoxCleaned
            // 
            this.textBoxCleaned.Font = new System.Drawing.Font("Mikhak Medium", 9.900001F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(178)));
            this.textBoxCleaned.Location = new System.Drawing.Point(157, 764);
            this.textBoxCleaned.Multiline = true;
            this.textBoxCleaned.Name = "textBoxCleaned";
            this.textBoxCleaned.RightToLeft = System.Windows.Forms.RightToLeft.Yes;
            this.textBoxCleaned.Size = new System.Drawing.Size(1818, 637);
            this.textBoxCleaned.TabIndex = 3;
            this.textBoxCleaned.TextChanged += new System.EventHandler(this.textBoxCleaned_TextChanged);
            // 
            // btnCopy
            // 
            this.btnCopy.BackColor = System.Drawing.Color.Khaki;
            this.btnCopy.Font = new System.Drawing.Font("Trebuchet MS", 8.1F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCopy.Location = new System.Drawing.Point(2025, 979);
            this.btnCopy.Name = "btnCopy";
            this.btnCopy.Size = new System.Drawing.Size(173, 84);
            this.btnCopy.TabIndex = 4;
            this.btnCopy.Text = "Copy";
            this.btnCopy.UseVisualStyleBackColor = false;
            this.btnCopy.Click += new System.EventHandler(this.btnCopy_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // btnSaveAs
            // 
            this.btnSaveAs.BackColor = System.Drawing.Color.Khaki;
            this.btnSaveAs.Font = new System.Drawing.Font("Trebuchet MS", 8.1F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnSaveAs.Location = new System.Drawing.Point(500, 31);
            this.btnSaveAs.Name = "btnSaveAs";
            this.btnSaveAs.Size = new System.Drawing.Size(158, 83);
            this.btnSaveAs.TabIndex = 5;
            this.btnSaveAs.Text = "Save";
            this.btnSaveAs.UseVisualStyleBackColor = false;
            this.btnSaveAs.Click += new System.EventHandler(this.buttonSaveAs_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.label1.ForeColor = System.Drawing.SystemColors.AppWorkspace;
            this.label1.Location = new System.Drawing.Point(2048, 1389);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(245, 32);
            this.label1.TabIndex = 6;
            this.label1.Text = "Abdollah Mohajeri";
            // 
            // btnClear
            // 
            this.btnClear.BackColor = System.Drawing.Color.Khaki;
            this.btnClear.Font = new System.Drawing.Font("Trebuchet MS", 8.1F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClear.Location = new System.Drawing.Point(327, 31);
            this.btnClear.Name = "btnClear";
            this.btnClear.Size = new System.Drawing.Size(148, 83);
            this.btnClear.TabIndex = 7;
            this.btnClear.Text = "Clear";
            this.btnClear.UseVisualStyleBackColor = false;
            this.btnClear.Click += new System.EventHandler(this.btnClear_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(16F, 31F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ActiveCaptionText;
            this.ClientSize = new System.Drawing.Size(2338, 1443);
            this.Controls.Add(this.btnClear);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnSaveAs);
            this.Controls.Add(this.btnCopy);
            this.Controls.Add(this.textBoxCleaned);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.textBoxOriginal);
            this.Controls.Add(this.btnOpen);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Text Cleaner";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnOpen;
        private System.Windows.Forms.TextBox textBoxOriginal;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.TextBox textBoxCleaned;
        private System.Windows.Forms.Button btnCopy;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button btnSaveAs;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnClear;
    }
}

