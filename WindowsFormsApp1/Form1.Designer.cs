namespace WindowsFormsApp1
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
            this.btnFile = new System.Windows.Forms.Button();
            this.txtBox = new System.Windows.Forms.TextBox();
            this.btnImport = new System.Windows.Forms.Button();
            this.listView1 = new System.Windows.Forms.ListView();
            this.btnPrint = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnFile
            // 
            this.btnFile.Location = new System.Drawing.Point(73, 204);
            this.btnFile.Name = "btnFile";
            this.btnFile.Size = new System.Drawing.Size(75, 23);
            this.btnFile.TabIndex = 0;
            this.btnFile.Text = "Select File to Upload";
            this.btnFile.UseVisualStyleBackColor = true;
            this.btnFile.Click += new System.EventHandler(this.btnFile_Click);
            // 
            // txtBox
            // 
            this.txtBox.Location = new System.Drawing.Point(73, 247);
            this.txtBox.Multiline = true;
            this.txtBox.Name = "txtBox";
            this.txtBox.Size = new System.Drawing.Size(594, 140);
            this.txtBox.TabIndex = 2;
            // 
            // btnImport
            // 
            this.btnImport.Location = new System.Drawing.Point(182, 204);
            this.btnImport.Name = "btnImport";
            this.btnImport.Size = new System.Drawing.Size(107, 23);
            this.btnImport.TabIndex = 3;
            this.btnImport.Text = "Import Data";
            this.btnImport.UseVisualStyleBackColor = true;
            // 
            // listView1
            // 
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(73, 23);
            this.listView1.Name = "listView1";
            this.listView1.Size = new System.Drawing.Size(594, 157);
            this.listView1.TabIndex = 4;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.Click += new System.EventHandler(this.listView1_Click);
            // 
            // btnPrint
            // 
            this.btnPrint.Location = new System.Drawing.Point(516, 204);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(107, 23);
            this.btnPrint.TabIndex = 5;
            this.btnPrint.Text = "Print Excel File";
            this.btnPrint.UseVisualStyleBackColor = true;
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.btnImport);
            this.Controls.Add(this.txtBox);
            this.Controls.Add(this.btnFile);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button btnFile;
        private System.Windows.Forms.TextBox txtBox;
        private System.Windows.Forms.Button btnImport;
        private System.Windows.Forms.ListView listView1;
        private System.Windows.Forms.Button btnPrint;
    }
}

