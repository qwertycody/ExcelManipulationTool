namespace ExcelManipulationTool
{
    partial class Main
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
            this.button_fileToRead = new System.Windows.Forms.Button();
            this.textBox_fileToRead = new System.Windows.Forms.TextBox();
            this.button_execute = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // button_fileToRead
            // 
            this.button_fileToRead.Location = new System.Drawing.Point(425, 12);
            this.button_fileToRead.Name = "button_fileToRead";
            this.button_fileToRead.Size = new System.Drawing.Size(112, 34);
            this.button_fileToRead.TabIndex = 0;
            this.button_fileToRead.Text = "Select File";
            this.button_fileToRead.UseVisualStyleBackColor = true;
            this.button_fileToRead.Click += new System.EventHandler(this.button_fileToRead_Click);
            // 
            // textBox_fileToRead
            // 
            this.textBox_fileToRead.Font = new System.Drawing.Font("Microsoft Sans Serif", 13.8F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.textBox_fileToRead.Location = new System.Drawing.Point(12, 12);
            this.textBox_fileToRead.Name = "textBox_fileToRead";
            this.textBox_fileToRead.ReadOnly = true;
            this.textBox_fileToRead.Size = new System.Drawing.Size(407, 34);
            this.textBox_fileToRead.TabIndex = 1;
            // 
            // button_execute
            // 
            this.button_execute.Location = new System.Drawing.Point(12, 52);
            this.button_execute.Name = "button_execute";
            this.button_execute.Size = new System.Drawing.Size(525, 27);
            this.button_execute.TabIndex = 2;
            this.button_execute.Text = "Execute";
            this.button_execute.UseVisualStyleBackColor = true;
            this.button_execute.Click += new System.EventHandler(this.button_execute_Click);
            // 
            // Main
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(548, 91);
            this.Controls.Add(this.button_execute);
            this.Controls.Add(this.textBox_fileToRead);
            this.Controls.Add(this.button_fileToRead);
            this.Name = "Main";
            this.Text = "Main Window";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button_fileToRead;
        private System.Windows.Forms.TextBox textBox_fileToRead;
        private System.Windows.Forms.Button button_execute;
    }
}

