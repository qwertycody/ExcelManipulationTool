using ExcelManipulationTool.Modules;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelManipulationTool
{
    public partial class Main : Form
    {
        public String filePath_fileToRead = "";

        public Boolean testMode = true;

        public Main()
        {
            InitializeComponent();

            if (this.testMode == true)
            {
                ExecutionModule module = new ExecutionModule("../../Assets/ExampleFile.xlsx", this.testMode);
            }
        }

        private String getFilePath(String titleText)
        {
            String filePathToReturn = null;

            Boolean successful = false;

            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Title = titleText;

            if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                successful = openFileDialog.CheckFileExists;
                filePathToReturn = openFileDialog.FileName;
            }
            else
            {
                successful = false;
            }


            if (successful != true)
            {
                MessageBox.Show("File Path Not Valid!");
                return null;
            }
            else
            {
                return filePathToReturn;
            }
        }

        private void button_fileToRead_Click(object sender, EventArgs e)
        {
            String filePath = getFilePath("Select File To Read");

            if (filePath != null)
            {
                this.filePath_fileToRead = filePath;
                textBox_fileToRead.Text = filePath;
            }
            else
            {
                textBox_fileToRead.Text = "";
            }
        }

        private void button_execute_Click(object sender, EventArgs e)
        {
            ExecutionModule module = new ExecutionModule(this.filePath_fileToRead, this.testMode);
            this.Close();
        }
    }
}
