using System;
using System.Windows.Forms;

namespace ConsoleApp1
{
    public partial class UserInputForm : Form
    {
        // 👇 This is the property Program.cs will use
        public string ModelPath { get; private set; }

        public UserInputForm()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            // 👇 txtModelPath is the TextBox where the user types or pastes the path
            ModelPath = txtModelPath.Text;
            this.DialogResult = DialogResult.OK; // tells Program.cs "user clicked Start"
            this.Close();
        }

        private void UserInputForm_Load(object sender, EventArgs e)
        {
            // Optional: pre-fill with an example path
            txtModelPath.Text = @"C:\Temp\Part1.ipt";
        }

        private void button1_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                // 👇 configure dialog
                openFileDialog.InitialDirectory = "C:\\";
                openFileDialog.Filter = "Inventor Files (*.ipt;*.iam)|*.ipt;*.iam|All files (*.*)|*.*";
                openFileDialog.FilterIndex = 1;
                openFileDialog.RestoreDirectory = true;

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    // 👇 set the textbox and property
                    txtModelPath.Text = openFileDialog.FileName;
                    ModelPath = openFileDialog.FileName;
                }
            }
        }
    }
}