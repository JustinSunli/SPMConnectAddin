using System;
using System.Windows.Forms;

namespace SPMConnectAddin
{
    public partial class OpenItem : Form
    {
        public string ValueIWant { get; set; }

        public OpenItem()
        {
            InitializeComponent();

        }

        bool IsDigitsOnly(string str)
        {
            foreach (char c in str)
            {
                if (c < '0' || c > '9')
                    return false;
            }

            return true;
        }


        private void Dialog_Load(object sender, EventArgs e)
        {
        }

        private void bunifuMaterialTextbox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (itmtxt.Text.Length == 6 && e.KeyChar != 8)
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }

        }

        private void bunifuFlatButton1_Click(object sender, EventArgs e)
        {
            errorProvider1.Clear();

            if (itmtxt.Text.Length == 6 && !String.IsNullOrEmpty(itmtxt.Text) && Char.IsLetter(itmtxt.Text[0]) && IsDigitsOnly(itmtxt.Text.Substring(1, 5)))
            {
                ValueIWant = itmtxt.Text.Trim();
                this.DialogResult = System.Windows.Forms.DialogResult.OK;
                this.Close();
            }
            else
            {
                errorProvider1.SetError(itmtxt, "Not a valid part number. Please enter a valid six digit SPM item number (starting with 'A', 'B', 'C') to open solidworks model.");
            }

        }

        private void bunifuFlatButton2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void itmtxt_KeyDown(object sender, KeyEventArgs e)
        {
            errorProvider1.Clear();
            if (e.KeyCode == Keys.Enter)
            {
                if (itmtxt.Text.Length == 6 && !String.IsNullOrEmpty(itmtxt.Text) && Char.IsLetter(itmtxt.Text[0]) && IsDigitsOnly(itmtxt.Text.Substring(1, 5)))
                {
                    ValueIWant = itmtxt.Text.Trim();
                    this.DialogResult = System.Windows.Forms.DialogResult.OK;
                    this.Close();
                }
                else
                {
                    errorProvider1.SetError(itmtxt, "Not a valid part number. Please enter a valid six digit SPM item number (starting with 'A', 'B', 'C') to open solidworks model.");
                }

            }
        }
    }
}
