using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Diagnostics;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;

namespace SPMConnectAddin
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            swApp = (SldWorks)System.Runtime.InteropServices.Marshal.GetActiveObject("SldWorks.Application");
            retVal = swApp.GetUserProgressBar(out pb);
        }
        UserProgressBar pb;
        int Position;
        int lRet;
        bool retVal;
        bool boolstatus;
        public SldWorks swApp;

        public void cmdExit_Click(System.Object sender, System.EventArgs e)
        {
            progressBar1.Value = 0;
            label1.Text = "";
            closebttn.Visible = false;
            this.Close();
        }

        public void UpdateProgressBar(int value, string labeltext)
        {
            this.Invoke((MethodInvoker)delegate
            {
                progressBar1.Value = value;
                label1.Text = labeltext;
                pb.UpdateTitle(labeltext);
                if (value == 100)
                {
                    closebttn.Visible = true;
                }
            });
        }


        public void cmdStartPB_Click(System.Object sender, System.EventArgs e)
        {
            boolstatus = pb.Start(0, 160, "Status");
            while (!(Position == 160))
            {
                Position = Position + 10;
                lRet = pb.UpdateProgress(Position);
            }
            Position = 0;
        }


        public void cmdStopPB_Click(System.Object sender, System.EventArgs e)
        {
            pb.End();
        }

        public void cmdUpdatePB_Click(System.Object sender, System.EventArgs e)
        {
            Position = Position + 10;
            if ((Position == 160))
                Position = 0;
            lRet = pb.UpdateProgress(Position);
            if (lRet != 2)
            {
                Debug.Print(" Result " + lRet);
            }
            else
            {
                MessageBox.Show(" User pressed Esc to cancel ", " API");
                pb.End();
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            progressBar1.Value = 0;
            label1.Text = "";
            closebttn.Visible = false;
        }
    }
}

