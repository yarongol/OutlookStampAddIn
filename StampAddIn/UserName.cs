using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;


namespace StampAddIn
{
    public partial class dlgUserName : Form
    {
        public dlgUserName()
        {
            InitializeComponent();
        }


        private void button2_Click(object sender, EventArgs e)
        {
            ColorDialog colorDlg = new ColorDialog();
            if (colorDlg.ShowDialog() == DialogResult.OK)
            {
                txtUserName.ForeColor = colorDlg.Color;
            }

        }

    }
}