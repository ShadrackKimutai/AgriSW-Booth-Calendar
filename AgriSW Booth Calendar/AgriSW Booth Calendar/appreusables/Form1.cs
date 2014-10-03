using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Booth_Calendar.Appinit;
using Booth_Calendar;

namespace Booth_Calendar
{
    public partial class Form1 : Form
    {
       public int activa;
        doordie k= new doordie();
        public Form1()
        {
            InitializeComponent();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {
           
        }

        private void initialiser_Tick(object sender, EventArgs e)
        {
            if (progressBar1.Value != 100)
            {
                if (activa == 0)
                {

                    {
                        progressBar1.PerformStep();

                    }
                }

            }
            else
            {
                Form mainP = new appmain.Boomain();
                activa = 1;
                mainP.Show();
                initialiser.Enabled = false;
                this.Hide();


            }

            if (k.diestate == true)
            {
                this.Close();
            
            
            
            }


        }
    }
}
