using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.OleDb;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml;
using Booth_Calendar.Properties;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Booth_Calendar;
using Send2Sql;



namespace Booth_Calendar.appmain
{
    public partial class Importxls : Form
    {
        public DataSet ds = new DataSet();
        public Importxls()
        {
            InitializeComponent();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;
        }

        private void Importxls_Load(object sender, EventArgs e)
        {

            impxls();
            // splitContainer1.Panel2Collapsed = true;


        }

        private int impxls()
        {
            String full = null, fpath = "", Fname = null, xsheet = "", tstring = "";
            String[] HeadARR = { "LocID", "Business", "Address", "Address2", "City", "State", "Zip", "Store (Full Name)", "Contact", "Phone", "Fax", "Email", "Date", "Start Time", "End Time", "Time Allotment (mins)", "Available Slots", "Troop", "Time Slot", "Slot Start", "Slot End", "Send eMail", "Send Fax" };
            int i,  n;
            DataSet emptyda = new DataSet();

        Reshow:
            Openxls.ShowDialog();

            full = Openxls.FileName;
            full.Trim();
            i = (full.Length) - 1;
            while (i > 0)
            {
                if (full.Substring(i, 1) != "\\")
                {
                    tstring = full.Substring(i, 1);
                    //extract the name of the file
                    Fname = tstring + Fname;
                }
                else
                {
                    i = 0;
                }
                i--;
            }
            if ((Fname == null) || (Fname == ""))
            {
                MessageBox.Show("The path does not contain a valid filename. The task is to be repeated", "Import Failure", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                goto Reshow;
            }

            fpath = full.Substring(0, (full.Length - Fname.Length));

        Resheet:
            xsheet = Microsoft.VisualBasic.Interaction.InputBox("Enter Sheet Name", "Sheet Name", "Sheet1");
            if ((xsheet == null) || (xsheet == "")) goto Resheet;
            ds = LoadSheet(fpath, Fname, xsheet);
            if (ds == null)
            {
                MessageBox.Show("Something Went Wrong Somewhere");
                this.Close();
                return 0;
            }

            dataGridView1.DataSource = ds.Tables[0];
            n = dataGridView1.Columns.Count;
            //global templates must have atleast 23 columns   
            if (n != 23)
            {
                dataGridView1.DataSource = emptyda;

                MessageBox.Show("The datasheet you have provided is not in as Defined in the global worksheet template. This may indicate that the current sheet has fewer columns than expected.\n Datasheet has been Discarded", "Datasheet Rejected");

                return 0;
            }

            while (n > 0)
            {
                dataGridView1.Columns[n - 1].HeaderText = HeadARR[n - 1];
                n -= 1;
            }


            return 0;

        }

        public DataSet LoadSheet(String str1, String str2, String str3)
        {
            OleDbCommand cmd;
            OleDbConnection conn;
            OleDbDataAdapter da;
            String conString = null, sqlString = null;
            FileInfo flnfo;
            DataSet datast = new DataSet();
            // in order to connect to older versions of excel i.e excel 2003

            conString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source={0}", str1 + str2) + @";Extended Properties=""Excel 8.0;HDR=Yes;IMEX=1;""";
            sqlString = "SELECT * FROM [" + str3 + "$]";

            flnfo = new FileInfo(str2);
            if (flnfo.Extension.ToLower() == ".xlsx")
            {
                // * to connect to  Excel 2007 or excel 2010
                conString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source={0}", str1 + str2) + @";Extended Properties=""Excel 12.0 Xml;HDR=YES;""";
            }

            conn = new OleDbConnection(conString);
            try
            {
                conn.Open();
                cmd = new OleDbCommand();
                cmd.Connection = conn;
                cmd.CommandText = sqlString;
                da = new OleDbDataAdapter();
                da.SelectCommand = cmd;
                da.Fill(datast);


            }
            catch (Exception ex)
            {

                MessageBox.Show("Something went wrong somewhere"+ ex.Message , "Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
                str1 = str2 = str3 = null;
                datast = null;
            }
            finally
            {
                conn.Close();


            }




            return datast;

        }
        private void checkdatetime()
        {




        }


        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void somethingToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void exportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (exporter.IsBusy != true)
            {
                exporter.RunWorkerAsync();
            }
            else
            {
                MessageBox.Show("Have some patience. I am already processing it anyway.", "Patience Please", MessageBoxButtons.OK, MessageBoxIcon.Information);
            
            
            }

        }
        private String stringfix(String sender)
        {
            int strL, m;
            String tempstr, tempsbstr;
            tempstr = sender;
            sender = null;
            strL = tempstr.Length;
            //tempstr = String.Replace(tempstr, "'", "`", , , CompareMethod.Text);
            m = strL - 1;
            while (m > 0)
            {
                if (tempstr.Substring(m, 1) != "'")
                {
                    tempsbstr = tempstr.Substring(m, 1);
                    //extract the name of the file
                    sender = tempsbstr + sender;
                }
                else
                {
                    sender = "`" + sender;

                }
                m -= 1;

            }


            return sender;
        }
        private void exporter_DoWork(object sender, DoWorkEventArgs e)
        {

            Send2sql k = new Send2sql();
            int i, j, rowc;
            SqlCommand cmd = new SqlCommand();
            SqlConnection conn = new SqlConnection();
            String tempstr;
            String[] clin = { null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null, null };
            rowc = dataGridView1.RowCount;

            j = 0;
            k.connectionstring = Settings.Default.cnn;
            while (j <= (rowc - 2))
            {
                i = 0;

                while (i <= 22)
                {
                    tempstr = dataGridView1.Rows[j].Cells[i].Value.ToString();
                    //' stringfix(tempstr);
                    k.ContentArray[i] = tempstr;
                    //  exporter.ReportProgress(i);
                    //MessageBox.Show(tempstr);

                    i += 1;
                    exporter.ReportProgress(i*4);
                }
                k.Sendthere();
                j += 1;
            }
            exporter.ReportProgress(100);
        }



        private void exporter_ProgressChanged(Object sender, ProgressChangedEventArgs e)
        {
            toolStripProgressBar1.Value = e.ProgressPercentage;
            
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            splitContainer1.Panel2Collapsed = true;
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            dataGridView1.CurrentRow.Selected = true;
        }

        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {

            dataGridView1.CurrentRow.Selected = true;
        }
    }
}
