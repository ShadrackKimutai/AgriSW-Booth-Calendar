using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Booth_Calendar.appreusables;
using Booth_Calendar;

namespace Booth_Calendar.appmain
{
    public partial class Boomain : Form
    {
        private int [] activa={0,0,0,0,0};
      //  private int childFormNumber = 0;
        Form Impxls = new appmain.Importxls();

        public Boomain()
        {
            InitializeComponent();
        }

        private void ShowNewForm(object sender, EventArgs e)
        {
            toolStripMenuItem3.PerformClick();
        }

        private void OpenFile(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            openFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = openFileDialog.FileName;
            }
        }

        private void SaveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            saveFileDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*";
            if (saveFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                string FileName = saveFileDialog.FileName;
            }
        }

        private void ExitToolsStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void CutToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void CopyToolStripMenuItem_Click(object sender, EventArgs e)
        {
        }

        private void PasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Appinit.Config k = new Appinit.Config();
            k.Upgrade();
        }

        private void ToolBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            toolStrip.Visible = toolBarToolStripMenuItem.Checked;
        }

        private void StatusBarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            statusStrip.Visible = statusBarToolStripMenuItem.Checked;
        }

        private void CascadeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.Cascade);
        }

        private void TileVerticalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileVertical);
        }

        private void TileHorizontalToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.TileHorizontal);
        }

        private void ArrangeIconsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LayoutMdi(MdiLayout.ArrangeIcons);
        }

        private void CloseAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            foreach (Form childForm in MdiChildren)
            {
                childForm.Close();
            }
        }

        private void menuStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void Boomain_FormClosing(object sender, FormClosingEventArgs e)
        {
            doordie ex = new doordie();
            ex.diestate = true;
            
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            
                Impxls.MdiParent = this;
                Impxls.Show();
                toolStripMenuItem3.Enabled = false;
                newToolStripButton.Enabled = false;
                activa[0] = 1;
            
        }

        private void Boomain_Load(object sender, EventArgs e)
        {

        }

        private void menuAssist_Tick(object sender, EventArgs e)
        {


            if (activa[0] == 0)
            {
                toolStripMenuItem3.Enabled = true;
             
            
            }
               
               

            

        }

        private void importFromSqlServerToolStripMenuItem_Click(object sender, EventArgs e)
        {
          
        }

        private void fileMenu_Click(object sender, EventArgs e)
        {

        }

        private void generateCalendarsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GenerateCal genkal = new GenerateCal();
          genkal.MdiParent = this;
           genkal.Show();
           generateCalendarsToolStripMenuItem.Enabled = false;
           toolStripButton2.Enabled = false;
            activa[3] = 1;
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            generateCalendarsToolStripMenuItem.PerformClick();
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void pdfReaderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            reader readpdf = new reader();
            readpdf.MdiParent = this;
            readpdf.Show();
            toolStripButton1.Enabled = pdfReaderToolStripMenuItem.Enabled = false;
        }
    }
}
