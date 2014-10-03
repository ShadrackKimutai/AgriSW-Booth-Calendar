using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.IO;
using System.Net;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Booth_Calendar.Properties;
using PdfSharp;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using Formater;
using Send2Sql;

namespace Booth_Calendar.appmain
{
    public partial class GenerateCal : Form
    {
        private int spst = 0;
        public GenerateCal()
        {
            InitializeComponent();
        
        }
        private void GenScal()
        { 
            
            DateTime day1,day2;
            day1=new DateTime(2012,8,28) ;
            day2=new DateTime(2012,11,24);
            DataSet tempDS=new DataSet() ;
            PdfDocument docu = new PdfDocument();
            PdfPage Page =docu.AddPage();
             Page.Size = PageSize.A4;
             Page.Orientation = PageOrientation.Landscape;
             XGraphics gfx;
             XFont f1ss,f1, f2, f3, f4,f5,f6;
             XPen t, tt, ttt;
             XRect topR, dayR;
             int int1,ddf, int2, int3,PageN=3,weekL,CurrP,Header=0,int5,int6;
             TimeSpan TSpan;
             Send2Sql.Send2sql N=new Send2sql();
             String x1, x2, x3,coverpage;
             String [] Weekdays = { "Sun", "Mon", "Tue", "Wed", "Thur", "Fri", "Sat" };
            String [] ScheduleTimes =new String[14];
            coverpage="";

             //XPdfForm form = XPdfForm.FromFile(coverpage);
             // Determine width and height
             //double extWidth = form.PixelWidth;
            // double extHeight = form.PixelHeight;
            f1ss = new XFont("Tahoma", 6, XFontStyle.Regular);
             f1 = new XFont("Tahoma", 7, XFontStyle.Regular);
             f2 = new XFont("Arial", 10, XFontStyle.Regular);
             f3 = new XFont("Arial", 11, XFontStyle.Bold);
             f4 = new XFont("Arial", 12, XFontStyle.Regular);
             f5 = new XFont("Arial",12,XFontStyle.Bold);
             f6 = new XFont("Arial",16,XFontStyle.Bold);
             t = new XPen(XColor.FromArgb(0, 0, 0));
             tt = new XPen(XColor.FromArgb(0,0,0),1);
             ttt = new XPen(XColor.FromArgb(230, 0, 0));
            Formater.Formating K=new Formater.Formating();
            ddf = 0;
             
            K.StartEndDates(ref day2, ref day1, ref ddf);

        
            
               
            if (ddf > 28)
             {
                 PageN = ddf / 28;
             }

             weekL = Weekdays.Length;
              gfx = XGraphics.FromPdfPage(Page);

              topR = new XRect(10, 10, 822, 40);
              dayR = new XRect(0, 0, 0, 0);
           
              for (CurrP = 0; CurrP <= PageN; CurrP++)
              {
                  gfx.DrawRectangle(t, topR);
                  gfx.DrawString("Haven", f3, XBrushes.Black, 15, 33);
                  gfx.DrawString("21st Street, NewPort", f3, XBrushes.Black, 14, 45);
                  gfx.DrawString("COOKIE BOOTH CALENDAR", f6, XBrushes.Black, 297, 25);
                  gfx.DrawString("Girl Scouts Of Northeast Texas", f3, XBrushes.Black, 665, 33);
                  gfx.DrawString("24-Hour Booth Sale Hotline, 254-XXX-XXX-XXX", f3, XBrushes.Black, 584, 45);
                  XImage letterhead  = Properties.Resources.letterhead;
                  gfx.DrawImage(letterhead, 320, 52);
                  gfx.DrawString("1912 - " + DateTime.Now.Year.ToString(), f5, XBrushes.Black, 351, 114);
                  int1 = 50;
                  int2 = 150;
              for (Header = 0; Header <= (weekL - 1); Header++)
              {
                  gfx.DrawString(Weekdays[Header], f3, XBrushes.Black, int1 + 10, 145);
                  int1 = int1 + 117;
              }

              int1 = 11;

              while (int2 < 505)
              {
                  while (int1 < 805)
                  {
                      dayR = new XRect(int1, int2,117, 100);
                      gfx.DrawRectangle(t, dayR);
                      x1=null;
                      x2=null;
                      parsedate(day1, ref x1,ref x2 );
                      gfx.DrawString("" + day1.Day.ToString() + "" + x1 + " " + x2.ToUpper().ToString() + "", f3, XBrushes.Black, int1 + 58, int2 + 12);
                      int6=int2 + 15;
                      tempDS = null;
                      //for (int5 = 0; int5 <= 12; int5++)
                      //{ 
                      //    gfx.DrawString("" +int5 +" Scouts", f1, XBrushes.Black, (int1+ 5), (int6+4 ));
                      //      int6 = int6 + 8;
                      //}

                      tempDS = fetchdate(Settings.Default.cnn, day1.Date, "1");
                     
                      MessageBox.Show(tempK.ToString());
                      if (tempK == 0)
                      {
                      }
                      else 
                      {
                          for (int5 = 0; int5 <= tempK-1; int5++)
                          {
                              gfx.DrawString(tempDS.Tables[0].Rows[int5].ItemArray[0].ToString(), f1, XBrushes.Black, (int1 + 5), (int6 + 4));
                              int6 = int6 + 8;
                          }


                      }


                      int1 = int1 + 117;
                      
                          day1 = day1.AddDays(1);
                         
                      
                  } 
                  if (day1.Date >= day2.Date)
                          {
                      
                              
                      break;
                          }
                  int1 = 11;
                  int2 = int2 + 105;
                 }

              if (PageN > CurrP)
              {
                  Page = docu.AddPage();
                  Page.Size = PageSize.A4;
                  Page.Orientation = PageOrientation.Landscape;
                  
                  gfx = XGraphics.FromPdfPage(Page);

                  int1 = 50;
                  int2 = 150;
           
              
              }
              


              }
            long k = DateTime.Now.Ticks;
              docu.Save("temp" + k.ToString() + ".pdf");
              System.Diagnostics.Process.Start("temp" + k.ToString() + ".pdf");
       
        }
        private DataSet fetchdate( String cnn,DateTime Day1,String SID)
        {
            DataSet ds=new DataSet();
          cnn= Settings.Default.cnn;
          SID = "1";
          int a1=0;
          SqlDataAdapter Da = new SqlDataAdapter("select SchTmslot from SCHEDULE where SchDate='" + Day1.Date.ToShortDateString() +"' and schlocid=" + SID + " ", cnn);
          Da.Fill(ds, "SCHEDULE");
          a1 = ds.Tables[0].Rows.Count;
            if (a1>0)
            {
            return ds;
            }
            else
            {
             return null;
            }
        }

        
        
       
        private void GenCal(string p1, DateTime p2, DateTime p3)
        {

     //             int ak, rc;
     //   SqlDataAdapter da;
     //   DataSet ds ;

     //  PdfDocument docu  = new PdfDocument();
     //  // PageSize pageSz;  
     //   //pageSz = PageSize.A4;
     //  PdfPage page = docu.AddPage();
     //XGraphics gfx; 
     //  XFont hfont, hhfont, smfont, tfont, regfont;
     //  XRect trect, drect;
     //   XPen tpen, thpen, rpen; 
     //   tpen= new PdfSharp.Drawing.XPen(XColor.FromArgb(0, 0, 0), 2);

     //   thpen = new PdfSharp.Drawing.XPen(XColor.FromArgb(0, 0, 0));
     //   rpen = new PdfSharp.Drawing.XPen(XColor.FromArgb(230, 0, 0));
     //   hfont = new XFont("Arial", 11, XFontStyle.Bold);
     //   hhfont = new XFont("Arial", 15, XFontStyle.Bold);
     //   regfont = new XFont("Arial", 11, XFontStyle.Regular);
     //   tfont = new XFont("Arial", 12, XFontStyle.Regular);
     //   smfont = new XFont("Tahoma", 7, XFontStyle.Regular);
     //   String k, s=null, t=null;
     //   k = DateAndTime.Now.Ticks.ToString();
     //  //page dimensions(5, 5, 783, 602) with 4' margin//
     //  /* End of Part 1 */
     //      int x=0, y=0, Weekl, headers, dd, pageN=0, xamm, tempy, retschid;
     //  String [] weekdays = {"Sun", "Mon", "Tue", "Wed", "Thur", "Fri", "Sat"};
     //   //Ensure the week start on sunday
     //   while (p2.DayOfWeek != 0)
     //   { 
     //       p2 = p2.AddDays(-1);
     //   }

     //   while (p3.DayOfWeek != 0)
     //   {
     //       p3 = p3.AddDays(1);
     //   }
                   
     //   dd =(int) DateAndTime.DateDiff(DateInterval.Day, p2, p3);
     //   if (dd > 28) 
     //   {
     //       pageN = dd / 28;
     //   }

     //   retschid = 0;
     //   retrieveSchid(retschid);
     //   x = 50;
     //   y = 150;


     //   Weekl = weekdays.Length;
     //   gfx = XGraphics.FromPdfPage(page);

     //  // '<Top rectangle>
     //   trect = new XRect(10, 10, 773, 40);

     //   xamm = 0;
     //for (xamm = 0 ;xamm == pageN ;xamm++)
     //{
     //    gfx.DrawRectangle(tpen, trect);
     //       gfx.DrawString("Textbox1", hfont, XBrushes.Black, 15, 33);
     //       gfx.DrawString("TextBox2", hfont, XBrushes.Black, 15, 45);
     //       gfx.DrawString("COOKIE BOOTH CALENDAR", hhfont, XBrushes.Black, 297, 25);
     //       gfx.DrawString("Girl Scouts Of Northeast Texas", hfont, XBrushes.Black, 575, 33);
     //       gfx.DrawString("24-Hour Booth Sale Hotline, 254-XXX-XXX-XXX", regfont, XBrushes.Black, 560, 45);
     //     //  '</top Rectangle>
           

     //       XImage letterhead  = Resources.letterhead;
     //       //gfx.DrawImage(letterhead, 320, 52);
     //       gfx.DrawString("1912 - " + DateTime.Now.Year.ToString(), hhfont, XBrushes.Black, 351, 118);
           
     //       for (headers = 0; headers == (Weekl - 1);headers++)
     //       { 
     //           gfx.DrawString(weekdays[headers], tfont, XBrushes.Black,( x + 35), 145);
     //           x = x + 100;
     //       }

     //       x = 50;
     //    /*
     //     * End of section two 
     //     * 
     //     * 
     //     *
     //     * */
     //    ds=new DataSet();
     //    while (y < 600)
     //    {
     //        while (x < 730)
     //        {
     //               //'rect
     //               drect = new XRect(x, y, 100, 110);
     //               gfx.DrawRectangle(thpen, drect);

     //               //'date
     //               parsedate(p2, ref s, ref t);
     //               gfx.DrawString(""+ (String)p2.Day.ToString() + "" + s +" " + t.ToUpper().ToString()+ "", hfont, XBrushes.Black, x + 28, y + 12);

     //               //Sch

     //               da = new SqlDataAdapter("select SchTmslot from SCHEDULE where SchDate='" + p2.Date.ToShortDateString() + "' and schlocid=" + retschid + " ", Settings.Default.cnn);
     //               da.Fill(ds, "SCHEDULE");
     //               rc = ds.Tables[0].Rows.Count;
     //            //   'MsgBox(rc)
     //                      if (rc == 0) 
     //                         { 
     //                               ds.Tables[0].Clear();
     //                         }
     //                        else
     //                       {
     //                               tempy = y + 15;
     //                                for( ak = 0 ; ak==(rc - 1);ak++)
     //                                     {
                                                              
     //                                              gfx.DrawString("" + ds.Tables[0].Rows[ak].ItemArray[0].ToString() +" Scouts", smfont, XBrushes.Black, x + 5, tempy + 10);
     //                                               tempy = tempy + 10;
     //                                     }

                  
     //                             ds.Tables[0].Clear();
     //                             p2 = p2.Date.AddDays(1);
     //                             x = x + 100;
     //                      }
     //        }
     //           if (p2 == p3) 
     //               break;
     //           else
     //           {
     //               x = 50;
     //               y = y + 115;
     //           }
     //       }
     //       if(pageN > xamm) 
     //       {
     //           page = docu.AddPage();
     //           gfx = XGraphics.FromPdfPage(page);
          
     //           x = 50;
     //           y = 150;
                
     //       }

        

     //   }
     //       docu.Save("tempdks"  + ".pdf");
     //       Process.Start("tempdks" + ".pdf");
        
        
        }
        private void parsedate(DateTime p1, ref String  xda,ref String xmo)
        {
            int em = (int)p1.Month;
            switch (em)
            {
                case 1:
                    xmo = "Jan"; break;

                case 2:
                    xmo = "Feb"; break;
                case 3:
                    xmo = "Mar"; break;
                case 4:
                    xmo = "Apr"; break;
                case 5:
                    xmo = "May"; break;
                case 6:
                    xmo = "Jun"; break;
                case 7:
                    xmo = "Jul"; break;
                case 8:
                    xmo = "Aug"; break;
                case 9:
                    xmo = "Sep"; break;
                case 10:
                    xmo = "Oct"; break;
                case 11:
                    xmo = "Nov"; break;
                case 12:
                    xmo = "Dec"; break;
            }
            em = (int)p1.Date.Day;
            switch (em)
            {
                case 1:

                    xda = "st";
                    return;
                case 21:
                    xda = "st";
                    return;
                case 31:
                    xda = "st";
                    return;
                case 2:
                    xda = "nd";
                    return;
                case 22:
                    xda = "nd";
                    return;
                case 3:
                    xda = "rd";
                    return;
                case 23:
                    xda = "rd";
                    return;
                default:
                    xda = "th";
                    return;
            }
        }
        private long retrieveSchid(long v1)
        {
            String V2 = null;
            SqlDataAdapter da = new SqlDataAdapter("select LocID from Locations where Locname='" + comboBox1.Text + "'", Settings.Default.cnn);
            DataSet ds = new DataSet();

            da.Fill(ds, "Locations");
            V2 = ds.Tables[0].Rows[0].ItemArray[0].ToString();
            v1 = long.Parse(V2);
            //

            return v1;
        }
        private void tabPage1_Click(object sender, EventArgs e)
        {
            //   String [] Item;

        }

        private void button1_Click(object sender, EventArgs e)
        {

            DateTime Day1 = DateTime.Now.Date;
            DateTime Day2 = DateTime.Now.Date;
            String Scmb1;
            String day1, day2;
            #region gettingdates
            Scmb1 = "SELECT  MIN(SchDate) AS Expr1, MAX(SchDate) AS Expr2 FROM  SCHEDULE WHERE (SchLocID LIKE  (SELECT  LocID  FROM LOCATIONS WHERE (LocName = '" + comboBox1.Text + "')))";

            SqlDataAdapter da = new SqlDataAdapter(Scmb1, Settings.Default.cnn);
            DataSet ds = new DataSet();
            da.Fill(ds, "SCHEDULE");
            day1 = ds.Tables[0].Rows[0].ItemArray[0].ToString();//day1.AddDays(-180);
            day2 = ds.Tables[0].Rows[0].ItemArray[1].ToString();//day1.AddDays(180);
             day1 = DateTime.Parse(day1).ToShortDateString();
            day2 = DateTime.Parse(day2).ToShortDateString();
            Day1 = DateTime.Parse(day1);
            Day2 = DateTime.Parse(day2);
            #endregion
            GenCal(comboBox1.Text, Day1, Day2);
        }

        private void GenerateCal_Load(object sender, EventArgs e)
        {
            filltabs();
            fillgrid();
            splitContainer2.Panel2Collapsed = true;
        }

        private void filltabs()
        {
            dateTimePicker2.Enabled = false;
            String CString = Settings.Default.cnn;
            SqlConnection conn = new SqlConnection(CString);
            SqlCommand cmd = new SqlCommand();
            DataSet ds;
            SqlDataAdapter da;
            SqlDataAdapter Reader;
            String Schlb1, Scmb1, Schlb2, Schlb3, Schlb4;
            Schlb1 = "Select DISTINCT LocName from LOCATIONS";

            #region ChecklistBx
            //checklist1
            conn.Open();
            Reader = new SqlDataAdapter(Schlb1, conn);
            ds = new DataSet();
            Reader.Fill(ds, "LOCATIONS");

            checkedListBox1.DataSource = ds.Tables[0];
            checkedListBox1.DisplayMember = "LocName";
            conn.Close();
            //checklist2
            conn.Open();
            Schlb2 = "Select DISTINCT LocCity from LOCATIONS";
            Reader = new SqlDataAdapter(Schlb2, conn);
            ds = new DataSet();
            Reader.Fill(ds, "LOCATIONS");

            checkedListBox2.DataSource = ds.Tables[0];
            checkedListBox2.DisplayMember = "LocCity";
            conn.Close();

            //checklist3
            conn.Open();
            Schlb3 = "Select DISTINCT LocState from LOCATIONS";
            Reader = new SqlDataAdapter(Schlb3, conn);
            ds = new DataSet();
            Reader.Fill(ds, "LOCATIONS");

            checkedListBox3.DataSource = ds.Tables[0];
            checkedListBox3.DisplayMember = "LocState";
            conn.Close();

            //checklist4
            conn.Open();
            Schlb4 = "Select DISTINCT SchTroop from SCHEDULE";
            Reader = new SqlDataAdapter(Schlb4, conn);
            ds = new DataSet();
            Reader.Fill(ds, "SCHEDULE");

            checkedListBox4.DataSource = ds.Tables[0];
            checkedListBox4.DisplayMember = "SchTroop";
            conn.Close();
            #endregion
            #region Combo
            Scmb1 = "Select DISTINCT LocName from LOCATIONS";
            da = new SqlDataAdapter(Scmb1, conn);
            ds = new DataSet();
            da.Fill(ds, "LOCATION");
            comboBox1.DataSource = ds.Tables["LOCATION"];
            comboBox1.DisplayMember = "LocName";
            comboBox1.ValueMember = "LocName";
            #endregion


        }
        private void fillgrid()
        {

            SqlDataAdapter da = new SqlDataAdapter("SELECT LocSendEmail,LocSendFax,LocName,LocAddr,LocCity FROM LOCATIONS", Settings.Default.cnn);
            DataSet ds = new DataSet();
            int a, b, x, y, y1;
            string Tempstring = null;
            da.Fill(ds, "LOCATIONS");
            a = ds.Tables[0].Rows.Count;
            b = ds.Tables[0].Columns.Count;
            x = y = 0;

            while (x <= (a - 1))
            {
                dataGridView1.Rows.Add(1);
                y = 0;
                while (y <= b)
                {
                    y1 = y + 1;
                    if (y == 0)
                    {
                        dataGridView1.Rows[x].Cells[y].Value = true;
                    }
                    if ((y > 0) && (y < 3))
                    {
                        Tempstring = ds.Tables[0].Rows[x].ItemArray[y - 1].ToString();

                        Tempstring = Tempstring.ToLower();

                        if (Tempstring.Contains("yes"))
                        {
                            dataGridView1.Rows[x].Cells[y].Value = true;

                        }
                        else
                        {

                            dataGridView1.Rows[x].Cells[y].Value = false;
                        }


                    }
                    if ((y > 2) && (y < y1))
                    {
                        dataGridView1.Rows[x].Cells[y].Value = ds.Tables[0].Rows[x].ItemArray[y - 1].ToString();

                    }
                    y += 1;

                }
                x += 1;
            }






        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            DateTime D1;

            D1 = dateTimePicker1.Value;
            dateTimePicker2.MinDate = D1;
            dateTimePicker2.Value = D1.AddDays(1);


        }

        private void splitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;
        }

        private void dataGridView1_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;
        }

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView1.CurrentRow.Selected = true;
        }
        private bool showhide()
        {
            int h, i, j;
            j = 0;
            i = 100;
            h = splitContainer2.Height;
            if ((splitContainer2.Panel2Collapsed == true) && (spst == 0))
            {
                splitContainer2.Panel2Collapsed = false;
                splitContainer2.SplitterDistance = h;
                while (i > j)
                {
                    splitContainer2.SplitterIncrement = 1;

                    j += 1;
                }
            }
            //else
            //{

            //    splitContainer2.SplitterDistance = h;
            //    while (i > j)
            //    {
            //        splitContainer2.SplitterIncrement = -1;

            //        j += 1;
            //    }
            //    splitContainer2.Panel2Collapsed = true;
            //}

            return true;

        }

        private void dataGridView1_MouseLeave(object sender, EventArgs e)
        {
            spst = 0;
        }

        private void dataGridView1_MouseHover(object sender, EventArgs e)
        {
            spst = 1;
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            
           
        }

        private void splitContainer2_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

       
        private void button3_Click_1(object sender, EventArgs e)
        {
GenScal();
        }

        private void button2_Click(object sender, EventArgs e)
        {

        }

    }
}
