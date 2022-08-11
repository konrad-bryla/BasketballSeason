using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace konradBryla_A12
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the '_2018_SeasonDataSet1.Season_2018' table. You can move, or remove it, as needed.
            this.season_2018TableAdapter.Fill(this._2018_SeasonDataSet1.Season_2018);

        }

        private void doExcel_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            application.Visible = true;

            Microsoft.Office.Interop.Excel.Worksheet worksheet = new Microsoft.Office.Interop.Excel.Worksheet();
            worksheet = application.Workbooks.Add().Worksheets[1];

            worksheet.Cells[1, 1] = "2018 Basketball Season Excel Sheet" + " " +DateTime.Today.ToString();

            //Title properties
            {
                var title = worksheet.Range["A1:B1"].Font;
                title.Size = 20;
                title.Bold = true;
                title.Color = Color.HotPink;
                
            }
            worksheet.Range["A1:K1"].Merge();
            //header properties
            {
                var header = worksheet.PageSetup;
                header.PrintTitleRows = "A3:G3";
                header.CenterHorizontally = true;
                header.CenterHeader = "༼ つ ◕_◕ ༽つ";
                header.CenterFooter = "Games from " + dtpStartDate.Value.ToString() + "to " + dtpEndDate.Value.ToString() + ", By Konrad Bryla.";
            }


            worksheet.Cells[3, 1] = "Line #";
            worksheet.Cells[3, 2] = "Date";
            worksheet.Cells[3, 3] = "Home Team";
            worksheet.Cells[3, 4] = "Away Team";
            worksheet.Cells[3, 5] = "Score";

            int gamesPlayed = 0;
            int count = 0;
            
            foreach(DataRowView row in bindingSource1.List)
            {
                if (((DateTime)row["Game Date"] > (DateTime)dtpStartDate.Value) && ((DateTime)row["Game Date"] < (DateTime)dtpEndDate.Value))
                {
                    worksheet.Cells[4 + count, 1] = "Line " + count;
                    worksheet.Cells[4 + count, 2] = (DateTime)row["Game Date"];
                    worksheet.Cells[4 + count, 3] = (String)row["Home Team"];
                    worksheet.Cells[4 + count, 4] = (String)row["Away Team"];
                    worksheet.Cells[4 + count, 5] = (String)row["Home score"].ToString() + " to " + (String)row["Away Score"].ToString();
                    ++count;
                    ++gamesPlayed;
                }
            }

            worksheet.Cells[5 + gamesPlayed, 1] = "The total number of games is: " + gamesPlayed.ToString();
            worksheet.Range["A:Z"].Columns.AutoFit();
        }
    }
}
