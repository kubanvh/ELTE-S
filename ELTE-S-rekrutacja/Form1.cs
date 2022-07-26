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
using System.Data.SqlClient;

namespace ELTE_S_rekrutacja
{
    public partial class Form1 : Form
    {
        SqlConnection con;
        SqlCommand cmd;
        SqlDataReader dr;
        ConnectionDB db = new ConnectionDB();
        public Form1()
        {
            InitializeComponent();

            con = new SqlConnection(db.GetConnection());
            Loadrecords();
        }

        public void Loadrecords()
        {
            dgv.Rows.Clear();
            int i = 0;
            con.Open();
            cmd = new SqlCommand("SELECT * FROM tblELTES", con);
            dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                i++;
                dgv.Rows.Add(i, dr["id"].ToString(),
                    dr["zdjęcie1"].ToString(),
                    dr["zdjęcie2"].ToString(),
                    dr["Kod"].ToString(),
                    dr["nazwa"].ToString(),
                    dr["EAN"].ToString(),
                    dr["producent"].ToString(),
                    dr["atrybut_stan"].ToString(),
                    dr["Vat"].ToString(),
                    dr["waga"].ToString(),
                    dr["opis"].ToString(),
                    dr["atrybut_min"].ToString(),
                    dr["nrkatalogowy"].ToString()
                    );
            }
            dr.Close();
            con.Close();
        }

        private void btnImport_Click(object sender, EventArgs e)
        {
            dgv.Rows.Clear();

            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkbook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorksheet;
            Microsoft.Office.Interop.Excel.Range xlRange;

            int xlRow;
            string strFileName;

            openFD.Filter = "Excel Office |*.xls; *xlsx";
            openFD.ShowDialog();
            strFileName = openFD.FileName;

            if(strFileName != "")
            {
                xlApp = new Microsoft.Office.Interop.Excel.Application();
                xlWorkbook = xlApp.Workbooks.Open(strFileName);
                xlWorksheet = xlWorkbook.Worksheets["cennikwzór"];
                xlRange = xlWorksheet.UsedRange;

                int i = 0;

                for (xlRow = 2; xlRow <= xlRange.Rows.Count; xlRow++)
                {
                    if (xlRange.Cells[xlRow, 1].Text != "")
                    {
                        i++;
                        dgv.Rows.Add(i, xlRange.Cells[xlRow, 1].Text,
                            xlRange.Cells[xlRow, 2].Text,
                            xlRange.Cells[xlRow, 3].Text,
                            xlRange.Cells[xlRow, 4].Text,
                            xlRange.Cells[xlRow, 5].Text,
                            xlRange.Cells[xlRow, 6].Text,
                            xlRange.Cells[xlRow, 7].Text,
                            xlRange.Cells[xlRow, 8].Text,
                            xlRange.Cells[xlRow, 9].Text,
                            xlRange.Cells[xlRow, 10].Text,
                            xlRange.Cells[xlRow, 11].Text,
                            xlRange.Cells[xlRow, 12].Text);
                    }
                }
                xlWorkbook.Close();
                xlApp.Quit();
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            for(int i = 0; i <dgv.Rows.Count; i++)
            {
                con.Open();
                cmd = new SqlCommand("INSERT INTO tblELTES (id, zdjęcie1, zdjęcie2, Kod, nazwa, EAN, producent, atrybut_stan, Vat, waga, opis, atrybut_min, nrkatalogowy) VALUES(@id, @zdjęcie1, @zdjęcie2, @Kod, @nazwa, @EAN, @producent, @atrybut_stan, @Vat, @waga, @opis, @atrybut_min, @nrkatalogowy);", con);
                cmd.Parameters.AddWithValue("@id", dgv.Rows[i].Cells[0].Value.ToString());
                cmd.Parameters.AddWithValue("@zdjęcie1", dgv.Rows[i].Cells[1].Value.ToString());
                cmd.Parameters.AddWithValue("@zdjęcie2", dgv.Rows[i].Cells[2].Value.ToString());
                cmd.Parameters.AddWithValue("@Kod", dgv.Rows[i].Cells[3].Value.ToString());
                cmd.Parameters.AddWithValue("@nazwa", dgv.Rows[i].Cells[4].Value.ToString());
                cmd.Parameters.AddWithValue("@EAN", dgv.Rows[i].Cells[5].Value.ToString());
                cmd.Parameters.AddWithValue("@producent", dgv.Rows[i].Cells[6].Value.ToString());
                cmd.Parameters.AddWithValue("@atrybut_stan", dgv.Rows[i].Cells[7].Value.ToString());
                cmd.Parameters.AddWithValue("@Vat", dgv.Rows[i].Cells[8].Value.ToString());
                cmd.Parameters.AddWithValue("@waga", dgv.Rows[i].Cells[9].Value.ToString());
                cmd.Parameters.AddWithValue("@opis", dgv.Rows[i].Cells[10].Value.ToString());
                cmd.Parameters.AddWithValue("@atrybut_min", dgv.Rows[i].Cells[11].Value.ToString());
                cmd.Parameters.AddWithValue("@nrkatalogowy", dgv.Rows[i].Cells[12].Value.ToString());
                cmd.ExecuteNonQuery();
                con.Close();
                MessageBox.Show("Rekordy poprawnie zapisane", "WIADOMOSC", MessageBoxButtons.OK, MessageBoxIcon.Information);
                Loadrecords();
            }
        }
    }
}
