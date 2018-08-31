using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
using System.Configuration;

namespace testproject
{
    
    public partial class Testdate : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void Button1_Click(object sender, EventArgs e)
        {

            DateTime Date;
            String Name;
           

            string path = Path.GetFileName(FileUpload1.FileName);
            path = path.Replace(" ", "");
            FileUpload1.SaveAs(Server.MapPath("~/ExcelFile/") + path);
            String ExcelPath = Server.MapPath("~/ExcelFile/") + path;
            OleDbConnection mycon = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ExcelPath + "; Extended Properties=Excel 8.0; Persist Security Info = False");
            mycon.Open();
            OleDbCommand cmd = new OleDbCommand("select * from [Sheet1$]", mycon);
            OleDbDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {

                
                
                Date = Convert.ToDateTime(dr[0].ToString());
                Name = dr[1].ToString();
                
                savedata(Date, Name);


            }
            Label1.Text = "Data Has Been Saved Successfully";

        }
        private void savedata(DateTime Date1, String Name1)
        {
            String query = "insert into Table(Date, name) values('" + Date1 + "','" + Name1 + "')";
            //String mycon = "Data Source=localhost\sqlexpress;Initial Catalog=dbGIN;Integrated Security=True";
            String mycon = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
            SqlConnection con = new SqlConnection(mycon);
            con.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = query;
            cmd.Connection = con;
            cmd.ExecuteNonQuery();
        }

    }
}
    
