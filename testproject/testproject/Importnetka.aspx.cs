using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Data.OleDb;
using System.Configuration;
using System.Globalization;
using System.Data.SqlTypes;

namespace testproject
{
    public partial class Importnetka : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }
        protected void Button2_Click(object sender, EventArgs e)
        {
            float ID;
            String Case_ID;
            String Created_Date;
            String Created_By;
            String Title;
            String Case_Status;
            String Case_Type;
            String Service_Type;
            String Case_Category;
            String Case_Sub_Category;
            String Engineer;
            String Team;
            String Customer;
            String Region;
            String Site;
            String Contact;
            String Channel;
            String Priority;
            String Response_Overdue;
            String Onsite_Overdue;
            String Resolve_Overdue;
            String Close_Overdue;
            String Response_Duration;
            String Onsite_Duration;
            String Resolve_Duration;
            String Close_Duration;
            float Case_Duration;
            String Response;
            String Onsite;
            String Resolve;
            String Auto_Close;
            String SLA;
            String Resolved_Time;
            float Hour_to_Resolve;
            float Hour_to_Resolve_Pending;
            String Closed_Time;
            float Hour_to_Closed;
            float Hour_to_Closed_Pending;
            String Root_Cause;
            String Resolved_Method;
            float New_to_response;
            float New_to_Assign;
            String Latest_Resolve_to_Close;
            float Latest_Response_to_Close;
            float Agent_UTL_Time;
            float Eng_UTL_Time;

            string path = Path.GetFileName(FileUpload2.FileName);
            path = path.Replace(" ", "");
            FileUpload2.SaveAs(Server.MapPath("~/ExcelFile2/") + path);
            String ExcelPath = Server.MapPath("~/ExcelFile2/") + path;
            OleDbConnection mycon = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ExcelPath + "; Extended Properties=Excel 8.0; Persist Security Info = False");
            mycon.Open();
            OleDbCommand cmd = new OleDbCommand("select * from [Sheet1$]", mycon);
            OleDbDataReader dr = cmd.ExecuteReader();
            while (dr.Read())
            {
                ID = Convert.ToInt32(dr[0].ToString());
                Case_ID = dr[1].ToString();
                Created_Date = dr[2].ToString();
                Created_By = dr[3].ToString();
                Title = dr[4].ToString();
                Case_Status = dr[5].ToString();
                Case_Type = dr[6].ToString();
                Service_Type = dr[7].ToString();
                Case_Category = dr[8].ToString();
                Case_Sub_Category = dr[9].ToString();
                Engineer = dr[10].ToString();
                Team = dr[11].ToString();
                Customer = dr[12].ToString();
                Region = dr[13].ToString();
                Site = dr[14].ToString();
                Contact = dr[15].ToString();
                Channel = dr[16].ToString();
                Priority = dr[17].ToString();
                Response_Overdue = dr[18].ToString();
                Onsite_Overdue = dr[19].ToString();
                Resolve_Overdue = dr[20].ToString();
                Close_Overdue = dr[21].ToString();
                Response_Duration = dr[22].ToString();
                Onsite_Duration = dr[23].ToString();
                Resolve_Duration = dr[24].ToString();
                Close_Duration = dr[25].ToString();
                Case_Duration = Convert.ToInt32(dr[26].ToString());
                Response = dr[27].ToString();
                Onsite = dr[28].ToString();
                Resolve = dr[29].ToString();
                Auto_Close = dr[30].ToString();
                SLA = dr[31].ToString();
                Resolved_Time = dr[32].ToString();
                Hour_to_Resolve = Convert.ToInt32(dr[33].ToString());
                Hour_to_Resolve_Pending = Convert.ToInt32(dr[34].ToString());
                Closed_Time = dr[35].ToString();
                Hour_to_Closed = Convert.ToInt32(dr[36].ToString());
                Hour_to_Closed_Pending = Convert.ToInt32(dr[37].ToString());
                Root_Cause = dr[38].ToString();
                Resolved_Method = dr[39].ToString();
                New_to_response = Convert.ToInt32(dr[40].ToString());
                New_to_Assign = Convert.ToInt32(dr[41].ToString());
                Latest_Resolve_to_Close = dr[42].ToString();
                Latest_Response_to_Close = Convert.ToInt32(dr[43].ToString());
                Agent_UTL_Time = Convert.ToInt32(dr[44].ToString());
                Eng_UTL_Time = Convert.ToInt32(dr[45].ToString());
                savedata(ID, Case_ID, Created_Date, Created_By, Title, Case_Status, Case_Type, Service_Type, Case_Category, Case_Sub_Category, Engineer, Team
                    , Customer, Region, Site, Contact, Channel, Priority, Response_Overdue, Onsite_Overdue, Resolve_Overdue, Close_Overdue, Response_Duration
                    , Onsite_Duration, Resolve_Duration, Close_Duration, Case_Duration, Response, Onsite, Resolve, Auto_Close, SLA, Resolved_Time, Hour_to_Resolve
                    , Hour_to_Resolve_Pending, Closed_Time, Hour_to_Closed, Hour_to_Closed_Pending, Root_Cause, Resolved_Method, New_to_response, New_to_Assign
                    , Latest_Resolve_to_Close, Latest_Response_to_Close, Agent_UTL_Time, Eng_UTL_Time);
            }
            Label2.Text = "Data Has Been Saved Successfully";

        }
        private void savedata(float ID1, String Case_ID1, String Created_Date1, String Created_By1, String Title1, String Case_Status1, String Case_Type1, String Service_Type1, String Case_Category1, String Case_Sub_Category1, String Engineer1, String Team1
                    , String Customer1, String Region1, String Site1, String Contact1, String Channel1, String Priority1, String Response_Overdue1, String Onsite_Overdue1, String Resolve_Overdue1, String Close_Overdue1, String Response_Duration1
                    , String Onsite_Duration1, String Resolve_Duration1, String Close_Duration1, float Case_Duration1, String Response1, String Onsite1, String Resolve1, String Auto_Close1, String SLA1, String Resolved_Time1, float Hour_to_Resolve1
                    , float Hour_to_Resolve_Pending1, String Closed_Time1, float Hour_to_Closed1, float Hour_to_Closed_Pending1, String Root_Cause1, String Resolved_Method1, float New_to_response1, float New_to_Assign1
                    , String Latest_Resolve_to_Close1, float Latest_Response_to_Close1, float Agent_UTL_Time1, float Eng_UTL_Time1)
        {
            String query = "insert into Netka1([ID],[Case ID],[Created Date],[Created By],[Title],[Case Status],[Case Type],[Service Type],[Case Category],[Case Sub Category],[Engineer],[Team]" +
                            ",[Customer],[Region],[Site],[Contact],[Channel],[Priority],[Response Overdue],[Onsite Overdue],[Resolve Overdue],[Close Overdue],[Response Duration],[Onsite Duration]" +
                            ",[Resolve Duration],[Close Duration],[Case Duration],[Response],[Onsite],[Resolve],[Auto Close],[SLA],[Resolved Time],[Hour to Resolve],[Hour to Resolve(Pending)],[Closed Time]" +
                            ",[Hour to Closed],[Hour to Closed(Pending)],[Root Cause],[Resolved Method],[New to response],[New to Assign],[Latest Resolve to Close],[Latest Response to Close],[Agent UTL Time],[Eng UTL Time]) " +
                    "values('" + ID1 + "','" + Case_ID1 + "','" + Created_Date1 + "','" + Created_By1 + "','" + Title1 + "','" + Case_Status1 + "','" + Case_Type1 + "','" + Service_Type1 + "','" + Case_Category1 +
                    "','" + Case_Sub_Category1 + "','" + Engineer1 + "','" + Team1 + "','" + Customer1 + "','" + Region1 + "','" + Site1 + "','" + Contact1 + "','" + Channel1 + "','" + Priority1 + "','"
                    + Response_Overdue1 + "','" + Onsite_Overdue1 + "','" + Resolve_Overdue1 + "','" + Close_Overdue1 + "','" + Response_Duration1 + "','" + Onsite_Duration1 + "','" + Resolve_Duration1 + "','" + Close_Duration1 + "','" + Case_Duration1 + "','"
                    + Response1 + "','" + Onsite1 + "','" + Resolve1 + "','" + Auto_Close1 + "','" + SLA1 + "','" + Resolved_Time1 + "','" + Hour_to_Resolve1 + "','" + Hour_to_Resolve_Pending1 + "','"
                    + Closed_Time1 + "','" + Hour_to_Closed1 + "','" + Hour_to_Closed_Pending1 + "','" + Root_Cause1 + "','" + Resolved_Method1 + "','" + New_to_response1 + "','" + New_to_Assign1 + "','"
                    + Latest_Resolve_to_Close1 + "','" + Latest_Response_to_Close1 + "','" + Agent_UTL_Time1 + "','" + Eng_UTL_Time1 + "')";

            String mycon = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
            SqlConnection con = new SqlConnection(mycon);
            con.Open();
            SqlCommand cmd = new SqlCommand();
            cmd.CommandText = query;
            cmd.Connection = con;
            cmd.ExecuteNonQuery();
            con.Close();


        }
    }
}