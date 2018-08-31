﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace testproject
{
    public partial class WebForm1 : System.Web.UI.Page
    {
        
        string strConn = ConfigurationManager.ConnectionStrings["dbGINConnectionString"].ConnectionString;
        SqlConnection conn = new SqlConnection();
        SqlCommand sqlQuery = new SqlCommand();
        DataTable dt = new DataTable();
        string sql = "";

        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            String txtDate1 = TextBox1.Text;
            String txtDate2 = TextBox2.Text;

            sql = "select [Circuit ID] as หมายเลขวงจร ";
            sql = sql + ",COALESCE(NULLIF(RIGHT([Subject],LEN([Subject])-charindex(')',[Subject])) , ''),SUBSTRING([Subject],1,case when charindex('(',[Subject]) = 0 then LEN([Subject]) else charindex('(',[Subject])-1 END) ) ";
            sql = sql + ",[SLA]*100  as SLA ";
            sql = sql + ",ISNULL(CONVERT(VARCHAR(10),[Link Down],103),CONVERT(VARCHAR(10),[Create Date],103))  as วันที่แจ้งเหตุ ";
            sql = sql + ",ISNULL(CONVERT(VARCHAR(5),[Link Down (Time)],108),CONVERT(VARCHAR(5),[Create Date],108)) as เวลาแจ้งเหตุ ";
            sql = sql + ",[Help Topic] as ประเภทเหตุขัดข้อง ";
            sql = sql + ",SUBSTRING([สรุปสาเหตุ # วิธีแก้ปัญหา],1,case when charindex('#',[สรุปสาเหตุ # วิธีแก้ปัญหา]) = 0 then LEN([สรุปสาเหตุ # วิธีแก้ปัญหา]) else charindex('#',[สรุปสาเหตุ # วิธีแก้ปัญหา])-1 END) as สาเหตุ ";
            sql = sql + ",RIGHT([สรุปสาเหตุ # วิธีแก้ปัญหา],LEN([สรุปสาเหตุ # วิธีแก้ปัญหา])-charindex('#',[สรุปสาเหตุ # วิธีแก้ปัญหา])) as การแก้ไข ";
            sql = sql + ",ISNULL (CONVERT(VARCHAR(10),[Link Up],103),CONVERT(VARCHAR(10),[Closed Date],103)) as วันที่แก้ไข ";
            sql = sql + ",ISNULL (CONVERT(VARCHAR(5),[Link Up (Time)],108),CONVERT(VARCHAR(5),[Closed Date],108)) as เวลาแก้ไข ";
            sql = sql + ",(select CONVERT(VARCHAR(10),Datediff(ss,a.t1,a.t2)/(60*60*24)) + ' Days ' ";
            sql = sql + "   +CONVERT(VARCHAR(5),DateAdd(SS,Datediff(ss,a.t1, a.t2)%(60*60*24),0),114) ";
            sql = sql + "from ";
            sql = sql + "(select ISNULL (CONVERT(VARCHAR(10),[Link Down],102),CONVERT(VARCHAR(10),[Create Date],102)) + ' ' +ISNULL (CONVERT(VARCHAR(5),[Link Down (Time)],108),CONVERT(VARCHAR(5),[Create Date],108)) t1 ";
            sql = sql + ",ISNULL (CONVERT(VARCHAR(10),[Link Up],102),CONVERT(VARCHAR(10),[Closed Date],102)) + ' ' +ISNULL (CONVERT(VARCHAR(5),[Link Up (Time)],108),CONVERT(VARCHAR(5),[Closed Date],108)) t2 ";
            sql = sql + ") a) as 'ระยะเวลาการแก้ไข' ";
            sql = sql + ",[Ticket Number] as 'OS Ticket Number' ";
            sql = sql + ",CONVERT(VARCHAR(3),DATENAME(dw, ISNULL ([Link Down],[Create Date]))) +'-'+CONVERT(VARCHAR(3),DATENAME(dw, ISNULL ([Link Up],[Closed Date]))) as 'หมายเหตุ' ";
            sql = sql + ",[Root Cause] as 'Root Cause' ";
            sql = sql + ",[Releated Hardware] as Hardware ";
            sql = sql + ",[ปัญหาเกิดที่] as ปัญหาที่เกิด ";
            sql = sql + ",(select (case when[Root Cause] in ('OFC','BGP / Link Flap','Change Config','Network Improvement','Hardware Failure','Lost / Unstable Connection','Hardware Error (Hard - Reset)','Hardware Error (Soft - Reset)') ";
            sql = sql + "and (select CONVERT(VARCHAR(10),Datediff(mi, a.t1, a.t2)) + ' min' from ";
            sql = sql + "(select ISNULL (CONVERT(VARCHAR(10),[Link Down],102),CONVERT(VARCHAR(10),[Create Date],102)) + ' ' +ISNULL (CONVERT(VARCHAR(5),[Link Down (Time)],108),CONVERT(VARCHAR(5),[Create Date],108)) t1 ";
            sql = sql + ",ISNULL (CONVERT(VARCHAR(10),[Link Up],102),CONVERT(VARCHAR(10),[Closed Date],102)) + ' ' +ISNULL (CONVERT(VARCHAR(5),[Link Up (Time)],108),CONVERT(VARCHAR(5),[Closed Date],108)) t2 ";
            sql = sql + ") a) > [min] ";
            sql = sql + "and [เอกสารใบเลื่อน] is null then 'Breach' ";
            sql = sql + "else 'Meet' end)) as 'Breach/Meet' ";
            sql = sql + ",[เอกสารใบเลื่อน]+' '+[OFC ขาดเนื่องจากสาเหตุ] as 'ข้อยกเว้น' ";
            sql = sql + ",[วิเคราะห์ Customer] as 'วิเคราะห์ Customer' ";
            sql = sql + "from [dbGIN].[dbo].[Ostickets] LEFT JOIN[dbGIN].[dbo].[Sla] ON[dbGIN].[dbo].[Sla].slamin = [dbGIN].[dbo].[Ostickets].SLA ";
            sql = sql + "where '" + txtDate1 + "' <= ISNULL(CONVERT(VARCHAR(10),[Link Down],103),CONVERT(VARCHAR(10),[Create Date],103)) and '" + txtDate2 + "' >= ISNULL(CONVERT(VARCHAR(10),[Link Down],103),CONVERT(VARCHAR(10),[Create Date],103)) ";
            sql = sql + "order by [Create Date] ";


            if (conn.State == ConnectionState.Open)
                conn.Close();
            conn.ConnectionString = strConn;
            conn.Open();

            DataTable datTable = new DataTable();
            using (SqlDataAdapter datAdapterTable = new SqlDataAdapter(sql, conn))
            {
                datAdapterTable.Fill(datTable);
            }
            conn.Close();

            GridView1.DataSource = datTable;
            GridView1.DataBind();


        }
    }
}