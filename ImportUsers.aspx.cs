using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

public partial class ImportUsers : System.Web.UI.Page
{
    protected void Page_Load(object sender, EventArgs e)
    {
        SqlDataSource ds = new SqlDataSource(ConfigurationManager.ConnectionStrings["CSiSQLExpress"].ConnectionString, "SELECT * from reports.dbo.tblCollisionLogons WHERE Active <> 0");
        OldUsers.DataSource = ds;
        OldUsers.DataBind();

        SqlConnection cn = new SqlConnection(ConfigurationManager.ConnectionStrings["CSiSQLExpress"].ConnectionString);
        SqlDataAdapter da = new SqlDataAdapter("SELECT * from reports.dbo.tblCollisionLogons WHERE Active <> 0", cn);
        DataSet users = new DataSet();
        da.Fill(users);
        foreach (DataRow r in users.Tables[0].Rows)
        {
            Response.Write(r["Logon"]+"<br/>\n");

            String email = Convert.ToString(r["EMail"]);
            if (email.Length == 0)
            {
                email = "dadkins@csicomplete.com";
            }
//            MembershipCreateStatus createStatus;
            try
            {
                //Membership.CreateUser(Convert.ToString(r["Logon"]), Convert.ToString(r["Password"]), email); //,"What is 1+1?","2",true,createStatus);

                // Create an empty Profile for the newly created user
                ProfileCommon p = (ProfileCommon)ProfileCommon.Create((String)r["Logon"], true);

                // Populate some Profile properties off of the create user wizard
                p.ShopID = (Int32)r["ShopID"];
                p.Alias = (String)r["Alias"];

                // Save the profile - must be done since we explicitly created this profile instance
                p.Save();
            }
            catch (Exception ex)
            {
                Response.Write(ex.Message+"<br/>\n");
            }
        }
    }
}