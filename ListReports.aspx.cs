using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;
using System.Security.Principal;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.WebControls;

public class LocalFile
{
    private String _name;
    private String _path;

    public LocalFile() { }
    public LocalFile(String name, String path)
    {
        _name = name;
        _path = path;
    }

    public String Name {
        get { return _name; }
        set { _name = value; }
    }
    public String Path {
        get { return _path; }
        set { _path = value; }
    }
}

public partial class Reports_ListReports : System.Web.UI.Page
{
    protected void Page_Init(object sender, EventArgs e)
    {
        if (!IsPostBack)
        {
            username.Text = Profile.GetProfile(HttpContext.Current.User.Identity.Name).Alias;
            
            ReportDocument doc = new ReportDocument();
            Dictionary<String, List<LocalFile>> folderlist = new Dictionary<string, List<LocalFile>>();
            String[] files = Directory.GetFiles(Server.MapPath("/Reports"), "survey*", SearchOption.AllDirectories);
            for (int i = 0; i < files.Length; i++)
            {
                if (!folderlist.ContainsKey("Survey"))
                {
                    folderlist["Survey"] = new List<LocalFile>();
                }
                String path = ReverseMapPath(files[i]);
                    LocalFile lf = new LocalFile();
                    lf.Name = "Survey";
                    lf.Path = path;
                    folderlist["Survey"].Add(lf);
            }
            String[] rpts = Directory.GetFiles(Server.MapPath("/Reports"), "*.rpt", SearchOption.AllDirectories);
            for (int i = 0; i < rpts.Length; i++)
            {
                String path = ReverseMapPath(rpts[i]);
                String[] folders = path.Split('/');
                String shoplogin = folders[folders.Length - 2];
                if (!folderlist.ContainsKey(shoplogin))
                {
                    folderlist[shoplogin] = new List<LocalFile>();
                }

                IPrincipal principal = HttpContext.Current.User;
                bool requiredAuthentication = UrlAuthorizationModule.CheckUrlAccessForPrincipal(path, principal, Request.HttpMethod);
                if (requiredAuthentication)
                {
                    LocalFile lf = new LocalFile();
                    doc.Load(Server.MapPath(path));
                    lf.Name = doc.SummaryInfo.ReportTitle;
                    lf.Path = path;
                    folderlist[shoplogin].Add(lf);
                }
            }
            List<LocalFile> list = new List<LocalFile>();
            foreach (String shoplogin in folderlist.Keys)
            {
                if (folderlist[shoplogin].Count > 0)
                {
                    LocalFile lf = new LocalFile();
                    if (shoplogin.Equals("Survey") || shoplogin.Equals("Corporate"))
                    {
                        lf.Name = shoplogin;
                    }
                    else
                    {
                        lf.Name = Profile.GetProfile(shoplogin).Alias;
                    }
                    lf.Path = "";
                    list.Add(lf);
                    list.AddRange(folderlist[shoplogin]);
                }
            }
            leftmenu.DataSource = list;
            leftmenu.DataBind();
        }
    }

    public void LoadReport(Object sender, CommandEventArgs e) //MenuEventArgs e)
    {
        intro.Visible = false;
        String path = (String)e.CommandArgument;

        if (!path.EndsWith(".rpt")) {
            Page.ClientScript.RegisterStartupScript(this.GetType(), "OpenSurveyScript", "window.open(\""+path+"\");", true);
            return;
        }

        String[] folders = path.Split('/');
        CrystalReportSource1.ReportDocument.Load(Server.MapPath((String)e.CommandArgument));
        ReportDocument doc = CrystalReportSource1.ReportDocument;

        foreach (InternalConnectionInfo cn in CrystalReportSource1.ReportDocument.DataSourceConnections)
        {
            String cs = ConfigurationManager.ConnectionStrings["CSiSQlExpressReports"].ConnectionString;
            SqlConnectionStringBuilder decoder = new SqlConnectionStringBuilder(cs);
            cn.SetConnection(decoder.DataSource, decoder.InitialCatalog, decoder.UserID, decoder.Password);
        }

        ReportViewer.ReportSourceID = "CrystalReportSource1";
        ReportViewer.RefreshReport();

        try
        {
            int ShopID = Profile.GetProfile(HttpContext.Current.User.Identity.Name).ShopID;
            if (!HttpContext.Current.User.Identity.Name.Equals(folders[folders.Length - 2]))
            {
                ShopID = Profile.GetProfile(folders[folders.Length-2]).ShopID;
            }
            ReportViewer.ParameterFieldInfo["ShopID"].CurrentValues.AddValue(ShopID);
        }
        catch (Exception ex)
        {
        }
        try
        {
            // First day of last month
            DateTime d = new DateTime(DateTime.Today.Year, DateTime.Today.AddMonths(-1).Month, 1);
            ReportViewer.ParameterFieldInfo["Report Month"].CurrentValues.AddValue(d);
        }
        catch (Exception ex)
        {
        }
    }
     
    public string ReverseMapPath(string path)
    {
        string appPath = HttpContext.Current.Server.MapPath("~");
        string res = string.Format("/{0}", path.Replace(appPath, "").Replace("\\", "/"));
        return res;
    }
}