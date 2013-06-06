using CrystalDecisions.CrystalReports.Engine;
using CrystalDecisions.Shared;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlClient;
using System.IO;
using System.Security.Principal;
using System.Web;
using System.Web.Security;
using System.Web.UI;
using System.Web.UI.HtmlControls;
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
            String lastshoplogin = "";
            int shops = 0;
            List<LocalFile> list = new List<LocalFile>();
            foreach (String shoplogin in folderlist.Keys)
            {
                if (folderlist[shoplogin].Count > 0)
                {
                    LocalFile lf = new LocalFile();
                    if (shoplogin.Equals("Survey"))
                    {
                        SurveyLink.NavigateUrl = folderlist[shoplogin][0].Path;
                        SurveyItem.Visible = true;
                    }
                    else if (shoplogin.Equals("Corporate"))
                    {
                        ExecutiveLink.CommandArgument = shoplogin;
                        ExecutiveItem.Visible = true;
                    }
                    else
                    {
                        lf.Name = Profile.GetProfile(shoplogin).Alias;
                        lf.Path = shoplogin;
                        list.Add(lf);

                        shops++;
                        lastshoplogin = shoplogin;
                    }
                    //list.AddRange(folderlist[shoplogin]);
                }
            }
            if (shops > 1)
            {
                ShopList.DataSource = list;
                ShopList.DataBind();
            }
            else
            {
                MultipleShops.Visible = false;
                ShopLink.Visible = true;
                ShopLink.CommandArgument = lastshoplogin;
            }

            Session["folderlist"] = folderlist;
        }
    }

    protected void ShowDashboard(object sender, CommandEventArgs e)
    {
        ChangeShopList("", DashboardItem);
    }
    protected void ShopList_SelectedIndexChanged(object sender, EventArgs e)
    {
        String shoplogin = ShopList.SelectedItem.Value;
        ChangeShopList(shoplogin,ShopListItem);
    }
    public void ChangeShopList(Object sender, CommandEventArgs e) //MenuEventArgs e)
    {
        String shoplogin = (String)e.CommandArgument;
        HtmlControl selectedItem = ShopListItem;
        if(e.CommandName.Equals("Executive")) {
            selectedItem = ExecutiveItem;
        }
        ChangeShopList(shoplogin,selectedItem);
    }

    protected void ChangeShopList(String shoplogin,HtmlControl selectedItem)
    {
        if (shoplogin == null || shoplogin.Length == 0)
        {
            ReportViewer.Visible = false;
            ReportList.DataSource = new List<LocalFile>();
            ReportList.DataBind();
        }
        try
        {
            Dictionary<String, List<LocalFile>> folderlist = (Dictionary<String, List<LocalFile>>)Session["folderlist"];
            if (folderlist != null)
            {
                ReportViewer.Visible = false;
                ReportList.DataSource = folderlist[shoplogin];
                ReportList.DataBind();
            }
        }
        catch (Exception ex) { }
        SelectPrimaryLink(selectedItem);
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
        CrystalReportSource1.ReportDocument.Load(Server.MapPath((String)e.CommandArgument),OpenReportMethod.OpenReportByDefault);
        ReportDocument doc = CrystalReportSource1.ReportDocument;

        //String cs = ConfigurationManager.ConnectionStrings["SQL1"].ConnectionString;
        //SqlConnectionStringBuilder decoder = new SqlConnectionStringBuilder(cs);
        //doc.SetDatabaseLogon(decoder.UserID, decoder.Password, decoder.DataSource, decoder.InitialCatalog);

        //foreach (InternalConnectionInfo cn in CrystalReportSource1.ReportDocument.DataSourceConnections)
        //{
        //    String cs = ConfigurationManager.ConnectionStrings["SQL1"].ConnectionString;
        //    SqlConnectionStringBuilder decoder = new SqlConnectionStringBuilder(cs);
        //    cn.SetConnection(decoder.DataSource, decoder.InitialCatalog, decoder.UserID, decoder.Password);
        //}

        String cs = ConfigurationManager.ConnectionStrings["SQL1"].ConnectionString;
        SqlConnectionStringBuilder decoder = new SqlConnectionStringBuilder(cs);
        ConnectionInfo ci = new ConnectionInfo();
        ci.ServerName = decoder.DataSource;
        ci.DatabaseName = decoder.InitialCatalog;
        ci.UserID = decoder.UserID;
        ci.Password = decoder.Password;

        TableLogOnInfo tli = new TableLogOnInfo();
        tli.ConnectionInfo = ci;

        foreach (CrystalDecisions.CrystalReports.Engine.Table t in CrystalReportSource1.ReportDocument.Database.Tables)
        {
            t.ApplyLogOnInfo(tli);
        }

        ReportViewer.Visible = true;
        ReportViewer.ReportSourceID = "CrystalReportSource1";
        ReportViewer.RefreshReport();

        ReportViewer.ToolPanelView = CrystalDecisions.Web.ToolPanelViewType.None;

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
            DateTime d = new DateTime(DateTime.Today.Year, DateTime.Today.AddMonths(-0).Month, 1);
            ReportViewer.ParameterFieldInfo["RptMonth"].CurrentValues.AddValue(d);
        }
        catch (Exception ex)
        {
        }

        String selectedId = (String)e.CommandName;
        SelectSecondaryLink(selectedId);
    }
     
    public string ReverseMapPath(string path)
    {
        string appPath = HttpContext.Current.Server.MapPath("~");
        string res = string.Format("/{0}", path.Replace(appPath, "").Replace("\\", "/"));
        return res;
    }

    private void SelectPrimaryLink(HtmlControl selected)
    {
        DashboardItem.Attributes["class"] = DashboardItem.Attributes["class"].Replace(" isActive", "");
        ExecutiveItem.Attributes["class"] = ExecutiveItem.Attributes["class"].Replace(" isActive", "");
        ShopListItem.Attributes["class"] = ShopListItem.Attributes["class"].Replace(" isActive", "");
        SurveyItem.Attributes["class"] = SurveyItem.Attributes["class"].Replace(" isActive", "");

        if(selected != null) {
            selected.Attributes["class"] = selected.Attributes["class"].TrimEnd() + " isActive";
        }
    }

    private void SelectSecondaryLink(String newSelectedId)
    {
        String selectedId = (String)Session["secondarySelectedId"];

        if (selectedId != null && selectedId.Length > 0)
        {
            HtmlControl lastSelected = (HtmlControl)ReportList.FindControl(selectedId);
            lastSelected.Attributes["class"] = lastSelected.Attributes["class"].Replace(" isActive", "");
        }
        try
        {
            HtmlControl selected = (HtmlControl)ReportList.FindControl(newSelectedId);
            if (selected != null)
            {
                selected.Attributes["class"] = selected.Attributes["class"].TrimEnd() + " isActive";
                Session["secondarySelectedId"] = selected.ID;
            }
            else
            {
                Session["secondarySelectedId"] = "";
            }
        }
        catch (Exception ex) { }
    }
}