using System;
using System.Data;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data.CData.GoogleSheets;

public partial class _Default : System.Web.UI.Page
{
    private String CLIENT_ID = "949443617993-ovh8gin9pbah77ritt0qor4e6ialqv8n.apps.googleusercontent.com";
    private String CLIENT_SECRET = "l0v3_gwXL_IzuejUi1dIvAXZ";

    private String connectionString = "";

    protected void btnOAuthConnect_Click(object sender, EventArgs e)
    {
        if (tbAccessToken.Text == "")
        {
            if (CLIENT_ID == "CLIENT_ID_HERE" || CLIENT_SECRET == "CLIENT_SECRET_HERE")
            {
                Response.Write("<font color=red><b>You must specify the CLIENT_ID and CLIENT_SECRET constants in code before calling Connect.</b></font>");
                return;
            }

            GoogleSheetsConnection conn = new GoogleSheetsConnection(String.Format("Offline=false;OAuth Client ID={0};OAuth Client Secret={1};", CLIENT_ID, CLIENT_SECRET));
            GoogleSheetsCommand cmd = new GoogleSheetsCommand("GetOAuthAuthorizationURL", conn);
            String Thispage = "http://" + Request.ServerVariables["SERVER_NAME"] + ":" + Request.ServerVariables["SERVER_PORT"] + Request.ServerVariables["URL"];
            Session["ThisPage"] = Thispage;
            cmd.Parameters.Add(new GoogleSheetsParameter("@AuthMode", "WEB"));
            cmd.Parameters.Add(new GoogleSheetsParameter("@CallbackURL", Thispage));
            cmd.Parameters.Add(new GoogleSheetsParameter("@ResponseType", "code"));
            cmd.Parameters.Add(new GoogleSheetsParameter("@ApprovalPrompt", "AUTO"));
            cmd.Parameters.Add(new GoogleSheetsParameter("@AccessType", "OFFLINE"));
            cmd.Parameters.Add(new GoogleSheetsParameter("@Scope", "https://docs.google.com/feeds/ https://spreadsheets.google.com/feeds"));
            cmd.CommandType = CommandType.StoredProcedure;
            GoogleSheetsDataReader dr = cmd.ExecuteReader();
            String URL = "";
            while (dr.Read())
            {
                URL = ((String)dr["URL"]);
            }

            Response.Redirect(URL, true);
        }
        else {
            connectionString = String.Format("Offline=False;OAuth Access Token={0}", tbAccessToken.Text);
            Session["connection"] = connectionString;
        }
    }

    protected void GetAccessToken()
    {
        connectionString = String.Format("OAuth Client Id={0};OAuth Client Secret={1}", CLIENT_ID, CLIENT_SECRET);
        using (GoogleSheetsConnection connection = new GoogleSheetsConnection(connectionString))
        {
            GoogleSheetsCommand cmd = new GoogleSheetsCommand("GetOAuthAccessToken", connection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add(new GoogleSheetsParameter("@AuthMode", "WEB"));
            cmd.Parameters.Add(new GoogleSheetsParameter("@Verifier", Request.QueryString["code"]));
            cmd.Parameters.Add(new GoogleSheetsParameter("@CallbackURL", Session["ThisPage"]));
            cmd.Parameters.Add(new GoogleSheetsParameter("@ResponseType", "code"));
            cmd.Parameters.Add(new GoogleSheetsParameter("@ApprovalPrompt", "AUTO"));
            cmd.Parameters.Add(new GoogleSheetsParameter("@AccessType", "OFFLINE"));
            cmd.Parameters.Add(new GoogleSheetsParameter("@Scope", "https://docs.google.com/feeds/ https://spreadsheets.google.com/feeds"));

            GoogleSheetsDataReader rdr = cmd.ExecuteReader();
            while (rdr.Read())
            {
                tbAccessToken.Text = rdr["OAuthAccessToken"] + "";
            }
            
            connectionString = String.Format("Offline=False;OAuth Access Token={0}", tbAccessToken.Text);

        }
    }

    protected void Page_Load(object sender, EventArgs e)
    {
        if (Request.QueryString["code"] != null && tbAccessToken.Text == "")
        {
            GetAccessToken();
        }
    }


    protected void btnProductSearch_Click(object sender, EventArgs e)
    {
        string query = "SELECT * FROM [Products] WHERE ";
        connectionString = String.Format("Offline=False;OAuth Access Token={0}", tbAccessToken.Text);
        GoogleSheetsConnection connection = new GoogleSheetsConnection(connectionString + ";Spreadsheet=Northwind;" + "header=true", tbAccessToken.Text);
        GoogleSheetsCommand cmd = null;

        if (productSearchDD.SelectedIndex == 1)
        {
            cmd = new GoogleSheetsCommand("SELECT * FROM [Categories] WHERE CategoryName ='" + tbSearch.Text + "';", connection);
        }
        else
        {
            cmd = new GoogleSheetsCommand("SELECT * FROM [Suppliers] WHERE Country = '" + tbSearch.Text +"';", connection);
        }
        DataTable table = new DataTable();
        GoogleSheetsDataAdapter da = new GoogleSheetsDataAdapter(cmd);

        da.Fill(table);

        foreach (DataRow row in table.Rows)
        {
            if (productSearchDD.SelectedIndex == 1)
            {
                query += "CategoryID =" + row["CategoryID"];
            }
            else
            { 
                query += "SupplierID =" + row["Supplierid"];
            }
            if (table.Rows.IndexOf(row)==(table.Rows.Count-1))
            {
                query += ";";
            }
            else
            {
                query += " OR ";
            }
        }
        GoogleSheetsConnection connection2 = new GoogleSheetsConnection(connectionString + ";Spreadsheet=Northwind;" + "header=true", tbAccessToken.Text);
        GoogleSheetsCommand cmd2 = new GoogleSheetsCommand(query, connection2);
        GoogleSheetsDataAdapter da2 = new GoogleSheetsDataAdapter(cmd2);
        DataTable table2 = new DataTable();
        da2.Fill(table2);


        gvSheets.DataSource = table2;
        gvSheets.DataBind();
        Session["Table"] = table2;


    }

    protected void btnAvgPrice_Click(object sender, EventArgs e)
    {
        Dictionary<string, double> price = new Dictionary<string, double>();        
        Dictionary<string, int> count = new Dictionary<string, int>();
        connectionString = String.Format("Offline=False;OAuth Access Token={0}", tbAccessToken.Text);
        GoogleSheetsConnection connection = new GoogleSheetsConnection(connectionString + ";Spreadsheet=Northwind;" + "header=true", tbAccessToken.Text);
        GoogleSheetsCommand cmd = new GoogleSheetsCommand("SELECT * FROM [Orders]", connection);
        GoogleSheetsDataAdapter da = new GoogleSheetsDataAdapter(cmd);
        DataTable table = new DataTable();

        da.Fill(table);

        foreach(DataRow row in table.Rows)
        {
            if (price.ContainsKey(row["ShipCountry"].ToString()))
            {
                price[row["ShipCountry"].ToString()] += ((double)row["OrderPrice"]);
                count[row["ShipCountry"].ToString()] += 1;
            }
            else
            {
                try
                {
                    price.Add(row["ShipCountry"].ToString(), (double)row["OrderPrice"]);
                    count.Add(row["ShipCountry"].ToString(), 1);
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                }
            }
        }


        var list = price.Keys.ToList();
        list.Sort();
        var output = new DataTable();
        DataRow rows;
        DataColumn column;
        column = new DataColumn();
        column.DataType = Type.GetType("System.String");
        column.ColumnName = "Country";
        output.Columns.Add(column);
        column = new DataColumn();
        column.DataType = Type.GetType("System.Double");
        column.ColumnName = "Average Price";
        output.Columns.Add(column);
        foreach (var key in list)
        {
            rows = output.NewRow();
            rows["Average Price"]= price[key] / count[key];
            rows["Country"] = key;
            output.Rows.Add(rows);
        }
        gvSheets.DataSource = output;
        gvSheets.DataBind();
        Session["Table"] = output;
        Session["Sort"] = "Country ASC";

    }

    protected void gvSheets_Sorting(object sender, GridViewSortEventArgs e)
    {
        string currentSort = (string)Session["Sort"];
        string direction = "ASC";
        if (currentSort == e.SortExpression + " " + direction)
            direction = "DESC";
        DataTable table = (DataTable)Session["Table"];
        DataView view = new DataView(table);
        view.Sort = e.SortExpression +" "+ direction;
        Session["Sort"] = e.SortExpression + " " + direction;
        gvSheets.DataSource = view;
        gvSheets.DataBind();
    }
}