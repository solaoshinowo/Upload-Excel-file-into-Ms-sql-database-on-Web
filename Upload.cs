using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using OfficeOpenXml;
using System.Data;
using System.IO;
using System.Data.SqlClient;
using yewoexcelupload;

namespace cooperativesocietysoftware
{
    public partial class excelupload : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            
            if (IsPostBack && Upload.HasFile)
            {
                if (Path.GetExtension(Upload.FileName).Equals(".xlsx"))
                {
                    var excel = new ExcelPackage(Upload.FileContent);
                    var dt = excel.ToDataTable();
                    var table = "excelimport";
                    using (var conn = new SqlConnection("Server=yourservername;Database=yourdatabasename;User Id=yourid;Password=yourpassword;Integrated Security=false"))

                    {
                        var bulkCopy = new SqlBulkCopy(conn);
                        bulkCopy.DestinationTableName = table;
                        conn.Open();
                        var schema = conn.GetSchema("Columns", new[] { null, null, table, null });
                        foreach (DataColumn sourceColumn in dt.Columns)
                        {
                            foreach (DataRow row in schema.Rows)
                            {
                                if (string.Equals(sourceColumn.ColumnName, (string)row["COLUMN_NAME"], StringComparison.OrdinalIgnoreCase))
                                {
                                    bulkCopy.ColumnMappings.Add(sourceColumn.ColumnName, (string)row["COLUMN_NAME"]);
                                    break;
                                }
                            }
                        }
                        SqlDataSource1.Delete();
                        bulkCopy.WriteToServer(dt);
                        Response.Write("<script>alert('Contribution file as been successfully uploaded..');window.location = 'excelupload.aspx';</script>");
                        SqlDataSource1.Update();
                        GridView1.Visible = true;
                    }
                }
            }
        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            SqlDataSourcepost.Insert();
            Response.Write("<script>alert('Contributions  as been successfully posted..');window.location = 'excelupload.aspx';</script>");
        }
    }
}
