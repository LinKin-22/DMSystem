using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace DMSystem
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                PopulateData();
                lblMessage.Text = "Current Database data!";
            }
        }

        private void PopulateData()
        {
            using (MuDatabaseEntities dc = new MuDatabaseEntities())
            {
                gvData.DataSource = dc.EmployeeMasters.ToList();
                gvData.DataBind();
            }
        }

        protected void btnImport_Click(object sender, EventArgs e)
        {
            if (FileUpload1.PostedFile.ContentType == "application/vnd.ms-excel" ||
                FileUpload1.PostedFile.ContentType ==
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
            {
                try
                {
                    string fileName = Path.Combine(Server.MapPath("~/ImportDocument"),
                        Guid.NewGuid().ToString() + Path.GetExtension(FileUpload1.PostedFile.FileName));
                    FileUpload1.PostedFile.SaveAs(fileName);

                    string conString = "";
                    string ext = Path.GetExtension(FileUpload1.PostedFile.FileName);
                    if (ext.ToLower() == ".xls")
                    {
                        conString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName +
                                    ";Extended Properties='Excel 8.0;HDR=Yes;IMEX=1;';";
                    }
                    else if (ext.ToLower() == ".xlsx")
                    {
                        conString = "Provider=Microsoft.ACE.OLEDB.12.0" +
                                    "" +
                                    "" +
                                    "" +
                                    "" +
                                    "" +
                                    "" +
                                    "" +
                                    "" +
                                    "" +
                                    ";Data Source=" + fileName +
                                    ";Extended Properties='Excel 12.0 xml;HDR=Yes;IMEX=1;';";
                    }

                    string query =
                        "Select [Student ID],[Student Roll],[Student Name],[Department Name],[Student Address],[Contact Title],[Student Email] from [EmployeeData$]";

                    OleDbConnection con = new OleDbConnection(conString);
                    if (con.State == System.Data.ConnectionState.Closed)
                    {
                        con.Open();
                    }

                    OleDbCommand cmd = new OleDbCommand(query, con);
                    OleDbDataAdapter da = new OleDbDataAdapter(cmd);

                    DataSet ds = new DataSet();
                    da.Fill(ds);
                    da.Dispose();
                    con.Close();
                    con.Dispose();


                    // Import to Database
                    using (MuDatabaseEntities dc = new MuDatabaseEntities())
                    {
                        foreach (DataRow dr in ds.Tables[0].Rows)
                        {
                            string stdID = dr["Student ID"].ToString();
                            var v = dc.EmployeeMasters.Where(a => a.StudentID.Equals(stdID)).FirstOrDefault();
                            if (v != null)
                            {
                                //update here
                                v.StudentRoll = dr["Student Roll"].ToString();
                                v.StudentName = dr["Student Name"].ToString();
                                v.DepartmentName = dr["Department Name"].ToString();
                                v.StudentAddress = dr["Student Address"].ToString();
                                v.ContactTitle = dr["Contact Title"].ToString();
                                v.StudentEmail = dr["Student Email"].ToString();
                            }
                            else
                            {
                                // Insert
                                dc.EmployeeMasters.Add(new EmployeeMaster
                                {
                                    StudentID = dr["Student ID"].ToString(),
                                    StudentRoll = dr["Student Roll"].ToString(),
                                    StudentName = dr["Student Name"].ToString(),
                                    DepartmentName = dr["Department Name"].ToString(),
                                    StudentAddress = dr["Student Address"].ToString(),
                                    ContactTitle = dr["Contact Title"].ToString(),
                                    StudentEmail = dr["Student Email"].ToString()
                                });
                            }
                        }
                        dc.SaveChanges();
                    }
                    PopulateData();
                    lblMessage.Text = "Successfully data import done!!!";
                }
                
                catch (Exception)
                {
                    throw;
                }
            }
        }
    }
}