using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Data.SqlClient;
using System.Configuration;
//using OfficeOpenXml;

namespace Web.UI
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void btnUpload_Click(object sender, EventArgs e)
        {
            if(flExcel.PostedFiles.Count > 0)
            {
                try
                {
                    if (flExcel.PostedFiles[0].FileName.ToLower().Contains(".xlsx"))
                    {
                        XSSFWorkbook xssfwb;
                        xssfwb = new XSSFWorkbook(flExcel.PostedFiles[0].InputStream);
                        ISheet xssSheet = xssfwb.GetSheetAt(0);

                        processing(xssSheet);
                    }
                    else if (flExcel.PostedFiles[0].FileName.ToLower().Contains(".xls"))
                    {
                        HSSFWorkbook hssfwb;
                        hssfwb = new HSSFWorkbook(flExcel.PostedFiles[0].InputStream);
                        ISheet hssSheet = hssfwb.GetSheetAt(0);

                        processing(hssSheet);
                    }
                    else if (flExcel.PostedFiles[0].FileName.ToLower().Contains(".docx"))
                    {
                        //we have to write for all types of excell files
                    }
                    else
                    {
                        lblStatus.Text = "Invalid File format.. Please upload a valid file, .xls or .xlsx";
                    }
                }
                catch (Exception ex)
                {
                    exceptionLog(ex);
                }
            }
            else
            {
                lblStatus.Text = "Please upload file";
            }
        }

        private void processing(ISheet sheet)
        {
            var excelDt = generateDatatableFromIsheet(sheet);
            var dbDt = getDataFromDataBase();

            // Comparing both data tables
            var records = compareTwoDataTables(excelDt, dbDt);
            if (records.Rows.Count > 0)
            {
                GridView1.DataSource = excelDt;
                GridView1.DataBind();
                lblStatus.Text = flExcel.PostedFiles[0].FileName;
            }
            else
            {
                lblStatus.Text = "No Matching records found against DB Table";
            }
        }

        private DataTable generateDatatableFromIsheet(ISheet sheet)
        {
            DataTable dt = new DataTable(sheet.SheetName);
            try
            {
                // write header row
                IRow headerRow = sheet.GetRow(0);
                foreach (ICell headerCell in headerRow)
                {
                    dt.Columns.Add(headerCell.ToString());
                }
                // write the rest
                int rowIndex = 0;
                foreach (IRow row in sheet)
                {
                    // skip header row
                    if (rowIndex++ == 0) continue;

                    // add row into datatable
                    var cells = new List<ICell>();
                    for (int i = 0; i < dt.Columns.Count; i++)
                    {
                        cells.Add(row.GetCell(i, MissingCellPolicy.CREATE_NULL_AS_BLANK));
                    }
                    dt.Rows.Add(cells.Select(c => c.ToString()).ToArray());
                }
            }
            catch (Exception ex)
            {
                exceptionLog(ex);
            }
            
            return dt;
        }

        private DataTable getDataFromDataBase()
        {
            DataTable dt = new DataTable();
            try
            {
                string queryString = "SELECT * FROM dbo.Employee;";
                string connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
                using (SqlConnection connection =  new SqlConnection(connectionString))
                {
                    SqlDataAdapter adapter = new SqlDataAdapter();
                    adapter.SelectCommand = new SqlCommand(
                        queryString, connection);
                    adapter.Fill(dt);
                    return dt;
                }

            }
            catch(Exception ex)
            {
                exceptionLog(ex);
            }
            
            return dt;
        }

        ///https://stackoverflow.com/questions/10984453/compare-two-datatables-for-differences-in-c
        private DataTable compareTwoDataTables(DataTable excelDt, DataTable dbDt)
        {
            DataTable newData = new DataTable();
            try
            {
                // get records ID matching AND Name non matching
                var results = from table1 in excelDt.AsEnumerable()
                              join table2 in dbDt.AsEnumerable() on table1.Field<string>("Name") equals table2.Field<string>("EName")
                              //where table1.Field<string>("Name") != table2.Field<string>("EName")
                              // && table1.Field<int>("ColumnB") != table2.Field<int>("ColumnB") && table1.Field<String>("ColumnC") != table2.Field<String>("ColumnC")
                              select table1;
                if(results.Count() > 0)
                {
                    newData = results.CopyToDataTable();

                    DataTable sessionDt = excelDt.Clone();
                    sessionDt.Columns.Add("STATUS");

                    foreach (DataRow dr in excelDt.Rows)
                    {
                        if (newData.AsEnumerable().Any(r => r.Field<string>("Name") == dr.Field<string>("Name")))
                        {
                            var itemArray = dr.ItemArray.ToList();
                            itemArray.Add("UPDATED");
                            sessionDt.Rows.Add(itemArray.ToArray());
                        }
                        else
                        {
                            var itemArray = dr.ItemArray.ToList();
                            itemArray.Add("NOT UPDATED");
                            sessionDt.Rows.Add(itemArray.ToArray());
                        }
                    }

                    // updating to DB 
                    updateEmployees(newData);

                    Session["sessionDt"] = sessionDt;
                }
                else
                {
                    lblStatus.Text = "No Matching records found";
                }
            }
            catch (Exception ex)
            {
                exceptionLog(ex);
            }
            return newData;
        }

        private void exceptionLog(Exception ex)
        {
            //Log to Database table 
        }

        protected void btnDownLoad_Click(object sender, EventArgs e)
        {
            if(Session["sessionDt"] != null)
            {
                ExportDataTableToExcel((DataTable)Session["sessionDt"], "employeesFile");
            }
            else
            {
                lblStatus.Text = "File Upload is not yet completed !!";
            }
        }

        /// <summary>
        /// Downaload DataTable to Excel File
        /// </summary>
        /// <param name="sourceTable">Source DataTable</param>
        /// <param name="fileName">Destination File name</param>
        private void ExportDataTableToExcel(DataTable sourceTable, string fileName)
        {
            try
            {
                HSSFWorkbook workbook = new HSSFWorkbook();
                MemoryStream memoryStream = new MemoryStream();
                ISheet sheet = workbook.CreateSheet("Sheet1");
                IRow headerRow = sheet.CreateRow(0);

                // handling header.
                foreach (DataColumn column in sourceTable.Columns)
                    headerRow.CreateCell(column.Ordinal).SetCellValue(column.ColumnName);

                // handling value.
                int rowIndex = 1;

                foreach (DataRow row in sourceTable.Rows)
                {
                    IRow dataRow = sheet.CreateRow(rowIndex);

                    foreach (DataColumn column in sourceTable.Columns)
                    {
                        dataRow.CreateCell(column.Ordinal).SetCellValue(row[column].ToString());
                    }

                    rowIndex++;
                }

                workbook.Write(memoryStream);
                memoryStream.Flush();

                HttpResponse response = HttpContext.Current.Response;
                response.ContentType = "application/vnd.ms-excel";
                response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}.xls", fileName));
                response.Clear();

                response.BinaryWrite(memoryStream.GetBuffer());
                response.End();
            }
            catch(Exception ex)
            {
                exceptionLog(ex);
            }
        }

        /// <summary>
        /// Update to the Database dbo.Employee from Datatable
        /// If not like this way of updating write as your expectations
        /// </summary>
        /// <param name="updatedDt"></param>
        private void updateEmployees(DataTable updatedDt)
        {
            try
            {
                string connectionString = ConfigurationManager.ConnectionStrings["myConnectionString"].ConnectionString;
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    using (SqlCommand command = new SqlCommand("", conn))
                    {
                        try
                        {
                            conn.Open();

                            //Creating temp table on database
                            command.CommandText = 
                                @"CREATE TABLE #TmpTable([ID] [int] , [EName][varchar](50), [Salary] [int] );";
                            command.ExecuteNonQuery();

                            //Bulk insert into temp table
                            using (SqlBulkCopy bulkcopy = new SqlBulkCopy(conn))
                            {
                                bulkcopy.BulkCopyTimeout = 660;
                                bulkcopy.DestinationTableName = "#TmpTable";
                                bulkcopy.WriteToServer(updatedDt);
                                bulkcopy.Close();
                            }

                            // Updating destination table, and dropping temp table
                            command.CommandTimeout = 300;
                            command.CommandText = "UPDATE e SET e.EName = tmp.Name, e.Salary = tmp.Salary  FROM dbo.Employee e INNER JOIN #TmpTable tmp ON e.ID = tmp.ID; DROP TABLE #TmpTable;";
                            command.ExecuteNonQuery();
                        }
                        catch (Exception ex)
                        {
                            exceptionLog(ex);
                        }
                        finally
                        {
                            conn.Close();
                        }
                    }
                }


            }
            catch (Exception ex)
            {
                exceptionLog(ex);
            }

        }
    }
}