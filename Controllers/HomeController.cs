using Microsoft.AspNetCore.Mvc;
using Npgsql;
using System.Data;
using System.Data.OleDb;
using System.Diagnostics;
using WebApplication2.Models;


namespace WebApplication2.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        private readonly IConfiguration Configuration;

        public HomeController(ILogger<HomeController> logger, IConfiguration configuration)
        {
            _logger = logger;
            Configuration = configuration;
        }

        public IActionResult Home()
        {
            return View();
        }

        [HttpPost]
        public bool Import(IFormFile file)
        {
            DataSet ds = new DataSet();

            if (file != null)
            {
                string path = Path.Combine("wwwroot", "Uploads");
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }

                string fileName = Path.GetFileName(file.FileName);
                string filePath = Path.Combine(path, fileName);
                using (FileStream stream = new FileStream(filePath, FileMode.Create))
                {
                    file.CopyTo(stream);
                }

                string conString = Configuration["ExcelConString"];

                DataTable dt = new DataTable();
                conString = string.Format(conString, filePath);

                using (var connExcel = new OleDbConnection(conString))
                {
                    using (OleDbCommand cmdExcel = new OleDbCommand())
                    {
                        using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                        {
                            cmdExcel.Connection = connExcel;

                            connExcel.Open();
                            DataTable dtExcelSchema;
                            dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                            string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                            connExcel.Close();

                            connExcel.Open();
                            cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                            odaExcel.SelectCommand = cmdExcel;
                            odaExcel.Fill(ds);
                            connExcel.Close();
                        }
                    }
                }

                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                {
                    string conn = Configuration["ConnectionString"];
                    NpgsqlConnection con = new NpgsqlConnection(conn);
                    string query = "Insert into public.\"TableName\"(\"Column1\", \"Column2\", \"Column3\",\"Column3\", \"Column4\", \"Column5\", \"Column6\", \"Column7\", \"Column8\", \"Column9\", \"Column10\", \"Column11\") Values('" + ds.Tables[0].Rows[i][0].ToString() + "','" + ds.Tables[0].Rows[i][1] + "','" + ds.Tables[0].Rows[i][2] + "','" + ds.Tables[0].Rows[i][3] + "','" + ds.Tables[0].Rows[i][4].ToString() + "','" + ds.Tables[0].Rows[i][5].ToString() + "','" + ds.Tables[0].Rows[i][6].ToString() + "','" + ds.Tables[0].Rows[i][7] + "','" + ds.Tables[0].Rows[i][8] + "','" + ds.Tables[0].Rows[i][9].ToString() + "','" + ds.Tables[0].Rows[i][10] + "','" + ds.Tables[0].Rows[i][11] + "')";
                    con.Open();
                    NpgsqlCommand cmd = new NpgsqlCommand(query, con);
                    cmd.ExecuteNonQuery();
                    con.Close();
                }
            }
            return true;
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}