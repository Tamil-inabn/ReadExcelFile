using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Primitives;
using Microsoft.Win32;
using ReadExcelFile.Models;
using System.Data;
using System.Diagnostics;
using System.Reflection.PortableExecutable;
using System.Configuration;
using System.Data.SqlClient;
using System.Xml.Linq;
using Microsoft.AspNetCore.Http;

namespace ReadExcelFile.Controllers
{
    public class HomeController : Controller
    {
        public IHostEnvironment _Environment;        
        public HomeController(IHostEnvironment environment)
        {
            _Environment = environment;
        }

        public IActionResult Index()
        {
            return View();
        }
        [HttpPost]
        public IActionResult Index(IFormFile file)
        {

            string path = Path.Combine(Directory.GetCurrentDirectory(),"wwwroot/Image");
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
            
            DataTable dt = new DataTable();
            using (StreamReader sdr = new StreamReader(file.OpenReadStream()))
            {
                string[] header = sdr.ReadLine().Split(',');
                foreach (string headers in header)
                {
                    dt.Columns.Add(headers);
                }               
                while (!sdr.EndOfStream)
                {
                    string[] row = sdr.ReadLine().Split(',');
                    DataRow dr = dt.NewRow();
                    for (int i = 0; i < header.Length; i++)
                    {
                        dr[i] = row[i];
                    }
                    dt.Rows.Add(dr);
                }

                string constring = "server=CIPL1309_DOTNET\\MSSQLSERVER19;database=ExcelFileUpload;trusted_connection=true;";
                using (SqlConnection con = new SqlConnection(constring))
                {
                    using (SqlBulkCopy sqlBulk=new SqlBulkCopy(con))
                    {
                        sqlBulk.DestinationTableName = "dbo.Employee";                        
                        con.Open();                       
                        sqlBulk.WriteToServer(dt);
                        con.Close();
                    }
                }
            }
            return View();
        }
        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}