using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ExcelToSqlServer.Models;
using System.IO;
using ExcelDataReader;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Cookies;
using System.Security.Claims;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace ExcelToSqlServer.Controllers
{
    [Authorize]
    public class HomeController : Controller
    {
        private IWebHostEnvironment _hostingEnvironment;
        private string _connectioString;

        public HomeController(
            IWebHostEnvironment environment,
            IConfiguration configuration
        )
        {
            _hostingEnvironment = environment;
            _connectioString = configuration.GetConnectionString("DbContext");
        }

        public IActionResult Index()
        {
            return View();
        }
        public IActionResult GLView()
        {
            return View(nameof(Index));
        }

        [HttpPost]
        public async Task<IActionResult> Post(IFormFile excelFile)
        {
            if (excelFile == null || excelFile.Length <= 0)
            {
                return View(nameof(Index));
            }

            System.IO.Directory.CreateDirectory(
                System.IO.Path.Combine(_hostingEnvironment.WebRootPath, "excel")
            );
            var filePath = System.IO.Path.Combine(
                _hostingEnvironment.WebRootPath, "excel", "excelFile.xlsx"
            );

            using (var stream = System.IO.File.Create(filePath))
            {
                await excelFile.CopyToAsync(stream);
            }

            // Process uploaded files
            // Don't rely on or trust the FileName property without validation.

            ViewData["Message"] = "File uploaded to the server successfully!";
            return View(nameof(Index));
        }

        [AllowAnonymous]
        public IActionResult Register()
        {
            // Clear the existing external cookie
            HttpContext.SignOutAsync(
                CookieAuthenticationDefaults.AuthenticationScheme
            );

            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        public IActionResult Register(User user)
        {
            using (var connection = new SqlConnection(_connectioString))
            {
                connection.Open();

                var command = new SqlCommand(
                    "INSERT INTO [EMAILING].[dbo].[Users] ([Email], [Password])"
                    + "VALUES ('" + user.Email + "', '" + user.Password + "');",
                    connection
                );
                command.ExecuteNonQuery();
            }

            return View(nameof(Login));
        }

        [AllowAnonymous]
        public IActionResult Login()
        {
            // Clear the existing external cookie
            HttpContext.SignOutAsync(
                CookieAuthenticationDefaults.AuthenticationScheme
            );

            return View();
        }

        [HttpPost]
        [AllowAnonymous]
        public async Task<IActionResult> Login(User model)
        {
            if (ModelState.IsValid)
            {
                if (!AuthenticateUser(model))
                {
                    ModelState.AddModelError(string.Empty, "Invalid login attempt.");
                    return View();
                }

                var claims = new List<Claim>
                {
                    new Claim(ClaimTypes.Email, model.Email),
                    new Claim("DateTime", DateTime.UtcNow.ToString())
                };

                var claimsIdentity = new ClaimsIdentity(
                    claims, CookieAuthenticationDefaults.AuthenticationScheme);

                var authProperties = new AuthenticationProperties
                {
                    //AllowRefresh = <bool>,
                    // Refreshing the authentication session should be allowed.

                    //ExpiresUtc = DateTimeOffset.UtcNow.AddMinutes(10),
                    // The time at which the authentication ticket expires. A 
                    // value set here overrides the ExpireTimeSpan option of 
                    // CookieAuthenticationOptions set with AddCookie.

                    //IsPersistent = true,
                    // Whether the authentication session is persisted across 
                    // multiple requests. When used with cookies, controls
                    // whether the cookie's lifetime is absolute (matching the
                    // lifetime of the authentication ticket) or session-based.

                    //IssuedUtc = <DateTimeOffset>,
                    // The time at which the authentication ticket was issued.

                    // RedirectUri = "/"
                    // The full path or absolute URI to be used as an http 
                    // redirect response value.
                };

                await HttpContext.SignInAsync(
                    CookieAuthenticationDefaults.AuthenticationScheme,
                    new ClaimsPrincipal(claimsIdentity),
                    authProperties);
            }


            return RedirectToAction(nameof(Index));
        }

        private bool AuthenticateUser(User user)
        {
            using (var connection = new SqlConnection(_connectioString))
            {
                connection.Open();

                var command = new SqlCommand(
                    "SELECT COUNT(*) FROM [EMAILING].[dbo].[Users]"
                    + "WHERE Email='" + user.Email + "' AND Password='" + user.Password + "';",
                    connection
                );

                return ((int)command.ExecuteScalar() > 0);
            }
        }

        public IActionResult SaveExcelFileToDatabase()
        {
            var filePath = System.IO.Path.Combine(
                    _hostingEnvironment.WebRootPath, "excel", "excelFile.xlsx"
                );

            try
            {
                using (System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read)) { };
            }
            catch (System.Exception)
            {
                return View(nameof(Index));
            };

            RecreateDatabase();

            using (var stream = System.IO.File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var result = reader.AsDataSet().Tables["Sheet1"];
                    for (int row = 8; row < result.Rows.Count - 1; row++)
                    {
                        var model = new GL
                        {
                            AnneeMois = result.Rows[row].ItemArray[1].ToString(),
                            Compte = result.Rows[row].ItemArray[2].ToString(),
                            Texte = result.Rows[row].ItemArray[3].ToString(),
                            NumPiece = result.Rows[row].ItemArray[4].ToString(),
                            DatePiece = result.Rows[row].ItemArray[5].ToString(),
                            Affectation = result.Rows[row].ItemArray[6].ToString(),
                            Reference = result.Rows[row].ItemArray[7].ToString(),
                            TypeDePiece = result.Rows[row].ItemArray[8].ToString(),
                            MontantEnDeviseInterne = Convert.ToSingle(result.Rows[row].ItemArray[9]),
                            PieceRapprochement = result.Rows[row].ItemArray[10].ToString()
                        };

                        InsertGL(model);
                    }
                }

                System.IO.File.Delete(filePath);

                ViewData["Message"] = "File saved to database successfully.";
                return View("GLView", ReadGL());
            }
        }

        private void RecreateDatabase()
        {
            using (var connection = new SqlConnection(_connectioString))
            {
                connection.Open();

                var command = new SqlCommand("IF OBJECT_ID ('EMAILING.dbo.GL') IS NOT NULL DROP TABLE EMAILING.dbo.GL;", connection);
                command.ExecuteNonQuery();

                command = new SqlCommand(
                    @"CREATE TABLE [EMAILING].[dbo].[GL](
                        [Client] [varchar](50) NULL,
                        [Date pièce] [varchar](50) NULL,
                        [Année Mo] [varchar](50) NULL,
                        [Texte] [varchar](50) NULL,
                        [N° pièce] [varchar](50) NULL,
                        [Affectation] [varchar](50) NULL,
                        [Référence] [varchar](50) NULL,
                        [Typ] [varchar](50) NULL,
                        [Mtant] [real] NULL,
                        [EC] [varchar](50) NULL
                    )",
                connection);
                command.ExecuteNonQuery();
            }

        }

        private void InsertGL(GL gl)
        {
            using (var connection = new SqlConnection(_connectioString))
            {
                connection.Open();
                var queryString = "INSERT INTO [EMAILING].[dbo].[GL]"
                    + "([Client], [Date pièce], [Année Mo], [Texte], [N° pièce], [Affectation], [Référence], [Typ], [Mtant], [EC])"
                    + " VALUES ("
                    + "'" + gl.Compte + "', "
                    + "'" + gl.DatePiece + "', "
                    + "'" + gl.AnneeMois + "', "
                    + "'" + gl.Texte.Replace("'", "''") + "', "
                    + "'" + gl.NumPiece + "', "
                    + "'" + gl.Affectation + "', "
                    + "'" + gl.Reference + "', "
                    + "'" + gl.TypeDePiece + "', "
                    + "'" + gl.MontantEnDeviseInterne + "', "
                    + "'" + gl.PieceRapprochement + "'"
                    + ");";
                var command = new SqlCommand(queryString, connection);
                command.ExecuteNonQuery();
            }
        }

        private IEnumerable<GL> ReadGL()
        {
            var glList = new List<GL>();

            using (var connection = new SqlConnection(_connectioString))
            {
                var queryString = "SELECT * FROM [EMAILING].[dbo].[GL];";
                var command = new SqlCommand(queryString, connection);
                command.Connection.Open();

                var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    var gl = new GL
                    {
                        Compte = reader["Client"].ToString(),
                        DatePiece = reader["Date pièce"].ToString(),
                        AnneeMois = reader["Année Mo"].ToString(),
                        Texte = reader["Texte"].ToString(),
                        NumPiece = reader["N° pièce"].ToString(),
                        Affectation = reader["Affectation"].ToString(),
                        Reference = reader["Référence"].ToString(),
                        TypeDePiece = reader["Typ"].ToString(),
                        MontantEnDeviseInterne = Convert.ToSingle(reader["Mtant"]),
                        PieceRapprochement = reader["EC"].ToString()
                    };

                    glList.Add(gl);
                }
            }

            return glList;
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}
