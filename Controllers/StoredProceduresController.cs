using System.Collections.Generic;
using System.Data;
using ExcelToSqlServer.Models;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace ExcelToSqlServer.Controllers
{
    [Authorize]
    public class StoredProceduresController : Controller
    {
        private IWebHostEnvironment _hostingEnvironment;
        private string _connectioString;

        public StoredProceduresController(
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

        public IActionResult Solde_calc()
        {
            using (var connection = new SqlConnection(_connectioString))
            {
                connection.Open();

                var storedProcedureCall = "Solde_calc";

                var command = new SqlCommand(storedProcedureCall, connection);
                command.CommandType = CommandType.StoredProcedure;

                var result = new List<Solde_calcModel>();
                var reader = command.ExecuteReader();
                while (reader.Read())
                {
                    var scr = new Solde_calcModel
                    {
                        Client = reader["client"].ToString(),
                        Echu = reader["Echu"].ToString(),
                        Echoir = reader["Echoir"].ToString(),
                        Lien = reader["lien"].ToString(),
                        Lien2 = reader["lien2"].ToString()
                    };

                    result.Add(scr);
                }

                return View(result);
            }
        }

    }
}