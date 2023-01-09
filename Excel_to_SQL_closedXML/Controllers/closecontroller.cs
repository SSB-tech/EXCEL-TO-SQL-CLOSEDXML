using ClosedXML.Excel;
using Dapper;
using Excel_to_SQL_closedXML.models;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Configuration;
using System.Data.SqlClient;

namespace Excel_to_SQL_closedXML.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class Closecontroller : ControllerBase
	{
		private readonly IConfiguration _configuration;

		public Closecontroller( IConfiguration configuration)
		{
			_configuration = configuration;
		}

		[HttpPost]
		public ActionResult import(IFormFile file)
		{
			List<close> list = new List<close>();

			MemoryStream stream = new MemoryStream();
			file.CopyTo(stream);

			XLWorkbook workbook = new XLWorkbook(stream);
			IXLWorksheet worksheet = workbook.Worksheet(1);
			var row = worksheet.RowsUsed().Count();

			for (int i = 2; i <= row; i++)
			{
				list.Add(new close
				{
					Id = worksheet.Cell(i, 1).GetValue<int>(), //First ma (int)worksheet.cells[i,1].value; yo garda null reference error ayo did it this way
					Customercode = worksheet.Cell(i,2).GetValue<int>(),
					FirstName = worksheet.Cell(i, 3).GetValue<string>(),
					LastName = worksheet.Cell(i, 4).GetValue<string>(),
					gender = worksheet.Cell(i, 5).GetValue<string>(),
					Country = worksheet.Cell(i, 6).GetValue<string>(),
					Age = worksheet.Cell(i, 7).GetValue<int>(),
				}
				);
			}

			var connection = new SqlConnection(_configuration.GetConnectionString("defaultconnection"));
			connection.ExecuteAsync("insert into closexmltbl (CustomerCode, FirstName, LastName, Gender, Country,Age) values (@CustomerCode, @FirstName, @LastName, @Gender, @Country, @Age)", list);
			return Ok(list);

		}
	}
}
