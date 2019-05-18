using api.CEP.Data;
using Microsoft.AspNetCore.Authorization;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System.Collections.Generic;

namespace api.CEP.Controllers
{
	[Route("api/[controller]")]
	[ApiController]
	public class FindController : ControllerBase
	{
		[HttpGet]
		public ActionResult<string> Get()
		{
			return "Try: $> cURL GET https://" + Request.Host.Value + "/api/find/<search-term>";
		}

		[HttpGet("{s}")]
		public ActionResult<string> Get([FromServices] CepDAO cepDAO, string s)
		{
			string exitCode = string.Empty;
			try
			{
				var code = long.Parse(s);
				var search = cepDAO.Search(code);
				exitCode = JsonConvert.SerializeObject(search);
			}
			catch
			{
				var search = cepDAO.Search(s);
				exitCode = JsonConvert.SerializeObject(search);
			}
			return exitCode;
		}

	}
}
