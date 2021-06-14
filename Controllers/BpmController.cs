using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using WebApplication2.Model;
using WebApplication2.Service;

namespace WebApplication2.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class BpmController : Controller
    {
        // GET: BpmController
        public ActionResult Index()
        {
            return View();
        }

        [Route("listworks")]
        [HttpPost]
        public IActionResult ListWorks([FromBody] DataRequest param)
        {
            return Ok(new BpmClass().getJsonResponse(param));
        }
    }   
}
