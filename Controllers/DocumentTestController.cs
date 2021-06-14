using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using WebApplication2.Model;
using WebApplication2.Service;

namespace WebApplication2.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentTestController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
            "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        };
        DataClass service;

        public DocumentTestController()
        {
            this.service = new DataClass();
        }

        [Route("getdocument")]
        [HttpPost]
        public IActionResult GetDocument([FromBody] ModelTest param)
        {
            return Ok(this.service.FillReport(param));
        }

        [Route("getpdf")]
        [HttpPost]
        public IActionResult GetPdf(ModelTest param)
        {
            return Ok(this.service.GetPdf(param));
        }

        [Route("getDecode")]
        [HttpGet]
        public IActionResult GetDecode()
        {
            return Ok(this.service.GetDecode());
        }
    }
}
