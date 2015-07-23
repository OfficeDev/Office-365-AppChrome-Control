using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Security.Claims;
using System.Web.Http;

namespace LoginControlForSPASolution.Controllers
{
    [Authorize]
    public class WebAPIController : ApiController
    {
        // GET api/WebAPI
        public string Get()
        {
            string owner = ClaimsPrincipal.Current.FindFirst(ClaimTypes.Name).Value;
            return "Hello, Server know you, "+owner+"!";
        }

        // GET api/WebAPI/5
        public string Get(int id)
        {
            return "value";
        }

        // POST api/WebAPI
        public void Post([FromBody]string value)
        {
        }

        // PUT api/WebAPI/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/WebAPI/5
        public void Delete(int id)
        {
        }
    }
}