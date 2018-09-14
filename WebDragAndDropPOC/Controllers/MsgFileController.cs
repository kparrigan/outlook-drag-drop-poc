using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web;
using System.Web.Http;
using WebDragAndDropPOC.Middleware;
using WebDragAndDropPOC.Models;
using WebDragAndDropPOC.Providers;

namespace WebDragAndDropPOC.Controllers
{
    public class MsgFileController : ApiController
    {
        private readonly IMsgParser _msgParser;

        public MsgFileController(IMsgParser msgParser)
        {
            _msgParser = msgParser ?? throw new ArgumentNullException(nameof(msgParser));
        }

        // POST api/values
        public async Task<IHttpActionResult> Post()
        {
            // Check if the request contains multipart/form-data.
            if (!Request.Content.IsMimeMultipartContent())
            {
                return BadRequest();
            }

            var root = HttpContext.Current.Server.MapPath("~/App_Data");
            var provider = new MsgOnlyMultipartFormDataStreamProvider(root);

            try
            {
                // Read the form data.
                await Request.Content.ReadAsMultipartAsync(provider);
                var files = provider.FileData;

                if (files == null || files.Count == 0)
                {
                    return BadRequest("No valid msg files uploaded.");
                }

                var metadata = new List<MsgMetadata>();

                foreach(var file in files)
                {
                    metadata.Add(_msgParser.Parse(file.LocalFileName, file.Headers.ContentDisposition.FileName));
                }

                return Ok(metadata);
            }
            catch(Exception ex)
            {
                return InternalServerError(new Exception("Error uploading file"));
            }
        }


    }
}
