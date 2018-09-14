using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web;
using WebDragAndDropPOC.Extensions;

namespace WebDragAndDropPOC.Providers
{
    public class MsgOnlyMultipartFormDataStreamProvider : MultipartFormDataStreamProvider
    {
        public MsgOnlyMultipartFormDataStreamProvider(string rootPath)
            : base (rootPath)
        {

        }

        public override Stream GetStream(HttpContent parent, HttpContentHeaders headers)
        {
            const string extension = ".msg";
            var fileExtension = headers.ContentDisposition.FileName.GetFileExtension();

            return string.Equals(extension, fileExtension, StringComparison.InvariantCultureIgnoreCase) ? base.GetStream(parent, headers) : Stream.Null;
        }
    }
}