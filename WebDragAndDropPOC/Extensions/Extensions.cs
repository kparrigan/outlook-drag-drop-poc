using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace WebDragAndDropPOC.Extensions
{
    public static class Extensions
    {
        public static string GetFileExtension(this string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName))
            {
                return string.Empty;
            }

            fileName = fileName.Replace("\"", "");

            return Path.GetExtension(fileName);
        }
    }
}