using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace PDFConverter.Models
{
    public class FileModel
    {
        [DataType(DataType.Upload)]
        public HttpPostedFileBase file { get; set; }
        //public string Name { get; set; }


    }
}