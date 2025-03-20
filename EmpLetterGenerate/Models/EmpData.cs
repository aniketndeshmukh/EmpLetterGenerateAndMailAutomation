using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmpLetterGenerate.Models
{
    public class EmpData
    {
        public int EmpId { get; set; }
        public string EmpName { get; set; }
        public string EmpAddress { get; set; }
        public string EmpEmail { get; set; }
        public string Designation { get; set; }
    }
}