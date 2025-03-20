using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace EmpLetterGenerate.Models
{
    public class LetterTemplate
    {
        public int LetterId { get; set; }
        public string LetterType { get; set; }
        public string LetterSubject { get; set; }
        public string LetterBody { get; set; }
    }
}