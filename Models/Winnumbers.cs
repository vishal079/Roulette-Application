using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Data;

namespace Roullete.Models
{
    public class Winnumbers
    {
        public int winnumber { get; set; }
        public int Priority_High_Low { get; set; }
        public List<int> numbers = new List<int>();
        public List<int> PossibleWin = new List<int>();
        public List<int> GetPrecionsFrom = new List<int>();
        public int Section1 {get;set;}
        public int Section2 { get; set; }
        public int Section3 { get; set; }
        public int Section4 { get; set; }
        public int Section5 { get; set; }
        public int Section6 { get; set; }
        public int Section7 { get; set; }
        public int Section8 { get; set; }
        public int Section9 { get; set; }
        public DataTable Presion_Result { get; set; }
        public Winnumbers()
        {
            Section1 = 50;
            Section2 = 50;
            Section3 = 50;
            Section4 = 50;
            Section5 = 50;
            Section6 = 50;
            Section7 = 50;
            Section8 = 50;
            Section9 = 50;
        }
        public Winnumbers(DataRow dr)
        {
 
        }
    }
}