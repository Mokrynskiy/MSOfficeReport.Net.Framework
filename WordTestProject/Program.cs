using MSOfficeReport.Net.Framework.WordReport;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordTestProject
{
    internal class Program
    {
        static void Main(string[] args)
        {
            decimal sum = 2.23M;
            Header test = new Header { LongName = "Акционерное общество \"Вектор-Бест\"", ShortName = "AO \"Вектор-Бест\"", Address = "г.Новосибирск, ул. Арбузова, д.1/1"};
            

            WordTemplate template = new WordTemplate("CommercialOfferSpec.docx");
            template.AddVariable("Saller", test);           
            template.Generate();
            template.SaveAs("TestOutput.docx");
            Process.Start("TestOutput.docx");
        }
    }
    public class Header
    {
        public string LongName { get; set; }
        public string ShortName { get; set; }
        public string Address { get; set; }
        public string Rtf { get; set; }
        public List<Positions> Positions { get; set; }
        public decimal SumAll
        {
            get
            {
                return Positions.Sum(x => x.Sum);
            }
        }
    }
    public class Positions
    {
        public int Number { get; set; }
        public string CatNumber { get; set; }
        public string Name { get; set; }
        public string Unit { get; set; }
        public int Amount { get; set; }
        public decimal Price { get; set; }
        public int Vat { get; set; }
        public decimal SumVat { get; set; }
        public decimal Sum { get; set; }
    }
}
