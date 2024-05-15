using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Globalization;


namespace ExcelReader
{
    public class Article
    {
        public int Id { get; set; }
        public string Libelle { get; set; }
        public int PU { get; set; }

        public Article() { }

        public Article(string id, string libelle, string pu)
        {
            this.Id = int.Parse(id);
            this.Libelle = libelle;
            this.PU = int.Parse(pu);
        }
    }
    //Added Achat class here as it didn't work making it with it's own file, same thing for Vente
    public class Achat
    {
        public int Num { get; set; }
        public int Art_Id { get; set; }
        public int Qte { get; set; }
        public DateTime Date { get; set; }

        public Achat() { }

        public Achat(string num, string id, string qte, DateTime date)
        {
            this.Num = int.Parse(num);
            this.Art_Id = int.Parse(id);
            this.Qte = int.Parse(qte);
            this.Date = date;
        }
    }

    public class Vente
    {
        public int Num { get; set; }
        public int Art_Id { get; set; }
        public int Qte { get; set; }
        public DateTime Date { get; set; }

        public Vente() { }

        public Vente(string num, string id, string qte, DateTime date)
        {
            this.Num = int.Parse(num);
            this.Art_Id = int.Parse(id);
            this.Qte = int.Parse(qte);
            this.Date = date;

        }
    }
}
