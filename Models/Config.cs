using System.Collections.Generic;

namespace empilhador.Models
{
    public class Config
    {
        public List<string> LNomeBase = new List<string>();
        public List<string> LPathBase = new List<string>();
        public List<int> LLinhaIN = new List<int>();
        public List<int> LColunaIN = new List<int>();
        public List<int> LColunaKey = new List<int>();

        public List<string> LNomeCampo = new List<string>();
        public List<List<string>> LParametros= new List<List<string>> ();

        

        // public Config()
        // {
        //    ValorCampo.Add(valor);
        // }

    }
}