using System;
using System.Collections.Generic;
using System.IO;
using empilhador.Models;
using OfficeOpenXml;

namespace empilhador
{
    class Program
    {
        static void Main(string[] args)
        {

            int linha = 2;
            int coluna = 6; 
            string pathArquivoConfig = Directory.GetCurrentDirectory() + @"\Arquivos\Config.xlsx";
            string pathSaida = Directory.GetCurrentDirectory() + @"\Arquivos\Saida.xlsx";
            var fi = new FileInfo(pathArquivoConfig);
            
            // Ler o arquivo Config
            using (var p = new ExcelPackage(fi))
            {
                
                var ws = p.Workbook.Worksheets["Config"];
                var cf = new Config();

                // Armazena os nomes de campos em listas
                while (!(ws.Cells[1, coluna].Value == null))
                {
                    cf.LNomeCampo.Add(ws.Cells[1, coluna].Value.ToString());
                    coluna++;
                }

                // Armazena os parâmetros em listas
                while (!(ws.Cells[linha, 1].Value == null))
                {
                    coluna = 6;
                    cf.LNomeBase.Add(ws.Cells[linha, 1].Value.ToString());
                    cf.LPathBase.Add(ws.Cells[linha, 2].Value.ToString());
                    cf.LLinhaIN.Add(Int32.Parse(ws.Cells[linha, 3].Value.ToString()));
                    cf.LColunaIN.Add(Int32.Parse(ws.Cells[linha, 4].Value.ToString()));
                    cf.LColunaKey.Add(Int32.Parse(ws.Cells[linha, 5].Value.ToString()));
                    
                    cf.LParametros.Add(new List<string>());
                    
                    for (int i = 0; i<cf.LNomeCampo.Count; i++)
                    {
                        if (ws.Cells[linha, coluna].Value == null) break;
                        cf.LParametros[cf.LParametros.Count - 1].Add(ws.Cells[linha, coluna].Value.ToString());
                        coluna++;
                    }

                    linha++;
                }

            }

            //Gera Arquivo de saida
            using (var p = new ExcelPackage())
            {

                foreach (var item in )
                {
                    
                }

                var ws=p.Workbook.Worksheets.Add("Saida");
                ws.Cells["A1"].Value = "This is cell A1";
                p.SaveAs(new FileInfo(pathSaida));
            }
        
        }
    }
}
