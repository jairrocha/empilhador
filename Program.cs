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
            int coluna = 7;
            string pathArquivoConfig = Directory.GetCurrentDirectory() + @"\Arquivos\Config.xlsx";
            string pathSaida = Directory.GetCurrentDirectory() + @"\Arquivos\Saida.xlsx";
            var fi = new FileInfo(pathArquivoConfig);
            var cf = new Config();

            // Ler o arquivo Config
            using (var p = new ExcelPackage(fi))
            {

                var ws = p.Workbook.Worksheets["Config"];

                // Armazena os nomes de campos em listas
                cf.LNomeCampo.Add("Base");
                while (!(ws.Cells[1, coluna].Value == null))
                {
                    cf.LNomeCampo.Add(ws.Cells[1, coluna].Value.ToString());
                    coluna++;
                }

                // Armazena os parâmetros em listas
                while (!(ws.Cells[linha, 1].Value == null))
                {
                    coluna = 7;
                    cf.LNomeBase.Add(ws.Cells[linha, 1].Value.ToString());
                    cf.LSheet.Add(ws.Cells[linha, 2].Value.ToString());
                    cf.LPathBase.Add(ws.Cells[linha, 3].Value.ToString());
                    cf.LLinhaIN.Add(Int32.Parse(ws.Cells[linha, 4].Value.ToString()));
                    cf.LColunaIN.Add(Int32.Parse(ws.Cells[linha, 5].Value.ToString()));
                    cf.LColunaKey.Add(Int32.Parse(ws.Cells[linha, 6].Value.ToString()));

                    cf.LParametros.Add(new List<int>());

                    for (int i = 0; i < cf.LNomeCampo.Count; i++)
                    {
                        if (ws.Cells[linha, coluna].Value == null) break;
                        cf.LParametros[cf.LParametros.Count - 1].Add(Int32.Parse(ws.Cells[linha, coluna].Value.ToString()));
                        coluna++;
                    }

                    linha++;
                }

            }

            //Gera Arquivo de saida
            using (var p = new ExcelPackage())
            {
                int colunaSaida = 1;
                int linhaSaida = 2;

                List<int> sublista = null;

                var wssaida = p.Workbook.Worksheets.Add("Saida");

                //Escreve o nome das colunas no arquivo de saída
                foreach (var item in cf.LNomeCampo)
                {
                    wssaida.Cells[1, colunaSaida].Value = item.ToString();
                    colunaSaida++;
                }

                for (int i = 0; i < cf.LNomeBase.Count; i++)
                {
                    // Lê as bases
                    using (var b = new ExcelPackage(new FileInfo(cf.LPathBase[i])))
                    {

                        var wsbase = b.Workbook.Worksheets[cf.LSheet[i]];
                        linha = cf.LLinhaIN[i];
                        coluna = cf.LColunaKey[i];


                        while (!(wsbase.Cells[linha, coluna].Value == null))
                        {
                            sublista = cf.LParametros[i];
                            for (int j = 0; j < sublista.Count; j++)
                            {
                                wssaida.Cells[linhaSaida, 1].Value = cf.LNomeBase[i];
                                wssaida.Cells[linhaSaida, j + 2].Value = wsbase.Cells[linha, sublista[j]].Value;

                            }

                            linha++;
                            linhaSaida++;

                        }

                    }

                }

                p.SaveAs(new FileInfo(pathSaida));
            }




        }


    }

}

