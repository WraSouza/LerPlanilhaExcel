using ClosedXML.Excel;
using System;
using System.Linq;

namespace LerPlanilhaExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            int i;
            
                var xls = new XLWorkbook(@"C:\Users\wladi\Documents\PlanilhasBanco\Corretagem.xlsx");

                var planilha = xls.Worksheets.First(w => w.Name == "ORDENS FINALIZADAS");

                var totalLinhas = planilha.Rows().Count();

                for ( i = 2; i <= totalLinhas; i++)
                {
                    var dataHora = planilha.Cell($"A{i}").Value.ToString();

                    var cliente = planilha.Cell($"D{i}").Value.ToString();

                    var conta = planilha.Cell($"C{i}").Value.ToString();

                    var tipo = planilha.Cell($"F{i}").Value.ToString();

                    var ativo = planilha.Cell($"E{i}").Value.ToString();

                    var qtd = planilha.Cell($"G{i}").Value.ToString();

                    Console.WriteLine(dataHora);
                    Console.WriteLine(conta);
                    Console.WriteLine(cliente);
                    Console.WriteLine(ativo);
                    Console.WriteLine(tipo);
                    Console.WriteLine(qtd);
                    Console.WriteLine("");                    
                }

            Console.WriteLine("Quantidade de linhas = " + i);
            Console.ReadLine();
        }
    }
}
