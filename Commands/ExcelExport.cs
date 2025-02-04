using System;
using System.Diagnostics;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using NPOI.SS.Formula.Functions;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel; // Para arquivos .xlsx

namespace BaseOffice
{
    public class ExcelExport
    {
        #region Props
        private int LinhaAtual { get; set; }
        private int ColunaAtual { get; set; }
        private IRow Linha0 { get; set; }
        private IRow Linha1 { get; set; }

        private ISheet Sheet { get; set; }
        public string CaminhoArquivo { get; set; }
        public IWorkbook Workbook { get; set; }

        #endregion


        public void CriarArquivo()
        {
            // Criação do diálogo para salvar o arquivo
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Arquivos Excel (*.xlsx)|*.xlsx",
                Title = "Salvar arquivo Excel",
                FileName = "Planilha.xlsx"
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                this.CaminhoArquivo = saveFileDialog.FileName;

                // Criação do arquivo Excel usando NPOI
                this.Workbook = new XSSFWorkbook(); // Para .xlsx
                this.Sheet = this.Workbook.CreateSheet("Main");
            }
            else
            {
                TaskDialog.Show("Aviso", "Operação Cancelada");
            }
        }

        public void SalvarArquivo()
        {
            using (FileStream fs = new FileStream(this.CaminhoArquivo, FileMode.Create, FileAccess.Write))
            {
                this.Workbook.Write(fs);
            }
        }

        public void AbrirArquivo()
        {
            // Abrindo o arquivo salvo
            Process.Start(new ProcessStartInfo(this.CaminhoArquivo) { UseShellExecute = true });
        }

        public void CriarLinha(string nomeParametro, ProjectInfo projectInfo)
        {
            if(LinhaAtual < 1)
            {
            Linha0 = this.Sheet.CreateRow(0);
            Linha1 = this.Sheet.CreateRow(1);
            LinhaAtual = 1;}

            var parametro = projectInfo.LookupParameter(nomeParametro);
            var valorParametro = parametro.AsString();

            Linha0.CreateCell(ColunaAtual).SetCellValue(nomeParametro);
            Linha1.CreateCell(ColunaAtual).SetCellValue(valorParametro);
            Sheet.AutoSizeColumn(ColunaAtual);
            ColunaAtual++;
        }

        public void CriarLinha(string nomeParametro, ProjectInfo projectInfo, string nomeNaCelula, int posicao)
        {
            if (LinhaAtual < 1)
            {
                Linha0 = this.Sheet.CreateRow(0);
                Linha1 = this.Sheet.CreateRow(1);
                LinhaAtual = 1;
            }

            var parametro = projectInfo.LookupParameter(nomeParametro);
            var valorParametro = parametro.AsString();

            string[] palavras = valorParametro.Split(' ');
            int tamanho = palavras.Length;
            string palavraSelecionada;

            if (posicao == -1)
            {
                palavraSelecionada = palavras[tamanho - 1];
            }
            else
            {
                palavraSelecionada = palavras[posicao];
            }

            Linha0.CreateCell(ColunaAtual).SetCellValue(nomeNaCelula);
            Linha1.CreateCell(ColunaAtual).SetCellValue(palavraSelecionada);
            Sheet.AutoSizeColumn(ColunaAtual);
            ColunaAtual++;
        }
    }
}
