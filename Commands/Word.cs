using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;
using NPOI.SS.Formula.Functions;
using System.Windows.Forms;
using BaseOffice.UI;


namespace BaseOffice
{
    [Transaction(TransactionMode.Manual)]
    internal class Word : IExternalCommand
    {

        string Cidade;
        string Estado;


        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            RevitContext context = new RevitContext(commandData.Application);

            using (Transaction transacao = new Transaction(context.Doc, "Exportar MMD"))
            {

                transacao.Start();


                ProjectInfo projectInfo = context.Doc.ProjectInformation;

                WordHandler wordHandler = new WordHandler(context);

                string titleBlock = Sheets.GetTitleBlockName(context.Doc);
                var consorcio = Sheets.ValidateTitleBlock(titleBlock);

                TaskDialog.Show("Teste", consorcio);

                wordHandler.ObterCaminhoSalvar();

                //var inputHandler = new StateInput();
                //if (inputHandler.ShowDialog() == DialogResult.OK)
                //{
                //    Cidade = inputHandler.Input1;
                //    Estado = inputHandler.Input2;
                //}

                var WpfForm = new UI.UI(context.Doc);
                WpfForm.ShowDialog();
                Cidade = WpfForm.City;
                Estado = WpfForm.State;

                wordHandler.CarregarDocumento();
                wordHandler.ReplaceText("ProjectName", "Nome do projeto");
                wordHandler.ReplaceText("Contratante", "Nome do Contratante");
                wordHandler.ReplaceText("Date", "Data do Projeto");
                wordHandler.ReplaceText("TITLE", "Título do Arquivo");
                wordHandler.ReplaceText("City", Cidade);
                wordHandler.ReplaceText("State", Estado);
                wordHandler.ReplaceText("Consorcio", consorcio);
                wordHandler.AbrirDocumento();

                transacao.Commit();

                return Result.Succeeded;
            }
        }
    }
}