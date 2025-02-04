using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.ApplicationServices;


namespace BaseOffice
{
    [Transaction(TransactionMode.Manual)]
    internal class Memorial : IExternalCommand
    {
        //create global variables
        UIApplication _uiapp;
        Autodesk.Revit.ApplicationServices.Application _app;
        UIDocument _uidoc;
        Document _doc;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Creating App and doc objects.
            _uiapp = commandData.Application;
            _app = _uiapp.Application;
            _uidoc = _uiapp.ActiveUIDocument;
            _doc = _uidoc.Document;


            using (Transaction transacao = new Transaction(_doc, "Exportar Memorial"))
            {

                ProjectInfo projectInfo = _doc.ProjectInformation;


                transacao.Start();

                ExcelExport instanciaExcel = new ExcelExport();

                instanciaExcel.CriarArquivo();
                instanciaExcel.CriarLinha("Nome do projeto", projectInfo);
                instanciaExcel.CriarLinha("Nome do Contratante", projectInfo, "Cidade", -1);
                instanciaExcel.CriarLinha("Data do Projeto", projectInfo);
                instanciaExcel.CriarLinha("Nome do Contratante", projectInfo);

                instanciaExcel.CriarLinha("REV 01 - NÚM", projectInfo);
                instanciaExcel.CriarLinha("REV 01 - DATA", projectInfo);
                instanciaExcel.CriarLinha("REV 01 - TIPO", projectInfo);
                instanciaExcel.CriarLinha("REV 01 - DESC", projectInfo);
                instanciaExcel.CriarLinha("REV 01 - ELAB", projectInfo);
                instanciaExcel.CriarLinha("REV 01 - VERIF", projectInfo);

                instanciaExcel.CriarLinha("Título do Arquivo", projectInfo);

                instanciaExcel.SalvarArquivo();
                instanciaExcel.AbrirArquivo();

                transacao.Commit();

                return Result.Succeeded;
            }
        }
    }
}