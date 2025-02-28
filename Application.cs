﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.UI;

namespace BaseOffice
{

    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    [Autodesk.Revit.Attributes.Regeneration(Autodesk.Revit.Attributes.RegenerationOption.Manual)]

    public class Application : IExternalApplication
    {

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }


        public Result OnStartup(UIControlledApplication application)
        {

            StartApp.StartupMain(application);

            RibbonPanel panelMain = StartApp.CriarRibbonPanel(application, "Main");

            #region CriarPushButton DEV

            StartApp.CriarPushButton
                (
                "DEV",
                "DEV",
                "BaseOffice.Dev",
                panelMain,
                "Uso para testes",
                "dev.ico",
                true
                );

            #endregion

            #region CriarPushButton Excel

            StartApp.CriarPushButton
                (
                "Excel",
                "Excel",
                "BaseOffice.Memorial",
                panelMain,
                "Exportação de Tabela Excel",
                "excel.ico"
                );

            #endregion

            #region CriarPushButton Word

            StartApp.CriarPushButton
                (
                "Word",
                "Word",
                "BaseOffice.Word",
                panelMain,
                "Exportação de arquivo word",
                "word.ico"
                );


            #endregion

            return Result.Succeeded;
        }
    }
}
