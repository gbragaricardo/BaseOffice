using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;
using Autodesk.Revit.UI;

namespace BaseOffice
{
    internal static class StartApp
    {
        static readonly string tab = "BaseOffice";

        static readonly string thisAssemblyPath = Assembly.GetExecutingAssembly().Location;

        internal static void StartupMain(UIControlledApplication application)
        {           

            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) =>
            {
                string folderPath = @"C:\Users\Usuario\AppData\Roaming\Autodesk\Revit\Addins\2024\ProjetaHDR";
                string assemblyPath = System.IO.Path.Combine(folderPath, new AssemblyName(args.Name).Name + ".dll");
                return System.IO.File.Exists(assemblyPath) ? Assembly.LoadFrom(assemblyPath) : null;
            };

        // Cria a aba
            try
            {
                application.CreateRibbonTab(tab);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Aba '{tab}' já existe ou houve um erro ao tentar criá-la: {ex.Message}");
            }
        }

        // Criar Painel
        internal static RibbonPanel CriarRibbonPanel(UIControlledApplication application, string nomePainel)
        {
            RibbonPanel ribbonPanel = null;

            try
            {
                ribbonPanel = application.CreateRibbonPanel(tab, nomePainel);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Erro ao criar o painel '{nomePainel}': {ex.Message}");
            }

            return ribbonPanel;
        }


        // Criar e Detalhar PushButton
        internal static void CriarPushButton(string nomeInterno,
            string nomeExibido,
            string nomeClasse,
            RibbonPanel painel,
            string dica,
            string nomeImagem,
            bool enabled = false)
        {

           //Cria os dados do botao
            PushButtonData pushButtonData = new PushButtonData
            (
                nomeInterno,
                nomeExibido,
                thisAssemblyPath,
                nomeClasse
                
            );

            // Adiciona o botão ao painel
            PushButton pushButton = painel.AddItem(pushButtonData) as PushButton;

            pushButton.Enabled = enabled;

            if( pushButton != null)
            {
                // Define uma dica (tooltip) que aparecerá quando o usuário passar o mouse sobre o botão
                pushButton.ToolTip = dica;

                // Define o caminho para o ícone do botão
                string iconPath = Path.Combine(Path.GetDirectoryName(thisAssemblyPath),"Resources", nomeImagem);

                // Cria a imagem do ícone
                Uri uri = new Uri(iconPath);
                BitmapImage bitmap = new BitmapImage(uri);

                // Define a imagem como o ícone do botão
                pushButton.LargeImage = bitmap;
            }

            
        }
    }
}
