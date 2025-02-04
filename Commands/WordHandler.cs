using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Document = Autodesk.Revit.DB.Document;

namespace BaseOffice
{
    internal class WordHandler
    {
        private readonly RevitContext _context;
        private string _caminhoArquivo;
        private string _caminhoArquivoBase = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "mmd.docx");

        public WordHandler(RevitContext context)
        {
            _context = context;
        }

        public void ObterCaminhoSalvar()
        {
            // Abrir janela "Salvar Como"
            var saveFileDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Arquivos Word (*.docx)|*.docx",
                Title = "Salvar documento como",
                FileName = "MMD-XXXXX-EXE-HDS-0101-REV0X.docx" // Nome sugerido
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                _caminhoArquivo = saveFileDialog.FileName;
            }
            else
            {
                TaskDialog.Show("Aviso", "Operação Cancelada");
            }
        }

        public void CarregarDocumento()
        {
            if (!File.Exists(_caminhoArquivoBase))
                TaskDialog.Show("Aviso", "Arquivo Base Nao Encontrado");

            // Copiar o arquivo base para o destino (mantendo um documento original)
            File.Copy(_caminhoArquivoBase, _caminhoArquivo, overwrite: true);
        }

        public void AbrirDocumento()
        {
            if (File.Exists(_caminhoArquivo))
            {
                Process.Start(new ProcessStartInfo(_caminhoArquivo) { UseShellExecute = true });
            }
            else
            {
                TaskDialog.Show("Aviso", "Não foi possível abrir o arquivo");
            }
        }

        

        public void ReplaceText(string textoAntigo, string parametroReferencia)
        {
            using (var wordDocument = WordprocessingDocument.Open(_caminhoArquivo, true))
            {
                var body = wordDocument.MainDocumentPart.Document.Body;

                ProjectInfo projectInfo = _context.Doc.ProjectInformation;
                var parametro = projectInfo.LookupParameter(parametroReferencia);
                var textoNovo = parametro?.AsString();

                if (string.IsNullOrEmpty(textoNovo))
                {
                    textoNovo = parametroReferencia;
                }

                // Substituir no corpo do documento
                ReplaceInElements(body.Descendants<Run>(), textoAntigo, textoNovo);

                // Substituir nos cabeçalhos
                foreach (var headerPart in wordDocument.MainDocumentPart.HeaderParts)
                {
                    ReplaceInElements(headerPart.RootElement.Descendants<Run>(), textoAntigo, textoNovo);
                }

                // Substituir nos rodapés
                foreach (var footerPart in wordDocument.MainDocumentPart.FooterParts)
                {
                    ReplaceInElements(footerPart.RootElement.Descendants<Run>(), textoAntigo, textoNovo);
                }

                // Salvar o documento após as alterações
                wordDocument.MainDocumentPart.Document.Save();
            }
        }

        private void ReplaceInElements(IEnumerable<Run> runElements, string textoAntigo, string textoNovo)
        {
            foreach (var runElement in runElements)
            {
                var runTextElements = runElement.Elements<Text>().ToList();

                for (int i = 0; i < runTextElements.Count; i++)
                {
                    var textElement = runTextElements[i];
                    string searchText = textElement.Text;
                    string textLower = textElement.Text.ToLower();

                    // Verificar se o texto contém o valor a ser substituído
                    if (textLower.Contains(textoAntigo.ToLower()))
                    {
                        // Ajustar o formato do texto novo para corresponder ao formato do texto antigo
                        if (searchText == searchText.ToUpper())
                        {
                            textoNovo = textoNovo.ToUpper();
                        }
                        else if (searchText == searchText.ToLower())
                        {
                            textoNovo = textoNovo.ToLower();
                        }
                        else
                        {
                            textoNovo = Capitalizar.TitleCase(textoNovo);
                        }

                        // Substituir o texto no nó atual
                        textElement.Text = Regex.Replace(textElement.Text, Regex.Escape(textoAntigo),
                            textoNovo,
                            RegexOptions.IgnoreCase);

                        // Aplicar o destaque verde ao texto substituído
                        var runProperties = runElement.GetFirstChild<RunProperties>();
                        if (runProperties == null)
                        {
                            runProperties = new RunProperties();
                            runElement.PrependChild(runProperties);
                        }

                        runProperties.Highlight = new Highlight { Val = HighlightColorValues.Green };
                    }
                }
            }
        }


        public void ReplaceImage(string idImagemAntiga, string caminhoImagemNova)
        {
            using (var wordDocument = WordprocessingDocument.Open(_caminhoArquivo, true))
            {
                var mainPart = wordDocument.MainDocumentPart;

                // Encontra o ImagePart correspondente pelo ID
                var imagePart = (ImagePart)mainPart.GetPartById(idImagemAntiga);

                if (imagePart != null)
                {
                    // Substitui o conteúdo do ImagePart pela nova imagem
                    using (var stream = new FileStream(caminhoImagemNova, FileMode.Open))
                    {
                        imagePart.FeedData(stream);
                    }
                }

                // Salva o documento após a substituição
                mainPart.Document.Save();
            }
        }


    }
}
