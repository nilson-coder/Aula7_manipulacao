using System;
using Spire.Doc;
using Spire.Doc.Documents;

namespace EX
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Documento
                Document exemploDoc = new Document();
            #endregion

            #region seção no Documento
                Section secaocapa = exemploDoc.AddSection();
            #endregion

            #region Criar um paragrafo
                Paragraph titulo = secaocapa.AddParagraph();
            #endregion

            #region Adiciona texto ao paragrafo
                titulo.AppendText("Exercícios resulvido\n\n");
            #endregion

            #region texto
                titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;
            #endregion

            #region 
                
            #endregion



       
            #region Salvar arquivo
                exemploDoc.SaveToFile(@"saida\exemplo_arquivo_word.docx", FileFormat.Docx);
            #endregion
        }
    }
}
