using System;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace exemplo
{
    class Program
    {
        static void Main(string[] args)
        {
            
            #region Criacao do documento
                // Cria um documento com o nome exemploDoc
                Document exemploDoc = new Document();
            #endregion

            #region Criacao de secao no documento
                // Adiciona uma seção  com nome secaoCapa ao documento
                // Capa secao pode ser entedida como uma pagina do documento
                Section secaocapa = exemploDoc.AddSection();
            #endregion

            #region Criar um paragrafo
                //Cria um paragrafo com nome titulo e adiciona a seção secaocapa
                //Os paragrafos são necessarios para inserçao de texto, imagens, tabelas etc
                Paragraph titulo = secaocapa.AddParagraph();
            #endregion

            #region Adiciona texto ao paragrafo
                // Adiciona o texto exemplo de titulo ao paragrafo titulo
                titulo.AppendText("Exemplo de título\n\n");
            #endregion

            #region Formatar paragrafo
                // Através da propriedade HorizontalAlignment, é possivel alinha o paragrafo
                titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

                // cria um estilo com o nome estilo1 e adiciona ao documento
                ParagraphStyle estilo1 = new ParagraphStyle(exemploDoc);

                // Adiciona nome ao estilo1
                estilo1.Name = "Cor do titulo";

                // definir a cor do texto
                estilo1.CharacterFormat.TextColor = Color.DarkBlue;

                // define que o texto será em negrito
                estilo1.CharacterFormat.Bold = true;

                // adiciona o estilo1 ao documento exemploDoc
                exemploDoc.Styles.Add(estilo1);

                // Aplica o estilo1 ao paragrafo titulo
                titulo.ApplyStyle(estilo1.Name);
            #endregion

            #region Trabalhar com tabulação
                // adiciona um paragrafo textoCapa á seção secaocapa
                Paragraph textoCapa = secaocapa.AddParagraph();

                // adiciona um texto ao paragrafo com tabulação
                textoCapa.AppendText("\tEste é um exemplo de texto com tabulação\n");
                
                // adiciona um novo paragrafo a mesma seção(secaoCapa)
                Paragraph textoCapa2 = secaocapa.AddParagraph();

                // adiciona umtexto ao paragrafo textocapa2 com concatenação
                textoCapa2.AppendText("\tBasicamente, então, uma seção representa uma página do documento e os paragrafos dentro de uma mesma seção, "+" obviamente, aparecem na mesma página");
            #endregion

            #region Inserir imagens
                // adiciona um paragrafo a secão capa
                Paragraph imagemcapa = secaocapa.AddParagraph();

                // adiciona um texto ao paragrafo imagenscapa
                imagemcapa.AppendText("\n\n\tAgora vamos inserir uma imagem ao documento\n\n");

                // centraliza horizintalmente o paragrafo imagemcapa
                imagemcapa.Format.HorizontalAlignment = HorizontalAlignment.Center;

                // adiciona uma imagem com o nome imagemExemplo ao paragrafo imagemCapa
                DocPicture imagemExemplo = imagemcapa.AppendPicture(Image.FromFile(@"saida/img/logo_csharp.png"));

                // define uma largura e uma altura para a imagem
                imagemExemplo.Width = 300;
                imagemExemplo.Height = 300;
            #endregion

            #region Adicionar nova seção
                // Adiciona uma nova seção
                Section secaocorpo = exemploDoc.AddSection();

                // Adiciona um paragrafo a seção secaocorpo
                Paragraph paragrafocorpo1 = secaocorpo.AddParagraph();

                // Adiciona um texto ao paragrafo paragrafocorpo1
                paragrafocorpo1.AppendText("\tEste é um exemplo de parágrafo criado em uma nova seção." + "\tComo foi criada uma nova seção, perceba que este texto aparece em uma nova página.");
            #endregion

            #region Adicionar uma tebela
                // Adiciona uma tabela à seção secaocorpo
                Table tabela = secaocorpo.AddTable(true);

                // Cria o cabeçalho da tabela
                String[] cabecalho = {"Item", "Descrição", "Qtd.", "Preço Unit.", "Preço"};

                String[][] dados = {
                    new String[]{"Cenoura", "Vegetal muito nutritivo", "1", "R$ 4,00", "R$ 4,00"},
                    new String[]{"Batata", "Vegetal muito Consumido", "2", "R$ 5,00", "R$ 10,00"},
                    new String[]{"Alface", "Vegetal utilizado desde 500 a.c", "1", "R$ 1,50", "R$ 1,50"},
                    new String[]{"Tomate", "tomate é uma fruta", "2", "R$ 6,00", "R$ 12,00"},
                };

                // Adiciona as células na tabela
                tabela.ResetCells(dados.Length + 1, cabecalho.Length);
                
                // Adiciona uma linha na posição [0] do vetor de linha
                // e define que esta linha é o cabeçalho
                TableRow Linha1 = tabela.Rows[0];
                Linha1.IsHeader = true;

                // Define a altura da linha
                Linha1.Height = 23;

                // Formatação do cabeçalho
                Linha1.RowFormat.BackColor = Color.AliceBlue;

                // Percorre as colunas do cabeçalho
                for (int i = 0; i < cabecalho.Length; i++)
                {
                    // Alinhamento das células
                    Paragraph p = Linha1.Cells[i].AddParagraph();
                    Linha1.Cells[i].CellFormat.VerticalAlignment = VerticalAlignment.Middle;
                    p.Format.HorizontalAlignment = HorizontalAlignment.Center;

                    // Formatação dos dados do cabeçalho
                    TextRange TR = p.AppendText(cabecalho[i]);
                    TR.CharacterFormat.FontName = "Calibri";
                    TR.CharacterFormat.FontSize = 14;
                    TR.CharacterFormat.TextColor = Color.Teal;
                    TR.CharacterFormat.Bold = true;
                }

                // Adiciona as linhas do corpo da tabela
                for (int r = 0; r < dados.Length; r++)
                {
                    TableRow LinhaDados = tabela.Rows[r + 1];

                    // Define a altura da linha
                    LinhaDados.Height = 20;

                    // Percorre as colunas
                    for (int c = 0; c < dados[r].Length; c++)
                    {
                        //alinha as celulas
                        LinhaDados.Cells[c].CellFormat.VerticalAlignment = VerticalAlignment. Middle;

                        //Preenche os dados nas linhas
                        Paragraph p2 = LinhaDados.Cells[c].AddParagraph();
                        TextRange TR2 = p2.AppendText(dados[r][c]);

                        //Formata as células
                        p2.Format.HorizontalAlignment = HorizontalAlignment.Center;
                        TR2.CharacterFormat.FontName = "calibri";
                        TR2.CharacterFormat.FontSize = 12;
                        TR2.CharacterFormat.TextColor = Color.Brown;
                    }

                }
            #endregion

            #region Salvar arquivo
                //salva o arquivo em .Dock
                // Utiliza o metodo SaveToFile para salvar o arquivo no formato desejado
                // assim como no word, caso já exista um arquivo com este nome, é substituido
                exemploDoc.SaveToFile(@"saida\exemplo_arquivo_word.docx", FileFormat.Docx);
            #endregion

        }
    }
}
