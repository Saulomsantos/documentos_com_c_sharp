using System;
using System.Diagnostics;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace TechTalk.Word.Teste1
{
    class Program
    {
        static void Main(string[] args)
        {
            #region Criar documento
                // Cria um documento word
                // Instancia o objeto exemplodedoc na classe Document
                // Este objeto é o documento Word propriamente dito
                Document exemplodedoc = new Document();
            #endregion
            
            #region Criar seção no documento
                // Adicionando uma seção ao documento
                // Instancia o objeto secaoPagina1 na classe Section 
                // e utiliza o metodo AddSection() para adicionar esta seção ao documento
                // cada seção pode ser interpretada como uma quebra de seção do Word, continuando em outra página do documento
                Section secaoCapa = exemplodedoc.AddSection();
            #endregion
            
            #region Criar parágrafo
                // Adicionando um parágrafo
                // Instancia o objeto titulo na classe Paragraph
                // e utiliza o metodo AddParagraph() para adicionar este paragrafo ao objeto secao criado anteriormente
                Paragraph titulo = secaoCapa.AddParagraph();
            #endregion
            
            #region Adicionar texto ao parágrafo
                // Adicionando texto ao parágrafo criado
                // Utiliza o metodo AppendText() para adicionar um texto ao paragrafo do objeto titulo criado anteriormente 
                titulo.AppendText("Exemplo de título\n\n");
                // Alinhando o texto. Alinhamento centralizado
                titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;
                // Criando um estilo de formatação
                ParagraphStyle estilo1 = new ParagraphStyle(exemplodedoc);
                // Definindo o nome do estilo
                estilo1.Name = "Cor do título";
                // Definindo a cor do texto
                estilo1.CharacterFormat.TextColor = Color.DarkBlue;
                estilo1.CharacterFormat.Bold = true;
                exemplodedoc.Styles.Add(estilo1);
                // Aplicando o estilo
                titulo.ApplyStyle(estilo1.Name);
            #endregion
            

            // Adicionando um novo parágrafo
            Paragraph paragrafoCapa1 = secaoCapa.AddParagraph();

            // Adicionando texto a este parágrafo
            // '\n' pula linha e '\t' é uma tabulação (tecla tab)
            paragrafoCapa1.AppendText("\tEste é um exemplo de criação de um parágrafo utilizando a biblioteca Spire.Doc.\n");

            Paragraph paragrafoCapa2 = secaoCapa.AddParagraph();

            paragrafoCapa2.AppendText("\tBasicamente, então, uma seção representa uma página e os parágrafos dentro de uma mesma seção, " +
            "obviamente, aparecem na mesma página.");

            // Inserindo imagens
            #region Inserir imagens
                Paragraph paragrafoCapaImagem = secaoCapa.AddParagraph();
                paragrafoCapaImagem.AppendText("\n\n\tAgora vamos inserir uma imagem em um parágrafo\n\n");
                paragrafoCapaImagem.Format.HorizontalAlignment = HorizontalAlignment.Center;
                DocPicture imagemExemplo = paragrafoCapaImagem.AppendPicture(Image.FromFile(@"D:\SENAI Info\Lógica de programação\C#\TechTalk\Codigo\Teste\TechTalk.Word.Teste1\img\cSharpLogo.png"));
                imagemExemplo.Width = 300;
                imagemExemplo.Height = 300;
            #endregion

            // Adicionando uma nova seção (página 2)
            Section secaoCorpo = exemplodedoc.AddSection();

            //Adicionando um novo parágrafo
            Paragraph paragrafoCorpo1 = secaoCorpo.AddParagraph();

            //Adicionando texto a este novo parágrafo
            paragrafoCorpo1.AppendText("\tEste é um exemplo de criação de um parágrafo em uma nova página, após uma quebra de seção. " +
            "Assim como quando utilizamos variáveis, é possível fechar aspas, inserir um sinal '+' e continuar o parágrafo.\n\n" +
            "\tComo foi criada outra seção, percbeba que o parágrafo acima começou em outra página");

            #region Salvar arquivo
                // Salvando o arquivo
                // Utiliza o metodo SaveToFile para salvar o arquivo com o nome e o formato escolhido
                // Assim como no Word, caso já exista um arquivo com o mesmo nome este será substituído por um novo
                exemplodedoc.SaveToFile(@"C:\word\exemplo_de_arquivo_Word.docx", FileFormat.Docx);
            #endregion
            

            // Abrindo o arquivo
            ProcessStartInfo info = new ProcessStartInfo
            {
                FileName = @"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
                Arguments = @"C:\word\exemplo_de_arquivo_Word.docx"
            };
            Process.Start(info);
        }
    }
}
