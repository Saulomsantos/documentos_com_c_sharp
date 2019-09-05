using System;
using System.Diagnostics;
using System.Drawing;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;

namespace Exemplo
{
    class Program
    {
        static void Main(string[] args)
        {
            // Criando um documento
            Document exemploDoc = new Document();

            // Criando uma seção
            Section secaoCapa = exemploDoc.AddSection();

            // Criando um parágrafo
            Paragraph titulo = secaoCapa.AddParagraph();

            // Adicionando um texto ao parágrafo
            titulo.AppendText("Exemplo de título\n\n");

            // Alinhando o texto ao centro
            titulo.Format.HorizontalAlignment = HorizontalAlignment.Center;

            // Inserindo uma imagem
            Paragraph paragrafoImagem = secaoCapa.AddParagraph();

            DocPicture imagemExemplo = paragrafoImagem.AppendPicture(Image.FromFile(@"D:\SENAI Info\Lógica de programação\C#\TechTalk\Codigo\Teste\TechTalk.Word.Teste1\img\cSharpLogo.png"));

            // Criando uma nova seção
            Section secaoCorpo = exemploDoc.AddSection();

            Paragraph paragrafoCorpo = secaoCorpo.AddParagraph();

            paragrafoCorpo.AppendText("Um exemplo de um texto dentro de um parágrafo" +
            "em uma nova seção");

            exemploDoc.SaveToFile(@"D:\SENAI Info\Lógica de programação\C#\TechTalk\Codigo\Exemplo\exemplo_arquivo_doc.html", FileFormat.Html);
        }
    }
}
