using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;

class Program
{
    static void Main(string[] args)
    {
        string titulo = "Davi dispõe-se a pelejar contra o gigante - Davi encontra-se com o gigante e mata-o";
        List<string> palavras = new List<string> {
            "Guerra", "Filisteu", "Isrraelista", "Jessé", "Davi", "Saul", "Golias", "Pelejar", "Moço",
            "Apascenta", "Ovelha", "Urso", "Leão", "Cajado", "Seixos", "Funda", "Alforje", "Pedra",
            "Testa", "Espada", "Cabeça", "Israel", "Judá", "Jerusalém"
        };

        CriarCacaPalavrasDocx(titulo, palavras, 20);
    }

    static void CriarCacaPalavrasDocx(string titulo, List<string> palavras, int tamanho)
    {
        // Criar o documento .docx
        string filePath = "caca_palavras_davi.docx";
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            // Adicionar o título
            MainDocumentPart mainPart = wordDoc.AddMainDocumentPart();
            mainPart.Document = new Document();
            Body body = mainPart.Document.AppendChild(new Body());

            Paragraph heading = body.AppendChild(new Paragraph());
            Run headingRun = heading.AppendChild(new Run());
            headingRun.AppendChild(new Text(titulo));

            // Criar a matriz de letras
            char[,] grade = new char[tamanho, tamanho];
            PreencherMatrizComPalavras(grade, palavras, tamanho);

            // Preencher as células vazias com letras aleatórias
            Random random = new Random();
            string alfabeto = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            for (int i = 0; i < tamanho; i++)
            {
                for (int j = 0; j < tamanho; j++)
                {
                    if (grade[i, j] == '\0') // se a célula estiver vazia
                    {
                        grade[i, j] = alfabeto[random.Next(alfabeto.Length)];
                    }
                }
            }

            // Adicionar a matriz ao documento
            for (int i = 0; i < tamanho; i++)
            {
                string linha = "";
                for (int j = 0; j < tamanho; j++)
                {
                    linha += grade[i, j] + " ";
                }
                Paragraph para = body.AppendChild(new Paragraph());
                Run run = para.AppendChild(new Run());
                run.AppendChild(new Text(linha.Trim()));
            }

            // Adicionar as palavras a serem encontradas
            Paragraph palavraPara = body.AppendChild(new Paragraph());
            Run palavraRun = palavraPara.AppendChild(new Run());
            palavraRun.AppendChild(new Text("\nPalavras a serem encontradas:"));

            Paragraph listaPara = body.AppendChild(new Paragraph());
            Run listaRun = listaPara.AppendChild(new Run());
            listaRun.AppendChild(new Text(string.Join(", ", palavras)));
        }

        Console.WriteLine($"Caça-palavras gerado e salvo em {filePath}");
    }

    static void PreencherMatrizComPalavras(char[,] grade, List<string> palavras, int tamanho)
    {
        Random random = new Random();

        foreach (string palavra in palavras)
        {
            string palavraUpper = palavra.ToUpper();
            int lenPalavra = palavraUpper.Length;
            bool posicionado = false;

            while (!posicionado)
            {
                int direcao = random.Next(2); // 0 = horizontal, 1 = vertical
                int linhaInicio = random.Next(tamanho - lenPalavra);
                int colunaInicio = random.Next(tamanho - lenPalavra);

                if (direcao == 0) // Horizontal
                {
                    bool podePosicionar = true;
                    for (int i = 0; i < lenPalavra; i++)
                    {
                        if (grade[linhaInicio, colunaInicio + i] != '\0')
                        {
                            podePosicionar = false;
                            break;
                        }
                    }

                    if (podePosicionar)
                    {
                        for (int i = 0; i < lenPalavra; i++)
                        {
                            grade[linhaInicio, colunaInicio + i] = palavraUpper[i];
                        }
                        posicionado = true;
                    }
                }
                else // Vertical
                {
                    bool podePosicionar = true;
                    for (int i = 0; i < lenPalavra; i++)
                    {
                        if (grade[linhaInicio + i, colunaInicio] != '\0')
                        {
                            podePosicionar = false;
                            break;
                        }
                    }

                    if (podePosicionar)
                    {
                        for (int i = 0; i < lenPalavra; i++)
                        {
                            grade[linhaInicio + i, colunaInicio] = palavraUpper[i];
                        }
                        posicionado = true;
                    }
                }
            }
        }
    }
}
