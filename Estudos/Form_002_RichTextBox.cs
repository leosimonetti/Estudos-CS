using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Section = Microsoft.Office.Interop.Word.Section;
using DocumentFormat.OpenXml.Packaging;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml;
using Spire.Doc;
using Spire.Doc.Documents;

namespace Estudos
{
    public partial class Form_002_RichTextBox : Form
    {
        public Form_002_RichTextBox()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // Cria aplicação Word
            Word.Application appWord = new Word.Application();
            appWord.Visible = true;
            appWord.WindowState = Word.WdWindowState.wdWindowStateNormal;

            // Cria Documento do Word
            Word.Document docWord = appWord.Documents.Add();

            // Aciciona o cabeçalho.
            foreach (Section section in docWord.Sections)
            {
                // Procura o range do cabeçalho e adiciona os detalhes.
                Word.Range headerRange = section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                headerRange.Font.ColorIndex = Word.WdColorIndex.wdBlue;
                headerRange.Font.Size = 10;
                //headerRange.Text = richTextBox1.Text;
                headerRange.Text = richTextBox1.Text;
            }


            // Adiciona Rodapé
            foreach (Section wordSection in docWord.Sections)
            {
                //Consegue o range do rodapé e adiciona os detalhes.
                Word.Range footerRange = wordSection.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Word.WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Text = richTextBox3.Text;
            }

            //Adiciona texto do documento
            docWord.Content.SetRange(0, 0);
            docWord.Content.Text = "Teste de adição de uma linha." + Environment.NewLine;

            //Adiciona um parágrafo com estilo Heading 1 style
            Word.Paragraph para1 = docWord.Content.Paragraphs.Add();
            object styleHeading1 = "Título 1";
            para1.Range.set_Style(ref styleHeading1);
            para1.Range.Text = richTextBox2.Text;
            para1.Range.InsertParagraphAfter();

            //Adiciona um parágrafo com estilo Heading 2 style
            Word.Paragraph para2 = docWord.Content.Paragraphs.Add();
            object styleHeading2 = "Normal";
            para2.Range.set_Style(ref styleHeading2);
            para2.Range.Text = richTextBox2.Text;
            para2.Range.InsertParagraphAfter();

            //Cria uma tabela 5X5 e insere alguns dados
            Word.Table firstTable = docWord.Tables.Add(para1.Range, 5, 5, null, null);

            firstTable.Borders.Enable = 1;
            foreach (Word.Row row in firstTable.Rows)
            {
                foreach (Word.Cell cell in row.Cells)
                {
                    //Header row
                    if (cell.RowIndex == 1)
                    {
                        cell.Range.Text = "Coluna " + cell.ColumnIndex.ToString();
                        cell.Range.Font.Bold = 1;
                        //other format properties goes here
                        cell.Range.Font.Name = "verdana";
                        cell.Range.Font.Size = 10;
                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                            
                        cell.Shading.BackgroundPatternColor = Word.WdColor.wdColorGray25;
                        //Center alignment for the Header cells
                        cell.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    }
                    //Data row
                    else
                    {
                        cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                    }
                }
            }

            //Salva o documento e fecha o Word
            //object filename = @"c:\TEMP\temp1.docx";
            //docWord.SaveAs2(ref filename);
            //docWord.Close();
            //docWord = null;
            //appWord.Quit();
            //appWord = null;
            //MessageBox.Show("Documento criado com sucesso!");

        }

        private void button2_Click(object sender, EventArgs e)
        {

            //AddHeaderFromTo(@"C:\Temp\Teste1.docx", @"C:\Temp\Teste2.docx");

            using (WordprocessingDocument doc = WordprocessingDocument.Open(@"C:\Temp\Teste2.docx", true))
            {
                string altChunkId = "AltChunkId5";

                MainDocumentPart mainDocPart = doc.MainDocumentPart;
                AlternativeFormatImportPart chunk = mainDocPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Rtf, altChunkId);

                string rtfEncodedString = richTextBox1.Rtf;

                using (MemoryStream ms = new MemoryStream(Encoding.ASCII.GetBytes(rtfEncodedString)))
                {
                    chunk.FeedData(ms);
                }

                AltChunk altChunk = new AltChunk();
                altChunk.Id = altChunkId;

                // Funciona adicionando na sequencia
                //mainDocPart.Document.Body.InsertAfter(altChunk, mainDocPart.Document.Body.Elements<Paragraph>().Last());
                
                
                //teste para adicionar no cabeçalho
                //mainDocPart.Document.Body.InsertAfter(altChunk, mainDocPart.Document.Body.Elements<Paragraph>().Last());

                // itera os elementos do documento
                //var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>();
                // Iterate through paragraphs, runs, and text, finding the text we want and replacing it
                //foreach (Paragraph paragraph in paragraphs)
                //{
                //    foreach (Run run in paragraph.Elements<Run>())
                //    {
                //        foreach (Text text in run.Elements<Text>())
                //        {
                //            if (text.Text == "CABEÇALHO")
                //            {
                //                text.Text = "CABEÇALHO 2";
                //           }
                //        }
                //    }
                //}

                mainDocPart.Document.Save();

            }
        }

        public static void AddHeaderFromTo(string filepathFrom, string filepathTo)
        {
            // Replace header in target document with header of source document.
            using (WordprocessingDocument wdDoc = WordprocessingDocument.Open(filepathTo, true))
            {
                MainDocumentPart mainPart = wdDoc.MainDocumentPart;

                // Delete the existing header part.
                mainPart.DeleteParts(mainPart.HeaderParts);

                // Create a new header part.
                HeaderPart headerPart = mainPart.AddNewPart<HeaderPart>();

                // Get Id of the headerPart.
                string rId = mainPart.GetIdOfPart(headerPart);

                // Feed target headerPart with source headerPart.
                using (WordprocessingDocument wdDocSource = WordprocessingDocument.Open(filepathFrom, true))
                {
                    HeaderPart firstHeader = wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

                    wdDocSource.MainDocumentPart.HeaderParts.FirstOrDefault();

                    if (firstHeader != null)
                    {
                        headerPart.FeedData(firstHeader.GetStream());
                    }
                }

                // Get SectionProperties and Replace HeaderReference with new Id.
                IEnumerable<SectionProperties> sectPrs = mainPart.Document.Body.Elements<SectionProperties>();
                foreach (var sectPr in sectPrs)
                {
                    // Delete existing references to headers.
                    sectPr.RemoveAllChildren<HeaderReference>();

                    // Create the new header reference node.
                    sectPr.PrependChild<HeaderReference>(new HeaderReference() { Id = rId });
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Spire.Doc.Document docSpire = new Spire.Doc.Document(@"C:\TEMP\Teste6.docx");
            //Spire.Doc.Document docSpire = new Spire.Doc.Document();
            Spire.Doc.Section secao = docSpire.Sections[0];
            Spire.Doc.HeaderFooter cabecalho = secao.HeadersFooters.Header;
            Spire.Doc.Documents.Paragraph para1 = cabecalho.AddParagraph();
            para1.AppendRTF(richTextBox1.Rtf);

            //Spire.Doc.Section secao2 = docSpire.AddSection();
            Spire.Doc.Documents.Paragraph para2 = secao.AddParagraph();
            para2.AppendRTF(richTextBox2.Rtf);

            docSpire.SaveToFile(@"C:\TEMP\Teste6.docx");
            
        }

        private void button4_Click(object sender, EventArgs e)
        {
            int testeGit;
            testeGit = 1;

            richTextBox1.Text = testeGit.ToString();

        }
    }
}
