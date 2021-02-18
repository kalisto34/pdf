using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MigraDoc;
using PdfSharp;
using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using PdfSharp.Pdf;
using MigraDoc.DocumentObjectModel.Shapes;
using MigraDoc.DocumentObjectModel.Tables;

namespace pdf
{
    class Program
    {
        
        static public void AddHeader(Document document, string text)
        {
            Table Header_text = new Table();
            Section section = document.LastSection;
            Header_text.Borders.Width = .1;
            Header_text.Borders.Color = Colors.Gray;
            Header_text.Borders.Left.Visible = false;
            Header_text.Borders.Right.Visible = false;
            Header_text.Borders.Top.Visible = false;
            Column column2 = Header_text.AddColumn("17.2cm");
            Row row2 = Header_text.AddRow();
            row2.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row2.Cells[0].AddParagraph(text);
            row2.Style = "Heading1";
            row2.Height = 38.0;
            section.Add(Header_text); ///не работаетd
        }

        static void AddToList(Document document, string text, bool continueList)
        {
            ListInfo listinfo = new ListInfo();
            listinfo.ContinuePreviousList = continueList;
            listinfo.ListType = ListType.BulletList1;
            Paragraph paragraph = document.LastSection.AddParagraph(text);
            paragraph.Format.ListInfo = listinfo;
            continueList = true;
        }

        static public void AddImage_w_cap(Document document, string address, string caption_text)
        {
            Section section = document.LastSection;
            Image image = new Image(address);
            image.Left = ShapePosition.Center;
            section.Add(image);
            Paragraph caption = section.AddParagraph();
            caption.AddText(caption_text);
            caption.Format.Alignment = ParagraphAlignment.Center;
          
        }
        static void Main(string[] args)
        {
            Document document = new Document();
            Section Header = document.AddSection();
            document.DefaultPageSetup.LeftMargin = MigraDoc.DocumentObjectModel.Unit.FromCentimeter(2.1);
            document.DefaultPageSetup.RightMargin = MigraDoc.DocumentObjectModel.Unit.FromCentimeter(2.1);
            document.DefaultPageSetup.TopMargin = MigraDoc.DocumentObjectModel.Unit.FromCentimeter(1.7);

            Font Sub_Header = new Font();
            //Sub_Header.Color = Color.Parse("0x2BBE6");
            Sub_Header.Color = Colors.SkyBlue;
            Sub_Header.Size = 16.0;

            Style style = document.Styles["Normal"];
            style.Font.Name = "Times New Roman";

            style = document.Styles["Heading1"];
            style.Font = Sub_Header;
            style.Font.Name = "Tahoma";
            style.ParagraphFormat.PageBreakBefore = true;
            style.ParagraphFormat.SpaceAfter = 6;


            Image Logo = new Image("../img/iLeco_logo.png");
            Image Cordium = new Image("../img/Cordium.png");

            Paragraph paragraph_f = new Paragraph();
            paragraph_f.AddPageField();
            paragraph_f.AddText("/");
            paragraph_f.AddNumPagesField();

            Paragraph paragraph_f1 = new Paragraph();
            paragraph_f1.AddText("i.Leco ©");
            var dateAndTime = DateTime.Now;
            var date = dateAndTime.Date;
            paragraph_f1.AddText(date.ToShortDateString());
            Table Footer = new Table();

            Footer.Borders.Width = .1;
            Footer.Borders.Color = Colors.Gray;
            Footer.Borders.Left.Visible = false;
            Footer.Borders.Right.Visible = false;
            Footer.Borders.Top.Visible = false;
            Footer.Borders.Bottom.Visible = false;
            Column column_F = Footer.AddColumn("7.5cm");
            column_F = Footer.AddColumn("9.5cm");


            Row row_F = Footer.AddRow();
            row_F.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row_F.Cells[0].Add(paragraph_f1);
            row_F.Height = 38.0;
            row_F.Cells[1].Add(paragraph_f);
            row_F.Cells[1].VerticalAlignment = VerticalAlignment.Bottom;
            row_F.Cells[1].Format.Alignment = ParagraphAlignment.Right;
            Header.Footers.Primary.Add(Footer);

            Table Header_t = new Table();
            Header_t.Borders.Width = 1.5;
            Header_t.Borders.Left.Visible = false;
            Header_t.Borders.Right.Visible = false;
            Header_t.Borders.Top.Visible = false;
            Column column = Header_t.AddColumn("7.5cm");
            column.Format.Alignment = ParagraphAlignment.Left;
            column = Header_t.AddColumn("9.5cm");
            column.Format.Alignment = ParagraphAlignment.Left;



            Row row = Header_t.AddRow();
            row.Cells[0].MergeDown = 1;
            row.Cells[0].Add(Logo);
            row.Height = 38.0;
            row.Format.Font.Bold = true;
           // row.Format.Font.Color = Color.Parse("0x12406B");
            row.Format.Font.Color = Colors.DarkBlue;
            row.Format.Font.Size = 18.0;
            row.Cells[1].VerticalAlignment = VerticalAlignment.Bottom;
            row.Borders.Bottom.Visible = false;
            row.Cells[1].AddParagraph("Energy Service Report");
            row = Header_t.AddRow();
            row.Style = "Heading1";
            row.Cells[1].AddParagraph("Cordium: Real-time heat control");
            row.Cells[1].VerticalAlignment = VerticalAlignment.Center;

            Logo.WrapFormat.Style = WrapStyle.Through;
            Logo.RelativeHorizontal = RelativeHorizontal.Margin;
            //    Header_t.SetEdge(0, 0, Header_t.Columns.Count, 1, Edge.Box, BorderStyle.Single, 2, Colors.Black);

            Header.Add(Header_t);


            Table Header_info = new Table();

            Header_info.Borders.Width = .1;
            Header_info.Borders.Color = Colors.Gray;
            Header_info.Borders.Left.Visible = false;
            Header_info.Borders.Right.Visible = false;
            Header_info.Borders.Top.Visible = false;
            Column column1 = Header_info.AddColumn("8cm");
            column1 = Header_info.AddColumn("9cm");


            Row row1 = Header_info.AddRow();
            row1.Cells[0].VerticalAlignment = VerticalAlignment.Bottom;
            row1.Cells[0].AddParagraph("Peroid:");
            row1.Style = "Heading1";
            //row1.Cells[0].Format.Font.Color = Color.Parse("0x2BBE6");
            row1.Cells[0].Format.Font.Color = Colors.SkyBlue;
            row1.Height = 38.0;
            row1.Cells[1].AddParagraph("December, 2020");
            row1.Style = "Heading1";
            row1.Cells[1].Format.Font.Color = Colors.Gray;
            row1.Cells[1].VerticalAlignment = VerticalAlignment.Bottom;
            row1.Cells[1].Format.Alignment = ParagraphAlignment.Right;
            

            row1 = Header_info.AddRow();
            row1.Style = "Heading1";
            row1.Cells[0].AddParagraph("Project Details:"); // пропадает картинка
            row.Cells[0].VerticalAlignment = VerticalAlignment.Center;

            Header.Add(Header_info);
            Header.AddParagraph("   ");
            Header.Add(Cordium);
            AddToList(document, "Project Manager:Frank Louwet", false);
            AddToList(document, "Location: Crutzestraat, Hasselt", true);


            AddHeader(document, "Executive Summary:");
        
            Header.AddParagraph("");
            Header.AddParagraph("This month there werea total of 345 degree days");
            Header.AddParagraph("Phase 1 heating energy:");
            Header.AddParagraph("");
            AddToList(document, "2196 kWh of gas consumed (0.32 kWh per apartment per degree day)", false);
            
            Header.AddParagraph("");
            Header.AddParagraph("Phase 2 heating energy:");
            AddToList(document, "14515 kWh of gas consumed (2.1 kWh per apartment per degree day)", false);
            AddToList(document, "3.5 kWh of electricity consumed (0.00051 kWh per apartment per degree day)", true);

            
            Header.AddParagraph("");
            Header.AddParagraph("Phase 3 heating energy:");
            AddToList(document, "17791 kWh of gas consumed (1.8 kWh per apartment per degree day)", false);
            AddToList(document, "650 kWh of electricity produced", true);
           
            

            AddHeader(document, "Project overview:");
            Header.AddParagraph("Theadvanced control strategy is implemented in a district heating system for social housing in Crutzestraat, Hasselt.Thesocial housing is operated by Cordium, the operating manager for social housing in Flemish region.The projectconsists of three phases or buildings with 20, 20 and 28 apartments in each phase.Each building has its own central heating system with various technologies installed.Furthermore,central heating systems areinterconnected by an internal heat transfer network. i.Leco developed thecontrol strategy which sends hourly setpoints fo: maximum and minimum temperaturesetpoint in each building and/or distribution circuit, operation modes of installed technologies, and distribution statesettings between building/heating systems.");
          
            
            Section Body = document.AddSection();

            AddHeader(document, "Phase 1");
            Body.AddParagraph("Installed technologies:");
            Body.AddParagraph("");
            AddToList(document, "Geothermal/water gas absorption heat pumps – 2 pcs", false);
            
            //Image image = new Image("../img/img1.png");
            //Header.Add(image);
            //Paragraph caption = Header.AddParagraph();
            //caption.AddText("Figure 1: Phase 1 Energy Diagram");
            //caption.Format.Alignment = ParagraphAlignment.Center;

            AddImage_w_cap(document, "../img/img1.png", "Figure 1: Phase 1 Energy Diagram");

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true, PdfFontEmbedding.Always);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            pdfRenderer.PdfDocument.Save("../NewDoc.pdf");




        }

    }
}
