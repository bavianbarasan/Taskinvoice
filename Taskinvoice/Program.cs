using System;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;

class Program
{
    static void Main()
    {
        string outputFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Invoice.pdf");
        Document document = new Document(PageSize.A3);

        try
        {
            PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(outputFilePath, FileMode.Create));
            MyPageEventHandler pageEventHandler = new MyPageEventHandler();
            writer.PageEvent = pageEventHandler;
            document.Open();

            
            BaseColor lightBlueColor = new BaseColor(0, 0, 0);
            BaseColor darkGrayColor = new BaseColor(64, 64, 64);
            BaseColor lightGrayColor = new BaseColor(192, 192, 192);

           
            string logoPath = @"C:\Users\HP\Downloads\download.png";
            if (File.Exists(logoPath))
            {
                Image logo = Image.GetInstance(logoPath);
                logo.ScaleAbsolute(150, 100);
                logo.Alignment = Image.ALIGN_RIGHT;
                document.Add(logo);
            }
            else
            {
                Console.WriteLine("Logo file not found at " + logoPath);
            }

           
            Font titleFont = FontFactory.GetFont("Arial", 20, Font.BOLD, lightBlueColor);
            Paragraph title = new Paragraph("Invoice", titleFont);
            title.Alignment = Element.ALIGN_CENTER;
            document.Add(title);

            
            document.Add(new Paragraph("\n"));

           
            Font headerFont = FontFactory.GetFont("Arial", 16, Font.BOLD, darkGrayColor);
            Paragraph header = new Paragraph("STAR STAFFING", headerFont);
            header.Alignment = Element.ALIGN_LEFT;
            header.SpacingBefore = 10;
            document.Add(header);

           
            Font normalFont = FontFactory.GetFont("Arial", 12, darkGrayColor);
            Paragraph address = new Paragraph("2712 Okmulgee Ave.\nOklahoma City, \n73102\n405-1212-1221\naccounting@starstaffing.com", normalFont);
            address.SpacingBefore = 5;
            document.Add(address);

          
            PdfPTable invoiceTable = new PdfPTable(2);
            invoiceTable.HorizontalAlignment = Element.ALIGN_RIGHT;
            invoiceTable.SpacingBefore = 20;
            invoiceTable.SpacingAfter = 20;
            invoiceTable.WidthPercentage = 50;
            invoiceTable.DefaultCell.Border = Rectangle.NO_BORDER;

            invoiceTable.AddCell(new PdfPCell(new Phrase("Invoice No:", normalFont)) { Border = Rectangle.NO_BORDER });
            invoiceTable.AddCell(new PdfPCell(new Phrase("042023001", normalFont)) { Border = Rectangle.NO_BORDER });

            invoiceTable.AddCell(new PdfPCell(new Phrase("Invoice Date:", normalFont)) { Border = Rectangle.NO_BORDER });
            invoiceTable.AddCell(new PdfPCell(new Phrase("04/01/2024", normalFont)) { Border = Rectangle.NO_BORDER });

            invoiceTable.AddCell(new PdfPCell(new Phrase("Due Date:", normalFont)) { Border = Rectangle.NO_BORDER });
            invoiceTable.AddCell(new PdfPCell(new Phrase("05/01/2024", normalFont)) { Border = Rectangle.NO_BORDER });

            invoiceTable.AddCell(new PdfPCell(new Phrase("Terms:", normalFont)) { Border = Rectangle.NO_BORDER });
            invoiceTable.AddCell(new PdfPCell(new Phrase("2/10 Net 30", normalFont)) { Border = Rectangle.NO_BORDER });

            document.Add(invoiceTable);

            
            Paragraph billTo = new Paragraph("BILL TO\nJeremy Owens\nCentral Healthcare Services\n151 Anderson St.\nTulsa, OK\n74101", normalFont);
            document.Add(billTo);

           
            PdfPTable table = new PdfPTable(4);
            table.WidthPercentage = 100;
            table.SpacingBefore = 20f;
            table.SpacingAfter = 20f;
            table.SplitLate = false; 
            table.SplitRows = true; 

            float[] columnWidths = new float[] { 2f, 1f, 1f, 1f };
            table.SetWidths(columnWidths);

            Font tableHeaderFont = FontFactory.GetFont("Arial", 12, Font.BOLD, BaseColor.WHITE);
            PdfPCell headerCell = new PdfPCell(new Phrase("Description", tableHeaderFont));
            headerCell.BackgroundColor = darkGrayColor;
            headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
            table.AddCell(headerCell);

            headerCell = new PdfPCell(new Phrase("Hours", tableHeaderFont));
            headerCell.BackgroundColor = darkGrayColor;
            headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
            table.AddCell(headerCell);

            headerCell = new PdfPCell(new Phrase("Rate", tableHeaderFont));
            headerCell.BackgroundColor = darkGrayColor;
            headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
            table.AddCell(headerCell);

            headerCell = new PdfPCell(new Phrase("Amount", tableHeaderFont));
            headerCell.BackgroundColor = darkGrayColor;
            headerCell.HorizontalAlignment = Element.ALIGN_CENTER;
            table.AddCell(headerCell);

            table.HeaderRows = 1;

            
            for (int i = 1; i <= 200; i++)
            {
                PdfPCell dataCell = new PdfPCell(new Phrase($"Staffing - Employee {i}", normalFont));
                dataCell.HorizontalAlignment = Element.ALIGN_LEFT;
                table.AddCell(dataCell);
                dataCell = new PdfPCell(new Phrase("40", normalFont));
                dataCell.HorizontalAlignment = Element.ALIGN_CENTER;
                table.AddCell(dataCell);
                dataCell = new PdfPCell(new Phrase("$30", normalFont));
                dataCell.HorizontalAlignment = Element.ALIGN_CENTER;
                table.AddCell(dataCell);
                dataCell = new PdfPCell(new Phrase("$1200.00", normalFont));
                dataCell.HorizontalAlignment = Element.ALIGN_CENTER;
                table.AddCell(dataCell);
            }

            document.Add(table);

            
            PdfPTable totalsTable = new PdfPTable(2);
            totalsTable.WidthPercentage = 30;
            totalsTable.HorizontalAlignment = Element.ALIGN_RIGHT;
            totalsTable.SpacingBefore = 20f;

            PdfPCell totalCell = new PdfPCell(new Phrase("Sub Total", normalFont));
            totalCell.Border = Rectangle.NO_BORDER;
            totalsTable.AddCell(totalCell);
            totalCell.Phrase = new Phrase("$240000.00", normalFont);
            totalsTable.AddCell(totalCell);

            totalCell.Phrase = new Phrase("Total Due (With Discount)", normalFont);
            totalsTable.AddCell(totalCell);
            totalCell.Phrase = new Phrase("$2380000.50", normalFont);
            totalsTable.AddCell(totalCell);

            document.Add(totalsTable);

            
            Paragraph footer = new Paragraph("Thank you for your business!", normalFont);
            footer.Alignment = Element.ALIGN_CENTER;
            footer.SpacingBefore = 30;
            document.Add(footer);

            
            Paragraph paymentInfo = new Paragraph("Payment Information\n\nPayment by Mail:\nOliver Public Relations\n1470 Jackson St.\nNashville, TN, USA\n37250\n\nPayment by ACH:\nBank: Bank Name\nRouting #: 1000 111 222\nAccount #: 000 111 2222", normalFont);
            paymentInfo.SpacingBefore = 20;
            document.Add(paymentInfo);

            
            Paragraph termsHeader = new Paragraph("Terms and Conditions", headerFont);
            termsHeader.Alignment = Element.ALIGN_CENTER;
            termsHeader.SpacingBefore = 280;
            document.Add(termsHeader);

            string termsText = "1. Payment is due within 30 days from the date of the invoice.\n\n" +
                               "2. A late fee of 1.5% per month will be applied to overdue balances.\n\n" +
                               "3. Please include the invoice number on your check or ACH payment.\n\n" +
                               "4. All services provided are subject to the terms and conditions outlined in our agreement.\n\n" +
                               "5. If you have any questions regarding this invoice, please contact our accounting department at accounting@starstaffing.com.\n\n" +
                               "6. Our business hours are Monday to Friday, 9 AM to 5 PM CST.";

            Paragraph terms = new Paragraph(termsText, normalFont);
            terms.SpacingBefore = 10;
            document.Add(terms);
        }
        catch (DocumentException de)
        {
            Console.Error.WriteLine(de.Message);
        }
        catch (IOException ioe)
        {
            Console.Error.WriteLine(ioe.Message);
        }
        finally
        {
            document.Close();
        }

        Console.WriteLine("PDF Created successfully at " + outputFilePath);
    }
}

public class MyPageEventHandler : PdfPageEventHelper
{
    PdfTemplate total;
    BaseFont bf = null;
    PdfContentByte cb;

    public override void OnOpenDocument(PdfWriter writer, Document document)
    {
        cb = writer.DirectContent;
        total = cb.CreateTemplate(50, 50);
        try
        {
            bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
        }
        catch (DocumentException e)
        {
            throw new Exception(e.Message);
        }
        catch (IOException e)
        {
            throw new Exception(e.Message);
        }
    }

    public override void OnEndPage(PdfWriter writer, Document document)
    {
        int pageN = writer.PageNumber;
        String text = "Page " + pageN + " of ";
        float len = bf.GetWidthPoint(text, 12);
        Rectangle pageSize = document.PageSize;

        cb.BeginText();
        cb.SetFontAndSize(bf, 12);
        cb.SetTextMatrix(pageSize.GetLeft(40), pageSize.GetBottom(30));
        cb.ShowText(text);
        cb.EndText();
        cb.AddTemplate(total, pageSize.GetLeft(40) + len, pageSize.GetBottom(30));
    }

    public override void OnCloseDocument(PdfWriter writer, Document document)
    {
        total.BeginText();
        total.SetFontAndSize(bf, 12);
        total.SetTextMatrix(0, 0);
        total.ShowText((writer.PageNumber - 1).ToString());
        total.EndText();
    }
}
