using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Spire.Doc;
using Spire.Doc.Documents;
using Spire.Doc.Fields;
using Spire.Doc.Formatting;

namespace FreeSpire.Controllers
{
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class SpireController : ControllerBase
    {
        [HttpGet]
        public IActionResult CreateSpireWord()
        {
            try
            {
                var document = new Document();
                //var section = document.AddSection();
                //var para = section.AddParagraph();
                //para.AppendText("aaa");
                //para.AppendBreak(Spire.Doc.Documents.BreakType.LineBreak);
                //para.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.Toc1);

                var section_2 = document.AddSection();
                var para_2 = section_2.AddParagraph();
                var text = para_2.AppendText("bbbcmdlsjvkodajv;oeaihvj;");
                text.CharacterFormat.TextColor = System.Drawing.Color.Red;
                para_2.ApplyStyle(Spire.Doc.Documents.BuiltinStyle.BodyText2);

                document.SaveToFile("elif.docx", FileFormat.Docx);
                return Ok("elif");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        public IActionResult CheckFreeSpire()
        {
            try
            {
                var document = new Document();
                document.LoadFromFile("Nemati.docx");
                return Ok("ok");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        public IActionResult CheckOnline()
        {
            try
            {
                Document doc = new Document();
                Section sec = doc.AddSection();
                Paragraph par = sec.AddParagraph();
                TextBox textBox = par.AppendTextBox(180, 30);
                textBox.Format.VerticalOrigin = VerticalOrigin.Margin;
                textBox.Format.VerticalPosition = 100;
                textBox.Format.HorizontalOrigin = HorizontalOrigin.Margin;
                textBox.Format.HorizontalPosition = 50;
                textBox.Format.NoLine = true;
                CharacterFormat format = new CharacterFormat(doc);
                format.FontName = "Calibri";
                format.FontSize = 15;
                format.Bold = true;
                Paragraph par1 = textBox.Body.AddParagraph();
                par1.AppendText("This is my new string").ApplyCharacterFormat(format);
                doc.SaveToFile("result.docx", FileFormat.Docx);
                return Ok("ok");
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

        [HttpGet]
        public IActionResult NeverGiveUp()
        {
            try
            {
                Document doc = new Document();
                Section sec = doc.AddSection();

                Paragraph main = sec.AddParagraph();
                TextBox mainTextBox = main.AppendTextBox(100, 20);
                mainTextBox.Format.HorizontalAlignment = ShapeHorizontalAlignment.Center;
                mainTextBox.Format.NoLine = true;
                Paragraph ichki = mainTextBox.Body.AddParagraph();
                ichki.AppendText("Приказ № 34-к");

                 Paragraph forStaff = sec.AddParagraph();
                TextBox staffniIsmi = forStaff.AppendTextBox(100, 20);
                staffniIsmi.Format.NoLine = true;
                staffniIsmi.Format.HorizontalAlignment = ShapeHorizontalAlignment.Left;
                Paragraph staffNameValue = staffniIsmi.Body.AddParagraph();
                staffNameValue.AppendText("A.B.Jo'rayev");

                TextBox staffniSababi = forStaff.AppendTextBox(400, 20);
                staffniSababi.Width = 300;
                staffniSababi.HorizontalPosition = 1000;
                staffniSababi.AllowOverlap = true;
                staffniSababi.Format.NoLine = true;
                staffniSababi.Format.HorizontalAlignment = ShapeHorizontalAlignment.Inside;
                Paragraph staffFiredValue = staffniSababi.Body.AddParagraph();
                staffFiredValue.AppendText("this is test from baik web api lorem cnbudiw cnuiwhfuewri ncuiedehfuewi nciuwdhncvuib ncjehewui");

                doc.SaveToFile("never give up.docx", FileFormat.Docx);
                return Ok("never give up");
            }
            catch(Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }

    //    [HttpGet]
    //    public IActionResult Usta()
    //    {
    //        try
    //        {
    //            Workbook workbook = new Workbook();
    //            Worksheet sheet = workbook.Worksheets[0];
    //            sheet.Range["A1"].Text = "No";
    //            sheet.Range["A1"].ColumnWidth = 10;
    //            sheet.Range["A1"].Style.Font.IsBold = true;
    //            sheet.Range["B1"].Text = "Кому";
    //            sheet.Range["B1"].ColumnWidth = 35;
    //            sheet.Range["C1"].Text = "Дата";
    //            sheet.Range["C1"].ColumnWidth = 35;
    //            for (int i = 0; i < 10; i++)
    //            {
    //                sheet.Range[$"A{i}"].Text = $"{i - 1}";
    //                sheet.Range[$"A{i}"].ColumnWidth = 10;
    //                sheet.Range[$"B{i}"].Text = string.Format("{0} polat", i + 1);
    //                sheet.Range[$"B{i}"].ColumnWidth = 35;
    //                sheet.Range[$"C{i}"].Text = $"{DateTime.UtcNow:dd-MM-yyyy HH:mm}";
    //                sheet.Range[$"C{i}"].ColumnWidth = 35;
    //                i++;
    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            return BadRequest(ex.Message);
    //        }
    //    }
    }
}
