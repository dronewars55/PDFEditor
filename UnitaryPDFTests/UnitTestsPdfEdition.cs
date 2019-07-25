using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.IO;
using iTextSharp.text.pdf;
using iTextSharp.text;
using PDFEditor;

namespace UnitaryPDFTests {
    [TestClass]
    public class UnitTestsPdfEdition {
        [TestMethod]
        public void EditFieldValue_ShouldNotReturnException()
        {
            var tempFile = CreateTempDocumentForTests();
            var tempFileAux = Path.GetTempFileName();

            /*Following three lines are an example on how to edit fields
             *with this library. They will take the document in tempFile
             *(which currently has only one text field)
             *and take its first field (fields numbered from 0). It will
             *then change its value to "hey".
             *It will then create a new text field, in the first page
             *inside the rectangle with corners (10, 10) and (50, 50)
             *and call it "hello". It will then change its value to
             *"OK" accessing it through its name and save the result
             *to tempFileAux.
             */
            var edit = new PDFEdition(tempFile);
            edit.Fields[ 0 ] = "hey";
            edit.InsertTextField(1, 10, 10, 50, 50, "hello");
            edit.Fields[ "hello" ] = "OK";
            edit.SaveAs(tempFileAux);

            File.Delete(tempFile);
            File.Delete(tempFileAux);
        }

        [TestMethod]
        public void InsertText_ShouldNotReturnException()
        {
            var tempFile = CreateTempDocumentForTests();
            var tempFileAux = Path.GetTempFileName();

            /*Following three lines are an example of code to add text to a document.
             *It takes the file in tempFile, adds the text "HELLO" in the font "Helvetica"
             *to the coordinates (150, 300) of the first page with a size of 15.
             *Then, it saves the code to tempFileAux.
             */
            var edit = new PDFEdition(tempFile);
            edit.InsertText(1, PDFEdition.Alignment.Center, "HELLO", 150, 300, "Helvetica", 15);
            edit.SaveAs(tempFileAux);
            File.Delete(tempFile);
            File.Delete(tempFileAux);
        }


        //This method is to be ignored
        private string CreateTempDocumentForTests()
        {
            var tempFile = Path.GetTempFileName();
            Document doc = new Document();
            var writer = PdfWriter.GetInstance(doc, new FileStream(tempFile, FileMode.Create));
            doc.Open();
            doc.Add(new Paragraph("Text"));
            doc.Close();
            writer.Close();
            var reader = new PdfReader(tempFile);
            var tempFileAux = Path.GetTempFileName();
            var stamper = new PdfStamper(reader, new FileStream(tempFileAux, FileMode.Create));
            stamper.AcroFields.SetField("hey", "");
            TextField field = new TextField(stamper.Writer, new Rectangle(10, 10, 50, 50), "hey");
            stamper.AddAnnotation(field.GetTextField(), 1);
            stamper.Close();
            reader.Close();
            File.Delete(tempFile);
            File.Move(tempFileAux, tempFile);
            return tempFile;
        }
    }
}
