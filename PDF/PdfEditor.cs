using System;
using System.Collections.Generic;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;

namespace PDFEditor {

    /// <summary>
    /// The main class for the PDF edition.
    /// </summary>
    /// <remarks>
    /// This class has methods that enable it to open PDF files, write text on them,
    /// fill in text forms they contain and save the result on an output file.
    /// </remarks>
    /// <example> For example:
    /// <code>
    /// var edit = new PDFEdition(sourcePath);
    /// //Some editing of the file
    /// edit.SaveAs(outputPath);
    /// </code>
    /// </example>
    public class PDFEdition {

        /// <summary>
        /// Contains default alignments for when writing text on the pdf
        /// </summary>
        public enum Alignment {
            /// <summary>
            /// Aligns text to the left relative to the point specified.
            /// </summary>
            Left,
            /// <summary>
            /// Aligns the text centered on the point specified.
            /// </summary>
            Center,
            /// <summary>
            /// Aligns text to the right of the specified point.
            /// </summary>
            Right
        };


        /// <summary>
        /// Contains the fields that can be filled in and can be used to edit them.
        /// </summary>
        /// <example> For example:
        /// <code>
        /// edit.Fields[2] = "Text that will end up inside the field with index 2";
        /// edit.Fields["something"] = "Text that will end up inside a field called something";
        /// </code>
        /// </example>
        /// <remarks>
        /// Field edition has only been developed and tested for a document with exclusively
        /// text fields. If the document has any other kind of field in its acrofields, unexpected
        /// results may occur.
        /// </remarks>

        public FieldsEdition Fields {
            get
            {
                PdfReader reader = null;
                try {
                    reader = new PdfReader(originalFile);
                    var ret = new List<string>();
                    foreach (var field in reader.AcroFields.Fields.Keys)
                        ret.Add(field);
                    return new FieldsEdition(ret, fieldsToAddList, this);
                }
                catch {
                    throw new Exception("Error getting the fields of the file");
                }
                finally {
                    if (reader != null)
                        reader.Close();
                }
            }
        }


        private List<TextFieldAddition> fieldsToAddList;
        private List<BaseFont> FontList;
        private Dictionary<string, string> EditedFields;
        private List<TextAddition> textToAddList;
        private Dictionary<string, int> fonts;
        private string originalFile;
        private string temporalPath;
        private string temporalPathAux;

        /// <summary>
        /// The constructor of the class takes one string and at some point must be followed
        /// by a call to the SaveAs method after the editing for exporting the result.
        /// </summary>
        /// <param name="originalFilePath">Path of the original file you want to edit</param>
        public PDFEdition(string originalFilePath)
        {
            originalFile = originalFilePath;
            PdfReader reader = null;
            try {
                reader = new PdfReader(originalFile);
            }
            catch {
                try {
                    reader = new PdfReader(string.Concat(originalFile, ".pdf"));
                    originalFile = string.Concat(originalFile, ".pdf");
                }
                catch {
                    throw new Exception("Invalid path for the original file, or it's been opened by another process");
                }
            }
            finally {
                if (reader != null)
                    reader.Close();
            }

            temporalPath = null;
            temporalPathAux = null;
            FontList = GetFonts();
            EditedFields = new Dictionary<string, string>();
            textToAddList = new List<TextAddition>();
            fieldsToAddList = new List<TextFieldAddition>();

            fonts = new Dictionary<string, int>()
            {
                {
                    "Courier", 0
                },
                {
                    "CourierNegrita", 1
                },
                {
                    "Helvetica", 2
                },
                {
                    "HelveticaNegrita", 3
                },
                {
                    "TimesNewRoman", 4
                },
                {
                    "TimesNewRomanNegrita", 5
                }
            };

            reader = null;
            var fieldList = new List<string>();
            try {
                reader = new PdfReader(originalFile);
                foreach (var campo in reader.AcroFields.Fields)
                    fieldList.Add(campo.Key);
            }
            catch {
                throw new Exception("Original file fields couldn't be readen, or the file was opened by another process");
            }
            finally {
                if (reader != null)
                    reader.Close();
            }
        }

        /// <summary>
        /// Method for importing an external font so that it can be used while writing.
        /// </summary>
        /// <remarks>
        /// The method is able to read a font file and import it so it can be used to write.
        /// The name given to the font is used when referencing it in the writing methods.
        /// Some fonts such as Wingdings can lead to errors.
        /// </remarks>
        /// <param name="name">Name wanted for the font, with which it can be referenced later</param>
        /// <param name="path">Path of the .ttf file used to import the font</param>
        public void AddNewFont(string name, string path)
        {
            if (!FontFactory.IsRegistered(name)) {
                if (!File.Exists(path))
                    throw new Exception("Fount file not found in the specified path");
                FontFactory.Register(path);
                fonts.Add(name, fonts.Count);
            }
            var baseFont = FontFactory.GetFont(name, BaseFont.IDENTITY_H, BaseFont.EMBEDDED).GetCalculatedBaseFont(false);
            FontList.Add(baseFont);
        }

        private List<BaseFont> GetFonts()
        {
            var ret = new List<BaseFont>
            {
                BaseFont.CreateFont(BaseFont.COURIER, BaseFont.CP1252, false),
                BaseFont.CreateFont(BaseFont.COURIER_BOLD, BaseFont.CP1252, false),
                BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, false),
                BaseFont.CreateFont(BaseFont.HELVETICA_BOLD, BaseFont.CP1252, false),
                BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false),
                BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, false)
            };
            return ret;
        }

        private void AddMultipleLinesOrder(int pageNumber, Alignment alignment, string text, float x, float y, BaseFont font, int size)
        {
            var textAux = text.Split(Environment.NewLine.ToCharArray());
            float height = font.GetAscentPoint(text, size);
            for (int i = 0; i < text.Length; i++)
                InsertText(pageNumber, alignment, textAux[ i ], x, y - ( height * 0.75f ) * i, font, size);
        }

        private void AddMultipleLinesOrder(int pageNumber, Alignment alignment, string text, float x, float y, string font, int size)
        {
            try {
                var fontAux = FontList.ElementAt(fonts[ font ]);
                AddMultipleLinesOrder(pageNumber, alignment, text, x, y, fontAux, size);
            }
            catch {
                throw new Exception("Font wasn't found in the fonts dictionary");
            }
        }

        private void AddMultipleLinesOrder(int pageNumber, Alignment alignment, string text, float x, float y, BaseFont font, int size, int width)
        {
            var textAux = text.Split(Environment.NewLine.ToCharArray());
            float height = font.GetAscentPoint(text, size);
            for (int i = 0; i < text.Length; i++)
                InsertText(pageNumber, alignment, textAux[ i ], x, y - ( height * 0.75f ) * i, font, size, width);
        }

        private void AddMultipleLinesOrder(int pageNumber, Alignment alignment, string text, float x, float y, string font, int size, int width)
        {
            try {
                var fontAux = FontList.ElementAt(fonts[ font ]);
                AddMultipleLinesOrder(pageNumber, alignment, text, x, y, fontAux, size, width);
            }
            catch {
                throw new Exception("Font wasn't found in the fonts dictionary");
            }
        }

        private void FillField(string fieldName, string fieldValue)
        {
            if (EditedFields.ContainsKey(fieldName))
                EditedFields.Remove(fieldName);
            EditedFields.Add(fieldName, fieldValue);
        }

        /// <summary>
        /// This methods writes some text on the final file
        /// </summary>
        /// <param name="pageNumber">Number of the page in which it will be written (based at 1)</param>
        /// <param name="alignment">Alignment for the text, specified through an enum the class contains</param>
        /// <param name="text">String of text to be written. Can contain <c cref="Environment.NewLine">Environment.NewLine</c> and will make a new line</param>
        /// <param name="x">X coordinate for the reference point for the text in the page</param>
        /// <param name="y">Y coordinate for the reference point for the text in the page</param>
        /// <param name="font">String with the name of the font to be used. Will be searched for in a built in dictionary</param>
        /// <param name="size">Size of the font</param>
        /// <param name="width">A width reference for the paragraph. Will not cut through words</param>
        public void InsertText(int pageNumber, Alignment alignment, string text, float x, float y, string font, int size, int width)
        {
            bool singleLine = true;
            if (text.Contains(Environment.NewLine)) {
                AddMultipleLinesOrder(pageNumber, alignment, text, x, y, font, size, width);
                singleLine = false;
            }

            if (singleLine) {
                try {
                    var fontAux = FontList.ElementAt(fonts[ font ]);
                    InsertText(pageNumber, alignment, text, x, y, fontAux, size, width);
                }
                catch {
                    throw new Exception("Font wasn't found in the fonts dictionary");
                }
            }
        }

        private void InsertText(int pageNumber, Alignment alignment, string text, float x, float y, BaseFont font, int size, int width)
        {
            bool singleLine = true;
            if (text.Contains(Environment.NewLine)) {
                AddMultipleLinesOrder(pageNumber, alignment, text, x, y, font, size, width);
                singleLine = false;
            }

            if (singleLine) {

                var chains = new List<string>();
                int i = 1;
                var cad = "";
                var textAux = text.Split(' ');

                for (int j = 0; j < textAux.Length; j++)
                    textAux[ j ] = string.Concat(textAux[ j ], ' ');

                if (text.Length == 0)
                    throw new Exception("The text string was empty");
                chains.Add(string.Concat(cad, textAux.ElementAt(0)));
                while (i < textAux.Length) {
                    cad = ( string ) chains.Last().Clone();
                    if (font.GetWidthPoint(string.Concat(cad, textAux.ElementAt(i)), size) > width) {
                        string cadn = "";
                        cadn = string.Concat(cadn, textAux.ElementAt(i));
                        chains.Add(cadn);
                    }
                    else {
                        chains.RemoveAt(chains.Count - 1);
                        chains.Add(string.Concat(cad, textAux.ElementAt(i)));
                    }
                    i++;
                }

                for (int j = 0; j < chains.Count; j++) {
                    InsertText(pageNumber, alignment, chains.ElementAt(j), x, y - ( size + 3 ) * j, font, size);
                }
            }
        }

        /// <summary>
        /// This methods writes some text on the final file
        /// </summary>
        /// <param name="pageNumber">Number of the page in which it will be written (based at 1)</param>
        /// <param name="alignment">Alignment for the text, specified through an enum the class contains</param>
        /// <param name="text">String of text to be written. Can contain <c cref="Environment.NewLine">Environment.NewLine</c> and will make a new line</param>
        /// <param name="x">X coordinate for the reference point for the text in the page</param>
        /// <param name="y">Y coordinate for the reference point for the text in the page</param>
        /// <param name="font">String with the name of the font to be used. Will be searched for in a built in dictionary</param>
        /// <param name="size">Size of the font</param>
        public void InsertText(int pageNumber, Alignment alignment, string text, float x, float y, string font, int size)
        {
            bool singleLine = true;
            if (text.Contains(Environment.NewLine)) {
                AddMultipleLinesOrder(pageNumber, alignment, text, x, y, font, size);
                singleLine = false;
            }

            if (singleLine) {

                try {
                    var fontAux = FontList.ElementAt(fonts[ font ]);
                    var nuevoTexto = new TextAddition(alignment, pageNumber, text, x, y, fontAux, size);
                    textToAddList.Add(nuevoTexto);
                }
                catch {
                    throw new Exception("Font wasn't found in the fonts dictionary");
                }
            }
        }

        /// <summary>
        /// This method adds a new textbox for a new form field to the PDF.
        /// </summary>
        /// <param name="page">
        /// Number of page in which it will appear (starting from 1)
        /// </param>
        /// <param name="lowerLeftY">
        /// Y coordinate of the lower left corner of the textbox
        /// </param>
        /// <param name="lowerLeftX">
        /// X coordinate of the lower left corner of the textbox
        /// </param>
        /// <param name="upperRightY">
        /// Y coordinate of the upper right corner of the textbox
        /// </param>
        /// <param name="upperRightX">
        /// X coordinate of the upper right corner of the textbox
        /// </param>
        /// <param name="name">
        /// Name for the field. It can be used to edit its value.
        /// </param>
        public void InsertTextField(int page, float lowerLeftY, float lowerLeftX, float upperRightY, float upperRightX, string name)
        {
            fieldsToAddList.Add(new TextFieldAddition(upperRightX, upperRightY, lowerLeftX, lowerLeftY, name, page));
        }

        private void InsertText(int pageNumber, Alignment alignment, string text, float x, float y, BaseFont font, int size)
        {
            bool singleLine = true;
            if (text.Contains(Environment.NewLine)) {
                AddMultipleLinesOrder(pageNumber, alignment, text, x, y, font, size);
                singleLine = false;
            }

            if (singleLine) {
                var nuevoTexto = new TextAddition(alignment, pageNumber, text, x, y, font, size);
                textToAddList.Add(nuevoTexto);
            }
        }

        private void ExecuteText()
        {
            PdfStamper stamper = null;
            PdfReader reader = null;
            temporalPathAux = Path.GetTempFileName();
            try {
                reader = new PdfReader(temporalPath);
                stamper = new PdfStamper(reader, new FileStream(temporalPathAux, FileMode.Create));
                textToAddList.OrderBy(j => j.Page);
                for (int i = 1; i <= reader.NumberOfPages; i++) {
                    var content = stamper.GetUnderContent(i);
                    List<TextAddition> listOrdersPage = textToAddList.FindAll(j => j.Page == i);
                    foreach (var order in listOrdersPage) {
                        content.SetFontAndSize(order.Font, order.Size);
                        content.BeginText();
                        content.ShowTextAligned(order.Alignment, order.Text, order.X, order.Y, 0);
                        content.EndText();
                    }
                }
                stamper.Close();
                reader.Close();
            }
            catch {
                throw new Exception("Error while executing a text order");
            }

            finally {
                if (stamper != null)
                    stamper.Close();
                if (reader != null)
                    reader.Close();
                File.Delete(temporalPath);
                File.Move(temporalPathAux, temporalPath);
                temporalPathAux = null;
            }
        }

        private void ExecuteFields()
        {
            PdfStamper stamper = null;
            PdfReader reader = null;
            temporalPathAux = Path.GetTempFileName();

            try {
                reader = new PdfReader(temporalPath);
                stamper = new PdfStamper(reader, new FileStream(temporalPathAux, FileMode.Create));

                var form = stamper.AcroFields;
                foreach (var order in EditedFields)
                    form.SetField(order.Key, order.Value);

                stamper.FormFlattening = false;
                stamper.Close();
                reader.Close();

                File.Delete(temporalPath);
                File.Move(temporalPathAux, temporalPath);
                temporalPathAux = null;
            }
            catch {
                throw new Exception("Error while editing fields");
            }
            finally {
                if (stamper != null)
                    stamper.Close();
                if (reader != null)
                    reader.Close();
            }
        }

        private void ExecuteFields(bool flag)
        {
            PdfReader reader = null;
            PdfStamper stamper = null;
            temporalPathAux = Path.GetTempFileName();

            try {
                reader = new PdfReader(temporalPath);
                stamper = new PdfStamper(reader, new FileStream(temporalPathAux, FileMode.Create));

                var form = stamper.AcroFields;
                foreach (var order in EditedFields)
                    form.SetField(order.Key, order.Value);

                stamper.FormFlattening = flag;
                stamper.Close();
                reader.Close();

                File.Delete(temporalPath);
                File.Move(temporalPathAux, temporalPath);
                temporalPathAux = null;
            }
            catch {
                throw new Exception("Error while editing fields");
            }
            finally {
                if (stamper != null)
                    stamper.Close();
                if (reader != null)
                    reader.Close();
            }
        }

        /// <summary>
        /// Method expected after the last use of the class that saves the result to a new file
        /// </summary>
        /// <param name="finalFilePath">Path for the new file to create with the results.</param>
        public void SaveAs(string finalFilePath)
        {
            PdfStamper stamper = null;
            PdfReader reader = null;
            temporalPath = Path.GetTempFileName();

            try {
                reader = new PdfReader(originalFile);
                stamper = new PdfStamper(reader, new FileStream(temporalPath, FileMode.Create));
                stamper.Close();
                reader.Close();
            }
            catch {
                throw new Exception("Error creating temporal file");
            }
            finally {
                if (reader != null)
                    reader.Close();
                if (stamper != null)
                    stamper.Close();
            }

            ExecuteText();
            ExecuteTextFieldAdditions();
            ExecuteFields();

            try {
                reader = new PdfReader(temporalPath);
                stamper = new PdfStamper(reader, new FileStream(finalFilePath, FileMode.Create));
                stamper.Close();
                reader.Close();
            }
            catch {
                throw new Exception("Error creating final file");
            }
            finally {
                if (reader != null)
                    reader.Close();
                if (stamper != null)
                    stamper.Close();
            }
        }

        private void ExecuteTextFieldAdditions()
        {
            PdfReader reader = null;
            PdfStamper stamper = null;

            try {
                reader = new PdfReader(temporalPath);
                temporalPathAux = Path.GetTempFileName();
                stamper = new PdfStamper(reader, new FileStream(temporalPathAux, FileMode.Create));

                foreach (var fieldOrder in fieldsToAddList) {
                    TextField fieldToAdd = new TextField(stamper.Writer, new Rectangle(fieldOrder.urx, fieldOrder.ury, fieldOrder.llx, fieldOrder.lly), fieldOrder.name);
                    stamper.AddAnnotation(fieldToAdd.GetTextField(), fieldOrder.page);
                    stamper.AcroFields.SetField(fieldOrder.name, "");
                }
                stamper.Close();
                reader.Close();

                File.Delete(temporalPath);
                File.Move(temporalPathAux, temporalPath);
            }
            catch {
                throw new Exception("Error adding new fields");
            }
            finally {
                if (stamper != null)
                    stamper.Close();
                if (reader != null)
                    reader.Close();
            }
        }

        /// <summary>
        /// Method expected after the last use of the class that saves the result to a new file
        /// </summary>
        /// <param name="finalFilePath">Path for the new file to create with the results.</param>
        /// <param name="flattenFields">Bool flag that when true makes the method flatten the fields after filling them</param>
        public void SaveAs(string finalFilePath, bool flattenFields)
        {

            PdfStamper stamper = null;
            PdfReader reader = null;
            temporalPath = Path.GetTempFileName();

            try {
                reader = new PdfReader(originalFile);
                stamper = new PdfStamper(reader, new FileStream(temporalPath, FileMode.Create));
                stamper.Close();
                reader.Close();
            }
            catch {
                throw new Exception("Error creating temporal file");
            }
            finally {
                if (reader != null)
                    reader.Close();
                if (stamper != null)
                    stamper.Close();
            }

            ExecuteText();
            ExecuteTextFieldAdditions();
            ExecuteFields(flattenFields);

            try {
                reader = new PdfReader(temporalPath);
                stamper = new PdfStamper(reader, new FileStream(finalFilePath, FileMode.Create));
                stamper.Close();
                reader.Close();
            }
            catch {
                throw new Exception("Error creating final file");
            }
            finally {
                if (reader != null)
                    reader.Close();
                if (stamper != null)
                    stamper.Close();
            }
        }

        /// <summary>
        /// List of the font names available to the user for writing. These can be used in the InsertText method
        /// </summary>
        /// <value>
        /// List of string of each name of the fonts available for the user. To add more use the AddNewFont method.
        /// </value>
        public List<string> FontDictionary {
            get
            {
                return fonts.Keys.ToList();
            }
        }

        private float GetHeight(int pageNumber)
        {
            PdfReader reader = null;

            try {
                reader = new PdfReader(originalFile);
                var height = reader.GetPageSize(pageNumber).Height;
                return height;
            }
            catch {
                throw new Exception("Couldn't read the page height, or the original file was opened by another process");
            }
            finally {
                if (reader != null)
                    reader.Close();
            }
        }

        /// <summary>
        /// This class contains the names of the fields that can be filled in and methods for filling them in.
        /// <example> For example:
        /// <code>
        /// var edit = new PDFEdition(inputPath);
        /// edit.Fields[0] = "Text the field with index 0 will contain";
        /// edit.SaveAs(outputPath);
        /// </code>
        /// </example>
        /// </summary>
        public class FieldsEdition {

            private PDFEdition outerClass;
            /// <value>
            /// Returns a list of the names of all the fields in the document
            /// </value>
            public List<string> FieldsNames {
                get
                {
                    var ret = new List<string>();
                    for (int i = 0; i < fieldArray.Length; i++) {
                        ret.Add(fieldArray[ i ]);
                    }
                    return ret;
                }
            }
            private int FieldListSize { get { return fieldArray.Length; } }
            private string[] fieldArray;
            /// <summary>
            /// Used for filling in fields
            /// </summary>
            /// <param name="i">Index of the field chosen to fill in</param>
            /// <returns></returns>
            public string this[ int i ] {
                get
                {
                    if (i < fieldArray.Length)
                        return fieldArray[ i ];
                    else
                        throw new Exception("Value out of bounds of the field list");
                }
                set
                {
                    if (i < fieldArray.Length)
                        outerClass.FillField(fieldArray[ i ], value);
                    else
                        throw new Exception("Value out of bounds of the fields list");
                }
            }

            /// <summary>
            /// Used for filling in fields
            /// </summary>
            /// <param name="nCampo">Name of the field chosen to fill in</param>
            /// <returns></returns>
            public string this[ string nCampo ] {
                set
                {
                    bool done = false;
                    for (int i = 0; i < fieldArray.Length; i++) {
                        if (fieldArray[ i ] == nCampo) {
                            outerClass.FillField(fieldArray[ i ], value);
                            done = true;
                        }
                    }
                    if (!done)
                        throw new Exception("Field could not be found");
                }
            }

            internal FieldsEdition(List<string> listActualFields, List<TextFieldAddition> listToAddFields, PDFEdition editor)
            {
                int i = 0;
                foreach (var field in listToAddFields)
                    listActualFields.Add(field.name);
                foreach (var field in listActualFields) {
                    i++;
                }
                fieldArray = new string[ i ];
                for (int j = 0; j < i; j++) {
                    try {
                        fieldArray[ j ] = listActualFields.ElementAt(j);
                    }
                    catch {
                        fieldArray[ j ] = listToAddFields.ElementAt(j - listActualFields.Count).name;
                    }
                }
                outerClass = editor;
            }
        }

        private class TextAddition {
            public int Alignment { get; set; }
            public int Page { get; set; }
            public string Text { get; set; }
            public float X { get; set; }
            public float Y { get; set; }
            public BaseFont Font { get; set; }
            public int Size { get; set; }

            public TextAddition(PDFEdition.Alignment alignment, int pageNumber, string text, float xCoordinate, float yCoordinate, BaseFont font, int size)
            {

                Alignment = ( int ) alignment;
                Page = pageNumber;
                Text = text;
                X = xCoordinate;
                Y = yCoordinate;
                Font = font;
                Size = size;
            }
        }

        internal class TextFieldAddition {
            internal float urx;
            internal float llx;
            internal float lly;
            internal float ury;
            internal string name;
            internal int page;

            internal TextFieldAddition(float urx, float ury, float llx, float lly, string name, int page)
            {
                this.page = page;
                this.urx = urx;
                this.ury = ury;
                this.llx = llx;
                this.lly = lly;
                this.name = name;
            }
        }
    }
}

