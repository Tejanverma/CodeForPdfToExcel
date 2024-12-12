using System.Drawing.Imaging;
using Ghostscript.NET.Rasterizer;
using Tesseract;
using PdfToImageUtility;
//using static System.Runtime.InteropServices.JavaScript.JSType;
using OfficeOpenXml;
using System.Runtime.Serialization.DataContracts;
using static System.Runtime.InteropServices.JavaScript.JSType;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using System.ComponentModel;
class Program
{
    private static string tessDataPath = @"C:\Program Files\Tesseract-OCR\tessdata";
    private static string ghostscriptPath = @"C:\Program Files\gs\gs9.26\bin\gswin64c.exe";

    static void Main(string[] args)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        Console.WriteLine("Hello Tej");
        Console.WriteLine("Enter the Pdf path to Get Excel File");
        string pdfPath = Console.ReadLine();

        string imagePath = ConvertPdfTOPng(pdfPath);
        if (imagePath == null)
        {
            Console.WriteLine("Image path is empty");
        }

        string extractedText = GetTextFromImage(tessDataPath, imagePath);
        //extractedText = extractedText.Replace("AP", "ΔP").Replace("?", "³").Replace("°*", "³").Replace("*", "³").Replace("é/h", "³/h");
        extractedText = CorrectReadingError(extractedText);
        //Console.WriteLine(extractedText);
        if (string.IsNullOrEmpty(extractedText)) { Console.WriteLine("Extracted text is empty"); }
        Dictionary<string, string> OrganizedData = OrganizeExtractedText(extractedText);
        if (OrganizedData != null)
        {
            WriteTOExcel(OrganizedData);
        }
    }

    static string ConvertPdfTOPng(string pdfPath)
    {
        Console.WriteLine("Converting Pdf To Image....");
        try
        {
            /*This will replace the \ by \\ it is needed because when we are taking 
             * input of pdfpath then it have single backslash but in C# single blackslash is
             * treated as escape sequence to avoid treating it as escape sequence we add double \\
             */

            pdfPath = pdfPath.Replace(@"\", @"\\");
            /*This is used to dynamically get the path of .net8 i.e from where the program is running
             */
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
            /*This has been done inorder to get the project folder path dynamically
             */
            string projectDirectory = Directory.GetParent(baseDirectory).Parent.Parent.Parent.FullName;
            /*This dynamically combine the project folder and Images folder i.e inside the project
             * folder i.e FinalPdfToExcel to Images
             * in this way FinalPdfToExcel/Images
             */
            string imageDirectory = Path.Combine(projectDirectory, "Images");
            /*This check whether the directory exist or not if it exist this will return true
             * other wise return false. In our case if it gives false then we will make it true and assing
             * inside the block we will create a directory
             */
            if (!Directory.Exists(imageDirectory))
            {
                //Creation of Directory
                Directory.CreateDirectory(imageDirectory);
            }
            /* Important part
             * 
             * When we run the application for the first time an image will be created 
             * and when we run the application for the second time the previous image will be over written
             * resulting in a single image in image folder that'why this will provide the unique name to the image
             * and check for the previously  present image and create a new image instead of over writing
             */
            int counter = 1;
            string imagePath;
            do
            {
                /* string */
                imagePath = Path.Combine(imageDirectory, $@"image{counter}.png");
                counter++;

            } while (File.Exists(imagePath));
            /*The PdfToImage class is present in the PdfToImageUtility class library where
             * ConvertSinglePagePdfToImage method is define to create image of PDF
             */
            PdfToImage.ConvertSinglePagePdfToImage(pdfPath, imagePath, ghostscriptPath);
            Console.WriteLine($"Pdf converted to Image Successfully and Saved At {imagePath}");
            return imagePath;


        }
        catch (Exception e)
        {
            Console.WriteLine($"Error in ConvertPDfTOPng Method {e.Message}");
        }
        return string.Empty;

    }

    static string GetTextFromImage(string TessDataPath, string imagePath)
    {
        Console.WriteLine("Extracting Text From Image......");
        try
        {
            /*TesseractEngine is the class which provide all the functionality to the Tesseract-OCR
             * so in order to use it we have created instance of it
             * which takes tessData path ,language and Engine MOde
             */
            using (var engine = new TesseractEngine(TessDataPath, "eng", EngineMode.Default))
            {
                /*The image which we are having is not supported by the Tesseract engine so we are
                 * Converting it to the format which is supported by the tessercat
                 * The LoadFromFile is present in the Pix class which takes image and convert it to 
                 * the Pix Object format.
                 */
                using (var pix = Pix.LoadFromFile(imagePath))
                {
                    /*The Process method is responsible for OCR which process the image and read the
                     * image content which is present in Tesseract engine class.
                     */
                    var page = engine.Process(pix);
                    /*The GetText method return the ocr result as a string
                     */
                    var extractedText = page.GetText();
                    //The extracted data was having leading space thats why used Trim method.

                    extractedText = extractedText.Trim();
                    Console.WriteLine("Extraction of Text from Image Completed.");
                    //Console.WriteLine(extractedText);
                    return extractedText;
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error in GetTextFromImage{ex.Message}");

        }
        return string.Empty;

    }
    static Dictionary<string, string> OrganizeExtractedText(string data)
    {
        string textFile = @"C:\Users\TejBahadurVerma\source\repos\FinalPdfToExcel\WorkbookPath\Text.txt";
        File.WriteAllText(textFile, data);
        string[] headers =
        { "Tag Number", "P&Id Number", "Line/Equipment No." , "GA. Drawing No.",
          "Service","Quantity Type", "Fluid", "Vapour MW", "Density",
          "Maximum Pressure", "Vapour Pressure",
          "Inlet Pressure", "Inlet Temperature",
          "Viscosity", "Outlet Temperature", "Maximum Flow",
          "ΔP @ Maximum Flow","Normal Flow", "ΔP @ Normal Flow" ,
          "Minimum Flow", "ΔP @ Minimum Flow",
          "Calc. Cv Min Nor Max","Est. Noise dBA | Nor Max","Design Temperature | Pressure",
          "Calc. Cv With Reducers", "Peak Frequency Hz", "Valve Cv",
          "Body Size", "Connections", "Body Type",
          "Body Material", "Plug Size","Plug Type", "Trim Material",
          "Position", "Flange", "ΔP Actuator",
          "Actuator Type", "Air Pressure to Actuator",
          "Travel", "Set Pressure", "Valve FL (Cf)","Spring Range",
          "Positioner / Converter Type","Features","Sov",
          "Seat Leakage", "Manufacturer",
          "Model","Remark","M.R./P.O. No.","M.R.P.O Item"
        };
        var result = new Dictionary<string, string>();

        foreach (var header in headers)
        {
            /* To handle the case where
             * Line\Equipment No. i.e.
             * make Line No. different 
             * and Equipment No. different
             */
            int headerIndex = data.IndexOf(header, StringComparison.OrdinalIgnoreCase);

            if (header.Equals("Line/Equipment No."))
            {
                int valueStart = headerIndex + header.Length;
                int valueEnd = data.IndexOf('/', valueStart);
                if (valueEnd == -1) valueEnd = data.Length;
                // Extract value
                string value = data.Substring(valueStart, valueEnd - valueStart).Trim();
                result["Line No."] = value;
                /*Here valueEnd in case of Equipment will be Value start
                 */
                int valueEndOfEquip = data.IndexOf('\n', valueEnd);

                result["Equipment No."] = data.Substring(valueEnd + 1, valueEndOfEquip - valueEnd + 1).Trim();
                continue;
            }
            if (header.Equals("Design Temperature | Pressure"))
            {
                int valueStart = headerIndex + header.Length;
                int valueEnd = data.IndexOf("C", valueStart);
                valueEnd++;
                if (valueEnd == -1) valueEnd = data.Length;
                // Extract value
                string value = data.Substring(valueStart, valueEnd - valueStart).Trim();
                result["Design Temperature"] = value;
                /*Here valueEnd in case of Equipment will be Value start
                 */
                int valueEndOfPressure = data.IndexOf('\n', valueEnd);

                result["Design Pressure"] = data.Substring(valueEnd, valueEndOfPressure - valueEnd).Trim();
                continue;
            }


            if (header.Equals("Calc. Cv Min Nor Max"))
            {
                int startValue = headerIndex + header.Length;
                //Console.WriteLine(startValue);
                int searchingPartIndex = data.IndexOf('\n', startValue);
                //Console.WriteLine(searchingPartIndex);
                string searchingPart = data.Substring(startValue, searchingPartIndex - startValue);
                //Console.WriteLine(searchingPart);
                int minValueEnd = searchingPart.IndexOf(' ', 0);
                if (minValueEnd == -1) minValueEnd = searchingPart.Length;
                string minValue = searchingPart.Substring(0, minValueEnd - 0);
                int norValueEndInd = searchingPart.IndexOf(" ", minValueEnd);
                if (norValueEndInd == -1) norValueEndInd = searchingPart.Length;
                string norValue = searchingPart.Substring(minValueEnd, norValueEndInd - minValueEnd);
                int maxValueEnd = searchingPart.IndexOf('\n', norValueEndInd);
                if (maxValueEnd == -1) maxValueEnd = searchingPart.Length;
                string maxValue = searchingPart.Substring(norValueEndInd, maxValueEnd - norValueEndInd);
                result["Calc. Cv Min"] = minValue;
                result["Calc. Cv Nor"] = norValue;
                result["Calc. Cv Max"] = maxValue;
                continue;
            }
            if (header.Equals("Quantity Type"))
            {
                int startValue = headerIndex + "Quantity Type".Length;
                int quantityEndIndx = data.IndexOf('|', startValue);
                if (quantityEndIndx == -1) quantityEndIndx = data.Length;
                string quanValue = data.Substring(startValue, quantityEndIndx - startValue - 1);
                int typeEndIndx = data.IndexOf('\n', quantityEndIndx);
                if (typeEndIndx == -1) typeEndIndx = data.Length;
                string typeValue = data.Substring(quantityEndIndx + 1, typeEndIndx - quantityEndIndx);
                result["Quntity"] = quanValue;
                result["Type"] = typeValue;
                continue;

            }
            if (header.Equals("Est. Noise dBA | Nor Max"))
            {
                int startIndexNor = headerIndex + header.Length;
                int endIndexNor = data.IndexOf(' ', startIndexNor);
                if (endIndexNor == -1) endIndexNor = data.Length;
                string norValue = data.Substring(startIndexNor, endIndexNor - startIndexNor);
                int endIndexMax = data.IndexOf('\n', endIndexNor);
                if (endIndexMax == -1) endIndexMax = data.Length;
                string maxValue = data.Substring(endIndexNor, endIndexMax - endIndexNor);
                result["Est. Noise dBA Nor"] = norValue;
                result["Est. Noise dBA Max"] = maxValue;
                continue;
            }
            if (headerIndex != -1)
            {
                // Start after the header
                int valueStart = headerIndex + header.Length;

                // Find the end of the line or next separator
                int valueEnd = data.IndexOf('\n', valueStart);
                if (valueEnd == -1) valueEnd = data.Length;

                // Extract value
                string value = data.Substring(valueStart, valueEnd - valueStart).Trim();
                result[header] = value;
            }
            else
            {
                result[header] = "Not Found"; // If header not found
            }
        }
        return result;
    }
    static void WriteTOExcel(Dictionary<string, string> organizedText)
    {
        Console.WriteLine("Writing Data to Excel");
        if (organizedText == null) { Console.WriteLine("Organize Text is null"); }
        var baseDirectory = AppDomain.CurrentDomain.BaseDirectory;
        var ProjectFolderPath = Directory.GetParent(baseDirectory).Parent.Parent.Parent.FullName;
        var WorkBookPath = Path.Combine(ProjectFolderPath, @"WorkbookPath");
        if (!Directory.Exists(WorkBookPath))
        {
            Directory.CreateDirectory(WorkBookPath);
        }
        var XmlFilePath = Path.Combine(WorkBookPath, @"XmlData.xlsx");

        if (File.Exists(XmlFilePath))
        {
            using (var WB = new ExcelPackage(XmlFilePath))
            {
                var WS = WB.Workbook.Worksheets["DataSheet"];
                var rowStart = FindTheFirstEmptyRow(XmlFilePath);
                if (rowStart != 0)
                {
                    int row = rowStart;

                    // Add values (attribute values) below the headers in the second row
                    int column = 1;
                    foreach (var kvp in organizedText)
                    {
                        WS.Cells[row, column].Value = kvp.Value; // Value
                        column++;
                    }

                    // Save the Excel file to the specified path
                    WB.Save();
                    Console.WriteLine("Data Written to Excel");

                }

            }
        }
        else
        {
            using (var WB = new ExcelPackage(XmlFilePath))
            {
                var WS = WB.Workbook.Worksheets.Add("DataSheet");
                var column = 1;
                foreach (var kvp in organizedText)
                {
                    WS.Cells[1, column].Value = kvp.Key;
                    WS.Cells[1, column].Style.Font.Bold = true;
                    column++;

                }
                var column2 = 1;
                foreach (var kvp in organizedText)
                {
                    WS.Cells[2, column2].Value = kvp.Value;
                    column2++;

                }
                WB.Save();
                Console.WriteLine("Data Written to Excel");
            }

        }


    }
    static int FindTheFirstEmptyRow(string XmlFilePath)
    {
        try
        {
            using (var WB = new ExcelPackage(XmlFilePath))
            {
                var WK = WB.Workbook.Worksheets["DataSheet"];
                int i = 1;
                /*Finding the first row where the value is not present
                 * if the cell value is null then this means that cell is empty
                 */
                while (WK.Cells[i, 1].Value != null)
                {
                    i++;
                }
                //Console.WriteLine(i);
                return i;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);

        }
        return 0;

    }
    public static string CorrectReadingError(string readData)
    {
        string pattern = @"AP|[\?°\*]|°\*|é/h";

        // Replace matches with the appropriate substitutions
        string result = Regex.Replace(readData, pattern, match =>
        {
            // Match value and replace accordingly
            switch (match.Value)
            {
                case "AP":
                    return "ΔP";
                case "?":
                    return "³";
                case "°*":
                    return "³";
                case "*":
                    return "³";
                case "é/h":
                    return "³/h";
                default:
                    return match.Value;  // In case there's an unrecognized match
            }

        });
        return result;
    }
}