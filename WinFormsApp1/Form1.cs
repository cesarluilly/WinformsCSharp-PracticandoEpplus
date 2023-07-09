using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.IO.Packaging;

namespace WinFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //                                              //GetDateNow
            String strDate = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            String strRootDirectory = Directory.GetCurrentDirectory();

            String strNameExcelGenerate = "myWorkbook.xlsx";
            String strRelativePathGenerate = @"Generate\" + strDate + @"\";
            String strAbsolutePathGenerate = Path.Combine(strRootDirectory, strRelativePathGenerate,
                strNameExcelGenerate);

            //                                              //Create directory if not exist.
            Directory.CreateDirectory(Path.GetDirectoryName(strAbsolutePathGenerate));

            using (var package = new ExcelPackage(new FileInfo(strAbsolutePathGenerate)))
            {
                var sheet = package.Workbook.Worksheets.Add("My Sheet");
                sheet.Cells["A1"].Value = "Hello World!";

                var sheet2 = package.Workbook.Worksheets.Add("My Sheet Cesar");
                sheet2.Cells["A2"].Value = "Hola Soy Cesar";

                //                                          //Save file.
                package.Save();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button2_Click(object sender, EventArgs e)
        {
            String strRootDirectory = Directory.GetCurrentDirectory();
            String strNameExcelTest = "Test.xlsx";
            String strRelativePathMyExcels = @"MyExcel\";
            String strAbsolutePathMyExcels = Path.Combine(strRootDirectory, strRelativePathMyExcels,
                strNameExcelTest);

            using (var package = new ExcelPackage(strAbsolutePathMyExcels))
            {
                String strDate = DateTime.Now.ToString("yyyyMMdd_HHmmss");

                var a = package.Workbook.Worksheets.ToList();

                package.Workbook.Worksheets["MySheet"].Cells["B3"].Formula = "SUM(B1:B2)";
                package.Workbook.Worksheets["MySheet"].Cells["D1"].Value = "Fecha";
                package.Workbook.Worksheets["MySheet"].Cells["D2"].Value = strDate;

                //                                          //Save as a new file.
                var newFile = new FileInfo("AnotherWorkbook.xlsx");
                package.SaveAs(newFile);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            //                                              //GetDateNow
            String strDate = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            String strRootDirectory = Directory.GetCurrentDirectory();

            String strNameExcelGenerate = "myWorkbookWithStyles.xlsx";
            String strRelativePathGenerate = @"Generate\" + strDate + @"\";
            String strAbsolutePathGenerate = Path.Combine(strRootDirectory, strRelativePathGenerate,
                strNameExcelGenerate);

            //                                              //Create directory if not exist.
            Directory.CreateDirectory(Path.GetDirectoryName(strAbsolutePathGenerate));

            using (var package = new ExcelPackage(new FileInfo(strAbsolutePathGenerate)))
            {
                //                                          //Obtengo mi hoja
                var sheet = package.Workbook.Worksheets.Add("My Sheet");

                //                                          //Asigno valor a celda A1
                sheet.Cells["D1"].Value = "Celda [D1]";

                //                                          //Asigno un valor a traves de [1,3] equivale a ["C1"]
                sheet.Cells[2, 4].Value = "Celda [2,4]";

                //                                          //Asigno valor a una matriz de celdas de 2 formas
                //                                          //    y aplico estilo de format.
                sheet.Cells["A1:B3"].Value = 6174;
                sheet.Cells[1, 1, 3, 2].Style.Numberformat.Format = "#,##0"; //es lo mismo que,A1:B3
                sheet.Cells["A1:B3"].Style.Numberformat.Format = "#,##0";

                //                                          //Asigno valor a 2 rangos en una instruccion
                //                                          //    y aplico formato.
                sheet.Cells["A6:B9,D6:E9"].Value = 16383;
                sheet.Cells["A6:B9,D6:E9"].Style.Numberformat.Format = "#,##0";


                sheet.Cells["g:h, I1"].Value = 26;
                sheet.Cells["g:h, I1"].Style.Font.Bold = true; //sets font-bold to true for column a & b

                sheet.Cells["K:K"].Value = 16;
                sheet.Cells["K:K"].Style.Font.Bold = true; //Sets font-bold to true for row 1,column A and cell C3

                ////                                          Si quiero seleccionar todas las celdas para aplicarles un valor
                //sheet.Cells["A:XFD"].Style.Font.Name = "Arial"; //Sets font to Arial for all cells in a worksheet.
                //sheet.Cells.Style.Font.Name = "Arial"; //This is equal to the above.

                sheet.Cells[1, 13].Value = "Aplicando bold";
                sheet.Cells[1, 13].Style.Font.Bold = true;

                sheet.Cells[2, 13].Value = "Aplicando PatternType y Background, ambos de la mano";
                sheet.Cells[2, 13].Style.Fill.PatternType = ExcelFillStyle.Solid;
                sheet.Cells[2, 13].Style.Fill.BackgroundColor.SetColor(Color.YellowGreen);

                sheet.Cells[3, 13].Value = "Aplicando ColorText";
                sheet.Cells[3, 13].Style.Font.Color.SetColor(Color.Red);

                //                                          //Save file.
                package.Save();
            }
        }
    }
}