// FinalProjectAp

using csEquationSolver;
using Mathematics;
using OfficeOpenXml;
using SparseCollections;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using LinqToExcel;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using Syncfusion.XlsIO;
using static System.Net.Mime.MediaTypeNames;
using System.CodeDom;
using System.Runtime.InteropServices;

namespace UI
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<string> equations = new List<string>();
        public List<string> result = new List<string>();
        static string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        private string excelPath = desktopPath + @"\Information.xlsx";



        public MainWindow()
        {
         
            InitializeComponent();
        }

        private void RichTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Button_Click(object sender, RoutedEventArgs e)   // save method : ba in method dar excel zakhire mikonim
        {
            Information info = new Information();

            String[] lines =
        StringFromRichTextBox(rtb).Split(new[] { Environment.NewLine }
                                          , StringSplitOptions.RemoveEmptyEntries);

            if (lines.Length < 5)
            {
                MessageBox.Show("تعداد خطوط کافی نمیباشد!");
                return;
            }
            info.Name = lines[0];
            info.Family = lines[1];
            info.City = lines[2];
            try
            {
                info.Age = Convert.ToInt32(lines[3]);

            }
            catch(Exception ex)
            {
                MessageBox.Show("لطفا سن را به صورت عددی  وارد نمایید");
                return;
            }
            
            equations.Add(lines[4].RemoveWhitespace());
            equations.Add(lines[5].RemoveWhitespace());
            if (lines.Length == 7)
            {
                equations.Add(lines[6].RemoveWhitespace());
            }
            equations.Add("");

            bool isSolved=false;
            if (!System.IO.File.Exists(excelPath))
            {
                CreateExcel2();
            }
            if (equations.Count ==3)
            {
                 isSolved = Solve();
            }
            else
            {
                MessageBox.Show("تعداد معادلات باید دو  یا سه عدد باشد.");
            }
            if (isSolved)
            {
                String[] solveLines =
                    StringFromRichTextBox(solveRichTextBox).Split('\n');
                result.Add(solveLines[0].RemoveWhitespace());
                result.Add(solveLines[1].RemoveWhitespace());
                if (solveLines.Length == 3)
                {
                    result.Add(solveLines[2].RemoveWhitespace());
                }
                result.Add("");
                info.Equations = equations;
                info.Results = result;
                SaveToExcel(info);
            }
            
            equations.Clear();
            result.Clear();


        }
        string StringFromRichTextBox(RichTextBox rtb)
        {
            TextRange textRange = new TextRange(
              rtb.Document.ContentStart,
              rtb.Document.ContentEnd
            );

            return textRange.Text;
        }

        private void Button_Click_1(object sender, RoutedEventArgs e)
        {
            this.Close();
            Environment.Exit(1);
        } 

        private bool Solve()
        {
            
            Sparse2DMatrix<int, int, double> aMatrix = new Sparse2DMatrix<int, int, double>();
            SparseArray<int, double> bVector = new SparseArray<int, double>();
            SparseArray<string, int> variableNameIndexMap = new SparseArray<string, int>();
            int numberOfEquations = 0;
            bool returnValue = false;
            LinearEquationParser parser = new LinearEquationParser();
            LinearEquationParserStatus parserStatus = LinearEquationParserStatus.Success;

            foreach (string inputLine in equations)
            {
                parserStatus = parser.Parse(inputLine,
                                            aMatrix,
                                            bVector,
                                            variableNameIndexMap,
                                            ref numberOfEquations);

                if (parserStatus != LinearEquationParserStatus.Success)
                {
                    break;
                }
            }

            string mainStatusBarText = UI.Properties.Resources.IDS_EQUATIONS_SOLVED;

            if (parserStatus == LinearEquationParserStatus.Success)
            {
                if (numberOfEquations == variableNameIndexMap.Count)
                {
                    SparseArray<int, double> xVector = new SparseArray<int, double>();

                    LinearEquationSolverStatus solverStatus =
                        LinearEquationSolver.Solve(numberOfEquations,
                                                   aMatrix,
                                                   bVector,
                                                   xVector);

                    if (solverStatus == LinearEquationSolverStatus.Success)
                    {
                        string solutionString = "";

                        foreach (KeyValuePair<string, int> pair in variableNameIndexMap)
                        {
                            solutionString += string.Format("{0} = {1}", pair.Key, xVector[pair.Value]);
                            solutionString += "\n";
                        }
                        solveRichTextBox.Document.Blocks.Clear();
                        solveRichTextBox.Document.Blocks.Add(new Paragraph(new Run(solutionString)));
                        returnValue = true;
                    }
                    else if (solverStatus == LinearEquationSolverStatus.IllConditioned)
                    {
                        mainStatusBarText = UI.Properties.Resources.IDS_ILL_CONDITIONED_SYSTEM_OF_EQUATIONS;
                        MessageBox.Show(mainStatusBarText);

                    }
                    else if (solverStatus == LinearEquationSolverStatus.Singular)
                    {
                        mainStatusBarText = UI.Properties.Resources.IDS_SINGULAR_SYSTEM_OF_EQUATIONS;
                        MessageBox.Show(mainStatusBarText);

                    }
                }
                else if (numberOfEquations < variableNameIndexMap.Count)
                {
                    mainStatusBarText = string.Format(UI.Properties.Resources.IDS_TOO_FEW_EQUATIONS,
                                                      numberOfEquations, variableNameIndexMap.Count);
                    MessageBox.Show(mainStatusBarText);

                }
                else if (numberOfEquations > variableNameIndexMap.Count)
                {
                    mainStatusBarText = string.Format(UI.Properties.Resources.IDS_TOO_MANY_EQUATIONS,
                                                      numberOfEquations, variableNameIndexMap.Count);
                    MessageBox.Show(mainStatusBarText);

                }
            }
            else
            {
                mainStatusBarText = LinearEquationParserStatusInterpreter.GetStatusString(parserStatus);
                MessageBox.Show(mainStatusBarText);
            }
            return returnValue;
        } 

        private void CreateExcel2()
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;

         

                //Create a new workbook
                IWorkbook workbook = application.Workbooks.Create(1);
                IWorksheet sheet = workbook.Worksheets[0];
                sheet.Name = "Information";


                #region Define Styles
                IStyle pageHeader = workbook.Styles.Add("PageHeaderStyle");
                IStyle tableHeader = workbook.Styles.Add("TableHeaderStyle");

                pageHeader.Font.RGBColor = Color.FromArgb(0, 83, 141, 213);
                pageHeader.Font.FontName = "Calibri";
                pageHeader.Font.Size = 18;
                pageHeader.Font.Bold = true;
                pageHeader.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                pageHeader.VerticalAlignment = ExcelVAlign.VAlignCenter;

                tableHeader.Font.Color = ExcelKnownColors.White;
                tableHeader.Font.Bold = true;
                tableHeader.Font.Size = 11;
                tableHeader.Font.FontName = "Calibri";
                tableHeader.HorizontalAlignment = ExcelHAlign.HAlignCenter;
                tableHeader.VerticalAlignment = ExcelVAlign.VAlignCenter;
                tableHeader.Color = Color.FromArgb(0, 118, 147, 60);
                tableHeader.Borders[ExcelBordersIndex.EdgeLeft].LineStyle = ExcelLineStyle.Thin;
                tableHeader.Borders[ExcelBordersIndex.EdgeRight].LineStyle = ExcelLineStyle.Thin;
                tableHeader.Borders[ExcelBordersIndex.EdgeTop].LineStyle = ExcelLineStyle.Thin;
                tableHeader.Borders[ExcelBordersIndex.EdgeBottom].LineStyle = ExcelLineStyle.Thin;
                #endregion

                #region Apply Styles
                //Apply style to the header
                sheet["A1"].Text = "Information ";
                sheet["A1"].CellStyle = pageHeader;

                sheet["A2"].Text = "Information Of Persons with Equations";
                sheet["A2"].CellStyle = pageHeader;
                sheet["A2"].CellStyle.Font.Bold = false;
                sheet["A2"].CellStyle.Font.Size = 16;

                sheet["A1:J1"].Merge();
                sheet["A2:J2"].Merge();
                sheet["A3:A4"].Merge();
                sheet["B3:B4"].Merge();
                sheet["C3:C4"].Merge();
                sheet["D3:D4"].Merge();

              
                sheet["E3:g3"].Merge();
                sheet["H3:J3"].Merge();

                sheet["B3"].Text = "Family";
                sheet["A3"].Text = "Name";
                sheet["C3"].Text = "City";
                sheet["D3"].Text = "Age";
                sheet["E3"].Text = "Equations";
                sheet["H3"].Text = "Results";

                sheet["E4"].Text = "First Equation";
                sheet["F4"].Text = "Second Equation";
                sheet["G4"].Text = "Third Equation";
                sheet["H4"].Text = "First Variable";
                sheet["I4"].Text = "Second Variable";
                sheet["J4"].Text = "Third Variable";
               // sheet["G4"].Text = "Third Equation";
                sheet["A3:J4"].CellStyle = tableHeader;
                #endregion

                sheet.UsedRange.AutofitColumns();

                //Save the file in the given path
                Stream excelStream = File.Create(Path.GetFullPath(excelPath));
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
            }
        } 

        private void SaveToExcel(Information info)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = application.Workbooks.Open(excelPath);
                IWorksheet sheet = workbook.Worksheets[0];
                object[] expenseArray = new object[10]
                    {info.Name, info.Family, info.City, info.Age, info.Equations[0], info.Equations[1], info.Equations[2], info.Results[0],info.Results[1], info.Results[2]};
                int x = sheet.UsedRange.LastRow;
                sheet.ImportArray(expenseArray, x+1, 1, false);        
                Stream excelStream = File.Open(excelPath,FileMode.Open);
                workbook.SaveAs(excelStream);
                excelStream.Dispose();
                MessageBox.Show("اطلاعات با موفقیت در فایل ذخیره گردید.");
            }
        }

        private void Button_Click_2(object sender, RoutedEventArgs e)
        {
            rtb.Document.Blocks.Clear();
            solveRichTextBox.Document.Blocks.Clear();
            
        } 

        private void TextBox_TextChanged(object sender, TextChangedEventArgs e)
        {

        }

        private void Button_Click_3(object sender, RoutedEventArgs e)
        {
            List<Information> informations = getInformations();                             
           
            var ds = informations.FindAll(s =>
                ((nameCheckBox.IsChecked != true) || s.Name == nameTextBox.Text)
                && ((familyCheckBox.IsChecked != true) || s.Family == familyTextBox.Text)
                && ((cityCheckBox.IsChecked != true) || s.City == cityTextBox.Text)
                && ((ageBelowCheckBox.IsChecked != true) || s.Age < Convert.ToInt16(belowAgeTextBox.Text))
                && ((ageUpperCheckBox.IsChecked != true) || s.Age > Convert.ToInt16(upperAgeTextBox.Text))
               // && ((euationCheckBox.IsChecked != true) || s.Equations.Contains(equationsTextBox.Text.Trim()))
                && ((euationCheckBox.IsChecked != true) || s.FirstEq.RemoveWhitespace() == equationsTextBox.Text.Trim().RemoveWhitespace())
                && ((euationCheckBox.IsChecked != true) || s.SecondEq.RemoveWhitespace() == equationsTextBox.Text.Trim().RemoveWhitespace())
                && ((rsultCheckBox.IsChecked != true) || s.FirstRes.RemoveWhitespace() == firstResultTextBox.Text.Trim().RemoveWhitespace())
                && ((rsultCheckBox.IsChecked != true) || s.SecondRes.RemoveWhitespace() == secondresultTextBox.Text.Trim().RemoveWhitespace())
                && ((rsultCheckBox.IsChecked != true) || s.ThirdRes.RemoveWhitespace() == thirdresultTextBox.Text.Trim().RemoveWhitespace())
            ).Select( u=> new {u.Name, u.Family, u.City,u.Age,u.FirstEq,u.SecondEq,u.ThirdEq,u.FirstRes,u.SecondRes,u.ThirdRes}).Distinct();
            DGV.ItemsSource = ds;

              
            
            
        } 

        private int getLastRow()
        {
            int lastRow = -1;
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Excel2016;
                IWorkbook workbook = application.Workbooks.Open(excelPath);
                IWorksheet sheet = workbook.Worksheets[0];
                lastRow = sheet.UsedRange.LastRow;
            }
            return lastRow;
        }

        private List<Information> getInformations()
        {
            var excelFile = new ExcelQueryFactory(excelPath);
            var getData = from a in excelFile.WorksheetNoHeader("Information") select a;
            var xray = getData.ToList();
            int lastRow = getLastRow();
            List<Information> informations = new List<Information>();
          
            if (lastRow != -1)
            {
                for (int i = 4; i < lastRow; i++)
                {
                    List<string> eqList = new List<string>();
                    List<string> reList = new List<string>();
                    Information info = new Information();
                    info.Name = xray[i][0].ToString();
                    info.Family = xray[i][1].ToString();
                    info.City = xray[i][2].ToString();
                    info.Age = Convert.ToInt16(xray[i][3]);
                    info.FirstEq = xray[i][4].ToString().Trim();
                    info.SecondEq = xray[i][5].ToString().Trim();
                    info.ThirdEq = xray[i][6].ToString().Trim();
                    eqList.Add(xray[i][4].ToString().Trim());
                    eqList.Add(xray[i][5].ToString().Trim());
                    eqList.Add(xray[i][6].ToString().Trim());
                    info.Equations = eqList;

                    info.FirstRes = xray[i][7].ToString().Trim();
                    info.SecondRes = xray[i][8].ToString().Trim();
                    info.ThirdRes = xray[i][9].ToString().Trim();
                    reList.Add(xray[i][7].ToString().Trim());
                    reList.Add(xray[i][8].ToString().Trim());
                    reList.Add(xray[i][9].ToString().Trim());
                    info.Results = reList;
                    informations.Add(info);
                  
                }
            }
            return informations;
        }
    }
}

