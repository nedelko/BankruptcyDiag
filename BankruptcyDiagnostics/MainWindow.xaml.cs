using Microsoft.Win32;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace BankruptcyDiagnostics
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        ReportCollection reportCollection = new ReportCollection();
        TextBlock textDiag;
        DataGrid dataGrid1;
        bool ReportIsValid = true;
        public MainWindow()
        {
            InitializeComponent();
            var ib = new ImageBrush
            {
                ImageSource = new BitmapImage(new Uri(@"Images\logo.jpg", UriKind.Relative))
            };
            LogoPanel.Background = ib;
        }
        public bool CheckIfDouble(object value)
        {
            return Double.TryParse((value == null ? "" : value.ToString()), out double n);
        }
        public bool CheckIfInt(object value)
        {
            return Int32.TryParse((value == null ? "" : value.ToString()), out int n);
        }
        public void report_Validation(Excel.Range xlRange1, Excel.Range xlRange2)
        {
            if (CheckIfInt(xlRange1.Cells[2, 7].Value2) != true)
            {
                xlRange1.Cells[2, 7].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            else
            {
                if (reportCollection.Count > 0)
                {
                    for (int i=0; i<reportCollection.Count; i++)
                    {
                        if ((Double)xlRange1.Cells[2, 7].Value2 == reportCollection.GetReport(i).rep_year)
                        {
                            string msg = "Звіт " + (reportCollection.GetReport(i).rep_year).ToString() + " року вже є в базі";
                            MessageBox.Show(msg);
                            ReportIsValid = false;
                            break;
                        }
                    }
                }
            }
            if (CheckIfDouble(xlRange1.Cells[28, 4].Value2) != true)
            {
                xlRange1.Cells[28, 4].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange1.Cells[30, 4].Value2) != true)
            {
                xlRange1.Cells[30, 4].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange1.Cells[58, 4].Value2) != true)
            {
                xlRange1.Cells[58, 4].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange1.Cells[58, 3].Value2) != true)
            {
                xlRange1.Cells[58, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange1.Cells[13, 13].Value2) != true)
            {
                xlRange1.Cells[13, 13].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange1.Cells[17, 13].Value2) != true)
            {
                xlRange1.Cells[17, 13].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange1.Cells[35, 13].Value2) != true)
            {
                xlRange1.Cells[35, 13].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange1.Cells[54, 13].Value2) != true)
            {
                xlRange1.Cells[54, 13].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange1.Cells[57, 13].Value2) != true)
            {
                xlRange1.Cells[57, 13].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[5, 3].Value2) != true)
            {
                xlRange2.Cells[5, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[11, 3].Value2) != true)
            {
                xlRange2.Cells[11, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[23, 3].Value2) != true)
            {
                xlRange2.Cells[23, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[24, 3].Value2) != true)
            {
                xlRange2.Cells[24, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[29, 3].Value2) != true)
            {
                xlRange2.Cells[29, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[30, 3].Value2) != true)
            {
                xlRange2.Cells[30, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[40, 3].Value2) != true)
            {
                xlRange2.Cells[40, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[41, 3].Value2) != true)
            {
                xlRange2.Cells[41, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[45, 3].Value2) != true)
            {
                xlRange2.Cells[45, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[46, 3].Value2) != true)
            {
                xlRange2.Cells[46, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
            if (CheckIfDouble(xlRange2.Cells[65, 3].Value2) != true)
            {
                xlRange2.Cells[65, 3].Interior.Color = Excel.XlRgbColor.rgbRed;
                ReportIsValid = false;
            }
        }
        private void FitToContent(DataGrid dg)
        {
            foreach (DataGridColumn column in dg.Columns)
            {
                column.Width = new DataGridLength(1.0, DataGridLengthUnitType.Auto);
            }
        }
        
        private void Upload_On_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            // Set filter for file extension and default file extension
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "Excel documents (.xlsx)|*.xlsx";
            // Display OpenFileDialog by calling ShowDialog method
            Nullable<bool> result = dlg.ShowDialog();
            ReportIsValid = true;
            // Get the selected file name and display in a TextBox
            if (result == true)
            {
                // Open document
                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet1;
                Excel.Worksheet xlWorkSheet2;
                Excel.Range xlRange1;
                Excel.Range xlRange2;
                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(dlg.FileName);
                xlWorkSheet1 = xlWorkBook.Sheets[1];
                xlWorkSheet2 = xlWorkBook.Sheets[2];
                xlRange1 = xlWorkSheet1.UsedRange;
                xlRange2 = xlWorkSheet2.UsedRange;
                report_Validation(xlRange1, xlRange2);
                if (ReportIsValid == true)
                {
                    Report rep = new Report((Double)xlRange1.Cells[28, 4].Value2, (Double)xlRange1.Cells[30, 4].Value2, (Double)xlRange1.Cells[58, 4].Value2, (Double)xlRange1.Cells[58, 3].Value2, (Double)xlRange1.Cells[13, 13].Value2, (Double)xlRange1.Cells[17, 13].Value2, (Double)xlRange1.Cells[35, 13].Value2, (Double)xlRange1.Cells[54, 13].Value2, (Double)xlRange1.Cells[57, 13].Value2, (Double)xlRange2.Cells[5, 3].Value2, (Double)xlRange2.Cells[11, 3].Value2, (Double)xlRange2.Cells[23, 3].Value2, (Double)xlRange2.Cells[24, 3].Value2, (Double)xlRange2.Cells[29, 3].Value2, (Double)xlRange2.Cells[30, 3].Value2, (Double)xlRange2.Cells[40, 3].Value2, (Double)xlRange2.Cells[41, 3].Value2, (Double)xlRange2.Cells[45, 3].Value2, (Double)xlRange2.Cells[46, 3].Value2, (Double)xlRange2.Cells[65, 3].Value2, (Int32)xlRange1.Cells[2, 7].Value2);
                    reportCollection.AddReport(rep);
                }
                else
                {
                    MessageBox.Show("Перевірте правильність заповнення шаблону (обов'язкові рядки підкреслені червоним)");
                    xlWorkBook.Save();
                }
                GC.Collect();
                GC.WaitForPendingFinalizers();
                Marshal.ReleaseComObject(xlRange1);
                Marshal.ReleaseComObject(xlWorkSheet1);
                Marshal.ReleaseComObject(xlRange2);
                Marshal.ReleaseComObject(xlWorkSheet2);
                xlWorkBook.Close();
                Marshal.ReleaseComObject(xlWorkBook);
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);
            }
        }
        private void Download_Template_On_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            xlApp = new Excel.Application();
            var path = System.IO.Path.GetFullPath(@"BDiag.xlsx");
            xlWorkBook = xlApp.Workbooks.Open(path);
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.FileName = "Diag.xlsx";
            saveFileDialog.DefaultExt = ".xlsx";
            saveFileDialog.Filter = "Excel documents (.xlsx)|*.xlsx";
            saveFileDialog.ShowDialog();
            xlWorkBook.SaveCopyAs(System.IO.Path.GetFullPath(saveFileDialog.FileName));
            GC.Collect();
            GC.WaitForPendingFinalizers();
            xlWorkBook.Close();
            Marshal.ReleaseComObject(xlWorkBook);
            xlApp.Quit();
            Marshal.ReleaseComObject(xlApp);
        }
        private void Two_factor_on_Click(object sender, RoutedEventArgs e)
        {
            DiagPanel.Visibility = Visibility.Visible;
            StackPanel stackPanel1 = new StackPanel();
            double k2parameter;
            double k3parameter;
            double bankruptcyProb;
            double lastBPresult=0;
            textDiag = new TextBlock();
            TextBlock textK2parameter = new TextBlock();
            textK2parameter.FontSize = 18;
            textK2parameter.FontFamily = new FontFamily("Times New Roman");
            textK2parameter.Text = "\n Нормативним значенням показника поточної ліквідності є значення в рамках 1-3, однак більш бажаним є значення 2-3. Показник нижче нормативного свідчить про проблемний стан платоспроможності, адже оборотних активів недостатньо для того, щоб відповісти за поточними зобов'язаннями. Це веде до зниження довіри до компанії з боку кредиторів, постачальників, інвесторів і партнерів. Крім цього, проблеми з платоспроможністю ведуть до збільшення вартості позикових коштів і, як результат, до прямих фінансових втрат.\n";
            TextBlock textK3parameter = new TextBlock();
            textK3parameter.FontSize = 18;
            textK3parameter.FontFamily = new FontFamily("Times New Roman");
            textK3parameter.Text = "\n\n Показник фінансової залежності є індикатором фінансової стійкості, який також вказує на здатність компанії проводити прогнозовану діяльність в довгостроковій перспективі. Значення показника говорить про те, скільки фінансових ресурсів використовує компанія на кожну гривню власного капіталу.\n";
            stackPanel1.HorizontalAlignment = HorizontalAlignment.Center;
            TextBlock textFinalBankruptcyParameter = new TextBlock();
            textFinalBankruptcyParameter.FontSize = 18;
            textFinalBankruptcyParameter.FontFamily = new FontFamily("Times New Roman");
            textFinalBankruptcyParameter.Text = "\n\n Ймовірність настання банкрутства за двофакторною моделлю.";
            dataGrid1 = new DataGrid();
            DataGridTextColumn c1 = new DataGridTextColumn();
            c1.Header = "Рік";
            c1.Binding = new Binding("K1");
            dataGrid1.Columns.Add(c1);
            DataGridTextColumn c2 = new DataGridTextColumn();
            c2.Header = "Коефіцієнт\nпоточної\nлівкідності";
            c2.Binding = new Binding("K2");
            dataGrid1.Columns.Add(c2);
            DataGridTextColumn c3 = new DataGridTextColumn();
            c3.Header = "Коефіцієнт\nфінансової\nзалежності";
            c3.Binding = new Binding("K3");
            dataGrid1.Columns.Add(c3);
            DataGridTextColumn c4 = new DataGridTextColumn();
            c4.Header = "Ступінь\nімовірності\nбанкрутства";
            c4.Binding = new Binding("K4");
            dataGrid1.Columns.Add(c4);
            for (int i = 0; i < reportCollection.Count; i++)
            {
                k2parameter = Math.Round(reportCollection.GetReport(i).elem_1_1195 / reportCollection.GetReport(i).elem_1_1695, 4);
                textK2parameter.Text = textK2parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році значення показника поточної ліквідності становило " + k2parameter.ToString();
                if (k2parameter < 1)
                {
                    textK2parameter.Text = textK2parameter.Text + ". Це свідчить про проблемний стан платоспроможності. ";
                }
                else
                {
                    textK2parameter.Text = textK2parameter.Text + ". Це свідчить про хороший стан платоспроможності. ";
                }
                if (i > 0)
                {
                    textK2parameter.Text = textK2parameter.Text + "У порівнянні з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком значення показника ";
                    if (reportCollection.GetReport(i - 1).elem_1_1195 / reportCollection.GetReport(i - 1).elem_1_1695 > reportCollection.GetReport(i).elem_1_1195 / reportCollection.GetReport(i).elem_1_1695)
                    {
                        textK2parameter.Text = textK2parameter.Text + "зменшилось на " + (Math.Round((100 * (k2parameter - reportCollection.GetReport(i - 1).elem_1_1195 / reportCollection.GetReport(i - 1).elem_1_1695) / k2parameter),2)).ToString() + " відсотків.";
                    }
                    else if (reportCollection.GetReport(i - 1).elem_1_1195 / reportCollection.GetReport(i - 1).elem_1_1695 < reportCollection.GetReport(i).elem_1_1195 / reportCollection.GetReport(i).elem_1_1695)
                    {
                        textK2parameter.Text = textK2parameter.Text + "збільшилось на " + (Math.Round((100 * (k2parameter - reportCollection.GetReport(i - 1).elem_1_1195 / reportCollection.GetReport(i - 1).elem_1_1695) / k2parameter), 2)).ToString() + " відсотків.";
                    }
                    else
                    {
                        textK2parameter.Text = textK2parameter.Text + "залишилось незмінним.";
                    }
                }
                k3parameter = Math.Round((reportCollection.GetReport(i).elem_1_1595 + reportCollection.GetReport(i).elem_1_1695) / reportCollection.GetReport(i).elem_1_1900, 4);
                textK3parameter.Text = textK3parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році значення показника фінансової залежності становило " + k3parameter.ToString();
                if (k3parameter < 1.67)
                {
                    textK3parameter.Text = textK3parameter.Text + ". Це свідчить про неповне використання фінансових можливостей компанією. ";
                }
                else if (k3parameter > 2.5)
                {
                    textK3parameter.Text = textK3parameter.Text + ". Це свідчить про високий рівень фінансових ризиків. ";
                }
                else
                {
                    textK3parameter.Text = textK3parameter.Text + ". Показник досяг свого нормативного (бажаного) значення. ";
                }
                if (i > 0)
                {
                    textK3parameter.Text = textK3parameter.Text + "У порівнянні з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком значення показника ";
                    if (((reportCollection.GetReport(i - 1).elem_1_1595 + reportCollection.GetReport(i - 1).elem_1_1695) / reportCollection.GetReport(i - 1).elem_1_1900) > k3parameter)
                    {
                        textK3parameter.Text = textK3parameter.Text + "зменшилось на " + (Math.Round((100 * (k3parameter - (reportCollection.GetReport(i - 1).elem_1_1595 + reportCollection.GetReport(i - 1).elem_1_1695) / reportCollection.GetReport(i - 1).elem_1_1900) / k2parameter), 2)).ToString() + " відсотків.";
                    }
                    else if (reportCollection.GetReport(i - 1).elem_1_1195 / reportCollection.GetReport(i - 1).elem_1_1695 < reportCollection.GetReport(i).elem_1_1195 / reportCollection.GetReport(i).elem_1_1695)
                    {
                        textK3parameter.Text = textK3parameter.Text + "збільшилось на " + (Math.Round((100 * (k3parameter - (reportCollection.GetReport(i - 1).elem_1_1595 + reportCollection.GetReport(i - 1).elem_1_1695) / reportCollection.GetReport(i - 1).elem_1_1900) / k2parameter), 2)).ToString() + " відсотків.";
                    }
                    else
                    {
                        textK3parameter.Text = textK3parameter.Text + "залишилось незмінним.";
                    }
                }
                bankruptcyProb = Math.Round((-0.3877 - 1.0736 * k2parameter + 0.0579 * k3parameter), 4);
                textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році ймовірність настання банкрутства була ";
                if (bankruptcyProb > 0.3)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "високою (" + bankruptcyProb.ToString() + " > 0.3).";
                }
                else if (bankruptcyProb < -0.3)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "низькою (" + bankruptcyProb.ToString() + " < -0.3).";
                }
                else
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "середньою (-0.3 < " + bankruptcyProb.ToString() + " < 0.3).";
                }
                if (i > 0)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " В порівнянні з " + reportCollection.GetReport(i).rep_year.ToString() + " роком ймовірність банкрутства ";
                    if (lastBPresult < bankruptcyProb)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " збільшилась.";
                    }
                    else if (lastBPresult > bankruptcyProb)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " зменшилась.";
                    }
                    else
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " залишилась незмінною.";
                    }
                }
                lastBPresult = bankruptcyProb;
                dataGrid1.Items.Add(new { K1 = reportCollection.GetReport(i).rep_year, K2 = k2parameter, K3 = k3parameter, K4 = bankruptcyProb });
            }
            dataGrid1.HorizontalAlignment = HorizontalAlignment.Center;
            FitToContent(dataGrid1);
            stackPanel1.Children.Add(dataGrid1);
            textDiag.FontSize = 18;
            textDiag.FontFamily = new FontFamily("Times New Roman");
            textDiag.TextAlignment = TextAlignment.Justify;
            textDiag.TextWrapping = TextWrapping.Wrap;
            textDiag.Padding = new Thickness(5,5,5,5);
            textDiag.Text = textK2parameter.Text + textK3parameter.Text + textFinalBankruptcyParameter.Text;
            stackPanel1.Children.Add(textDiag);
            scrollDiag.Content = stackPanel1;
        }

        private void Altman_on_Click(object sender, RoutedEventArgs e)
        {
            DiagPanel.Visibility = Visibility.Visible;
            StackPanel stackPanel1 = new StackPanel();
            double k2parameter;
            double k3parameter;
            double k4parameter;
            double lastk4parameter = 0;
            double k5parameter;
            double lastk5parameter = 0;
            double k6parameter;
            double bankruptcyProb;
            double lastBPparameter = 0;
            TextBlock textK4parameter = new TextBlock();
            textK4parameter.FontSize = 18;
            textK4parameter.FontFamily = new FontFamily("Times New Roman");
            textK4parameter.Text = "\n Рентабельність активів (відношення прибутку до сплати відсотків і податків (EBIT) до загальної вартості активів), відображає ефективність операційної діяльності підприємства, показує, скільки прибутку припадає на 1 грн вкладених активів (інвестицій). \n";
            TextBlock textK5parameter = new TextBlock();
            textK5parameter.FontSize = 18;
            textK5parameter.FontFamily = new FontFamily("Times New Roman");
            textK5parameter.Text = "\n\n Коефіцієнт фінансової стійкості характеризує фінансову стійкість підприємства. Він показує скільки грн. власного капіталу припадає на 1 грн залученого капіталу. \n";
            TextBlock textFinalBankruptcyParameter = new TextBlock();
            textFinalBankruptcyParameter.FontSize = 18;
            textFinalBankruptcyParameter.FontFamily = new FontFamily("Times New Roman");
            textFinalBankruptcyParameter.Text = "\n\n Ймовірність настання банкрутства за показником Альтмана.";
            stackPanel1.HorizontalAlignment = HorizontalAlignment.Center;
            dataGrid1 = new DataGrid();
            DataGridTextColumn c1 = new DataGridTextColumn();
            c1.Header = "Рік";
            c1.Binding = new Binding("K1");
            dataGrid1.Columns.Add(c1);
            DataGridTextColumn c2 = new DataGridTextColumn();
            c2.Header = "Частка оборотного\nкапіталу в активах\nпідприємства";
            c2.Binding = new Binding("K2");
            dataGrid1.Columns.Add(c2);
            DataGridTextColumn c3 = new DataGridTextColumn();
            c3.Header = "Відношення накопиченого\n(нерозподіленого) прибутку\nдо суми активів\nпідприємства";
            c3.Binding = new Binding("K3");
            dataGrid1.Columns.Add(c3);
            DataGridTextColumn c4 = new DataGridTextColumn();
            c4.Header = "Рентабельність\nактивів";
            c4.Binding = new Binding("K4");
            dataGrid1.Columns.Add(c4);
            DataGridTextColumn c5 = new DataGridTextColumn();
            c5.Header = "Коефіцієнт\nфінансової\nстійкості";
            c5.Binding = new Binding("K5");
            dataGrid1.Columns.Add(c5);
            DataGridTextColumn c6 = new DataGridTextColumn();
            c6.Header = "Відношення\nобсягу продажів до\nзагальної величини\nактивів підприємства";
            c6.Binding = new Binding("K6");
            dataGrid1.Columns.Add(c6);
            DataGridTextColumn c7 = new DataGridTextColumn();
            c7.Header = "Показник\nАльтмана";
            c7.Binding = new Binding("K7");
            dataGrid1.Columns.Add(c7);
            for (int i = 0; i < reportCollection.Count; i++)
            {
                k2parameter = Math.Round((reportCollection.GetReport(i).elem_1_1195 - reportCollection.GetReport(i).elem_1_1695)/ reportCollection.GetReport(i).elem_1_1900, 4);
                if (reportCollection.GetReport(i).elem_2_2350 != 0)
                {
                    k3parameter = Math.Round(reportCollection.GetReport(i).elem_2_2350 / reportCollection.GetReport(i).elem_1_1900, 4);
                }
                else
                {
                    k3parameter = Math.Round(reportCollection.GetReport(i).elem_2_2355 / reportCollection.GetReport(i).elem_1_1900, 4);
                }
                if (reportCollection.GetReport(i).elem_2_2290 != 0)
                {
                    k4parameter = Math.Round(reportCollection.GetReport(i).elem_2_2290 / reportCollection.GetReport(i).elem_1_1900, 4);
                }
                else
                {
                    k4parameter = Math.Round(reportCollection.GetReport(i).elem_2_2295 / reportCollection.GetReport(i).elem_1_1900, 4);
                }
                textK4parameter.Text = textK4parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році рентабельність активів становила " + k4parameter.ToString() + ".";
                if (i > 0)
                {
                    textK4parameter.Text = textK4parameter.Text + " В порівнянні з " + reportCollection.GetReport(i-1).rep_year.ToString() + " роком рентабельність активів ";
                    if (k4parameter > lastk4parameter)
                    {
                        textK4parameter.Text = textK4parameter.Text + "збільшилась на " + (Math.Round((100*(k4parameter-lastk4parameter)/k4parameter), 2)).ToString() +" відсотків.";
                    }
                    else if(k4parameter < lastk4parameter)
                    {
                        textK4parameter.Text = textK4parameter.Text + "зменшилась на " + (Math.Round((100 * (k4parameter - lastk4parameter) / k4parameter), 2)).ToString() + " відсотків.";
                    }
                    else
                    {
                        textK4parameter.Text = textK4parameter.Text + "залишилась незмінною.";
                    }
                }
                lastk4parameter = k4parameter;
                k5parameter = Math.Round(reportCollection.GetReport(i).elem_1_1495 / (reportCollection.GetReport(i).elem_1_1695 + reportCollection.GetReport(i).elem_1_1595),4);
                textK5parameter.Text = textK5parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році коефіцієнт фінансової стійкості становив " + k5parameter.ToString() + ", тобто на 1 грн. власного капіталу припадало " + k5parameter.ToString() + " грн. залученого.";
                if (i > 0)
                {
                    textK5parameter.Text = textK5parameter.Text + " Порівняно з " + reportCollection.GetReport(i-1).rep_year.ToString() + " роком коефіцієнт фінансової стійкості ";
                    if (k5parameter > lastk5parameter)
                    {
                        textK5parameter.Text = textK5parameter.Text + " збільшився на " + (Math.Round((k5parameter-lastk5parameter)/k5parameter,4)).ToString() + " відсотків.";
                    }
                    else if (k5parameter < lastk5parameter)
                    {
                        textK5parameter.Text = textK5parameter.Text + " зменшився на " + (Math.Round((k5parameter - lastk5parameter) / k5parameter, 4)).ToString() + " відсотків.";
                    }
                    else
                    {
                        textK5parameter.Text = textK5parameter.Text + " залишився незмінним.";
                    }
                }
                lastk5parameter = k5parameter;
                k6parameter = Math.Round(reportCollection.GetReport(i).elem_2_2000 / reportCollection.GetReport(i).elem_1_1900, 4);
                bankruptcyProb = Math.Round((1.2 * k2parameter + 1.4 * k3parameter + 3.3 * k4parameter + 0.6 * k5parameter + k6parameter),4);
                textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році ймовірність настання банкрутства була ";
                if (bankruptcyProb <= 1.8)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "дуже високою (" + bankruptcyProb.ToString()+" <= 1.8).";
                }
                else if (bankruptcyProb>1.8 && bankruptcyProb<=2.7)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "високою (1.8 < " + bankruptcyProb.ToString() + " <= 2.7).";
                }
                else if (bankruptcyProb>2.7 && bankruptcyProb < 3)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " середньою (2.7 < " + bankruptcyProb.ToString() + " < 3).";
                }
                else
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " дуже низькою (3 <= " + bankruptcyProb.ToString() + ").";
                }
                if (i > 0)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "Порівняно з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком ймовірність банкрутства ";
                    if (lastBPparameter < bankruptcyProb)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " зменшилась.";
                    }
                    else if (lastBPparameter > bankruptcyProb)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " збільшилась.";
                    }
                    else
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " залишилась незмінною.";
                    }
                }
                lastBPparameter = bankruptcyProb;
                dataGrid1.Items.Add(new { K1 = reportCollection.GetReport(i).rep_year, K2 = k2parameter, K3 = k3parameter, K4 = k4parameter, K5 = k5parameter, K6 = k6parameter, K7 = bankruptcyProb });
            }
            dataGrid1.HorizontalAlignment = HorizontalAlignment.Center;
            FitToContent(dataGrid1);
            stackPanel1.Children.Add(dataGrid1);
            textDiag = new TextBlock();
            textDiag.FontSize = 18;
            textDiag.FontFamily = new FontFamily("Times New Roman");
            textDiag.TextAlignment = TextAlignment.Justify;
            textDiag.TextWrapping = TextWrapping.Wrap;
            textDiag.Padding = new Thickness(5, 5, 5, 5);
            textDiag.Text = textK4parameter.Text + textK5parameter.Text + textFinalBankruptcyParameter.Text;
            stackPanel1.Children.Add(textDiag);
            scrollDiag.Content = stackPanel1;
        }

        private void Lis_on_Click(object sender, RoutedEventArgs e)
        {
            DiagPanel.Visibility = Visibility.Visible;
            StackPanel stackPanel1 = new StackPanel();
            double k2parameter;
            double k3parameter;
            double k4parameter;
            double k5parameter;
            double lastk5parameter = 0;
            double bankruptcyProb;
            TextBlock textK5parameter = new TextBlock();
            textK5parameter.FontSize = 18;
            textK5parameter.FontFamily = new FontFamily("Times New Roman");
            textK5parameter.Text = "\n Коефіцієнт фінансової стійкості характеризує фінансову стійкість підприємства. Він показує скільки грн. власного капіталу припадає на 1 грн залученого капіталу. \n";
            TextBlock textFinalBankruptcyParameter = new TextBlock();
            textFinalBankruptcyParameter.FontSize = 18;
            textFinalBankruptcyParameter.FontFamily = new FontFamily("Times New Roman");
            textFinalBankruptcyParameter.Text = "\n\n Ймовірність настання банкрутства за показником Ліса.";
            stackPanel1.HorizontalAlignment = HorizontalAlignment.Center;
            dataGrid1 = new DataGrid();
            DataGridTextColumn c1 = new DataGridTextColumn();
            c1.Header = "Рік";
            c1.Binding = new Binding("K1");
            dataGrid1.Columns.Add(c1);
            DataGridTextColumn c2 = new DataGridTextColumn();
            c2.Header = "Відношення\nоборотних активів\nдо балансу";
            c2.Binding = new Binding("K2");
            dataGrid1.Columns.Add(c2);
            DataGridTextColumn c3 = new DataGridTextColumn();
            c3.Header = "Відношення\nприбутку від основної\nдіяльності до балансу";
            c3.Binding = new Binding("K3");
            dataGrid1.Columns.Add(c3);
            DataGridTextColumn c4 = new DataGridTextColumn();
            c4.Header = "Відношення\nнерозподіленого прибутку\nдо балансу";
            c4.Binding = new Binding("K4");
            dataGrid1.Columns.Add(c4);
            DataGridTextColumn c5 = new DataGridTextColumn();
            c5.Header = "Коефіцієнт\nфінансової\nстійкості";
            c5.Binding = new Binding("K5");
            dataGrid1.Columns.Add(c5);
            DataGridTextColumn c6 = new DataGridTextColumn();
            c6.Header = "Показник\nЛіса";
            c6.Binding = new Binding("K6");
            dataGrid1.Columns.Add(c6);
            for (int i = 0; i < reportCollection.Count; i++)
            {
                k2parameter = Math.Round(reportCollection.GetReport(i).elem_1_1195 / reportCollection.GetReport(i).elem_1_1900, 4);
                if (reportCollection.GetReport(i).elem_2_2190 != 0)
                {
                    k3parameter = Math.Round(reportCollection.GetReport(i).elem_2_2190 / reportCollection.GetReport(i).elem_1_1900, 4);
                }
                else
                {
                    k3parameter = Math.Round(reportCollection.GetReport(i).elem_2_2195 / reportCollection.GetReport(i).elem_1_1900, 4);
                }
                k4parameter = Math.Round(reportCollection.GetReport(i).elem_1_1420 / reportCollection.GetReport(i).elem_1_1900, 4);
                k5parameter = Math.Round(reportCollection.GetReport(i).elem_1_1495 / (reportCollection.GetReport(i).elem_1_1900 - reportCollection.GetReport(i).elem_1_1495), 4);
                textK5parameter.Text = textK5parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році коефіцієнт фінансової стійкості становив " + k5parameter.ToString() + ", тобто на 1 грн. власного капіталу припадало " + k5parameter.ToString() + " грн. залученого.";
                if (i > 0)
                {
                    textK5parameter.Text = textK5parameter.Text + " Порівняно з " + reportCollection.GetReport(i-1).rep_year.ToString() + " роком коефіцієнт фінансової стійкості ";
                    if (k5parameter > lastk5parameter)
                    {
                        textK5parameter.Text = textK5parameter.Text + " збільшився на " + (Math.Round((k5parameter-lastk5parameter)/k5parameter,4)).ToString() + " відсотків.";
                    }
                    else if (k5parameter < lastk5parameter)
                    {
                        textK5parameter.Text = textK5parameter.Text + " зменшився на " + (Math.Round((k5parameter - lastk5parameter) / k5parameter, 4)).ToString() + " відсотків.";
                    }
                    else
                    {
                        textK5parameter.Text = textK5parameter.Text + " залишився незмінним.";
                    }
                }
                lastk5parameter = k5parameter;
                bankruptcyProb = Math.Round(0.063 * k2parameter + 0.092 * k3parameter + 0.057 * k4parameter + 0.001 * k5parameter, 4);
                textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році показник Ліса становив " + bankruptcyProb.ToString() + " - підприємству ";
                if (bankruptcyProb < 0.037)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "загрожує банкрутство.";
                }
                else
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "не загрожує банкрутство.";
                }
                dataGrid1.Items.Add(new { K1 = reportCollection.GetReport(i).rep_year, K2 = k2parameter, K3 = k3parameter, K4 = k4parameter, K5 = k5parameter, K6 = bankruptcyProb });
            }
            dataGrid1.HorizontalAlignment = HorizontalAlignment.Center;
            FitToContent(dataGrid1);
            stackPanel1.Children.Add(dataGrid1);
            textDiag = new TextBlock();
            textDiag.FontSize = 18;
            textDiag.FontFamily = new FontFamily("Times New Roman");
            textDiag.TextAlignment = TextAlignment.Justify;
            textDiag.TextWrapping = TextWrapping.Wrap;
            textDiag.Padding = new Thickness(5, 5, 5, 5);
            textDiag.Text = textK5parameter.Text + textFinalBankruptcyParameter.Text;
            stackPanel1.Children.Add(textDiag);
            scrollDiag.Content = stackPanel1;
        }

        private void Taffler_on_Click(object sender, RoutedEventArgs e)
        {
            DiagPanel.Visibility = Visibility.Visible;
            StackPanel stackPanel1 = new StackPanel();
            double k2parameter;
            double k3parameter;
            double k4parameter;
            double k5parameter;
            double bankruptcyProb;
            double lastBPparameter = 0;
            TextBlock textFinalBankruptcyParameter = new TextBlock();
            textFinalBankruptcyParameter.FontSize = 18;
            textFinalBankruptcyParameter.FontFamily = new FontFamily("Times New Roman");
            textFinalBankruptcyParameter.Text = "\n Ймовірність настання банкрутства за показником Таффлера.";
            stackPanel1.HorizontalAlignment = HorizontalAlignment.Center;
            dataGrid1 = new DataGrid();
            DataGridTextColumn c1 = new DataGridTextColumn();
            c1.Header = "Рік";
            c1.Binding = new Binding("K1");
            dataGrid1.Columns.Add(c1);
            DataGridTextColumn c2 = new DataGridTextColumn();
            c2.Header = "Відношення\nопераційного прибутку до\nпоточних зобов'язань";
            c2.Binding = new Binding("K2");
            dataGrid1.Columns.Add(c2);
            DataGridTextColumn c3 = new DataGridTextColumn();
            c3.Header = "Відношення\nоборотних активів\nдо зобов'язань";
            c3.Binding = new Binding("K3");
            dataGrid1.Columns.Add(c3);
            DataGridTextColumn c4 = new DataGridTextColumn();
            c4.Header = "Відношення\nпоточних зобов'язань\nдо балансу";
            c4.Binding = new Binding("K4");
            dataGrid1.Columns.Add(c4);
            DataGridTextColumn c5 = new DataGridTextColumn();
            c5.Header = "Відношення\nчистого доходу\nдо балансу";
            c5.Binding = new Binding("K5");
            dataGrid1.Columns.Add(c5);
            DataGridTextColumn c6 = new DataGridTextColumn();
            c6.Header = "Показник\nТаффлера";
            c6.Binding = new Binding("K6");
            dataGrid1.Columns.Add(c6);
            for (int i = 0; i < reportCollection.Count; i++)
            {
                k2parameter = Math.Round((reportCollection.GetReport(i).elem_2_2000 + reportCollection.GetReport(i).elem_2_2050 + reportCollection.GetReport(i).elem_2_2130 + reportCollection.GetReport(i).elem_2_2150)/reportCollection.GetReport(i).elem_1_1695, 4);
                k3parameter = Math.Round(reportCollection.GetReport(i).elem_1_1195/(reportCollection.GetReport(i).elem_1_1595 + reportCollection.GetReport(i).elem_1_1695),4);
                k4parameter = Math.Round(reportCollection.GetReport(i).elem_1_1695 / reportCollection.GetReport(i).elem_1_1900, 4);
                k5parameter = Math.Round(reportCollection.GetReport(i).elem_2_2000 / reportCollection.GetReport(i).elem_1_1900, 4);
                bankruptcyProb = Math.Round(0.53*k2parameter + 0.13*k3parameter + 0.18*k4parameter + 0.16*k5parameter, 4);
                textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році показник Таффлера становив " + bankruptcyProb.ToString() + ".";
                if (bankruptcyProb < 0.2)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " Підприємству загрожувало банкрутство.";
                }
                else if (bankruptcyProb>0.3)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " Підприємство перебувало в хорошому фінансовому стані.";
                }
                else
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " Ймовірність банкрутства була на середньому рівні.";
                }
                if (i > 0)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "Порівняно з " + reportCollection.GetReport(i-1).rep_year.ToString() + " роком ймовірність банкрутства ";
                    if (bankruptcyProb > lastBPparameter)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " збільшилась.";
                    }
                    else if(bankruptcyProb < lastBPparameter)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " зменшилась.";
                    }
                    else
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " не змінилась.";
                    }
                }
                lastBPparameter = bankruptcyProb;
                dataGrid1.Items.Add(new { K1 = reportCollection.GetReport(i).rep_year, K2 = k2parameter, K3 = k3parameter, K4 = k4parameter, K5 = k5parameter, K6 = bankruptcyProb });
            }
            dataGrid1.HorizontalAlignment = HorizontalAlignment.Center;
            FitToContent(dataGrid1);
            stackPanel1.Children.Add(dataGrid1);
            textDiag = new TextBlock();
            textDiag.FontSize = 18;
            textDiag.FontFamily = new FontFamily("Times New Roman");
            textDiag.TextAlignment = TextAlignment.Justify;
            textDiag.TextWrapping = TextWrapping.Wrap;
            textDiag.Padding = new Thickness(5, 5, 5, 5);
            textDiag.Text = textFinalBankruptcyParameter.Text;
            stackPanel1.Children.Add(textDiag);
            scrollDiag.Content = stackPanel1;
        }

        private void Beaver_on_Click(object sender, RoutedEventArgs e)
        {
            DiagPanel.Visibility = Visibility.Visible;
            StackPanel stackPanel1 = new StackPanel();
            double k2parameter;
            double lastk2parameter = 0;
            double k3parameter;
            double lastk3parameter = 0;
            double k4parameter;
            double lastk4parameter = 0;
            double k5parameter;
            double lastk5parameter = 0;
            double k6parameter;
            TextBlock textK2parameter = new TextBlock();
            textK2parameter.FontSize = 18;
            textK2parameter.FontFamily = new FontFamily("Times New Roman");
            textK2parameter.Text = "\n Ймовірність настання банкрутства за коефіцієнтом Бівера.";
            TextBlock textK3parameter = new TextBlock();
            textK3parameter.FontSize = 18;
            textK3parameter.FontFamily = new FontFamily("Times New Roman");
            textK3parameter.Text = "\n\n Рентабельність активів відображає ефективність операційної діяльності підприємства і показує скільки гривень чистого прибутку припадає на 1 грн вкладених активів (інвестицій). \n";
            TextBlock textK4parameter = new TextBlock();
            textK4parameter.FontSize = 18;
            textK4parameter.FontFamily = new FontFamily("Times New Roman");
            textK4parameter.Text = "\n\n Показник фінансового левереджу демонструє залежність підприємства від зовнішніх кредиторів.";
            TextBlock textK5parameter = new TextBlock();
            textK5parameter.FontSize = 18;
            textK5parameter.FontFamily = new FontFamily("Times New Roman");
            textK5parameter.Text = "\n\n Коефіцієнт маневреності власного капіталу показує, яка частина власного капіталу використовується для фінансування поточної діяльності, тобто вкладена в оборотні засоби, а яка — капіталізована. \n";
            TextBlock textK6parameter = new TextBlock();
            textK6parameter.FontSize = 18;
            textK6parameter.FontFamily = new FontFamily("Times New Roman");
            textK6parameter.Text = "\n\n Коефіцієнт покриття показує, яку частину поточних зобов’язань підприємство спроможне погасити за рахунок оборотних активів. \n";
            stackPanel1.HorizontalAlignment = HorizontalAlignment.Center;
            dataGrid1 = new DataGrid();
            DataGridTextColumn c1 = new DataGridTextColumn();
            c1.Header = "Рік";
            c1.Binding = new Binding("K1");
            dataGrid1.Columns.Add(c1);
            DataGridTextColumn c2 = new DataGridTextColumn();
            c2.Header = "Коефіцієнт\nБівера";
            c2.Binding = new Binding("K2");
            dataGrid1.Columns.Add(c2);
            DataGridTextColumn c3 = new DataGridTextColumn();
            c3.Header = "Рентабельність\nактивів";
            c3.Binding = new Binding("K3");
            dataGrid1.Columns.Add(c3);
            DataGridTextColumn c4 = new DataGridTextColumn();
            c4.Header = "Фінансовий\nлевередж";
            c4.Binding = new Binding("K4");
            dataGrid1.Columns.Add(c4);
            DataGridTextColumn c5 = new DataGridTextColumn();
            c5.Header = "Коефіцієнт\nманврування";
            c5.Binding = new Binding("K5");
            dataGrid1.Columns.Add(c5);
            DataGridTextColumn c6 = new DataGridTextColumn();
            c6.Header = "Коефіцієнт\nпокриття";
            c6.Binding = new Binding("K6");
            dataGrid1.Columns.Add(c6);
            for (int i = 0; i < reportCollection.Count; i++)
            {
                if (reportCollection.GetReport(i).elem_2_2350 != 0)
                {
                    k2parameter = Math.Round((reportCollection.GetReport(i).elem_2_2350 + reportCollection.GetReport(i).elem_2_2515) / (reportCollection.GetReport(i).elem_1_1595 + reportCollection.GetReport(i).elem_1_1695), 4);
                    k3parameter = Math.Round(reportCollection.GetReport(i).elem_2_2350 / reportCollection.GetReport(i).elem_1_1900, 4);
                }
                else
                {
                    k2parameter = Math.Round((reportCollection.GetReport(i).elem_2_2355 + reportCollection.GetReport(i).elem_2_2515) / (reportCollection.GetReport(i).elem_1_1595 + reportCollection.GetReport(i).elem_1_1695), 4);
                    k3parameter = Math.Round(reportCollection.GetReport(i).elem_2_2355 / reportCollection.GetReport(i).elem_1_1900, 4);
                }
                textK2parameter.Text = textK2parameter.Text + " Згідно з коефіцієнтом Бівера, у " + reportCollection.GetReport(i).rep_year.ToString() + " році підприємство знаходилось ";
                if (k2parameter > 0.4)
                {
                    textK2parameter.Text = textK2parameter.Text + "у доброму фінансовому стані.";
                }
                else if (k2parameter > 0.17 && k2parameter < 0.4)
                {
                    textK2parameter.Text = textK2parameter.Text + "за 5 років до банкрутства.";
                }
                else
                {
                    textK2parameter.Text = textK2parameter.Text + "за 1 рік до банкрутства.";
                }
                if (i > 0)
                {
                    textK2parameter.Text = textK2parameter.Text + " В порівнянні з " + reportCollection.GetReport(i-1).rep_year.ToString() + " роком ймовірність банкрутства ";
                    if (k2parameter < lastk2parameter)
                    {
                        textK2parameter.Text = textK2parameter.Text + "збільшилась.";
                    }
                    else if (k2parameter > lastk2parameter)
                    {
                        textK2parameter.Text = textK2parameter.Text + "зменшилась";
                    }
                    else
                    {
                        textK2parameter.Text = textK2parameter.Text + "не змінилась.";
                    }
                }
                lastk2parameter = k2parameter;
                textK3parameter.Text = textK3parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році рентабельність активів становила " + k3parameter.ToString() + ", тобто на 1 гривню вкладених активів припадало " + k3parameter.ToString() + " грн. ";
                if (k3parameter < 0)
                {
                    textK3parameter.Text = textK3parameter.Text + "збитку.";
                }
                else
                {
                    textK3parameter.Text = textK3parameter.Text + "прибутку.";
                }
                if (i > 0)
                {
                    textK3parameter.Text = textK3parameter.Text + " У порівнянні з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком показник ";
                    if (k3parameter > lastk3parameter)
                    {
                        textK3parameter.Text = textK3parameter.Text + "збільшився на " + Math.Round((k3parameter-lastk3parameter)/k3parameter,2) + " відсотків.";
                    }
                    else if (k3parameter<lastk3parameter)
                    {
                        textK3parameter.Text = textK3parameter.Text + "зменшився на " + Math.Round((k3parameter - lastk3parameter) / k3parameter, 2) + " відсотків.";
                    }
                    else
                    {
                        textK3parameter.Text = textK3parameter.Text + "не змінився.";
                    }
                }
                lastk3parameter = k3parameter;
                k4parameter = Math.Round((reportCollection.GetReport(i).elem_1_1595 + reportCollection.GetReport(i).elem_1_1695) / reportCollection.GetReport(i).elem_1_1900, 4);
                textK4parameter.Text = textK4parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році показник фінансового левереджу становив " + k4parameter.ToString() + ".";
                if (i > 0)
                {
                    textK4parameter.Text = textK4parameter.Text + " В порівнянні з " + reportCollection.GetReport(i-1).rep_year.ToString() + " роком показник ";
                    if (k4parameter > lastk4parameter)
                    {
                        textK4parameter.Text = textK4parameter.Text + "збільшився на " + Math.Round((k4parameter - lastk4parameter) / k4parameter, 2) + " відсотків.";
                    }
                    else if (k4parameter < lastk4parameter)
                    {
                        textK4parameter.Text = textK4parameter.Text + "зменшився на " + Math.Round((k4parameter - lastk4parameter) / k4parameter, 2) + " відсотків.";
                    }
                    else
                    {
                        textK4parameter.Text = textK4parameter.Text + "залишився незмінним.";
                    }
                }
                lastk4parameter = k4parameter;
                k5parameter = Math.Round((reportCollection.GetReport(i).elem_1_1495 + reportCollection.GetReport(i).elem_1_1595 - reportCollection.GetReport(i).elem_1_1095) / reportCollection.GetReport(i).elem_1_1900, 4);
                textK5parameter.Text = textK5parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році коефіцієнт маневрування становив " + k5parameter.ToString() + ". Підприємство знаходилось ";
                if (k5parameter >= 0.4)
                {
                    textK5parameter.Text = textK5parameter.Text + "у доброму фінансовому стані.";
                }
                else if (k5parameter >= 0.3 && k5parameter<0.4)
                {
                    textK5parameter.Text = textK5parameter.Text + "за 5 років до банкрутства.";
                }
                else
                {
                    textK5parameter.Text = textK5parameter.Text + "за 1 рік до банкрутства.";
                }
                lastk5parameter = k5parameter;
                k6parameter = Math.Round(reportCollection.GetReport(i).elem_1_1195 / reportCollection.GetReport(i).elem_1_1695, 4);
                textK6parameter.Text = textK6parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році коефіцієнт покриття становив " + k6parameter.ToString() + " відсотків. Підприємство знаходилось ";
                if (k6parameter >= 3.2)
                {
                    textK6parameter.Text = textK6parameter.Text + "у доброму фінансовому стані.";
                }
                else if (k6parameter<3.2 && k6parameter >= 2)
                {
                    textK6parameter.Text = textK6parameter.Text + "за 5 років до банкрутства.";
                }
                else
                {
                    textK6parameter.Text = textK6parameter.Text + "за 1 рік до банкрутства.";
                }
                dataGrid1.Items.Add(new { K1 = reportCollection.GetReport(i).rep_year, K2 = k2parameter, K3 = k3parameter, K4 = k4parameter, K5 = k5parameter, K6 = k6parameter });
            }
            dataGrid1.HorizontalAlignment = HorizontalAlignment.Center;
            FitToContent(dataGrid1);
            stackPanel1.Children.Add(dataGrid1);
            textDiag = new TextBlock();
            textDiag.FontSize = 18;
            textDiag.FontFamily = new FontFamily("Times New Roman");
            textDiag.TextAlignment = TextAlignment.Justify;
            textDiag.TextWrapping = TextWrapping.Wrap;
            textDiag.Padding = new Thickness(5, 5, 5, 5);
            textDiag.Text = textK2parameter.Text + textK3parameter.Text + textK4parameter.Text + textK5parameter.Text + textK6parameter.Text;
            stackPanel1.Children.Add(textDiag);
            scrollDiag.Content = stackPanel1;
        }

        private void Tereshchenko_on_Click(object sender, RoutedEventArgs e)
        {

        }

        private void Springate_on_Click(object sender, RoutedEventArgs e)
        {
            DiagPanel.Visibility = Visibility.Visible;
            StackPanel stackPanel1 = new StackPanel();
            double k2parameter;
            double k3parameter;
            double lastk3parameter = 0;
            double k4parameter;
            double k5parameter;
            double bankruptcyProb;
            double lastBPparameter = 0;
            stackPanel1.HorizontalAlignment = HorizontalAlignment.Center;
            TextBlock textK3parameter = new TextBlock();
            textK3parameter.FontSize = 18;
            textK3parameter.FontFamily = new FontFamily("Times New Roman");
            textK3parameter.Text = "\n Рентабельність активів (відношення прибутку до сплати відсотків і податків (EBIT) до загальної вартості активів), відображає ефективність операційної діяльності підприємства, показує, скільки прибутку припадає на 1 грн вкладених активів (інвестицій). \n";
            TextBlock textFinalBankruptcyParameter = new TextBlock();
            textFinalBankruptcyParameter.FontSize = 18;
            textFinalBankruptcyParameter.FontFamily = new FontFamily("Times New Roman");
            textFinalBankruptcyParameter.Text = "\n\n Ймовірність настання банкрутства за показником Спрінгейта.";
            dataGrid1 = new DataGrid();
            DataGridTextColumn c1 = new DataGridTextColumn();
            c1.Header = "Рік";
            c1.Binding = new Binding("K1");
            dataGrid1.Columns.Add(c1);
            DataGridTextColumn c2 = new DataGridTextColumn();
            c2.Header = "Відношення\nсередньорічних оборотних\nактивів до балансу";
            c2.Binding = new Binding("K2");
            dataGrid1.Columns.Add(c2);
            DataGridTextColumn c3 = new DataGridTextColumn();
            c3.Header = "Рентабельність\nактивів";
            c3.Binding = new Binding("K3");
            dataGrid1.Columns.Add(c3);
            DataGridTextColumn c4 = new DataGridTextColumn();
            c4.Header = "Відношення фінансового\nрезультату (до оподаткування)\nдо поточних зобов'язань";
            c4.Binding = new Binding("K4");
            dataGrid1.Columns.Add(c4);
            DataGridTextColumn c5 = new DataGridTextColumn();
            c5.Header = "Відношення\nчистого доходу\nдо балансу";
            c5.Binding = new Binding("K5");
            dataGrid1.Columns.Add(c5);
            DataGridTextColumn c6 = new DataGridTextColumn();
            c6.Header = "Показник\nСпрінгейта";
            c6.Binding = new Binding("K6");
            dataGrid1.Columns.Add(c6);
            for (int i = 0; i < reportCollection.Count; i++)
            {
                k2parameter = Math.Round((reportCollection.GetReport(i).elem_1_1195 + reportCollection.GetReport(i).elem_1_1195_first) /(2*reportCollection.GetReport(i).elem_1_1900), 4);
                if (reportCollection.GetReport(i).elem_2_2290!=0)
                {
                    k3parameter = Math.Round(reportCollection.GetReport(i).elem_2_2290 / reportCollection.GetReport(i).elem_1_1900, 4);
                    k4parameter = Math.Round(reportCollection.GetReport(i).elem_2_2290 / reportCollection.GetReport(i).elem_1_1695, 4);
                }
                else
                {
                    k3parameter = Math.Round(reportCollection.GetReport(i).elem_2_2295 / reportCollection.GetReport(i).elem_1_1900, 4);
                    k4parameter = Math.Round(reportCollection.GetReport(i).elem_2_2295 / reportCollection.GetReport(i).elem_1_1695, 4);
                }
                textK3parameter.Text = textK3parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році рентабельність активів становила " + k3parameter.ToString() + ", тобто на 1 гривню вкладених активів припадало " + k3parameter.ToString() + " грн. ";
                if (k3parameter < 0)
                {
                    textK3parameter.Text = textK3parameter.Text + "збитку.";
                }
                else
                {
                    textK3parameter.Text = textK3parameter.Text + "прибутку.";
                }
                if (i > 0)
                {
                    textK3parameter.Text = textK3parameter.Text + " У порівнянні з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком показник ";
                    if (k3parameter > lastk3parameter)
                    {
                        textK3parameter.Text = textK3parameter.Text + "збільшився на " + Math.Round((k3parameter - lastk3parameter) / k3parameter, 2) + " відсотків.";
                    }
                    else if (k3parameter < lastk3parameter)
                    {
                        textK3parameter.Text = textK3parameter.Text + "зменшився на " + Math.Round((k3parameter - lastk3parameter) / k3parameter, 2) + " відсотків.";
                    }
                    else
                    {
                        textK3parameter.Text = textK3parameter.Text + "не змінився.";
                    }
                }
                lastk3parameter = k3parameter;
                k5parameter = Math.Round(reportCollection.GetReport(i).elem_2_2000 / reportCollection.GetReport(i).elem_1_1900, 4);
                bankruptcyProb = Math.Round(1.03*k2parameter + 3.07*k3parameter + 0.66*k4parameter + 0.4*k5parameter, 4);
                textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році показник Спрінгейта становив " + bankruptcyProb.ToString() + " - ";
                if (bankruptcyProb < 0.862)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "підприємству загрожувало банкрутство.";
                }
                else
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "підприємство знаходилось у хорошому фінансовому стані.";
                }
                if (i > 0)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "Порівняно з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком ймовірність банкрутства ";
                    if (bankruptcyProb > lastBPparameter)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " зменшилась.";
                    }
                    else if (bankruptcyProb < lastBPparameter)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " збільшилась.";
                    }
                    else
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " не змінилась.";
                    }
                }
                lastBPparameter = bankruptcyProb;
                dataGrid1.Items.Add(new { K1 = reportCollection.GetReport(i).rep_year, K2 = k2parameter, K3 = k3parameter, K4 = k4parameter, K5 = k5parameter, K6 = bankruptcyProb });
            }
            dataGrid1.HorizontalAlignment = HorizontalAlignment.Center;
            FitToContent(dataGrid1);
            stackPanel1.Children.Add(dataGrid1);
            textDiag = new TextBlock();
            textDiag.FontSize = 18;
            textDiag.FontFamily = new FontFamily("Times New Roman");
            textDiag.TextAlignment = TextAlignment.Justify;
            textDiag.TextWrapping = TextWrapping.Wrap;
            textDiag.Padding = new Thickness(5, 5, 5, 5);
            textDiag.Text = textK3parameter.Text + textFinalBankruptcyParameter.Text;
            stackPanel1.Children.Add(textDiag);
            scrollDiag.Content = stackPanel1;
        }

        private void Safulin_Kadikov_on_Click(object sender, RoutedEventArgs e)
        {
            DiagPanel.Visibility = Visibility.Visible;
            StackPanel stackPanel1 = new StackPanel();
            double k2parameter;
            double lastk2parameter = 0;
            double k3parameter;
            double lastk3parameter = 0;
            double k4parameter;
            double k5parameter;
            double lastk5parameter = 0;
            double k6parameter;
            double lastk6parameter = 0;
            double bankruptcyProb;
            double lastBPparameter = 0;
            stackPanel1.HorizontalAlignment = HorizontalAlignment.Center;
            TextBlock textK2parameter = new TextBlock();
            textK2parameter.FontSize = 18;
            textK2parameter.FontFamily = new FontFamily("Times New Roman");
            textK2parameter.Text = "\n Показник забезпечення власними оборотними засобами запасів є індикатором достатності довгострокових коштів компанії для забезпечення безперебійного виробничо-збутового процесу. Нормативним значенням є 0,5 і вище. \n";
            TextBlock textK3parameter = new TextBlock();
            textK3parameter.FontSize = 18;
            textK3parameter.FontFamily = new FontFamily("Times New Roman");
            textK3parameter.Text = "\n\n Нормативним значенням показника поточної ліквідності є значення в рамках 1-3, однак більш бажаним є значення 2-3. Показник нижче нормативного свідчить про проблемний стан платоспроможності, адже оборотних активів недостатньо для того, щоб відповісти за поточними зобов'язаннями. Це веде до зниження довіри до компанії з боку кредиторів, постачальників, інвесторів і партнерів. Крім цього, проблеми з платоспроможністю ведуть до збільшення вартості позикових коштів і, як результат, до прямих фінансових втрат.\n";
            TextBlock textK5parameter = new TextBlock();
            textK5parameter.FontSize = 18;
            textK5parameter.FontFamily = new FontFamily("Times New Roman");
            textK5parameter.Text = "\n\n Рентабельність реалізації продукції за чистим прибутком показує скільки гривень чистого прибутку генерує кожна гривня продажів.\n";
            TextBlock textK6parameter = new TextBlock();
            textK6parameter.FontSize = 18;
            textK6parameter.FontFamily = new FontFamily("Times New Roman");
            textK6parameter.Text = "\n\n Рентабельність власного капіталу вказує, наскільки ефективно використовується власний капітал, тобто скільки прибутку було згенеровано на кожну гривню залучених власних коштів.\n";
            TextBlock textFinalBankruptcyParameter = new TextBlock();
            textFinalBankruptcyParameter.FontSize = 18;
            textFinalBankruptcyParameter.FontFamily = new FontFamily("Times New Roman");
            textFinalBankruptcyParameter.Text = "\n\n Ймовірність настання банкрутства за коефіцієнтом Сайфуліна-Кадикова.";
            dataGrid1 = new DataGrid();
            DataGridTextColumn c1 = new DataGridTextColumn();
            c1.Header = "Рік";
            c1.Binding = new Binding("K1");
            dataGrid1.Columns.Add(c1);
            DataGridTextColumn c2 = new DataGridTextColumn();
            c2.Header = "Коефіцієнт забезпечення\nвласними оборотними\nзасобами запасів";
            c2.Binding = new Binding("K2");
            dataGrid1.Columns.Add(c2);
            DataGridTextColumn c3 = new DataGridTextColumn();
            c3.Header = "Коефіцієнт\nпоточної\nлівкідності";
            c3.Binding = new Binding("K3");
            dataGrid1.Columns.Add(c3);
            DataGridTextColumn c4 = new DataGridTextColumn();
            c4.Header = "Відношення\nчистого доходу\nдо балансу";
            c4.Binding = new Binding("K4");
            dataGrid1.Columns.Add(c4);
            DataGridTextColumn c5 = new DataGridTextColumn();
            c5.Header = "Рентабельність\nреалізації продукції\nза чистим прибутком";
            c5.Binding = new Binding("K5");
            dataGrid1.Columns.Add(c5);
            DataGridTextColumn c6 = new DataGridTextColumn();
            c6.Header = "Рентабельність\nвласного\nкапіталу";
            c6.Binding = new Binding("K6");
            dataGrid1.Columns.Add(c6);
            DataGridTextColumn c7 = new DataGridTextColumn();
            c7.Header = "Коефіцієнт\nСайфуліна-\nКадикова";
            c7.Binding = new Binding("K7");
            dataGrid1.Columns.Add(c7);
            for (int i = 0; i < reportCollection.Count; i++)
            {
                k2parameter = Math.Round((reportCollection.GetReport(i).elem_1_1195 - reportCollection.GetReport(i).elem_1_1695) / reportCollection.GetReport(i).elem_1_1100, 4);
                textK2parameter.Text = textK2parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році коефіцієнт забезпечення власними оборотними засобами запасів становив " + k2parameter + ".";
                if (k2parameter >= 0.5)
                {
                    textK2parameter.Text = textK2parameter.Text + " Коефіцієнт досяг свого нормативного значення ("+ k2parameter.ToString() + " >= 0.5), що свідчить про підвищення стійкості компанії в середньостроковій перспективі і про зниження залежності від короткострокових джерел фінансування.";
                }
                else
                {
                    textK2parameter.Text = textK2parameter.Text + " Коефіцієнт не досяг свого нормативного значення (" + k2parameter.ToString() + " < 0.5) - без короткострокового та довгострокового позикового капіталу компанія не зможе забезпечити безперебійний виробничо-збутової процес.";
                }
                if (i > 0)
                {
                    textK2parameter.Text = textK2parameter.Text + " Порівняно з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком коефіцієнт ";
                    if (k2parameter > lastk2parameter)
                    {
                        textK2parameter.Text = textK2parameter.Text + " збільшився на " + (Math.Round((k2parameter - lastk2parameter) / k2parameter, 4)).ToString() + " відсотків.";
                    }
                    else if (k2parameter < lastk2parameter)
                    {
                        textK2parameter.Text = textK2parameter.Text + " зменшився на " + (Math.Round((k2parameter - lastk2parameter) / k2parameter, 4)).ToString() + " відсотків.";
                    }
                    else
                    {
                        textK2parameter.Text = textK2parameter.Text + " залишився незмінним.";
                    }
                }
                lastk2parameter = k2parameter;
                k3parameter = Math.Round(reportCollection.GetReport(i).elem_1_1195 / reportCollection.GetReport(i).elem_1_1695, 4);
                textK3parameter.Text = textK3parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році значення показника поточної ліквідності становило " + k3parameter.ToString();
                if (k3parameter < 1)
                {
                    textK3parameter.Text = textK3parameter.Text + ". Це свідчить про проблемний стан платоспроможності. ";
                }
                else
                {
                    textK3parameter.Text = textK3parameter.Text + ". Це свідчить про хороший стан платоспроможності. ";
                }
                if (i > 0)
                {
                    textK3parameter.Text = textK3parameter.Text + " В порівнянні з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком значення показника ";
                    if (k3parameter > lastk3parameter)
                    {
                        textK3parameter.Text = textK3parameter.Text + " збільшилось на " + (Math.Round((k3parameter - lastk3parameter) / k3parameter, 4)).ToString() + " відсотків.";
                    }
                    else if (k3parameter < lastk3parameter)
                    {
                        textK3parameter.Text = textK3parameter.Text + " зменшилось на " + (Math.Round((k3parameter - lastk3parameter) / k3parameter, 4)).ToString() + " відсотків.";
                    }
                    else
                    {
                        textK3parameter.Text = textK3parameter.Text + " залишилось незмінним.";
                    }
                }
                lastk3parameter = k3parameter;
                k4parameter = Math.Round(reportCollection.GetReport(i).elem_2_2000 / reportCollection.GetReport(i).elem_1_1900, 4);
                if (reportCollection.GetReport(i).elem_2_2350 != 0)
                {
                    k5parameter = Math.Round(reportCollection.GetReport(i).elem_2_2350 / reportCollection.GetReport(i).elem_2_2000, 4);
                    k6parameter = Math.Round(reportCollection.GetReport(i).elem_2_2350 / reportCollection.GetReport(i).elem_1_1495, 4);
                }
                else
                {
                    k5parameter = Math.Round(reportCollection.GetReport(i).elem_2_2355 / reportCollection.GetReport(i).elem_2_2000, 4);
                    k6parameter = Math.Round(reportCollection.GetReport(i).elem_2_2355 / reportCollection.GetReport(i).elem_1_1495, 4);
                }
                textK5parameter.Text = textK5parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році показник рентабельності реалізації продукції за чистим прибутком становив " + k5parameter.ToString() + ", тобто кожна гривня продажів генерувала " + k5parameter.ToString() + " гривень ";
                if (k5parameter >= 0)
                {
                    textK5parameter.Text = textK5parameter.Text + "чистого прибутку.";
                }
                else
                {
                    textK5parameter.Text = textK5parameter.Text + "збитку.";
                }
                if (i > 0)
                {
                    textK5parameter.Text = textK5parameter.Text + " Порівняно з " + reportCollection.GetReport(i-1).rep_year.ToString() + " роком показник ";
                    if (k5parameter > lastk5parameter)
                    {
                        textK5parameter.Text = textK5parameter.Text + "збільшився на " + (Math.Round((k5parameter-lastk5parameter)/k5parameter, 2)).ToString() + " відсотків.";
                    }
                    else if (k5parameter < lastk5parameter)
                    {
                        textK5parameter.Text = textK5parameter.Text + "зменшився на " + (Math.Round((k5parameter - lastk5parameter) / k5parameter, 2)).ToString() + " відсотків.";
                    }
                    else
                    {
                        textK5parameter.Text = textK5parameter.Text + "залишився незмінним.";
                    }
                }
                lastk5parameter = k5parameter;
                textK6parameter.Text = textK6parameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році показник рентабельності власного капіталу становив " + k6parameter.ToString() + ", тобто на кожну гривню залучених власних коштів було згенеровано " + k6parameter.ToString() + " гривень ";
                if (k6parameter >= 0)
                {
                    textK6parameter.Text = textK6parameter.Text + "чистого прибутку.";
                }
                else
                {
                    textK6parameter.Text = textK6parameter.Text + "збитку.";
                }
                if (i > 0)
                {
                    textK6parameter.Text = textK6parameter.Text + " Порівняно з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком показник ";
                    if (k6parameter > lastk6parameter)
                    {
                        textK6parameter.Text = textK6parameter.Text + "збільшився на " + (Math.Round((k6parameter - lastk6parameter) / k6parameter, 2)).ToString() + " відсотків.";
                    }
                    else if (k6parameter < lastk6parameter)
                    {
                        textK6parameter.Text = textK6parameter.Text + "зменшився на " + (Math.Round((k6parameter - lastk6parameter) / k6parameter, 2)).ToString() + " відсотків.";
                    }
                    else
                    {
                        textK6parameter.Text = textK6parameter.Text + "залишився незмінним.";
                    }
                }
                lastk6parameter = k6parameter;
                bankruptcyProb = Math.Round(2*k2parameter + 0.1*k3parameter + 0.08*k4parameter + 0.45*k5parameter + k6parameter, 4);
                textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " У " + reportCollection.GetReport(i).rep_year.ToString() + " році коефіцієнт Сайфуліна-Кадикова становив " + bankruptcyProb.ToString() + " - ";
                if (bankruptcyProb < 1)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "підприємство знаходилось у незадовільному стані.";
                }
                else
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "підприємство знаходилось у задовільному стані.";
                }
                if (i > 0)
                {
                    textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + "Порівняно з " + reportCollection.GetReport(i - 1).rep_year.ToString() + " роком ймовірність банкрутства ";
                    if (bankruptcyProb > lastBPparameter)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " зменшилась.";
                    }
                    else if (bankruptcyProb < lastBPparameter)
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " збільшилась.";
                    }
                    else
                    {
                        textFinalBankruptcyParameter.Text = textFinalBankruptcyParameter.Text + " не змінилась.";
                    }
                }
                lastBPparameter = bankruptcyProb;
                dataGrid1.Items.Add(new { K1 = reportCollection.GetReport(i).rep_year, K2 = k2parameter, K3 = k3parameter, K4 = k4parameter, K5 = k5parameter, K6 = k6parameter, K7 = bankruptcyProb });
            }
            dataGrid1.HorizontalAlignment = HorizontalAlignment.Center;
            FitToContent(dataGrid1);
            stackPanel1.Children.Add(dataGrid1);
            textDiag = new TextBlock();
            textDiag.FontSize = 18;
            textDiag.FontFamily = new FontFamily("Times New Roman");
            textDiag.TextAlignment = TextAlignment.Justify;
            textDiag.TextWrapping = TextWrapping.Wrap;
            textDiag.Padding = new Thickness(5, 5, 5, 5);
            textDiag.Text = textK2parameter.Text + textK3parameter.Text + textK5parameter.Text + textK6parameter.Text + textFinalBankruptcyParameter.Text;
            stackPanel1.Children.Add(textDiag);
            scrollDiag.Content = stackPanel1;
        }

        private void Close_Diag_Stack_Panel(object sender, RoutedEventArgs e)
        {
            scrollDiag.Content = "";
            DiagPanel.Visibility = Visibility.Collapsed;
        }

        private void Clear_All_on_Click(object sender, RoutedEventArgs e)
        {
            reportCollection.ClearReports();
            scrollDiag.Content = "";
            DiagPanel.Visibility = Visibility.Collapsed;
        }

        private void Close_App_on_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void Open_Instruction_On_Click(object sender, RoutedEventArgs e)
        {
            InstructionWindow instructionWindow = new InstructionWindow();
            instructionWindow.Owner = this;
            instructionWindow.Show();
        }
    }
}
