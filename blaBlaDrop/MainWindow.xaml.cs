using System;
using System.Collections.Generic;
using System.Linq;
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
using Microsoft.Office.Interop.Excel;
using System.IO;
using System.Windows.Forms;

namespace blaBlaDrop
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow :System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void bGenerateDocs_Click(object sender, RoutedEventArgs e)
        {
            r = new Random();
            string pathToDeviceList = tbPathToDeviceList.Text;
            bool needPDF = (bool)cbIsPDF.IsChecked;

            addToAllAutopsyText = "";
            if (cbConstantSnoring.IsChecked==true)
            {
                addToAllAutopsyText += tbConstantSnoring.Text;
            }

            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            dialog.Description = "Выберите папку для сохранения готовых заключений:";
            System.Windows.Forms.DialogResult result = dialog.ShowDialog();

            specialDeviceTypeList = GetSpecialDeviceTypes();

            string rootPath = AppDomain.CurrentDomain.BaseDirectory;
            string pathToSave = dialog.SelectedPath + "\\";
            var deviceList = GetDevicesToDrop(pathToDeviceList);

            

            while (deviceList.Count > 0)
            {
                if (deviceList.Count == 1)
                {
                    GenerateFullAutopsyReport(deviceList[0], pathToSave, true, needPDF);
                    deviceList.RemoveAt(0);
                }
                else
                {
                    GenerateFullAutopsyReport(deviceList[0], deviceList[1], pathToSave, true, needPDF);
                    deviceList.RemoveAt(0);
                    deviceList.RemoveAt(0);
                }
            }
            System.Windows.Forms.MessageBox.Show("Генерация завершена!");
        }

        public List<string> GetSpecialDeviceTypes()
        {
            specialDeviceReasonsDictionary = new Dictionary<string, string>();
            DirectoryInfo d = new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory + "reasons\\");
            FileInfo[] Files = d.GetFiles("*.txt");
            List<string> specialDeviceType = new List<string>();

            foreach (FileInfo file in Files)
            {
                string tmp = file.Name.Substring(0,file.Name.Length - 4);
                string[] tmpSpecialDeviceTypes = tmp.Split(',');
                foreach(string i in tmpSpecialDeviceTypes)
                {
                    if(i.Length>2)
                    {
                        specialDeviceType.Add(i);
                        specialDeviceReasonsDictionary.Add(i, file.Name);
                    }
                }
            }
            return specialDeviceType;
        }

        public void GenerateFullAutopsyReport(deviceFullDetails device1, string pathToSave, bool needXLSX, bool needPDF)
        {
            string templatePath = AppDomain.CurrentDomain.BaseDirectory + "template.xlsx";
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            Workbook myWorkbook = excelApp.Workbooks.Open(templatePath);

            Sheets myWorksheets = myWorkbook.Worksheets;
            Worksheet myWorksheet = (Worksheet)myWorksheets.get_Item("sheet");

            Range myRange = myWorksheet.get_Range(deviceNamePos[0]);
            myRange.Value = device1.deviceName;

            myRange = myWorksheet.get_Range(deviceYearWasMadePos[0]);
            myRange.Value = device1.deviceYearWasMade;

            myRange = myWorksheet.get_Range(deviceFactoryWasMadePos[0]);
            myRange.Value = device1.deviceFactoryWasMade;

            myRange = myWorksheet.get_Range(deviceIdPos[0]);
            myRange.Value = device1.deviceId;

            myRange = myWorksheet.get_Range(deviceWasUsedSincePos[0]);
            myRange.Value = device1.deviceWasUsedSince;

            myRange = myWorksheet.get_Range(deviceYearWasMadePos[0]);
            myRange.Value = device1.deviceYearWasMade;

            myRange = myWorksheet.get_Range(deviceAmount[0]);
            myRange.Value = device1.deviceAmount;

            myRange = myWorksheet.get_Range(devicePlacePos[0]);
            myRange.Value = device1.devicePlace;

            myRange = myWorksheet.get_Range(deviceBossPos[0]);
            myRange.Value = device1.deviceBoss;

            myRange = myWorksheet.get_Range(deviceAutopsyReportPos[0]);
            myRange.Value = GenerateAutopsyReport(device1);

            //Random r = new Random();
            string futureFileName = device1.deviceName.Substring(0, 5) + r.Next(100).ToString();
            myWorkbook.SaveAs(pathToSave + futureFileName + ".xlsx");
            myWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pathToSave + futureFileName + ".pdf");
            excelApp.Quit();
        }

        public void GenerateFullAutopsyReport(deviceFullDetails device1, deviceFullDetails device2, string pathToSave, bool needXLSX, bool needPDF)
        {
            string templatePath = AppDomain.CurrentDomain.BaseDirectory + "template.xlsx";
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            Workbook myWorkbook = excelApp.Workbooks.Open(templatePath);

            Sheets myWorksheets = myWorkbook.Worksheets;
            Worksheet myWorksheet = (Worksheet)myWorksheets.get_Item("sheet");

            Range myRange = myWorksheet.get_Range(deviceNamePos[0]);
            myRange.Value = device1.deviceName;

            myRange = myWorksheet.get_Range(deviceYearWasMadePos[0]);
            myRange.Value = device1.deviceYearWasMade;

            myRange = myWorksheet.get_Range(deviceFactoryWasMadePos[0]);
            myRange.Value = device1.deviceFactoryWasMade;

            myRange = myWorksheet.get_Range(deviceIdPos[0]);
            myRange.Value = device1.deviceId;

            myRange = myWorksheet.get_Range(deviceWasUsedSincePos[0]);
            myRange.Value = device1.deviceWasUsedSince;

            myRange = myWorksheet.get_Range(deviceYearWasMadePos[0]);
            myRange.Value = device1.deviceYearWasMade;

            myRange = myWorksheet.get_Range(deviceAmount[0]);
            myRange.Value = device1.deviceAmount;

            myRange = myWorksheet.get_Range(devicePlacePos[0]);
            myRange.Value = device1.devicePlace;

            myRange = myWorksheet.get_Range(deviceBossPos[0]);
            myRange.Value = device1.deviceBoss;

            myRange = myWorksheet.get_Range(deviceAutopsyReportPos[0]);
            myRange.Value = GenerateAutopsyReport(device1);

            ////////////////////////////////////////////////////

            myRange = myWorksheet.get_Range(deviceNamePos[1]);
            myRange.Value = device2.deviceName;

            myRange = myWorksheet.get_Range(deviceYearWasMadePos[1]);
            myRange.Value = device2.deviceYearWasMade;

            myRange = myWorksheet.get_Range(deviceFactoryWasMadePos[1]);
            myRange.Value = device2.deviceFactoryWasMade;

            myRange = myWorksheet.get_Range(deviceIdPos[1]);
            myRange.Value = device2.deviceId;

            myRange = myWorksheet.get_Range(deviceWasUsedSincePos[1]);
            myRange.Value = device2.deviceWasUsedSince;

            myRange = myWorksheet.get_Range(deviceYearWasMadePos[1]);
            myRange.Value = device2.deviceYearWasMade;

            myRange = myWorksheet.get_Range(devicePlacePos[1]);
            myRange.Value = device2.devicePlace;

            myRange = myWorksheet.get_Range(deviceAmount[1]);
            myRange.Value = device2.deviceAmount;

            myRange = myWorksheet.get_Range(deviceBossPos[1]);
            myRange.Value = device2.deviceBoss;

            myRange = myWorksheet.get_Range(deviceAutopsyReportPos[1]);
            myRange.Value = GenerateAutopsyReport(device2);

            //Random r = new Random();
            if(device1.deviceName.Length<6)
            {
                device1.deviceName += "      ";
            }
            if (device2.deviceName.Length < 6)
            {
                device2.deviceName += "      ";
            }
            string futureFileName = device1.deviceName.Substring(0, 5) + "_" + device2.deviceName.Substring(0, 5) + r.Next(100).ToString();
            myWorkbook.SaveAs(pathToSave + futureFileName + ".xlsx");
            myWorkbook.ExportAsFixedFormat(XlFixedFormatType.xlTypePDF, pathToSave + futureFileName + ".pdf");
            excelApp.Quit();
        }

        public string GenerateAutopsyReport(deviceFullDetails device)
        {
            string rootPath = AppDomain.CurrentDomain.BaseDirectory;


            string niceDeviceName = "";
            //if(device.deviceName.Length > 25)
            //{
            //    niceDeviceName = "\"" + device.deviceName.Substring(0, 25) + ".\"";
            //}
            //else
            //{
            niceDeviceName = "\"" + device.deviceName + "\"";
            //}

            string[] reasons;
            if (specialDeviceTypeList.Any(device.deviceName.ToLower().Contains))
            {
                reasons = File.ReadAllLines(rootPath + "reasons\\" + specialDeviceReasonsDictionary[specialDeviceTypeList.FirstOrDefault(device.deviceName.ToLower().Contains)]);
            }

            else
            {
                reasons = File.ReadAllLines(rootPath + "reasons\\" + "general.txt", Encoding.ASCII);
                
            }
            string[] morale = File.ReadAllLines(rootPath + "reasons\\" + "moral.txt", Encoding.ASCII);

            int n1 = r.Next(reasons.Length);
            int n2 = r.Next(morale.Length);
            string result = reasons[n1].Replace("$object$", niceDeviceName) + " " + morale[n2] + " " + addToAllAutopsyText;
            return result;
        }

        public List<deviceFullDetails> GetDevicesToDrop(string pathToExcelList)
        {
            Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelApp.Visible = false;

            Workbook myWorkbook = excelApp.Workbooks.Open(pathToExcelList);
            Sheets myWorksheets = myWorkbook.Worksheets;
            Worksheet myWorksheet = (Worksheet)myWorksheets.get_Item("list");

            List<deviceFullDetails> deviceList = new List<deviceFullDetails>();

            for (int i = 2; i < 1000; i++)
            {
                deviceFullDetails tmp = new deviceFullDetails();
                Range myRange = myWorksheet.get_Range("A" + i.ToString());
                if (myRange.Value != null)
                {
                    tmp.deviceName = myRange.Value.ToString();
                }
                else
                {
                    break;
                }

                myRange = myWorksheet.get_Range("B" + i.ToString());
                if (myRange.Value != null)
                {
                    if (myRange.Value.ToString().IndexOf("шт") < 0)
                    {
                        tmp.deviceAmount = myRange.Value.ToString() + " шт.";
                    }
                    else
                    {
                        tmp.deviceAmount = myRange.Value.ToString();
                    }
                }
                else
                {
                    tmp.deviceAmount = "1 шт.";
                }

                myRange = myWorksheet.get_Range("C" + i.ToString());
                if (myRange.Value != null)
                {
                    tmp.deviceYearWasMade = myRange.Value.ToString();
                }
                else
                {
                    tmp.deviceYearWasMade = "";
                }

                myRange = myWorksheet.get_Range("D" + i.ToString());
                if (myRange.Value != null)
                {
                    tmp.deviceFactoryWasMade = myRange.Value.ToString();
                }
                else
                {
                    tmp.deviceFactoryWasMade = "";
                }

                myRange = myWorksheet.get_Range("E" + i.ToString());
                if (myRange.Value != null)
                {
                    tmp.deviceId = myRange.Value.ToString();
                }
                else
                {
                    tmp.deviceId = "";
                }

                myRange = myWorksheet.get_Range("F" + i.ToString());
                if (myRange.Value != null)
                {
                    tmp.deviceWasUsedSince = myRange.Value.ToString();
                }
                else
                {
                    tmp.deviceWasUsedSince = "";
                }

                myRange = myWorksheet.get_Range("G" + i.ToString());
                if (myRange.Value != null)
                {
                    tmp.devicePlace = myRange.Value.ToString();
                }
                else
                {
                    tmp.devicePlace = "";
                }

                myRange = myWorksheet.get_Range("H" + i.ToString());
                if (myRange.Value != null)
                {
                    tmp.deviceBoss = myRange.Value.ToString();
                }
                else
                {
                    tmp.deviceBoss = "";
                }

                deviceList.Add(tmp);
            }
            myWorkbook.Close();
            excelApp.Quit();
            return deviceList;
        }

        List<string> specialDeviceTypeList;
        Dictionary<string, string> specialDeviceReasonsDictionary;

        string addToAllAutopsyText;
        Random r;

        string[] deviceNamePos = new string[] { "D6", "D30" };
        string[] deviceAmount = new string[] { "B7", "B31" };
        string[] deviceYearWasMadePos = new string[] { "B8", "B32" };
        string[] deviceFactoryWasMadePos = new string[] { "F7", "F31" };
        string[] deviceIdPos = new string[] { "C9", "C33" };
        string[] deviceWasUsedSincePos = new string[] { "G8", "G32" };
        string[] devicePlacePos = new string[] { "G9", "G33" };
        string[] deviceBossPos = new string[] { "E10", "E34" };
        string[] deviceAutopsyReportPos = new string[] { "A13", "A37" };

        public class deviceFullDetails
        {
            public string deviceName;
            public string deviceAmount;
            public string deviceYearWasMade;
            public string deviceFactoryWasMade;
            public string deviceId;
            public string deviceWasUsedSince;
            public string devicePlace;
            public string deviceBoss;
            public string deviceAutopsyReport;
        }

        private void bLoadDeviceList_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "Excel files (*.xlsx)|*.xlsx";
            openFileDialog1.FilterIndex = 2;
            openFileDialog1.RestoreDirectory = true;

            openFileDialog1.ShowDialog();

            tbPathToDeviceList.Text = openFileDialog1.FileName;
        }

        private void cbConstantSnoring_Checked(object sender, RoutedEventArgs e)
        {
            tbConstantSnoring.IsEnabled = true;
        }

        private void cbConstantSnoring_Unchecked(object sender, RoutedEventArgs e)
        {
            tbConstantSnoring.IsEnabled = false;
        }
    }
}
