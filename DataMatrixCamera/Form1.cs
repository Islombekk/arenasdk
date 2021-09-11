using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using ClosedXML.Excel;
using System.IO;
using System.Reflection;
using Vintasoft.Barcode;
using Vintasoft.Barcode.QualityTests;
using Vintasoft.Barcode.BarcodeInfo;
using DataMatrix;

namespace DataMatrixCamera
{
    public partial class Form1 : Form
    {
        public class TreeItem
        {
            public string Title { get; set; } // string to display
            public string UId { get; set; } = ""; // serial number to identify device
        }

        private ArenaNET.ISystem m_system = null;
        private List<ArenaNET.IDeviceInfo> m_deviceInfos;
        private ArenaNET.IDevice m_connectedDevice = null;
        private ArenaNET.IImage m_converted;
        List<string> deviceUIds;
        ArenaNET.IImage image = null;
        List<DataMatrixQuality> dataMatrixQuality = new List<DataMatrixQuality>();

        public Form1()
        {
            InitializeComponent();
        }

        private void checkLicence()
        {

            Vintasoft.Barcode.BarcodeGlobalSettings.Register("Saken Alimkulov", "sakenalimkulov.dev@gmail.com", "2021-08-30", "oh1+VSDfAwVWQk+EYVCCQkwOXlWnOm4blCrVBGSN0lJ27BlUgkl0ePnlGgxc+VU7b90YYlLZbpAzFHKMlI3fVYInfT/14N7LsEgnfYXuRCTFy3aRMMdy44fKhqwMBNwbFoPqS0xzTawKSrju6Vp63rM5AJ6YoKqAvljEQf9Gmp0");
        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void подключитьсяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                deviceUIds = GetDeviceUIds();//ToList
                //string connectedDeviceUId = ConnectedDeviceUId(); // String
                ConnectDevice(deviceUIds[0]);
                StartStream();
                pictureBox1.Image = GetImage();
                обУстройствеToolStripMenuItem.Enabled = true;
                отключитьсяToolStripMenuItem.Enabled = true;

            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка подключения к устройству : " + ex.Message);
            }
        }


        public void ConnectDevice(String UId)
        {
            if (m_connectedDevice == null)
            {
                UpdateDevices();

                for (int i = 0; i < m_deviceInfos.Count; i++)
                {
                    if (m_deviceInfos[i].SerialNumber == UId)
                    {
                        m_connectedDevice = m_system.CreateDevice(m_deviceInfos[i]);
                        return;
                    }
                }
            }

            throw new Exception();
        }

        public void DisconnectDevice()
        {
            m_system.DestroyDevice(m_connectedDevice);
            m_connectedDevice = null;
        }

        // returns if connected to a device
        public bool DeviceConnected()
        {
            return (m_connectedDevice != null);
        }

        // gets uid of connected device
        public String ConnectedDeviceUId()
        {
            return ((ArenaNET.IString)m_connectedDevice.NodeMap.GetNode("DeviceSerialNumber")).Value;
        }


        // gets list of uids of available devices
        public List<String> GetDeviceUIds()
        {
            UpdateDevices();

            List<String> uids = new List<String>();

            for (int i = 0; i < m_deviceInfos.Count; i++)
            {
                uids.Add(m_deviceInfos[i].SerialNumber);
            }

            return uids;
        }

        // gets name of device
        public String GetDeviceName(String UId, String node)
        {
            for (int i = 0; i < m_deviceInfos.Count; i++)
            {
                if (m_deviceInfos[i].SerialNumber == UId)
                {
                    if (node == "DeviceUserID" || node == "UserDefinedName")
                        return m_deviceInfos[i].UserDefinedName;
                    else if (node == "DeviceModelName" || node == "ModelName")
                        return m_deviceInfos[i].ModelName;
                }
            }

            return "Invalid argument";
        }

        // check if ip addresses between device and interface match based on their subnets
        public bool IsNetworkValid(String UId)
        {
            UpdateDevices();

            for (int i = 0; i < m_deviceInfos.Count; i++)
            {
                if (m_deviceInfos[i].SerialNumber == UId)
                {
                    UInt32 ip = (UInt32)m_deviceInfos[i].IpAddress;
                    UInt32 subnet = (UInt32)m_deviceInfos[i].SubnetMask;

                    ArenaNET.IInteger ifipNode = (ArenaNET.IInteger)m_system
                            .GetTLInterfaceNodeMap(m_deviceInfos[i]).GetNode("GevInterfaceSubnetIPAddress");
                    ArenaNET.IInteger ifsubnetNode = (ArenaNET.IInteger)m_system
                            .GetTLInterfaceNodeMap(m_deviceInfos[i]).GetNode("GevInterfaceSubnetMask");
                    UInt32 ifip = (UInt32)ifipNode.Value;
                    UInt32 ifsubnet = (UInt32)ifsubnetNode.Value;

                    if (subnet != ifsubnet)
                        return false;

                    if ((ip & subnet) != (ifip & ifsubnet))
                        return false;

                    return true;
                }
            }

            throw new Exception();
        }

        // gets image as bitmap
        public System.Drawing.Bitmap GetImage(UInt32 timeout = 2000)
        {
            if (m_converted != null)
            {
                ArenaNET.ImageFactory.Destroy(m_converted);
                m_converted = null;
            }

            ArenaNET.IImage image = m_connectedDevice.GetImage(timeout);
            m_converted = ArenaNET.ImageFactory.Convert(image, (ArenaNET.EPfncFormat)0x02200017);
            m_connectedDevice.RequeueBuffer(image);


            ReadBarcodesAndTestBarcodePrintQuality(m_converted.Bitmap);

            return m_converted.Bitmap;
        }

        private void ReadBarcodesAndTestBarcodePrintQuality(System.Drawing.Image imageWithBarcodes)
        {
            checkLicence();
            using (BarcodeReader reader = new BarcodeReader())
            {
                reader.Settings.CollectTestInformation = true;

                reader.Settings.ScanBarcodeTypes = BarcodeType.Aztec | BarcodeType.DataMatrix | BarcodeType.QR | BarcodeType.MicroQR | BarcodeType.HanXinCode;


                IBarcodeInfo[] barcodeInfos = reader.ReadBarcodes(imageWithBarcodes);

                ISO15415QualityTest test = new ISO15415QualityTest();
                for (int i = 0; i < barcodeInfos.Length; i++)
                {
                    test.CalculateGrades((BarcodeInfo2D)barcodeInfos[i], imageWithBarcodes);
                    var additionalGradesKeys = "";
                    var quietZone = "";
                    foreach (string name in test.AdditionalGrades.Keys)
                    {
                        additionalGradesKeys = name.PadRight(40, ' ') + " " + GradeToString(test.AdditionalGrades[name]);
                    }

                    if (test.QuietZone >= 0)
                    {
                        quietZone = string.Format("{0} ({1} %)", GradeToString(test.QuietZoneGrade), test.QuietZone);
                    }

                    DataMatrixQuality dataMatrix = new DataMatrixQuality(
                         "",
                        (GradeToString(test.DecodeGrade)),
                        (GradeToString(test.UnusedErrorCorrectionGrade) + " " + test.UnusedErrorCorrection + "%"),
                        (GradeToString(test.SymbolContrastGrade) + " " + test.SymbolContrast),
                        (GradeToString(test.AxialNonuniformityGrade) + " " + test.AxialNonuniformity),
                        (GradeToString(test.GridNonuniformityGrade) + " " + test.GridNonuniformity + " cell"),
                        (GradeToString(test.ModulationGrade)),
                        (GradeToString(test.ReflectanceMarginGrade)),
                        (GradeToString(test.FixedPatternDamageGrade)),
                        additionalGradesKeys,
                        quietZone,
                        (test.DistortionAngle + "°"),
                        (GradeToString(test.ScanGrade))
                        );

                    dataMatrixQuality.Add(dataMatrix);

                }
            }
        }

        private string GradeToString(ISO15415QualityGrade grade)
        {
            return string.Format("{0}({1})", ((int)grade).ToString(), grade);
        }

        // gets image as ArenaNET.IImage
        public ArenaNET.IImage GetIImage(UInt32 timeout = 2000)
        {
            if (m_converted != null)
            {
                ArenaNET.ImageFactory.Destroy(m_converted);
                m_converted = null;
            }

            if (image != null)
            {
                m_connectedDevice.RequeueBuffer(image);
                image = null;
            }

            image = m_connectedDevice.GetImage(timeout);
            return image;
        }

        // starts stream
        public void StartStream()
        {
            if (m_connectedDevice != null)
            {
                if ((m_connectedDevice.TLStreamNodeMap.GetNode("StreamIsGrabbing") as ArenaNET.IBoolean).Value == false)
                {
                    (m_connectedDevice.TLStreamNodeMap.GetNode("StreamBufferHandlingMode") as ArenaNET.IEnumeration).Symbolic = "NewestOnly";
                    m_connectedDevice.StartStream();
                }
            }
        }

        // stops stream
        public void StopStream()
        {
            if (m_connectedDevice != null)
            {
                if ((m_connectedDevice.TLStreamNodeMap.GetNode("StreamIsGrabbing") as ArenaNET.IBoolean).Value == true)
                {
                    m_connectedDevice.StopStream();
                }
            }
        }

        // updates list of available devices
        private void UpdateDevices()
        {
            m_system.UpdateDevices(100);
            m_deviceInfos = m_system.Devices;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            обУстройствеToolStripMenuItem.Enabled = false;
            отключитьсяToolStripMenuItem.Enabled = false;
            m_system = ArenaNET.Arena.OpenSystem();
        }

        private void выйтиToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void обУстройствеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                MessageBox.Show("Подключённое устройство: " + ConnectedDeviceUId() + "\n" +
                                "Модель:" + m_deviceInfos[0].ModelName + "\n" +
                                "IP адрес устройства:" + m_deviceInfos[0].IpAddress + "\n" +
                                "Версия устройства:" + m_deviceInfos[0].DeviceVersion + "\n" +
                                "Серийный номер устройства:" + m_deviceInfos[0].SerialNumber + "\n" +
                                "MAC-адрес устройства:" + m_deviceInfos[0].MacAddress
                                );
            }
            catch (Exception)
            {
                MessageBox.Show("Камера не подключена или отсутствует!");
            }

        }

        private void отключитьсяToolStripMenuItem_Click(object sender, EventArgs e)
        {
            StopStream();
            DisconnectDevice();
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }
       
        public Boolean writeToExcel()
        {
            var appDir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);

            try
            {
                for (int k = 0; k < dataMatrixQuality.Count; k++)
                {
                    MessageBox.Show(
                   ("Decode:" + dataMatrixQuality[k].Decode + "\n" +
                   "Unused error correction:" + dataMatrixQuality[k].UnusedErrorCorrection + "\n" +
                   "Symbol contrast:" + dataMatrixQuality[k].SymbolContrast + "\n" +
                   "Axial nonuniformity:" + dataMatrixQuality[k].AxialNonuniformity + "\n" +
                   "Grid nonuniformity:" + dataMatrixQuality[k].GridNonuniformity + "\n" +
                   "Modulation:" + dataMatrixQuality[k].Modulation + "\n" +
                   "Reflectance margin:" + dataMatrixQuality[k].ReflectanceMargin + "\n" +
                   "Fixed pattern damage:" + dataMatrixQuality[k].FixedPatternDamage + "\n" +
                   "Additional Grades Keys:" + dataMatrixQuality[k].AdditionalGradesKeys + "\n" +
                   "Quiet zone:" + dataMatrixQuality[k].QuietZone + "\n" +
                   "Distortion angle (informative):" + dataMatrixQuality[k].DistortionAngle + "\n" +
                   "Scan grade:" + dataMatrixQuality[k].ScanGrade
                   )
                    );
                }


                using (var wbook = new XLWorkbook())
                {
                    var ws = wbook.AddWorksheet("Analize result");

                    ws.Cell(1, 1).Value = "File name";
                    ws.Cell(1, 2).Value = "Decode";
                    ws.Cell(1, 3).Value = "Unused error correction";
                    ws.Cell(1, 4).Value = "Symbol contrast";
                    ws.Cell(1, 5).Value = "Axial nonuniformity";
                    ws.Cell(1, 6).Value = "Grid nonuniformity";
                    ws.Cell(1, 7).Value = "Modulation";
                    ws.Cell(1, 8).Value = "Reflectance margin";
                    ws.Cell(1, 9).Value = "Fixed pattern damage";
                    ws.Cell(1, 10).Value = "Additional Grades Keys";
                    ws.Cell(1, 11).Value = "Quiet zone";
                    ws.Cell(1, 12).Value = "Distortion angle (informative)";
                    ws.Cell(1, 13).Value = "Scan grade";

                    for (int k = 0; k < dataMatrixQuality.Count; k++)
                    {
                        for (int l = 0; l < 13; l++)
                        {
                            ws.Cell(k + 2, 1).Value = "";
                            ws.Cell(k + 2, 2).Value= dataMatrixQuality[k].Decode;
                            ws.Cell(k + 2, 3).Value= dataMatrixQuality[k].UnusedErrorCorrection;
                            ws.Cell(k + 2, 4).Value= dataMatrixQuality[k].SymbolContrast;
                            ws.Cell(k + 2, 5).Value= dataMatrixQuality[k].AxialNonuniformity;
                            ws.Cell(k + 2, 6).Value= dataMatrixQuality[k].GridNonuniformity;
                            ws.Cell(k + 2, 7).Value= dataMatrixQuality[k].Modulation;
                            ws.Cell(k + 2, 8).Value= dataMatrixQuality[k].ReflectanceMargin;
                            ws.Cell(k + 2, 9).Value= dataMatrixQuality[k].FixedPatternDamage;
                            ws.Cell(k + 2, 10).Value = dataMatrixQuality[k].AdditionalGradesKeys;
                            ws.Cell(k + 2, 11).Value = dataMatrixQuality[k].QuietZone;
                            ws.Cell(k + 2, 12).Value = dataMatrixQuality[k].DistortionAngle;
                            ws.Cell(k + 2, 13).Value = dataMatrixQuality[k].ScanGrade;
                        }

                    }

                    string filename = "datamatrix-" + DateTime.Now.Day.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + "-" + DateTime.Now.Second.ToString() + ".xlsx";
                    wbook.SaveAs(Path.Combine(appDir, "exported/ " + filename));

                }



                /*Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
                Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                xlWorkSheet.Cells[1, 1] = "File name";
                xlWorkSheet.Cells[1, 1].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 2] = "Decode";
                xlWorkSheet.Cells[1, 2].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 3] = "Unused error correction";
                xlWorkSheet.Cells[1, 3].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 4] = "Symbol contrast";
                xlWorkSheet.Cells[1, 4].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 5] = "Axial nonuniformity";
                xlWorkSheet.Cells[1, 5].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 6] = "Grid nonuniformity";
                xlWorkSheet.Cells[1, 6].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 7] = "Modulation";
                xlWorkSheet.Cells[1, 7].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 8] = "Reflectance margin";
                xlWorkSheet.Cells[1, 8].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 9] = "Fixed pattern damage";
                xlWorkSheet.Cells[1, 9].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 10] = "Additional Grades Keys";
                xlWorkSheet.Cells[1, 10].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 11] = "Quiet zone";
                xlWorkSheet.Cells[1, 11].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 12] = "Distortion angle (informative)";
                xlWorkSheet.Cells[1, 12].EntireRow.Font.Bold = true;
                xlWorkSheet.Cells[1, 13] = "Scan grade";
                xlWorkSheet.Cells[1, 13].EntireRow.Font.Bold = true;


                for (int k = 0; k < dataMatrixQuality.Count; k++)
                {
                    for (int l = 0; l < 13; l++)
                    {
                        xlWorkSheet.Cells[k + 2, 1] = fileNames[k];
                        xlWorkSheet.Cells[k + 2, 2] = dataMatrixQuality[k].Decode;
                        xlWorkSheet.Cells[k + 2, 3] = dataMatrixQuality[k].UnusedErrorCorrection;
                        xlWorkSheet.Cells[k + 2, 4] = dataMatrixQuality[k].SymbolContrast;
                        xlWorkSheet.Cells[k + 2, 5] = dataMatrixQuality[k].AxialNonuniformity;
                        xlWorkSheet.Cells[k + 2, 6] = dataMatrixQuality[k].GridNonuniformity;
                        xlWorkSheet.Cells[k + 2, 7] = dataMatrixQuality[k].Modulation;
                        xlWorkSheet.Cells[k + 2, 8] = dataMatrixQuality[k].ReflectanceMargin;
                        xlWorkSheet.Cells[k + 2, 9] = dataMatrixQuality[k].FixedPatternDamage;
                        xlWorkSheet.Cells[k + 2, 10] = dataMatrixQuality[k].AdditionalGradesKeys;
                        xlWorkSheet.Cells[k + 2, 11] = dataMatrixQuality[k].QuietZone;
                        xlWorkSheet.Cells[k + 2, 12] = dataMatrixQuality[k].DistortionAngle;
                        xlWorkSheet.Cells[k + 2, 13] = dataMatrixQuality[k].ScanGrade;
                        xlWorkSheet.Columns.AutoFit();
                    }

                }


                dataMatrixQuality = new List<DataMatrixQuality>();
                string filename = "datamatrix-" + DateTime.Now.Day.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Year.ToString() + "-" + DateTime.Now.Hour.ToString() + "-" + DateTime.Now.Minute.ToString() + "-" + DateTime.Now.Second.ToString() + ".xls";
                xlWorkBook.SaveAs(Path.Combine(appDir, "exported/ " + filename), Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();*/
                // pictureBox1.SizeMode = PictureBoxSizeMode.StretchImage;
                // pictureBox1.Load(Path.Combine(appDir, "asset/success.png"));
                return true;
            }
            catch (Exception e)
            {
                // pictureBox1.Load(Path.Combine(appDir, "asset/fail.png"));
                // MessageBox.Show(e.Message);
                return false;
            }
        }
    }
}
