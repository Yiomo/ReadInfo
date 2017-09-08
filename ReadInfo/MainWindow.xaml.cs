using System;
using System.Collections.Generic;
using System.IO;
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
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace ReadInfo
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void button_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog openfile = new Microsoft.Win32.OpenFileDialog();
            openfile.DefaultExt = ".txt";
            openfile.Filter = "Text documents (.txt)|*.txt";
            bool? result = openfile.ShowDialog();
            if (result == true)
            {
                textBlock.Text = openfile.FileName;
                StreamReader readfile = new StreamReader(openfile .FileName);
                textBlock.Text = readfile.ReadToEnd();
            }
         }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            FolderBrowserDialog openfolder = new FolderBrowserDialog();
            DialogResult result = openfolder.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.Cancel)
            {
                return;
            }
            //textBlock.Text = openfolder.SelectedPath.Trim();
            DirectoryInfo TheFolder = new DirectoryInfo(openfolder.SelectedPath);
            foreach (FileInfo NextFile in TheFolder.GetFiles())
            {
                if (NextFile.Name .Contains(".txt")||NextFile.Name.Contains(".TXT"))
                {
                    StreamReader readfile = new StreamReader(openfolder .SelectedPath+"\\"+NextFile.Name);
                    listBox.Items.Add(NextFile.Name);
                }
            }
            int a = listBox.Items.Count;//统计多少文件

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            app.Visible = false;
            app.UserControl = true;
            Workbook workbook;
            Worksheet worksheet;
            workbook = app.Workbooks.Add(System.Reflection.Missing.Value);
            worksheet = workbook.Worksheets.Add(System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            worksheet.Name = "test";
            for (int i = 0; i < a; i++)//逐行写入
            {
                //textBlock1.Text = "";
                textBlock.Text = a.ToString();
                string[] lines;
                lines = File.ReadAllLines(openfolder.SelectedPath + "\\" + listBox.Items[i].ToString());
                for (int j = 0; j < lines.Length; j++)
                {
                    //textBlock1.Text = textBlock1.Text + lines[j] + "\n";
                    //ws.cells[i + 1, j + 1] = dv[i][j].tostring();
                    worksheet.Cells[i + 1, j + 1] = lines[j];
                }
            }
            workbook.SaveAs(openfolder.SelectedPath+"\\"+"sum.xls", System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
            workbook.Close(System.Reflection.Missing.Value, System.Reflection.Missing.Value, System.Reflection.Missing.Value);
        }
    }
}
                    