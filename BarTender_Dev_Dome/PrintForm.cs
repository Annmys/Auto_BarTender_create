using Seagull.BarTender.Print;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using UnityEngine;
using Application = System.Windows.Forms.Application;
using Color = System.Drawing.Color;
using Resolution = Seagull.BarTender.Print.Resolution;
using OfficeOpenXml; // EPPlus的命名空间
using System.Runtime.InteropServices;
using System.Threading.Tasks;  // 包含这个指令以使用 Task 和 async/await
using MyLibrary; // 引入命名空间
using maitou;
using OfficeOpenXml.FormulaParsing.ExpressionGraph.FunctionCompilers;//唛头.cs






namespace BarTender_Dev_Dome

{
    public partial class PrintForm : Form
    {
        string bq_ipdj = null;
        private string _bmp_path = Application.StartupPath + @"\exp.jpg";
        private string _btw_path = "";
        string 模板地址 = "";
        //string 软件版本 = "";
        string 灯带系列 = "";
        string _sjk_path = "";
        string _PrinterName = "";
        string _wjm_ = "";
        string output_name = "Name:LED Flex Linear Light";
        string output_灯带型号 = "";
        string output_电压 = "";
        string output_功率 = "";
        string output_灯数 = "";
        string output_剪切单元 = "";
        string output_长度 = "Length:";
        string output_色温 = "";
        string output_尾巴 = "";



        //程序开始运行
        public PrintForm()
        {
            InitializeComponent();
            //MessageBox.Show("程序打开", "操作提示");
            //this.Load += new EventHandler(MainForm_Load); // 订阅窗体加载事件-读取配置文件

            //读取excel文件必须增加声明才能运行正常
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;



        }

        //打印被点击
        private void print_btn_Click(object sender, EventArgs e)
        {
            PrintBar();
        }

        //预览被点击
        private void preview_btn_Click(object sender, EventArgs e)
        {
            PrintBar(true);
        }



        //筛选标签规格类型
        private void button_筛选_Click_1(object sender, EventArgs e)
        {
            if (checkBox_常规型号.Checked)
            {
                // 如果选中，设置 comboBox_标签规格 为可见
                comboBox_标签规格.Visible = true;

                // 重置 comboBox_标签规格
                comboBox_标签规格.Items.Clear();

                // 重新加载标签规格
                获取标签规格("常规型号");
            }
            else if(checkBox_客制型号.Checked)
            {
                // 如果选中，设置 comboBox_客户编号 为可见
                comboBox_标签规格.Visible = true;

                // 重置 comboBox_标签规格
                comboBox_标签规格.Items.Clear();

                // 重新加载标签规格
                获取标签规格("客制型号");
            }
            else
            {
                // 如果未选中，设置 comboBox_客户编号 为不可见
                comboBox_标签规格.Visible = false;
            }


            

            // 执行筛选逻辑
            //FilterSpecifications();
        }



        // 执行筛选逻辑
        private void FilterSpecifications()
        {
            // 创建一个集合来存储需要保留的标签规格
            HashSet<string> specsToKeep = new HashSet<string>();

            // 根据复选框的状态添加对应的规格到集合中
            if (checkBox_常规型号.Checked)
            {
                specsToKeep.Add("常规型号");
            }
            if (checkBox_客制型号.Checked)
            {
                specsToKeep.Add("客制型号");
            }
            

            // ... 为其他复选框添加相应的条件

            // 如果没有任何复选框被选中，则不执行任何操作
            if (specsToKeep.Count == 0)
            {
                return; // 退出方法
            }

            // 遍历 comboBox_标签规格 的项
            for (int i = comboBox_标签规格.Items.Count - 1; i >= 0; i--)
            {
                string item = comboBox_标签规格.Items[i].ToString();
                // 检查当前项是否包含所有需要保留的标签规格
                bool keepItem = true; // 假设当前项需要保留
                foreach (string spec in specsToKeep)
                {
                    if (!item.Contains(spec))
                    {
                        keepItem = false; // 如果当前项缺少某个规格，则不需要保留
                        break; // 退出当前循环，检查下一项
                    }
                }

                // 如果当前项不需要保留，则删除它
                if (!keepItem)
                {
                    comboBox_标签规格.Items.RemoveAt(i);
                }
            }
        }




        private void 获取标签种类()
        {
            // 软件运行目录
            string baseDirectory = "\\\\192.168.1.33\\Annmy\\订单标签自动生成软件";

            // 指定moban文件夹路径
            string templatesDirectory = Path.Combine(baseDirectory, "moban");

            // 清空ComboBox中的现有项
            标签种类_comboBox.Items.Clear();

            // 判断moban文件夹是否存在
            if (Directory.Exists(templatesDirectory))
            {
                // 获取moban文件夹内的所有文件夹
                string[] directories = Directory.GetDirectories(templatesDirectory);

                // 遍历所有文件夹
                foreach (string directory in directories)
                {
                    // 获取文件夹名称并添加到ComboBox中
                    string folderName = Path.GetFileName(directory);
                    标签种类_comboBox.Items.Add(folderName);
                }
            }
            else
            {
                MessageBox.Show("moban文件夹不存在。");
            }
        }

        private void 获取标签规格(string aa)
        {
            string templatesDirectory;
            string _btw_path_1;
            // 软件运行目录
            string baseDirectory = "\\\\192.168.1.33\\Annmy\\订单标签自动生成软件\\moban";

            // 拼接标签种类_comboBox.Text 到基础目录
            if (aa == "常规型号")
            {
                templatesDirectory = Path.Combine(baseDirectory, @"\常规型号" + @"\" );
                
                _btw_path_1 =  @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\" + aa + @"\";
                // 清空ComboBox中的现有项
                comboBox_标签规格.Items.Clear();

                // 判断moban文件夹是否存在
                if (Directory.Exists(_btw_path_1))
                {
                    // 获取moban文件夹内的所有文件夹
                    string[] directories = Directory.GetDirectories(_btw_path_1);


                    // 遍历所有文件夹
                    foreach (string directory in directories)
                    {
                        // 获取文件夹名称并添加到ComboBox中
                        string folderName = Path.GetFileName(directory);
                        comboBox_标签规格.Items.Add(folderName);
                    }
                }
                else
                {
                    // 如果文件夹不存在，弹出消息框提示用户
                    MessageBox.Show("文件夹不存在。");
                }
            }
            if(aa =="客制型号")
            {
                templatesDirectory = Path.Combine(baseDirectory, @"\客制型号" );
                _btw_path_1 =  @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\" + aa + @"\";
                // 清空ComboBox中的现有项
                comboBox_标签规格.Items.Clear();

                // 判断moban文件夹是否存在
                if (Directory.Exists(_btw_path_1))
                {
                    // 获取moban文件夹内的所有文件夹
                    string[] directories = Directory.GetDirectories(_btw_path_1);


                    // 遍历所有文件夹
                    foreach (string directory in directories)
                    {
                        // 获取文件夹名称并添加到ComboBox中
                        string folderName = Path.GetFileName(directory);
                        comboBox_标签规格.Items.Add(folderName);
                    }
                }
                else
                {
                    // 如果文件夹不存在，弹出消息框提示用户
                    MessageBox.Show("文件夹不存在。");
                }
            }
            

            
        }

        //软件开启加载标签种类
        private void PrintForm_Load(object sender, EventArgs e)
        {

            获取标签种类();


            Printers printers = new Printers();
            foreach (Printer printer in printers)
            {
                printer_comboBox.Items.Add(printer.PrinterName);
                
            }

            if (printers.Count > 0)
            {
                // Automatically select the default printer.
                printer_comboBox.SelectedItem = printers.Default.PrinterName;
            }
        }


        private void printer_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            _PrinterName = printer_comboBox.Text;
        }

        //标签种类被选择时
        private void 标签种类_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            //获取标签规格("常规型号");

              //string currentDirectory = Directory.GetCurrentDirectory();
              _btw_path =  @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\";
              fileNametBox.Text = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\";


            // 检查 comboBox 是否有选中项
            if (标签种类_comboBox.SelectedIndex != -1)
            {
                // 如果有选择项，则显示 groupBox4
                groupBox4.Visible = true;
            }
            else
            {
                // 如果没有选择项，则隐藏 groupBox4
                groupBox4.Visible = false;
            }




        }


        private string 重构产品信息_工字标(string name_CPXXBox, string textBox_剪切长度)
        {
            // 检查 name_CPXXBox 是否包含特定的标识 ")\nLength:"，并找到它的位置
            int lengthIndex = name_CPXXBox.IndexOf(")\nLength:");
            if (lengthIndex != -1)
            {
                // 调整 lengthIndex 以确保它位于 ")\nLength:" 后面
                lengthIndex += ")\nLength:".Length;
            }
            else
            {
                // 如果没有找到 ")\nLength:"，可以决定如何处理这种情况
                // 例如，可以返回原始的 name_CPXXBox 或者返回一个错误消息
                return "指定的标识 ')\nLength:' 不存在。";
            }

            // 使用 StringBuilder 创建新的文本
            StringBuilder newText = new StringBuilder(name_CPXXBox);

            // 移除从找到的位置到字符串末尾的所有内容
            newText.Remove(lengthIndex, newText.Length - lengthIndex);

            // 在找到的位置后面插入 textBox_剪切长度 的内容
            newText.Insert(lengthIndex, textBox_剪切长度);

            // 在文本末尾追加其他信息
            newText.Append("\n" + output_色温);
            newText.Append("\n" + output_尾巴);

            // 返回构建好的字符串
            return newText.ToString();
        }



            //预览和打印
            void PrintBar(bool isPreView = false)
        {
            // 假设这是从某个文本框获取的字符串
            string cpxx_text =cpxxBox.Text;
            判断产品信息(cpxx_text);

            //执行工字标和品名标
            if (标签种类_comboBox.Text == "工字标" || 标签种类_comboBox.Text == "品名标")
            {
                if (_btw_path.Length < 5)
                {
                    fileNametBox.BackColor = Color.Red;
                    return;
                }

                using (Engine btEngine = new Engine(true))
                {

                    if (!_btw_path.Contains(comboBox_标签规格.Text))
                    {
                       // _btw_path = _btw_path + comboBox_标签规格.Text; // 如果不包含，则拼接

                        if (checkBox_常规型号.Checked  )
                        {
                            _btw_path = _btw_path +@"常规型号\"+ comboBox_标签规格.Text ;
                        }
                        else if (checkBox_客制型号.Checked  )
                        {
                            _btw_path = _btw_path +@"客制型号\"+ comboBox_标签规格.Text ;
                        }


                    }

                    //MessageBox.Show(_btw_path);   

                    //寻找文件名_单字匹配("正弯", "侧弯");
                    string 复选框 = 判断复选框内容(cpxxBox.Text, comboBox_标签规格.Text);
                    //MessageBox.Show(复选框);

                    //读取excel文件必须增加声明才能运行正常
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                    // 检查文本框内容是否不为空
                    if (!string.IsNullOrEmpty(Box_数据库.Text))
                    {
                        // 读取Excel文件的B、C、D、E、F列内容
                        string filePath = Box_数据库.Text; // 假设这是Excel文件路径的文本框
                        Dictionary<int, bool> columnHasContent = new Dictionary<int, bool>();

                        using (var package = new ExcelPackage(new FileInfo(filePath)))
                        {
                            var worksheet = package.Workbook.Worksheets[0];
                            for (int col = 2; col <= 6; col++) // B到F列的索引是2到6
                            {
                                var cell = worksheet.Cells[2, col]; // 假设从第二行开始读取
                                columnHasContent[col] = !string.IsNullOrEmpty(cell.Text?.Trim());
                            }
                        }

                        // 根据列内容设置 _wjm_
                        if (columnHasContent[2] && columnHasContent[3] && columnHasContent[4] && columnHasContent[5] && columnHasContent[6])
                        {
                            _wjm_ = "5.btw";
                        }
                        else if (columnHasContent[2] && columnHasContent[3] && columnHasContent[4] && columnHasContent[5] && !columnHasContent[6])
                        {
                            _wjm_ = "4.btw";
                        }
                        else if (columnHasContent[2] && columnHasContent[3] && columnHasContent[4] && !columnHasContent[5] && !columnHasContent[6])
                        {
                            _wjm_ = "3.btw";
                        }
                        else if (columnHasContent[2] && (!columnHasContent[3] || (columnHasContent[3] && !columnHasContent[4])) && !columnHasContent[5] && !columnHasContent[6])
                        {
                            _wjm_ = "1.btw";
                        }
                        else
                        {
                            // 如果不符合上述任一条件，可以设置默认值或者抛出异常
                            _wjm_ = "1.btw"; // 默认值
                                             // 或者 throw new Exception("Excel文件的列内容不符合要求");
                        }
                    }
                    else
                    {
                        _wjm_ = "1.btw";
                    }



                    模板地址 = _btw_path + @"\" + _wjm_;
                    //MessageBox.Show(模板地址, "操作提示");
                    模板地址 = 模板地址.Replace("\n", "").Replace("\r", "");  //去除换行符，否则下面会报错
                    //MessageBox.Show(模板地址);


                    if (_wjm_.Length > 2)
                    {
                        LabelFormatDocument labelFormat = btEngine.Documents.Open(模板地址);

                        try
                        {
                            //构建的新文本赋值给 name_CPXXBox 的 Text 属性

                            //int lengthIndex = name_CPXXBox.Text.IndexOf(")\nLength:") + ")\nLength:".Length;
                            //StringBuilder newText = new StringBuilder(name_CPXXBox.Text);
                            //newText.Remove(lengthIndex, newText.Length - lengthIndex);
                            //newText.Insert(lengthIndex, textBox_剪切长度.Text);
                            //newText.Append("\n"+output_色温);
                            //newText.Append("\n"+ output_尾巴 );
                            //MessageBox.Show(newText.ToString());
                            //name_CPXXBox.Text = newText.ToString();

                            // 调用方法时，可以这样使用返回的字符串
                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, textBox_剪切长度.Text);
                            //MessageBox.Show(name_CPXXBox.Text);




                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                            //labelFormat.SubStrings.SetSubString("CPCD", textBox_剪切长度.Text);
                            labelFormat.SubStrings.SetSubString("CPCD", " ");
                            labelFormat.SubStrings.SetSubString("FXK", 复选框);
                            labelFormat.SubStrings.SetSubString("XLH", " ");

                            if (comboBox_标签规格.Text.Contains("水下"))
                            {
                                labelFormat.SubStrings.SetSubString("IPDJ", "IP68 5m");
                            }
                            else
                            {
                                labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj);
                            }

                            //高压情况
                            if (comboBox_标签规格.Text.Contains("高压")) { double.TryParse(textBox_剪切长度.Text, out double length); double result = length * 3.28;labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                                double washu = length * 10;double anpai = length * 0.093;labelFormat.SubStrings.SetSubString("CPXX-2", washu .ToString()+"W,"+anpai.ToString()+"A");
                            }
                            

                            //写入2排标识码时候的内容


                            //判断是否增加标识码
                            if (checkBox_标识码01.Checked)
                            {
                                labelFormat.SubStrings.SetSubString("BSM-01", textBox_标识码01.Text);
                            }
                            else
                            {
                                labelFormat.SubStrings.SetSubString("BSM-01", " ");
                            }
                            if (checkBox_标识码02.Checked)
                            {
                                labelFormat.SubStrings.SetSubString("BSM-02", textBox_标识码02.Text);
                            }
                            else
                            {
                                labelFormat.SubStrings.SetSubString("BSM-02", " ");
                            }

                            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                            // 检查数据库地址不为空时
                            if (!string.IsNullOrEmpty(Box_数据库.Text))
                            {

                                string b2Data, c2Data, g2Data, h2Data, a2Data, d2Data, e2Data, f2Data, i2Data;

                                // 使用EPPlus打开Excel文件
                                using (var package = new ExcelPackage(new FileInfo(Box_数据库.Text)))
                                {
                                    // 假设Excel工作表名为"Sheet1"
                                    var worksheet = package.Workbook.Worksheets["Sheet1"];

                                    // 读取B2和C2单元格的数据
                                    a2Data = worksheet.Cells["A2"].Value?.ToString() ?? string.Empty;
                                    b2Data = worksheet.Cells["B2"].Value?.ToString() ?? string.Empty;
                                    c2Data = worksheet.Cells["C2"].Value?.ToString() ?? string.Empty;
                                    d2Data = worksheet.Cells["D2"].Value?.ToString() ?? string.Empty;
                                    e2Data = worksheet.Cells["E2"].Value?.ToString() ?? string.Empty;
                                    f2Data = worksheet.Cells["F2"].Value?.ToString() ?? string.Empty;
                                    g2Data = worksheet.Cells["G2"].Value?.ToString() ?? string.Empty;
                                    h2Data = worksheet.Cells["H2"].Value?.ToString() ?? string.Empty;
                                    h2Data = worksheet.Cells["H2"].Value?.ToString() ?? string.Empty;
                                    i2Data = worksheet.Cells["I2"].Value?.ToString() ?? string.Empty;
                                }

                                switch (_wjm_)
                                {
                                    case "1.btw":
                                        labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                        textBox1.Text = g2Data;
                                        //labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        //高压情况
                                        if (comboBox_标签规格.Text.Contains("高压"))
                                        {
                                            double.TryParse(h2Data, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                                            double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString() + "A");
                                        }

                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                        labelFormat.SubStrings.SetSubString("CPCD", " ");

                                        //客户型号被选择时
                                        if (checkBox_客户型号.Checked)
                                        {
                                            //output_灯带型号 = "ART. No.: " + i2Data;
                                            int ai = i2Data.Length;
                                            if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                            else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                            else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                            name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);

                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }

                                        break;
                                    case "3.btw":
                                        labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-03", d2Data);
                                        textBox1.Text = g2Data;
                                        //labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                        //高压情况
                                        if (comboBox_标签规格.Text.Contains("高压"))
                                        {
                                            double.TryParse(h2Data, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                                            double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString() + "A");
                                        }

                                        labelFormat.SubStrings.SetSubString("CPCD", " ");

                                        //客户型号被选择时
                                        if (checkBox_客户型号.Checked)
                                        {
                                            //output_灯带型号 = "ART. No.: " + i2Data;
                                            int ai = i2Data.Length;
                                            if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                            else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                            else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                            name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }

                                        break;
                                    case "4.btw":
                                        labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-03", d2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-04", e2Data);
                                        textBox1.Text = g2Data;
                                        //labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                        //高压情况
                                        if (comboBox_标签规格.Text.Contains("高压"))
                                        {
                                            double.TryParse(h2Data, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                                            double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString() + "A");
                                        }

                                        labelFormat.SubStrings.SetSubString("CPCD", " ");

                                        //客户型号被选择时
                                        if (checkBox_客户型号.Checked)
                                        {
                                            //output_灯带型号 = "ART. No.: " + i2Data;
                                            int ai = i2Data.Length;
                                            if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                            else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                            else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                            name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        break;
                                    case "5.btw":
                                        labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-03", d2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-04", e2Data);
                                        labelFormat.SubStrings.SetSubString("BSM-05", f2Data);
                                        textBox1.Text = g2Data;
                                        //labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);



                                        labelFormat.SubStrings.SetSubString("CPCD", " ");


                                        //客户型号被选择时
                                        if (checkBox_客户型号.Checked)
                                        {
                                            //output_灯带型号 = "ART. No.: " + i2Data;
                                            int ai = i2Data.Length;
                                            if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                            else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                            else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                            name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }

                                        break;
                                    default:
                                        MessageBox.Show("未知的模板文件。");
                                        return;
                                }






                            }


                            //判断显指内容
                            // 重置 XZ 字段的内容
                            labelFormat.SubStrings.SetSubString("XZ", "");
                            string BPrefixContent = string.Empty; // 用于存储 "B-" 前面的内容

                            // 查找 "B-" 并获取它之前的所有内容
                            int BIndex = cpxxBox.Text.IndexOf("\r\nB-");
                            if (BIndex != -1 && BIndex > 0) // 确保 "B-" 存在且不是在字符串开头
                            {
                                BPrefixContent = cpxxBox.Text.Substring(0, BIndex).Trim();
                            }
                            //MessageBox.Show(灯带系列);
                            // 检查是否存在 "Ra90" 或 "Ra95"
                            bool containsRa90 = BPrefixContent.Contains("Ra90");
                            bool containsRa95 = BPrefixContent.Contains("Ra95");

                            if (BPrefixContent.Contains("三面发光"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", 灯带系列 + @"T");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                                // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                                if (containsRa90 || containsRa95)
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                    string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                    labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                                }

                            }
                            else if (BPrefixContent.Contains("高亮"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "BH");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                                // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                                if (containsRa90 || containsRa95)
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                    string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                    labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                                }

                            }
                            else if (BPrefixContent.Contains("翻边"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "BF");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                                // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                                if (containsRa90 || containsRa95)
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                    string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                    labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                                }

                            }
                            else if (BPrefixContent.Contains("DTW"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "DTW");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                                // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                                if (containsRa90 || containsRa95)
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                    string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                    labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                                }

                            }
                            else if (灯带系列 == "D")
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "D");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                                if (containsRa90 || containsRa95)
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                    string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                    labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                                }
                            }
                            else if (BPrefixContent.Contains("Ra90"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "Ra90");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");

                            }
                            else if (BPrefixContent.Contains("Ra95"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "Ra95");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }
                            else
                            {
                                // 如果没有找到上述任何关键字，则设置为空
                                labelFormat.SubStrings.SetSubString("XZ", "");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-2", "空.png");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }

                            //检查是否为非水下内容
                            // 检查 cpxxBox.Text 中是否同时包含 "IP68" 和 "非水下"
                            bool containsIP68 = cpxxBox.Text.Contains("IP68");
                            bool containsNonUnderwater = cpxxBox.Text.Contains("非水下");


                            if (标签种类_comboBox.Text == "品名标")
                            {
                                // 检查cpxxBox文本中是否包含"非水下方案"
                                if (cpxxBox.Text.Contains("非水下方案"))
                                {
                                    // 如果包含"非水下方案"，则设置labelFormat的"SX"子字符串为"Not suitable for underwater use"
                                    labelFormat.SubStrings.SetSubString("SX", "Not suitable for underwater use");
                                }
                                else
                                {
                                    // 否则，将"SX"子字符串设置为空字符串
                                    labelFormat.SubStrings.SetSubString("SX", " ");
                                }

                            }
                            else
                            {
                                if (comboBox_标签规格.Text.Contains("水下"))
                                {
                                    labelFormat.SubStrings.SetSubString("SX", "空.png");
                                }
                                else
                                {
                                    labelFormat.SubStrings.SetSubString("SX", containsIP68 && containsNonUnderwater ? "非水下.png" : "正常.png");
                                }

                            }
                        }

                        catch (Exception ex)
                        {
                            MessageBox.Show("修改内容出错 " + ex.Message, "操作提示");
                        }

                        //生成预览图
                        pictureBox.Image = null;
                        if (labelFormat != null)
                        {
                            //MessageBox.Show(_bmp_path, "操作提示");
                            //Generate a thumbnail for it.
                            labelFormat.ExportImageToFile(_bmp_path, ImageType.BMP, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(407, 407), OverwriteOptions.Overwrite);
                            System.Drawing.Image image = System.Drawing.Image.FromFile(_bmp_path);
                            Bitmap NmpImage = new Bitmap(image);
                            pictureBox.Image = NmpImage;
                            image.Dispose();

                        }
                        else
                        {
                            MessageBox.Show("生成图片错误", "操作提示");
                        }

                        if (isPreView) return;



                        // 确保打印机已选择
                        if (_PrinterName == "")
                        {
                            MessageBox.Show("请先选择打印机。");
                            return;
                        }

                        //读取excel文件必须增加声明才能运行正常
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                        if (!string.IsNullOrEmpty(Box_数据库.Text))
                        {
                            // 确保filePath是有效的Excel文件路径
                            string filePath = Box_数据库.Text;
                            // 使用EPPlus打开Excel文件
                            using (var package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                // 读取工作表
                                var worksheet = package.Workbook.Worksheets[0];

                                // 从第二行开始遍历，假设第一行是标题行
                                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    // 读取每列的数据
                                    var aData = worksheet.Cells[row, 1].Value?.ToString() ?? string.Empty;
                                    var bData = worksheet.Cells[row, 2].Value?.ToString() ?? string.Empty;
                                    var cData = worksheet.Cells[row, 3].Value?.ToString() ?? string.Empty;
                                    var dData = worksheet.Cells[row, 4].Value?.ToString() ?? string.Empty;
                                    var eData = worksheet.Cells[row, 5].Value?.ToString() ?? string.Empty;
                                    var fData = worksheet.Cells[row, 6].Value?.ToString() ?? string.Empty;
                                    var gData = worksheet.Cells[row, 7].Value?.ToString() ?? string.Empty;
                                    var hData = worksheet.Cells[row, 8].Value?.ToString() ?? string.Empty;
                                    var iData = worksheet.Cells[row, 9].Value?.ToString() ?? string.Empty;



                                    if (_wjm_ == "1.btw")
                                    {
                                        // 检查labelFormat是否已初始化
                                        if (labelFormat != null)
                                        {
                                            // 使用读取的数据设置labelFormat的子字符串
                                            labelFormat.SubStrings.SetSubString("XLH", aData);
                                            labelFormat.SubStrings.SetSubString("BSM-01", bData);
                                            labelFormat.SubStrings.SetSubString("BSM-02", cData);
                                            textBox1.Text = gData; // 假设textBox1是WinForms的TextBox控件
                                            //labelFormat.SubStrings.SetSubString("CPCD", hData);
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                            //高压情况
                                            if (comboBox_标签规格.Text.Contains("高压"))
                                            {
                                                double.TryParse(hData, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                                                double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString() + "A");
                                            }

                                            labelFormat.SubStrings.SetSubString("CPCD", " ");

                                            //客户型号被选择时
                                            if (checkBox_客户型号.Checked)
                                            {
                                                //output_灯带型号 = "ART. No.: " + iData;
                                                int ai = iData.Length;
                                                if (ai <= 19) { output_灯带型号 = "ART. No.: " + iData; }
                                                else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + iData; }
                                                else { output_灯带型号 = "ART. No.: " + iData.Substring(0, 19) + Environment.NewLine + iData.Substring(19); }
                                                name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                            }

                                            // 设置打印机名称
                                            labelFormat.PrintSetup.PrinterName = _PrinterName;

                                            // 执行打印操作
                                            int 次数 = Convert.ToInt32(textBox1.Text);

                                            labelFormat.PrintSetup.PrinterName = _PrinterName;
                                            for (int i = 0; i < 次数; i++)
                                            {
                                                labelFormat.Print("BarPrint" + DateTime.Now, 300);
                                            }

                                        }
                                        else
                                        {
                                            MessageBox.Show("标签格式未初始化。");
                                            break; // 如果labelFormat未初始化，退出循环
                                        }
                                    }
                                    if (_wjm_ == "3.btw")
                                    {
                                        // 检查labelFormat是否已初始化
                                        if (labelFormat != null)
                                        {
                                            // 使用读取的数据设置labelFormat的子字符串
                                            labelFormat.SubStrings.SetSubString("XLH", aData);
                                            labelFormat.SubStrings.SetSubString("BSM-01", bData);
                                            labelFormat.SubStrings.SetSubString("BSM-02", cData);
                                            labelFormat.SubStrings.SetSubString("BSM-03", dData);
                                            textBox1.Text = gData; // 假设textBox1是WinForms的TextBox控件
                                            //labelFormat.SubStrings.SetSubString("CPCD", hData);
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                            //高压情况
                                            if (comboBox_标签规格.Text.Contains("高压"))
                                            {
                                                double.TryParse(hData, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                                                double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString() + "A");
                                            }

                                            labelFormat.SubStrings.SetSubString("CPCD", " ");

                                            //客户型号被选择时
                                            if (checkBox_客户型号.Checked)
                                            {
                                                //output_灯带型号 = "ART. No.: " + iData;
                                                int ai = iData.Length;
                                                if (ai <= 19) { output_灯带型号 = "ART. No.: " + iData; }
                                                else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + iData; }
                                                else { output_灯带型号 = "ART. No.: " + iData.Substring(0, 19) + Environment.NewLine + iData.Substring(19); }
                                                name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                            }

                                            // 设置打印机名称
                                            labelFormat.PrintSetup.PrinterName = _PrinterName;

                                            // 执行打印操作
                                            int 次数 = Convert.ToInt32(textBox1.Text);

                                            labelFormat.PrintSetup.PrinterName = _PrinterName;
                                            for (int i = 0; i < 次数; i++)
                                            {
                                                labelFormat.Print("BarPrint" + DateTime.Now, 300);
                                            }

                                        }
                                        else
                                        {
                                            MessageBox.Show("标签格式未初始化。");
                                            break; // 如果labelFormat未初始化，退出循环
                                        }
                                    }
                                    if (_wjm_ == "4.btw")
                                    {
                                        // 检查labelFormat是否已初始化
                                        if (labelFormat != null)
                                        {
                                            // 使用读取的数据设置labelFormat的子字符串
                                            labelFormat.SubStrings.SetSubString("XLH", aData);
                                            labelFormat.SubStrings.SetSubString("BSM-01", bData);
                                            labelFormat.SubStrings.SetSubString("BSM-02", cData);
                                            labelFormat.SubStrings.SetSubString("BSM-03", dData);
                                            labelFormat.SubStrings.SetSubString("BSM-04", eData);
                                            textBox1.Text = gData; // 假设textBox1是WinForms的TextBox控件
                                            //labelFormat.SubStrings.SetSubString("CPCD", hData);
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                            labelFormat.SubStrings.SetSubString("CPCD", " ");

                                            //客户型号被选择时
                                            if (checkBox_客户型号.Checked)
                                            {
                                                //output_灯带型号 = "ART. No.: " + iData;
                                                int ai = iData.Length;
                                                if (ai <= 19) { output_灯带型号 = "ART. No.: " + iData; }
                                                else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + iData; }
                                                else { output_灯带型号 = "ART. No.: " + iData.Substring(0, 19) + Environment.NewLine + iData.Substring(19); }
                                                name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                            }

                                            // 设置打印机名称
                                            labelFormat.PrintSetup.PrinterName = _PrinterName;

                                            // 执行打印操作
                                            int 次数 = Convert.ToInt32(textBox1.Text);

                                            labelFormat.PrintSetup.PrinterName = _PrinterName;
                                            for (int i = 0; i < 次数; i++)
                                            {
                                                labelFormat.Print("BarPrint" + DateTime.Now, 300);
                                            }

                                        }
                                        else
                                        {
                                            MessageBox.Show("标签格式未初始化。");
                                            break; // 如果labelFormat未初始化，退出循环
                                        }
                                    }
                                    if (_wjm_ == "5.btw")
                                    {
                                        // 检查labelFormat是否已初始化
                                        if (labelFormat != null)
                                        {
                                            // 使用读取的数据设置labelFormat的子字符串
                                            labelFormat.SubStrings.SetSubString("XLH", aData);
                                            labelFormat.SubStrings.SetSubString("BSM-01", bData);
                                            labelFormat.SubStrings.SetSubString("BSM-02", cData);
                                            labelFormat.SubStrings.SetSubString("BSM-03", dData);
                                            labelFormat.SubStrings.SetSubString("BSM-04", eData);
                                            labelFormat.SubStrings.SetSubString("BSM-05", fData);
                                            textBox1.Text = gData; // 假设textBox1是WinForms的TextBox控件
                                            //labelFormat.SubStrings.SetSubString("CPCD", hData);
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                            labelFormat.SubStrings.SetSubString("CPCD", " ");

                                            //客户型号被选择时
                                            if (checkBox_客户型号.Checked)
                                            {
                                                //output_灯带型号 = "ART. No.: " + iData;
                                                int ai = iData.Length;
                                                if (ai <= 19) { output_灯带型号 = "ART. No.: " + iData; }
                                                else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + iData; }
                                                else { output_灯带型号 = "ART. No.: " + iData.Substring(0, 19) + Environment.NewLine + iData.Substring(19); }
                                                name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                            }

                                            // 设置打印机名称
                                            labelFormat.PrintSetup.PrinterName = _PrinterName;

                                            // 执行打印操作
                                            int 次数 = Convert.ToInt32(textBox1.Text);

                                            labelFormat.PrintSetup.PrinterName = _PrinterName;
                                            for (int i = 0; i < 次数; i++)
                                            {
                                                labelFormat.Print("BarPrint" + DateTime.Now, 300);
                                            }

                                        }
                                        else
                                        {
                                            MessageBox.Show("标签格式未初始化。");
                                            break; // 如果labelFormat未初始化，退出循环
                                        }
                                    }
                                }
                            }

                        }
                        else
                        {
                            // 设置打印机名称
                            labelFormat.PrintSetup.PrinterName = _PrinterName;

                            // 执行打印操作
                            int 次数 = Convert.ToInt32(textBox1.Text);

                            labelFormat.PrintSetup.PrinterName = _PrinterName;
                            for (int i = 0; i < 次数; i++)
                            {
                                labelFormat.Print("BarPrint" + DateTime.Now, 300);
                            }
                        }

                    }




                }




            }
            if (标签种类_comboBox.Text == "唛头")
            {


                唛头 mt = new 唛头();
                mt.正常型号判断(cpxxBox.Text, checkBox_客户Name.Checked, checkBox_客户型号.Checked, textBox_客户资料.Text, comboBox_标签规格.Text, 标签种类_comboBox.Text, textBox_唛头数量.Text, textBox_唛头尺寸.Text);
                using (Engine btEngine = new Engine(true))
                {
                    if (!_btw_path.Contains(comboBox_标签规格.Text))
                    {
                        _btw_path = _btw_path + comboBox_标签规格.Text; // 如果不包含，则拼接
                    }

                    //寻找文件名_单字匹配("正弯", "侧弯");
                    string 复选框 = 判断复选框内容(cpxxBox.Text, comboBox_标签规格.Text);

                    模板地址 = _btw_path + @"\" + _wjm_;
                    //MessageBox.Show(模板地址, "操作提示");
                    模板地址 = 模板地址.Replace("\n", "").Replace("\r", "");  //去除换行符，否则下面会报错

                    if (_wjm_.Length > 2)
                    { 
                        LabelFormatDocument labelFormat = btEngine.Documents.Open(模板地址);

                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, textBox_剪切长度.Text);
                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text );
                        //labelFormat.SubStrings.SetSubString("CPCD", textBox_剪切长度.Text);
                        //高压情况
                        if (comboBox_标签规格.Text.Contains("高压"))
                        {
                            double.TryParse(textBox_剪切长度.Text, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                            double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString() + "A");
                        }

                        labelFormat.SubStrings.SetSubString("CPCD", " ");
                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                        labelFormat.SubStrings.SetSubString("XLH", " ");


                    }

                }

            }

        }

            
    

        //判断复选框内容
        private static string 判断复选框内容(string input,string 标签规格)
        {
            
            bool hasConstantCurrent = input.Contains("恒流");
            bool hasConstantVoltage = !hasConstantCurrent; // 如果没有恒流，则默认为恒压

            string firstField = hasConstantCurrent ? "恒流" : "恒压";

            // 检查是否有正弯或侧弯
            bool hasPositiveBend = input.Contains("正弯");
            bool hasSideBend = input.Contains("侧弯");

            string secondField = hasPositiveBend ? "正弯" : (hasSideBend ? "侧弯" : "");

           
            // 检查comboBox_标签规格的当前选择项目文本中是否包含"RCM"
            if (标签规格.Contains("RCM"))
            {
                
                // 如果包含"RCM"，检查output_灯带型号的文本
                var models = new[] { "F10", "F11", "F15", "F21", "F2222" }; // 这里可以添加或修改型号列表
                bool isSideBend = models.Any(model => input.Contains(model));

                // 如果是特定的型号，则默认为侧弯
                if (isSideBend)
                {
                    secondField = "侧弯";
                }
                else
                {
                    secondField = "正弯"; // 其他的都是正弯
                }

                // 构建结果
                return $"{firstField}-{secondField}.png";
            }
            else
            {
                // 构建结果
                if (!string.IsNullOrEmpty(secondField))
                {
                    return $"{firstField}-{secondField}.png";
                }
                else
                {
                    return $"{firstField}.png";
                }

            }



        }





        //判断产品信息
        private void 判断产品信息(string aa)
        {

            



            // 正则表达式模式，
            string pattern1 = @"^(\w+-\w+-\w+)";
            string pattern2 = @"D(\d+)V";
            //string pattern3 = @"额定功率(\d+)W";
            string pattern3 = @"额定功率(\d+(?:\.\d+)?)W";
            string pattern4 = @"-(\d+)-";
            string pattern5 = @"(\d+)灯\/(\d+\.?\d*)cm";
            string pattern6 = @"-IP(\d{2})";


            // 使用“-”字符分割输入字符串
            string[] parts = aa.Split('-');

            // 使用正则表达式匹配输入字符串
            Match match1 = Regex.Match(aa, pattern1);
            Match match2 = Regex.Match(aa, pattern2);
            Match match3 = Regex.Match(aa, pattern3);
            Match match4 = Regex.Match(aa, pattern4);
            Match match5 = Regex.Match(aa, pattern5);
            Match match6 = Regex.Match(aa, pattern6);

            

            //灯带型号
            if (match1.Success)   
            {
                // 构造输出字符串
                string artNo = match1.Groups[1].Value; // 第一个括号匹配的内容
                //output_灯带型号 = $"ART. No.: {artNo}";

                // 使用信息框输出结果
                //MessageBox.Show(output1, "提取结果");

                // 检查复选框是否同时被选中
                bool isCustomerNameChecked = checkBox_客户Name.Checked;
                bool isCustomerModelChecked = checkBox_客户型号.Checked;

                string originalString = textBox_客户资料.Text;
                int spaceIndex = originalString.LastIndexOf('	');

                int textLength = textBox_客户资料.Text.Length;

                if (isCustomerNameChecked && !isCustomerModelChecked)
                {
                    // 如果只有 checkBox_客户Name 被选中，则输出 1
                    //MessageBox.Show("1");
                    //output_name = "Name:" + textBox_客户资料.Text;
                    if (textLength <= 19)
                    {
                        // 如果长度在19以内
                        output_name = "Name: " + textBox_客户资料.Text;
                    }
                    else if (textLength > 19 && textLength <= 27)
                    {
                        // 如果长度大于19，小于27
                        // 在第20个字符的位置插入换行符
                        output_name = "Name: " + Environment.NewLine + textBox_客户资料.Text;
                    }
                    else
                    {
                        // 如果长度大于27
                        // 取前19个字符，然后加上剩余的字符
                        output_name = "Name: " + textBox_客户资料.Text.Substring(0, 19) +Environment.NewLine +textBox_客户资料.Text.Substring(19);
                    }

                }
                else if (!isCustomerNameChecked && isCustomerModelChecked)
                {
                    // 如果只有 checkBox_客户型号 被选中，则输出 2
                    //MessageBox.Show("2");
                    //output_灯带型号 = "ART. No.: " + textBox_客户资料.Text;   
                    if (textLength <= 19)
                    {
                        // 如果长度在19以内
                        output_灯带型号 = "ART. No.: " + textBox_客户资料.Text;
                    }
                    else if (textLength > 19 && textLength <= 27)
                    {
                        // 如果长度大于19，小于27
                        // 在第20个字符的位置插入换行符
                        output_灯带型号 = "ART. No.: " + Environment.NewLine + textBox_客户资料.Text;
                    }
                    else
                    {
                        // 如果长度大于27
                        // 取前19个字符，然后加上剩余的字符
                        output_灯带型号 = "ART. No.: " + textBox_客户资料.Text.Substring(0, 19) + Environment.NewLine + textBox_客户资料.Text.Substring(19);
                    }

                }
                else if (isCustomerNameChecked && isCustomerModelChecked)
                {
                    // 如果两个复选框都被选中，则按照制表符拆分字符串并显示
                    int tabIndex = originalString.LastIndexOf('\t');

                    if (tabIndex != -1)
                    {
                        string part1 = originalString.Substring(0, tabIndex);
                        string part2 = originalString.Substring(tabIndex + 1);
                        int a1 = part1.Length;
                        int a2 = part2.Length;
                        //MessageBox.Show(part1 + "\n" + part2);
                        //output_name = "Name: " + part1 ;
                        //output_灯带型号 = "ART. No.: " + part2 ;
                        if (a1 <= 19)
                        {
                            output_name = "Name: " + part1;
                        }
                        else if (a1 > 19 && a1 <= 27)
                        {
                            output_name = "Name: " + Environment.NewLine + part1;
                        }
                        else
                        {
                            output_name = "Name: " + part1.Substring(0, 19) + Environment.NewLine + part1.Substring(19);
                        }
                        if (a2 <= 19) { output_灯带型号 = "ART. No.: " + part2; }
                        else if (a2 > 19 && a2 <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + part2; }
                        else { output_灯带型号 = "ART. No.: " + part2.Substring(0, 19) + Environment.NewLine + part2.Substring(19);}


                    }
                    else
                    {
                        MessageBox.Show("未找到制表符分隔符。");
                    }
                }
                else
                {
                    // 如果两个复选框都没有被选中，则直接显示原始字符串
                    //MessageBox.Show(originalString);
                    output_name = "Name:LED Flex Linear Light";
                    //output_灯带型号 = $"ART. No.: {artNo}";
                    int a4 = artNo.Length;
                    if (a4 <= 19)
                    {
                        output_灯带型号 = "ART. No.: " + artNo;
                    }
                    else if (a4 > 19 && a4 <= 27)
                    {
                        output_灯带型号 = "ART. No.: " + Environment.NewLine + artNo;
                    }
                    else
                    {
                        output_灯带型号 = "ART. No.: " + artNo.Substring(0, 19) + Environment.NewLine + artNo.Substring(19);
                    }

                }

                string lightModel = output_灯带型号;
                // 获取最后一个字符
                char lastChar = lightModel[lightModel.Length - 1];
                // 是否为字母
                if (char.IsLetter(lastChar))
                {
                    灯带系列 = lastChar.ToString();
                    //MessageBox.Show(灯带系列);
                }


            }
            else
            {
                // 如果没有找到匹配项，则输出错误信息
                MessageBox.Show("未找到灯带型号匹配项。", "错误");
            }

            //电压
            if (match2.Success)  
            {
                // 从匹配结果中提取电压值
                string voltageValue = match2.Groups[1].Value; // 第一个捕获组匹配的内容

                // 检查 comboBox_标签规格.Text 的内容中是否包含“高压”这两个字
                if (comboBox_标签规格.Text.Contains("高压"))
                {
                    // 如果包含“高压”，则设置 output_电压 为 AC
                    output_电压 = $"Rated Voltage: AC {voltageValue}V";
                }
                else
                {
                    // 如果不包含“高压”，则设置 output_电压 为 DC
                    output_电压 = $"Rated Voltage: DC {voltageValue}V";
                }

                // 使用信息框输出结果
                //MessageBox.Show(output2, "提取结果");
            }
            else
            {
                // 如果没有找到匹配项，则输出错误信息
                MessageBox.Show("未找到电压匹配项。", "错误");
            }

            //功率
            if (match3.Success)
            {
                // 从匹配结果中提取功率值
                string powerValue = match3.Groups[1].Value; // 第一个捕获组匹配的内容

                // 构造输出字符串
                output_功率 = $"Rated Power: {powerValue}W/m";
            }
            else
            {
                MessageBox.Show("未找到功率匹配项。", "错误");
            }

            // 灯数
            if (match4.Success)
            {
                // 从匹配结果中提取数字
                string ledQtyValue = match4.Groups[1].Value; // 第一个捕获组匹配的内容

                // 构造输出字符串
                output_灯数 = $"LED Qty.: {ledQtyValue}LEDs/m";
            }
            else
            {
                MessageBox.Show("未找到灯数匹配项。", "错误");
            }

            // 剪切单元
            if (match5.Success)
            {
                // 从匹配结果中提取灯数和长度
                string ledQuantity = match5.Groups[1].Value; // 第一个捕获组匹配的内容，如 "7灯"
                string length = match5.Groups[2].Value; // 第二个捕获组匹配的内容，如 "5.56"

                // 构造输出字符串
                //output_剪切单元 = $"Min. Cutting Length: {ledQuantity}LEDs\n({length}cm)";
                if (标签种类_comboBox.Text == "品名标")
                {
                    output_剪切单元 = $"Min. Cutting Length: {ledQuantity}LEDs({length}cm)";
                }
                else
                {
                    output_剪切单元 = $"Min. Cutting Length: {ledQuantity}LEDs\n({length}cm)";
                }
            }
            else
            {
                // 如果没有找到匹配项，则输出错误信息
                MessageBox.Show( "未找到剪切单元信息匹配项","错误");

            }

            // 色温
            if (parts.Length >= 6)
            {
                string numericValue;
                // 第五个"-"和第六个"-"之间的内容是parts[5]，因为数组索引是从0开始的
                string contentBetweenFifthAndSixth = parts[5];

              

                // 检查灯带系列是否为"S"或"D"
                string series = 灯带系列; 
                if (series.Equals("S") || series.Equals("E"))
                {

                    // 如果灯带系列是"S"或"D"，执行新的逻辑

                    output_色温 = $"Color: {contentBetweenFifthAndSixth}";// 示例，表示输出所有色温内容

                }
                else
                {
                    // 如果灯带系列是"S"或"D"，按照现有逻辑判断色温
                    // 检查内容是否为纯字母
                    if (Regex.IsMatch(contentBetweenFifthAndSixth, @"^[a-zA-Z]*$"))
                    {
                        output_色温 = $"Color: {contentBetweenFifthAndSixth}";
                    }
                    else
                    {
                        // 如果包含数字，则提取数字部分
                        numericValue = Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", "");
                        output_色温 = $"Color: {numericValue}K";
                    }
                }

                // 检查 "全彩" 是否存在于 cpxxBox.Text 中
                bool containsFullColor = cpxxBox.Text.Contains("全彩");
                

                if (contentBetweenFifthAndSixth == "R" && containsFullColor) { output_色温 = $"Color: Red(Full color jacket)"; }
                else if(contentBetweenFifthAndSixth == "R" && ! containsFullColor) { output_色温 = $"Color: Red"; }
                else if (contentBetweenFifthAndSixth == "B" && containsFullColor) { output_色温 = $"Color: Blue(Full color jacket)"; }
                else if (contentBetweenFifthAndSixth == "B" && !containsFullColor) { output_色温 = $"Color: Blue"; }
                else if (contentBetweenFifthAndSixth == "G" && containsFullColor) { output_色温 = $"Color: Green(Full color jacket)"; }
                else if (contentBetweenFifthAndSixth == "G" && !containsFullColor) { output_色温 = $"Color: Green"; }
                else if (contentBetweenFifthAndSixth == "O" && containsFullColor) { output_色温 = $"Color: Orange(Full color jacket)"; }
                else if (contentBetweenFifthAndSixth == "O" && !containsFullColor) { output_色温 = $"Color: Orange"; }
                else if (contentBetweenFifthAndSixth == "Y" && containsFullColor) { output_色温 = $"Color: Yellow(Full color jacket)"; }
                else if (contentBetweenFifthAndSixth == "Y" && !containsFullColor) { output_色温 = $"Color: Yellow"; }
                else if (contentBetweenFifthAndSixth == "Y578" ) { output_色温 = $"Color: Yellow (Full color jacket) (Y578nm)"; }
                else if (contentBetweenFifthAndSixth == "Y580" ) { output_色温 = $"Color: Yellow (Full color jacket) (Y580nm)"; }
                else if (contentBetweenFifthAndSixth == "Y582" ) { output_色温 = $"Color: Yellow (Full color jacket) (Y582nm)"; }
                else if (cpxxBox.Text.Contains("黑色遮光+雾面发光") ) { output_色温 = $"Color: {Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", "")}K(Black jacket)"; }
                else if (cpxxBox.Text.Contains("黑色全彩") ) { output_色温 = $"Color: {Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", "")}K(Full Black jacket)"; }
                else if (cpxxBox.Text.Contains("白+暖白") ) { output_色温 = $"Color: Warm White+White"; }
                else if (cpxxBox.Text.Contains("暖白+暖白") ) { output_色温 = $"Color: Warm White+Warm White"; }






            }
            else
            {
                MessageBox.Show("未找到色温匹配项。", "错误");
            }

            // IP等级
            if (match6.Success)
            {
                // 从匹配结果中提取数字
                // string ipNumber = match6.Groups[1].Value; // 第一个捕获组匹配的内容
                string result = FindMinimumIPNumber(cpxxBox.Text);
                if (result != null)
                {
                    Console.WriteLine($"最小的IP等级数字是: {result}");
                    bq_ipdj = $"IP{result}"; // 假设这是全局变量或在类中定义的属性
                    Console.WriteLine($"bq_ipdj: {bq_ipdj}");
                }
                


            }
            else if (comboBox_标签规格.Text.Contains("直发") && 标签种类_comboBox.Text == "品名标")
            {
                bq_ipdj = "  ";

            }
            else
            {
                // 如果没有找到匹配项，则输出错误信息
                MessageBox.Show("未找到IP等级匹配项。", "错误");
            }

            //判断尾巴Made in China
            if (checkBox_结尾.Checked)
            {
                output_尾巴 =textBox_结尾.Text ;
            }
            else
            {
                output_尾巴 = " "; 
            }


            //MessageBox.Show(output_name +"\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴, "提取结果");
            name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;

        }

        //判断-IP
        private static string FindMinimumIPNumber(string text)
        {
            // 正则表达式，用于匹配 -IP 后面的数字
            string pattern = @"-IP(\d{2})";
            var matches = Regex.Matches(text, pattern);

            if (matches.Count > 0)
            {
                // 转换为数字，找出最小值
                int minIPNumber = matches.Cast<Match>().Select(m => int.Parse(m.Groups[1].Value)).Min();
                return minIPNumber.ToString();
            }

            return null; // 如果没有找到任何匹配项，返回 null
        }

        //加载excel数据路径
        private async void button_数据库_Click_1(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择数据库文件";
            dialog.Filter = "excel文件(*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dialog.InitialDirectory = Application.StartupPath + @"\数据库";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _sjk_path = dialog.FileName;
                Box_数据库.Text = dialog.FileName;
                Box_数据库.BackColor = System.Drawing.Color.LightGreen;

                
            }
        }

        //读取excel数据内容 
        private void button_test_Click_2(object sender, EventArgs e)
        {
            //唛头 mt = new 唛头();
            //mt.正常型号判断(cpxxBox.Text, checkBox_客户Name.Checked ,checkBox_客户型号.Checked ,textBox_客户资料.Text ,comboBox_标签规格.Text ,标签种类_comboBox.Text,textBox_唛头数量.Text ,textBox_唛头尺寸.Text );
            int textLength = textBox_text.Text.Length;
            MessageBox.Show("文本框中包含文本，长度为: " + textLength);
            
          
        }



    private void 寻找文件名_单字匹配(string aa1,string aa2)
        {
            // 检查文本中是否包含“aa1”或“aa2”
            bool containsPositiveBend = cpxxBox.Text.Contains(aa1);
            bool containsSideBend = cpxxBox.Text.Contains(aa2);

            string[] files = Directory.GetFiles(_btw_path);

            string message = "";

            if (containsPositiveBend || containsSideBend)
            {
                foreach (string file in files)
                {
                    if (file.Contains(aa1) && containsPositiveBend)
                    {
                        message += Path.GetFileName(file) + "\n";
                    }
                    else if (file.Contains(aa2) && containsSideBend)
                    {
                        message += Path.GetFileName(file) + "\n";
                    }
                }
            }

            if (!string.IsNullOrEmpty(message))
            {
                _wjm_ = message;
            }
            else
            {
                _wjm_ = "";
                MessageBox.Show("没有找到包含指定字眼的文件。", "未找到文件");
            }
        }


        private void 另存模板()
        {
            SaveFileDialog dialog = new SaveFileDialog();
            dialog.Title = "请选择要保存的文件地址";
            dialog.Filter = "bwt文件(*.btw)|*.btw";
            dialog.InitialDirectory = Application.StartupPath+ @"\输出文件";
            dialog.DefaultExt = "btw"; // 设置默认文件扩展名
            dialog.AddExtension = true; // 确保即使用户未指定扩展名也会添加扩展名

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                using (Engine btEngine = new Engine(true))
                {
                    string 复选框 = 判断复选框内容(cpxxBox.Text, comboBox_标签规格.Text );

                    LabelFormatDocument labelFormat = btEngine.Documents.Open(模板地址);
                    //LabelFormatDocument labelFormat = btEngine.Documents.Open(@"E:\正在进行项目\经管中心-标签打印\python\BarTender_Dev_Dome-master\BarTender_Dev_Dome-master\BarTender_Dev_Dome\bin\Debug\moban\工字标\tesssst-正弯.btw");
                    try
                    {
                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, textBox_剪切长度.Text);
                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                        //labelFormat.SubStrings.SetSubString("CPCD", textBox_剪切长度.Text);
                        labelFormat.SubStrings.SetSubString("CPCD", " ");

                        //高压情况
                        if (comboBox_标签规格.Text.Contains("高压"))
                        {
                            double.TryParse(textBox_剪切长度.Text, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                            double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString() + "A");
                        }

                        //labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj);
                        if (comboBox_标签规格.Text.Contains("水下"))
                        {
                            labelFormat.SubStrings.SetSubString("IPDJ", "IP68 5m");
                        }
                        else
                        {
                            labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj);
                        }

                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                        labelFormat.SubStrings.SetSubString("XLH", " ");

                        //写入2排标识码时候的内容
                        //判断显指内容
                        // 重置 XZ 字段的内容
                        labelFormat.SubStrings.SetSubString("XZ", "");
                        string BPrefixContent = string.Empty; // 用于存储 "B-" 前面的内容

                        // 查找 "B-" 并获取它之前的所有内容
                        int BIndex = cpxxBox.Text.IndexOf("\r\nB-");
                        if (BIndex != -1 && BIndex > 0) // 确保 "B-" 存在且不是在字符串开头
                        {
                            BPrefixContent = cpxxBox.Text.Substring(0, BIndex).Trim();
                        }
                        //MessageBox.Show(BPrefixContent, "操作提示");
                        // 检查是否存在 "Ra90" 或 "Ra95"
                        bool containsRa90 = BPrefixContent.Contains("Ra90");
                        bool containsRa95 = BPrefixContent.Contains("Ra95");

                        if (BPrefixContent.Contains("三面发光"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ", 灯带系列 + @"T");
                            labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                            if (containsRa90 || containsRa95)
                            {
                                labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                            }

                        }
                        else if (BPrefixContent.Contains("高亮"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ", "BH");
                            labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                            if (containsRa90 || containsRa95)
                            {
                                labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                            }

                        }
                        else if (BPrefixContent.Contains("翻边"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ", "BF");
                            labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                            if (containsRa90 || containsRa95)
                            {
                                labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                            }

                        }
                        else if (BPrefixContent.Contains("DTW"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ", "DTW");
                            labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                            if (containsRa90 || containsRa95)
                            {
                                labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                            }

                        }
                        else if (灯带系列 == "D")
                        {
                            labelFormat.SubStrings.SetSubString("XZ", "D");
                            labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            if (containsRa90 || containsRa95)
                            {
                                labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : "");
                                labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                            }
                        }
                        else if (BPrefixContent.Contains("Ra90"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ", "Ra90");
                            labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");

                        }
                        else if (BPrefixContent.Contains("Ra95"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ", "Ra95");
                            labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                        }
                        else
                        {
                            // 如果没有找到上述任何关键字，则设置为空
                            labelFormat.SubStrings.SetSubString("XZ", "");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-2", "空.png");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                        }

                        //检查是否为非水下内容
                        // 检查 cpxxBox.Text 中是否同时包含 "IP68" 和 "非水下"
                        bool containsIP68 = cpxxBox.Text.Contains("IP68");
                        bool containsNonUnderwater = cpxxBox.Text.Contains("非水下");
                        //labelFormat.SubStrings.SetSubString("SX", containsIP68 && containsNonUnderwater ? "非水下.png" : "正常.png");
                        if (标签种类_comboBox.Text == "品名标")
                        {
                            // 检查cpxxBox文本中是否包含"非水下方案"
                            if (cpxxBox.Text.Contains("非水下方案"))
                            {
                                // 如果包含"非水下方案"，则设置labelFormat的"SX"子字符串为"Not suitable for underwater use"
                                labelFormat.SubStrings.SetSubString("SX", "Not suitable for underwater use");
                            }
                            else
                            {
                                // 否则，将"SX"子字符串设置为空字符串
                                labelFormat.SubStrings.SetSubString("SX", " ");
                            }

                        }
                        else
                        {
                            if (comboBox_标签规格.Text.Contains("水下"))
                            {
                                labelFormat.SubStrings.SetSubString("SX", "空.png");
                            }
                            else
                            {
                                labelFormat.SubStrings.SetSubString("SX", containsIP68 && containsNonUnderwater ? "非水下.png" : "正常.png");
                            }

                        }


                        //判断是否增加标识码
                        if (checkBox_标识码01.Checked)
                        {
                            labelFormat.SubStrings.SetSubString("BSM-01", textBox_标识码01.Text);
                        }
                        else
                        {
                            labelFormat.SubStrings.SetSubString("BSM-01", " ");
                        }
                        if (checkBox_标识码02.Checked)
                        {
                            labelFormat.SubStrings.SetSubString("BSM-02", textBox_标识码02.Text);
                        }
                        else
                        {
                            labelFormat.SubStrings.SetSubString("BSM-02", " ");
                        }

                        // 检查数据库地址不为空时
                        if (!string.IsNullOrEmpty(Box_数据库.Text))
                        {

                            string b2Data, c2Data, g2Data, h2Data, a2Data, d2Data, e2Data, f2Data, i2Data;

                            // 使用EPPlus打开Excel文件
                            using (var package = new ExcelPackage(new FileInfo(Box_数据库.Text)))
                            {
                                // 假设Excel工作表名为"Sheet1"
                                var worksheet = package.Workbook.Worksheets["Sheet1"];

                                // 读取B2和C2单元格的数据
                                a2Data = worksheet.Cells["A2"].Value?.ToString() ?? string.Empty;
                                b2Data = worksheet.Cells["B2"].Value?.ToString() ?? string.Empty;
                                c2Data = worksheet.Cells["C2"].Value?.ToString() ?? string.Empty;
                                d2Data = worksheet.Cells["D2"].Value?.ToString() ?? string.Empty;
                                e2Data = worksheet.Cells["E2"].Value?.ToString() ?? string.Empty;
                                f2Data = worksheet.Cells["F2"].Value?.ToString() ?? string.Empty;
                                g2Data = worksheet.Cells["G2"].Value?.ToString() ?? string.Empty;
                                h2Data = worksheet.Cells["H2"].Value?.ToString() ?? string.Empty;
                                i2Data = worksheet.Cells["I2"].Value?.ToString() ?? string.Empty;

                            }

                            switch (_wjm_)
                            {
                                case "1.btw":
                                    labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                    textBox1.Text = g2Data;
                                    //labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                    //高压情况
                                    if (comboBox_标签规格.Text.Contains("高压"))
                                    {
                                        double.TryParse(h2Data, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                                        double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString() + "A");
                                    }

                                    labelFormat.SubStrings.SetSubString("CPCD", " ");

                                    //客户型号被选择时
                                    if (checkBox_客户型号.Checked)
                                    {
                                        //output_灯带型号 = "ART. No.: " + i2Data;
                                        int ai = i2Data.Length;
                                        if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                        else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                        else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }

                                    break;
                                case "3.btw":
                                    labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-03", d2Data);
                                    textBox1.Text = g2Data;
                                    //labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                    //高压情况
                                    if (comboBox_标签规格.Text.Contains("高压"))
                                    {
                                        double.TryParse(h2Data, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString());
                                        double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString() + "A");
                                    }

                                    labelFormat.SubStrings.SetSubString("CPCD", " ");

                                    //客户型号被选择时
                                    if (checkBox_客户型号.Checked)
                                    {
                                        //output_灯带型号 = "ART. No.: " + i2Data;
                                        int ai = i2Data.Length;
                                        if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                        else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                        else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }

                                    break;
                                case "4.btw":
                                    labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-03", d2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-04", e2Data);
                                    textBox1.Text = g2Data;
                                    //labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                    labelFormat.SubStrings.SetSubString("CPCD", " ");

                                    //客户型号被选择时
                                    if (checkBox_客户型号.Checked)
                                    {
                                        //output_灯带型号 = "ART. No.: " + i2Data;
                                        int ai = i2Data.Length;
                                        if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                        else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                        else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text,h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }

                                    break;
                                case "5.btw":
                                    labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-03", d2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-04", e2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-05", f2Data);
                                    textBox1.Text = g2Data;
                                    //labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                    labelFormat.SubStrings.SetSubString("CPCD", " ");

                                    //客户型号被选择时
                                    if (checkBox_客户型号.Checked)
                                    {
                                        //output_灯带型号 = "ART. No.: " + i2Data;
                                        int ai = i2Data.Length;
                                        if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                        else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                        else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }

                                    break;
                                default:
                                    MessageBox.Show("未知的模板文件。");
                                    return;
                            }


                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("修改内容出错 " + ex.Message, "操作提示");
                    }

                    // 获取用户指定的文件路径
                    string saveFilePath = dialog.FileName;
                    //MessageBox.Show(saveFilePath, "操作提示");
                    labelFormat.SaveAs(saveFilePath, true);

                    MessageBox.Show("文件输出完成！", "操作提示");
                }

            }

        }



        //另存为
        private void button_另存为_Click(object sender, EventArgs e)
        {
            另存模板();
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

       
    }

}
