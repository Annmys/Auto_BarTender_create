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







namespace BarTender_Dev_Dome

{
    public partial class PrintForm : Form
    {
    string bq_ipdj = null;
    private string _bmp_path = Application.StartupPath + @"\exp.jpg";
        private string _btw_path = "";
        string 模板地址 = "";
        string _sjk_path = "";
        string _PrinterName = "";
        string _wjm_ = "";
        

        //程序开始运行
        public PrintForm()
        {
            InitializeComponent();
            //MessageBox.Show("程序打开", "操作提示");
            this.Load += new EventHandler(MainForm_Load); // 订阅窗体加载事件-读取配置文件

        }


        //读取config配置文件1
        private void MainForm_Load(object sender, EventArgs e)
        {
            string configFilePath = Application.StartupPath + @"\config\标签规格.txt" ; // 配置标签规格文件相对路径
            string fullPath = Path.Combine(Application.StartupPath, configFilePath);

            if (File.Exists(fullPath))
            {
                try
                {
                    comboBox_标签规格.Items.Clear(); // 清空现有的项
                    // 读取所有行
                    string[] specifications = File.ReadAllLines(fullPath);

                    foreach (string spec in specifications)
                    {
                        if (!string.IsNullOrWhiteSpace(spec))
                        {
                            comboBox_标签规格.Items.Add(spec);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("读取配置文件时出错: " + ex.Message);
                }
            }
            else
            {
                MessageBox.Show("配置文件未找到: " + fullPath);
            }
        }

        //读取config配置文件2
        private void LoadLabelSpecifications()
        {
            string configFilePath = Application.StartupPath + @"\config\标签规格.txt"; // 配置文件相对路径
            string fullPath = Path.Combine(Application.StartupPath, configFilePath);

            if (File.Exists(fullPath))
            {
                // 从文件中读取所有行
                string[] specifications = File.ReadAllLines(fullPath);

                foreach (string spec in specifications)
                {
                    if (!string.IsNullOrWhiteSpace(spec))
                    {
                        comboBox_标签规格.Items.Add(spec);
                    }
                }
            }
            else
            {
                MessageBox.Show("配置文件未找到: " + fullPath);
            }
        }

        // 执行筛选逻辑
        private void FilterSpecifications()
        {
            // 创建一个集合来存储需要保留的标签规格
            HashSet<string> specsToKeep = new HashSet<string>();

            // 根据复选框的状态添加对应的规格到集合中
            if (checkBox_中性.Checked)
            {
                specsToKeep.Add("中性");
            }
            if (checkBox_Clear.Checked)
            {
                specsToKeep.Add("Clear");
            }
            if (checkBox_客制.Checked)
            {
                specsToKeep.Add("客制");
            }
            if (checkBox_低压.Checked)
            {
                specsToKeep.Add("低压");
            }
            if (checkBox_高压.Checked)
            {
                specsToKeep.Add("高压");
            }
            if (checkBox_水下.Checked)
            {
                specsToKeep.Add("水下");
            }
            if (checkBox_桑拿.Checked)
            {
                specsToKeep.Add("桑拿");
            }
            if (checkBox_温泉水.Checked)
            {
                specsToKeep.Add("温泉水");
            }
            if (checkBox_高温高湿.Checked)
            {
                specsToKeep.Add("高温高湿");
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

        //筛选标签规格类型
        private void button_筛选_Click(object sender, EventArgs e)
        {

            // 重置 comboBox_标签规格
            comboBox_标签规格.Items.Clear();

            // 重新加载标签规格
            LoadLabelSpecifications();

            // 执行筛选逻辑
            FilterSpecifications();

        }


        private void 获取标签种类()
        {
            // 软件运行目录
            string baseDirectory = AppDomain.CurrentDomain.BaseDirectory;

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


        private void PrintForm_Load(object sender, EventArgs e)
        {
            获取标签种类();

            //标签种类_comboBox.Items.Add("工字标");
            //标签种类_comboBox.Items.Add("测试标");

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


        private void openFilebtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;//多个文件
            dialog.Title = "请选择要烧录的文件";
            dialog.Filter = "bwt文件(*.btw)|*.btw";
            dialog.InitialDirectory = Application.StartupPath;

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _btw_path = dialog.FileName;
                fileNametBox.Text = dialog.FileName;
                // fileNametBox.Text = dialog.SafeFileName;
                fileNametBox.BackColor = System.Drawing.Color.LightGreen;

                pictureBox.Image = null;
                using (Engine btEngine = new Engine(true))
                {
                    LabelFormatDocument labelFormat = btEngine.Documents.Open(_btw_path);

                    if (labelFormat != null)
                    {
                        Seagull.BarTender.Print.Messages m;
                        labelFormat.ExportPrintPreviewToFile(Application.StartupPath, @"\exp.bmp", ImageType.JPEG, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(300, 300), System.Drawing.Color.White, OverwriteOptions.Overwrite, true, true, out m);
                        labelFormat.ExportImageToFile(_bmp_path, ImageType.JPEG, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(300, 300), OverwriteOptions.Overwrite);

                        Image image = Image.FromFile(_bmp_path);
                        Bitmap NmpImage = new Bitmap(image);
                        pictureBox.Image = NmpImage;
                        image.Dispose();
                    }
                    else
                    {
                        MessageBox.Show("生成图片错误", "操作提示");
                    }
                }
            }
        }

        private void printer_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            _PrinterName = printer_comboBox.Text;
        }

        private void 标签种类_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            
              string currentDirectory = Directory.GetCurrentDirectory();
              _btw_path = currentDirectory +@"\moban\"+ 标签种类_comboBox.Text + @"\";
              fileNametBox.Text = currentDirectory+ @"\moban\"+ 标签种类_comboBox.Text + @"\";
            

        }

        //预览和打印
        void PrintBar(bool isPreView = false)
        {
            // 假设这是从某个文本框获取的字符串
            string cpxx_text =cpxxBox.Text;
            判断产品信息(cpxx_text);
           


            if (_btw_path.Length < 5)
            {
                fileNametBox.BackColor = Color.Red;
                return;
            }
            using (Engine btEngine = new Engine(true))
            {

                寻找文件名_单字匹配("正弯", "侧弯");

                模板地址  = _btw_path + _wjm_;
                模板地址 = 模板地址.Replace("\n", "").Replace("\r", "");  //去除换行符，否则下面会报错
                //MessageBox.Show(模板地址, "操作提示");
                //cpxxBox.Text = _btw_path;

                

                if (_wjm_.Length > 2)
                {
                    LabelFormatDocument labelFormat = btEngine.Documents.Open(模板地址);

                    //LabelFormatDocument labelFormat = btEngine.Documents.Open(@"E:\正在进行项目\经管中心-标签打印\python\BarTender_Dev_Dome-master\BarTender_Dev_Dome-master\BarTender_Dev_Dome\bin\Debug\moban\工字标\tesssst-正弯.btw");
                    try
                    {
                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                        labelFormat.SubStrings.SetSubString("CPCD", TextBox_sku.Text);
                        labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj);

                        labelFormat.SubStrings.SetSubString("test", "实心.jpg");

                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("修改内容出错 " + ex.Message, "操作提示");
                    }

                    //labelFormat.SaveAs(@"E:\正在进行项目\经管中心-标签打印\python\BarTender_Dev_Dome-master\BarTender_Dev_Dome-master\BarTender_Dev_Dome\bin\Debug\moban\工字标\1111.btw", true);

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

                    if (_PrinterName != "")
                    {
                        int 次数 = Convert.ToInt32(textBox1 .Text) ;

                        labelFormat.PrintSetup.PrinterName = _PrinterName;
                        for (int i = 0; i < 次数; i++)
                        {
                            labelFormat.Print("BarPrint" + DateTime.Now,3 * 1000);
                        }
                            
                        //labelFormat.Print("OK");
                        //labelFormat.Print("BarPrint" + DateTime.Now, 3 * 1000);

                    }
                    else
                    {
                        MessageBox.Show("请先选择打印机", "操作提示");
                    }
                }


                
            }
        }

        private void print_btn_Click(object sender, EventArgs e)
        {
            PrintBar();
        }

        private void preview_btn_Click(object sender, EventArgs e)
        {
            PrintBar(true);
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void name_textBox_TextChanged(object sender, EventArgs e)
        {

        }

        //判断产品信息
        private void 判断产品信息(string aa)
        {
            string output = "";
            string output_name = "Name:"+Box_Name .Text ;
            string output_灯带型号 = "";
            string output_电压 = "";
            string output_功率 = "";
            string output_灯数 = "";
            string output_剪切单元 = "";
            string output_长度 = "Length:";
            string output_色温 = "";
            string output_尾巴 = "Made in China";



            // 正则表达式模式，
            string pattern1 = @"^(\w+-\w+-\w+)";
            string pattern2 = @"D(\d+)V";
            string pattern3 = @"额定功率(\d+)W";
            string pattern4 = @"-(\d+)-";
            string pattern5 = @"(\d+)灯\/(\d+\.?\d*)cm";
            string pattern6 = @"-IP(\d+)-";


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
                output_灯带型号 = $"ART. No.: {artNo}";

                // 使用信息框输出结果
                //MessageBox.Show(output1, "提取结果");

                // 检查复选框是否同时被选中
                bool isCustomerNameChecked = checkBox_客户Name.Checked;
                bool isCustomerModelChecked = checkBox_客户型号.Checked;

                string originalString = textBox_客户资料.Text;
                int spaceIndex = originalString.LastIndexOf('	');


                if (isCustomerNameChecked && !isCustomerModelChecked)
                {
                    // 如果只有 checkBox_客户Name 被选中，则输出 1
                    //MessageBox.Show("1");
                    output_name = "Name:" + textBox_客户资料.Text;
                }
                else if (!isCustomerNameChecked && isCustomerModelChecked)
                {
                    // 如果只有 checkBox_客户型号 被选中，则输出 2
                    //MessageBox.Show("2");
                    output_灯带型号 = "ART. No.: " + textBox_客户资料.Text;
                }
                else if (isCustomerNameChecked && isCustomerModelChecked)
                {
                    // 如果两个复选框都被选中，则按照制表符拆分字符串并显示
                    int tabIndex = originalString.LastIndexOf('\t');

                    if (tabIndex != -1)
                    {
                        string part1 = originalString.Substring(0, tabIndex);
                        string part2 = originalString.Substring(tabIndex + 1);
                        //MessageBox.Show(part1 + "\n" + part2);
                        output_name = "Name:" + part1 ;
                        output_灯带型号 = "ART. No.: " + part2 ;

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
                    output_灯带型号 = $"ART. No.: {artNo}";
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

                // 构造输出字符串
                output_电压 = $"Rated Voltage: {voltageValue}V";

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
                output_剪切单元 = $"Min. Cutting Length: {ledQuantity}LEDs\n({length}cm)";
            }
            else
            {
                // 如果没有找到匹配项，则输出错误信息
                MessageBox.Show( "未找到剪切单元信息匹配项","错误");

            }

            // 色温
            if (parts.Length >= 6)
            {
                // 第五个"-"和第六个"-"之间的内容是parts[5]，因为数组索引是从0开始的
                string contentBetweenFifthAndSixth = parts[5];



                // 检查内容是否为纯字母
                if (Regex.IsMatch(contentBetweenFifthAndSixth, @"^[a-zA-Z]*$"))
                {
                    output_色温 = $"Color: {contentBetweenFifthAndSixth}";
                }
                else
                {
                    // 如果包含字母，则提取数字部分
                    string numericValue = Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", "");
                    output_色温 = $"Color: {numericValue}K";
                }

            }
            else
            {
                MessageBox.Show("未找到色温匹配项。", "错误");
            }

            // IP等级
            if (match6.Success)
            {
                // 从匹配结果中提取数字
                string ipNumber = match6.Groups[1].Value; // 第一个捕获组匹配的内容

                // 构造输出字符串
                bq_ipdj  = $"IP{ipNumber}";
            }
            else
            {
                // 如果没有找到匹配项，则输出错误信息
                MessageBox.Show("未找到IP等级匹配项。", "错误");
            }




            //MessageBox.Show(output_name +"\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴, "提取结果");
            name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;

        }



        //链接excel数据
        //private void button_数据库_Click(object sender, EventArgs e)
        //{
        //    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;  //避免许可证错误


        //    OpenFileDialog dialog = new OpenFileDialog();
        //    dialog.Multiselect = false;//多个文件
        //    dialog.Title = "请选择数据库文件";
        //    dialog.Filter = "excel文件(*.xlsx)|*.xlsx|All files (*.*)|*.*";
        //    dialog.InitialDirectory = Application.StartupPath + @"\数据库";

        //    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        //    {
        //        _sjk_path = dialog.FileName;
        //        Box_数据库 .Text = dialog.FileName;
        //        // fileNametBox.Text = dialog.SafeFileName;
        //        Box_数据库.BackColor = System.Drawing.Color.LightGreen;

        //        try
        //        {
        //            // 加载Excel文件
        //            using (var package = new ExcelPackage(new FileInfo(_sjk_path)))
        //            {
        //                // 获取工作表
        //                var worksheet = package.Workbook.Worksheets["Sheet1"];

        //                int row = 2;

        //                while (worksheet.Cells[row, 1].Value != null)
        //                {
        //                    // 读取列单元格的值
        //                    var cellValue1 = worksheet.Cells[row, 1].Value.ToString();
        //                    var cellValue2 = worksheet.Cells[row, 2].Value.ToString();
        //                    var cellValue3 = worksheet.Cells[row, 3].Value.ToString();
        //                    var cellValue4 = worksheet.Cells[row, 4].Value.ToString();



        //                    // 将读取的内容赋值给 TextBox 控件
        //                    MessageBox.Show(cellValue1,"打印序列");  //不能删，删掉会卡死
        //                    textBox_序号.Text = cellValue1;
        //                    TextBox_sku.Text = cellValue2;
        //                    textBox_type.Text = cellValue3;
        //                    textBox_剪切长度.Text = cellValue4;

        //                    // 等待10秒
        //                    System.Threading.Thread.Sleep(60);

        //                    // 移到下一行
        //                    row++;
        //                }



        //            }


        //        }
        //        catch (Exception ex)
        //        {
        //            MessageBox.Show("无法读取Excel文件: " + ex.Message);
        //        }


        //    }
        //}
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

                try
                {
                    // 加载Excel文件
                    using (var package = new ExcelPackage(new FileInfo(_sjk_path)))
                    {
                        var worksheet = package.Workbook.Worksheets["Sheet1"];

                        int row = 2;
                        while (worksheet.Cells[row, 1].Value != null)
                        {
                            var cellValue1 = worksheet.Cells[row, 1].Value.ToString();
                            var cellValue2 = worksheet.Cells[row, 2].Value.ToString();
                            var cellValue3 = worksheet.Cells[row, 3].Value.ToString();
                            var cellValue4 = worksheet.Cells[row, 4].Value.ToString();

                            // 使用Task.Run来在后台线程处理数据读取
                            await Task.Run(() =>
                            {
                                // 模拟数据处理延迟
                                System.Threading.Thread.Sleep(1000);
                            });

                            // 使用Invoke确保UI更新在UI线程上执行
                            this.Invoke((MethodInvoker)delegate
                            {
                                // 将读取的内容赋值给 TextBox 控件
                                textBox_序号.Text = cellValue1;
                                TextBox_sku.Text = cellValue2;
                                textBox_type.Text = cellValue3;
                                textBox_剪切长度.Text = cellValue4;
                            });

                            // 移到下一行
                            row++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("无法读取Excel文件: " + ex.Message);
                }
            }
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
                    LabelFormatDocument labelFormat = btEngine.Documents.Open(模板地址);
                    //LabelFormatDocument labelFormat = btEngine.Documents.Open(@"E:\正在进行项目\经管中心-标签打印\python\BarTender_Dev_Dome-master\BarTender_Dev_Dome-master\BarTender_Dev_Dome\bin\Debug\moban\工字标\tesssst-正弯.btw");
                    try
                    {
                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                        labelFormat.SubStrings.SetSubString("CPCD", TextBox_sku.Text);
                        labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj);


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

        private void button_test_Click_1(object sender, EventArgs e)
        {
            

        }

       
    }

}
