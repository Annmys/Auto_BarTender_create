using IniParser;
using IniParser.Model;
using OfficeOpenXml; // EPPlus的命名空间
using Seagull.BarTender.Print;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Color = System.Drawing.Color;
using Resolution = Seagull.BarTender.Print.Resolution;







namespace BarTender_Dev_Dome

{
    public partial class PrintForm : Form
    {
        string bq_ipdj = null;
        private string _bmp_path = Application.StartupPath + @"\exp.jpg";
        private string _btw_path = string.Empty;
        string 模板地址 = string.Empty;
        string 软件版本 = string.Empty;
        string 灯带系列 = string.Empty;
        string _sjk_path = string.Empty;
        string _PrinterName = string.Empty;
        string _wjm_ = string.Empty;
        string output_name = "Name:LED Flex Linear Light";
        string output_灯带型号 = string.Empty;
        string output_电压 = string.Empty;
        string output_功率 = string.Empty;
        string output_灯数 = string.Empty;
        string output_剪切单元 = string.Empty;
        string output_长度 = "Length:";
        string output_色温 = string.Empty;
        string output_尾巴 = string.Empty;
        string 最小ip等级= string.Empty;
        private string configFilePath;
        private FileIniDataParser parser;
        private IniData data;



        public enum biaoqian // 枚举名称建议使用大写，以符合C#的命名规范，例如：ActionType
        {
            dayin,
            lingcun,
            yulan

        }

        //程序开始运行
        public PrintForm()
        {
            InitializeComponent();
            //MessageBox.Show("程序打开", "操作提示");


            //读取excel文件必须增加声明才能运行正常
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            LoadConfiguration();



        }

        //读取操作记录
        private void LoadConfiguration()
        {
            try
            {
                string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;
                configFilePath = Path.Combine(currentDirectory, "config.ini");
                parser = new FileIniDataParser();

                if (File.Exists(configFilePath))
                {
                    data = parser.ReadFile(configFilePath);
                    if (data.Sections.ContainsSection("Settings"))
                    {
                        textBox_客户资料.Text = data["Settings"]["客户资料"];
                        cpxxBox.Text = data["Settings"]["产品信息"].Replace("\\n", Environment.NewLine); // 处理换行符
                        textBox_剪切长度.Text = data["Settings"]["剪切长度"];
                        textBox_标识码01.Text = data["Settings"]["标识码01"];
                        textBox_标识码02.Text = data["Settings"]["标识码02"];
                        textBox_标识码03.Text = data["Settings"]["标识码03"];
                        textBox_标识码04.Text = data["Settings"]["标识码04"];

                    }
                    else
                    {
                        MessageBox.Show("配置文件中未找到 'Settings' 部分。");
                    }
                }
                else
                {
                    MessageBox.Show("未找到配置文件: " + configFilePath);
                    // 如果文件不存在，则创建一个新的 IniData 对象
                    data = new IniData();
                    data["Settings"]["客户资料"] = "默认客户资料";
                    data["Settings"]["产品信息"] = "默认产品信息";
                    data["Settings"]["剪切长度"] = "默认剪切长度";
                    data["Settings"]["标识码01"] = "默认标识码01";
                    data["Settings"]["标识码02"] = "默认标识码02";
                    data["Settings"]["标识码03"] = "默认标识码03";
                    data["Settings"]["标识码04"] = "默认标识码04";

                    SaveConfiguration(); // 保存默认配置
                }
            }
            catch (IniParser.Exceptions.ParsingException ex)
            {
                MessageBox.Show("加载配置文件时出错: " + ex.Message);
            }
            catch (Exception ex)
            {
                MessageBox.Show("发生未知错误: " + ex.Message);
            }
        }

        private void SaveConfiguration()
        {
            try
            {
                if (data == null)
                {
                    data = new IniData();
                }

                data["Settings"]["客户资料"] = textBox_客户资料.Text;
                data["Settings"]["产品信息"] = cpxxBox.Text.Replace(Environment.NewLine, "\\n"); // 保存换行符
                data["Settings"]["剪切长度"] = textBox_剪切长度.Text;
                data["Settings"]["标识码01"] = textBox_标识码01.Text;
                data["Settings"]["标识码02"] = textBox_标识码02.Text;
                data["Settings"]["标识码03"] = textBox_标识码03.Text;
                data["Settings"]["标识码04"] = textBox_标识码04.Text;
                parser.WriteFile(configFilePath, data);
            }
            catch (Exception ex)
            {
                MessageBox.Show("保存配置文件时出错: " + ex.Message);
            }
        }



        //打印被点击
        private void print_btn_Click(object sender, EventArgs e)
        {
            //PrintBar();
            //if (标签种类_comboBox.Text == "工字标" || 标签种类_comboBox.Text == "品名标") { PrintBar(); }
            if (标签种类_comboBox.Text == "工字标" || 标签种类_comboBox.Text == "品名标") { shengcheng_biaoqian(biaoqian.dayin); }

            if (标签种类_comboBox.Text == "唛头")
            {
                shengcheng_maitou(biaoqian.dayin);
            }
        }

        //预览被点击
        private void preview_btn_Click(object sender, EventArgs e)
        {

            //if (标签种类_comboBox.Text == "工字标" || 标签种类_comboBox.Text == "品名标") { PrintBar(true); }
            if (标签种类_comboBox.Text == "工字标" || 标签种类_comboBox.Text == "品名标") { shengcheng_biaoqian(biaoqian.yulan); }

            if (标签种类_comboBox.Text == "唛头")
            {
                shengcheng_maitou(biaoqian.yulan);
            }
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
            else if (checkBox_客制型号.Checked)
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
            //MessageBox.Show(aa);

            // 拼接标签种类_comboBox.Text 到基础目录
            if (aa == "常规型号")
            {
                templatesDirectory = Path.Combine(baseDirectory, @"\常规型号" + @"\");

                _btw_path_1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\" + aa + @"\";
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
            if (aa == "客制型号")
            {
                templatesDirectory = Path.Combine(baseDirectory, @"\客制型号");
                _btw_path_1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\" + aa + @"\";
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
            _btw_path = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\";
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

            if (标签种类_comboBox.Text == "唛头")
            {
                tabControl_唛头.Visible = true;
                comboBox_唛头规格.Visible = true;
                groupBox_唛头显指.Visible = true;
            }
            else
            {
                tabControl_唛头.Visible = false;
                comboBox_唛头规格.Visible = false;
                groupBox_唛头显指.Visible = false;
            }


        }


        private string 重构产品信息_工字标(string name_CPXXBox, string textBox_剪切长度)
        {  
            int lengthIndex = 0;
            //MessageBox.Show(name_CPXXBox );
            if (comboBox_标签规格.Text.Contains("BIS"))
            {
                lengthIndex = name_CPXXBox.IndexOf("m\nLength:");
                if (lengthIndex != -1)
                {
                    // 调整 lengthIndex 以确保它位于 ")\nLength:" 后面
                    lengthIndex += "m\nLength:".Length;
                }
                else
                {
                    // 如果没有找到 ")\nLength:"，可以决定如何处理这种情况
                    // 例如，可以返回原始的 name_CPXXBox 或者返回一个错误消息
                    return "指定的标识 'm\nLength:' 不存在。";
                }
            }
            else if (comboBox_标签规格.Text.Contains("高压") && comboBox_标签规格.Text.Contains("短剪"))
            {
                lengthIndex = name_CPXXBox.IndexOf("m\nLength:");
                if (lengthIndex != -1)
                {
                    // 调整 lengthIndex 以确保它位于 ")\nLength:" 后面
                    lengthIndex += "m\nLength:".Length;
                }
                else
                {
                    // 如果没有找到 ")\nLength:"，可以决定如何处理这种情况
                    // 例如，可以返回原始的 name_CPXXBox 或者返回一个错误消息
                    return "指定的标识 'm\nLength:' 不存在。";
                }
            }
            else
            {
                lengthIndex = name_CPXXBox.IndexOf(")\nLength:");
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
            }
            // 检查 name_CPXXBox 是否包含特定的标识 ")\nLength:"，并找到它的位置


            // 使用 StringBuilder 创建新的文本
            StringBuilder newText = new StringBuilder(name_CPXXBox);

            // 移除从找到的位置到字符串末尾的所有内容
            newText.Remove(lengthIndex, newText.Length - lengthIndex);

            // 在找到的位置后面插入 textBox_剪切长度 的内容
            newText.Insert(lengthIndex, textBox_剪切长度);

            if (comboBox_标签规格.Text.Contains("高压") && comboBox_标签规格.Text.Contains("短剪")) {
                // 在文本末尾追加其他信息
                newText.Append("\n" + output_色温);
                newText.Append("\n" + " ");
                newText.Append("\n" + output_尾巴);
            }
            else {
                // 在文本末尾追加其他信息
                newText.Append("\n" + output_色温);
                newText.Append("\n" + output_尾巴);
            }



            if (comboBox_标签规格.Text.Contains("BIS"))
            {
                newText.Append("\n" + "Ta:-40 to 55℃");
            }

            // 返回构建好的字符串
            return newText.ToString();
        }




        public void shengcheng_biaoqian(biaoqian actionType)
        {

            // 假设这是从某个文本框获取的字符串
            string cpxx_text = cpxxBox.Text;
            判断产品信息(cpxx_text);


            using (Engine btEngine = new Engine(true))
            {

                if (!_btw_path.Contains(comboBox_标签规格.Text))
                {
                    // _btw_path = _btw_path + comboBox_标签规格.Text; // 如果不包含，则拼接

                    if (checkBox_常规型号.Checked)
                    {
                        _btw_path = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\";
                        _btw_path = _btw_path + @"常规型号\" + comboBox_标签规格.Text;
                    }
                    else if (checkBox_客制型号.Checked)
                    {
                        _btw_path = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\";
                        _btw_path = _btw_path + @"客制型号\" + comboBox_标签规格.Text;
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

                if (checkBox_标识码01.Checked) { _wjm_ = "1.btw"; }
                if (checkBox_标识码02.Checked) { _wjm_ = "1.btw"; }
                if (checkBox_标识码03.Checked) { _wjm_ = "3.btw"; }
                if (checkBox_标识码04.Checked) { _wjm_ = "4.btw"; }


                模板地址 = _btw_path + @"\" + _wjm_;
                //MessageBox.Show(模板地址, "操作提示");
                模板地址 = 模板地址.Replace("\n", string.Empty).Replace("\r", string.Empty);  //去除换行符，否则下面会报错
                                                                                      //MessageBox.Show(_wjm_);


                if (_wjm_.Length > 2)
                {
                    LabelFormatDocument labelFormat = btEngine.Documents.Open(模板地址);

                    try
                    {


                        // 调用方法时，可以这样使用返回的字符串
                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, textBox_剪切长度.Text);
                        //MessageBox.Show(name_CPXXBox.Text);




                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                        //labelFormat.SubStrings.SetSubString("CPCD", textBox_剪切长度.Text);
                        labelFormat.SubStrings.SetSubString("CPCD", " ");
                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                        labelFormat.SubStrings.SetSubString("XLH", " ");


                        if (comboBox_标签规格.Text.Contains("BIS"))
                        {
                            labelFormat.SubStrings.SetSubString("IPDJ", "IP68 2m");
                        }
                        else if (comboBox_标签规格.Text.Contains("水下"))
                        {
                            labelFormat.SubStrings.SetSubString("IPDJ", "IP68 5m");
                        }

                        else
                        {
                            labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj);
                        }


                        //高压情况
                        if (comboBox_标签规格.Text.Contains("高压"))
                        {
                            double.TryParse(textBox_剪切长度.Text, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString("F2"));
                            double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString("F3") + "A");
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
                        if (checkBox_标识码03.Checked)
                        {
                            labelFormat.SubStrings.SetSubString("BSM-03", textBox_标识码03.Text);
                        }

                        if (checkBox_标识码04.Checked)
                        {
                            labelFormat.SubStrings.SetSubString("BSM-04", textBox_标识码04.Text);
                        }


                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                        // 检查数据库地址不为空时
                        if (!string.IsNullOrEmpty(Box_数据库.Text))
                        {

                            string b2Data, c2Data, g2Data, h2Data, h1Data, a2Data, d2Data, e2Data, f2Data, i2Data;

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
                                h1Data = worksheet.Cells["H1"].Value?.ToString() ?? string.Empty;
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
                                        double.TryParse(h2Data, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString("F2"));
                                        double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString("F3") + "A");
                                        if (h1Data == "英尺长度")
                                        {
                                        }

                                    }

                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                    labelFormat.SubStrings.SetSubString("CPCD", " ");

                                    //客户型号被选择时
                                    if (checkBox_客户型号.Checked)
                                    {
                                        //output_灯带型号 = "ART. No.: " + i2Data;
                                        if (标签种类_comboBox.Text == "工字标")
                                        {
                                            int ai = i2Data.Length;
                                            if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                            else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                            else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                        }
                                        else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + i2Data; }

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
                                        double.TryParse(h2Data, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString("F2"));
                                        double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString("F3") + "A");
                                    }

                                    labelFormat.SubStrings.SetSubString("CPCD", " ");

                                    //客户型号被选择时
                                    if (checkBox_客户型号.Checked)
                                    {
                                        //output_灯带型号 = "ART. No.: " + i2Data;
                                        if (标签种类_comboBox.Text == "工字标")
                                        {
                                            int ai = i2Data.Length;
                                            if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                            else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                            else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                        }
                                        else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + i2Data; }


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
                                        double.TryParse(h2Data, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString("F2"));
                                        double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString("F3") + "A");
                                    }

                                    labelFormat.SubStrings.SetSubString("CPCD", " ");

                                    //客户型号被选择时
                                    if (checkBox_客户型号.Checked)
                                    {
                                        //output_灯带型号 = "ART. No.: " + i2Data;
                                        if (标签种类_comboBox.Text == "工字标")
                                        {
                                            int ai = i2Data.Length;
                                            if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                            else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                            else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                        }
                                        else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + i2Data; }


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
                                        if (标签种类_comboBox.Text == "工字标")
                                        {
                                            int ai = i2Data.Length;
                                            if (ai <= 19) { output_灯带型号 = "ART. No.: " + i2Data; }
                                            else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + i2Data; }
                                            else { output_灯带型号 = "ART. No.: " + i2Data.Substring(0, 19) + Environment.NewLine + i2Data.Substring(19); }
                                        }
                                        else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + i2Data; }

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
                        labelFormat.SubStrings.SetSubString("XZ", string.Empty);

                        string BPrefixContent = string.Empty; // 用于存储 "B-" 前面的内容


                        string text = cpxxBox.Text;
                        string pattern = @"(?<=^|\r\n)(C-(SFR|FR|SFB)-.*)(?=\r\n)";
                        Regex regex = new Regex(pattern);
                        MatchCollection matches = regex.Matches(text);
                        foreach (Match match in matches)
                        {
                            if (match.Success) // 确保找到了匹配项
                            {
                                BPrefixContent = match.Groups[1].Value.Trim();
                                //MessageBox.Show(BPrefixContent);
                            }
                        }


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
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
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
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
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
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
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
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
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
                                string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
                                labelFormat.SubStrings.SetSubString("XZ-2", raValue);
                            }
                        }
                        else if (BPrefixContent.Contains("Ra90"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ", "Ra90");
                            labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");

                            if (comboBox_标签规格.Text.Contains("13013"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", " ");
                                labelFormat.SubStrings.SetSubString("FXK-2", "空.png");
                            }

                        }
                        else if (BPrefixContent.Contains("Ra95"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ", "Ra95");
                            labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");

                            if (comboBox_标签规格.Text.Contains("13013"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", " ");
                                labelFormat.SubStrings.SetSubString("FXK-2", "空.png");
                            }
                        }
                        else
                        {
                            // 如果没有找到上述任何关键字，则设置为空
                            labelFormat.SubStrings.SetSubString("XZ", string.Empty);
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-2", "空.png");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                        }

                        if (comboBox_标签规格.Text.Contains("13013"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ-2", " ");
                            labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                        }


                        //检查是否为非水下内容
                        // 检查 cpxxBox.Text 中是否同时包含 "IP68" 和 "非水下"
                        bool containsIP68 = bq_ipdj.Contains("IP68");
                        bool containsNonUnderwater = cpxxBox.Text.Contains("非水下");




                        if (标签种类_comboBox.Text == "品名标")
                        {
                            // 检查cpxxBox文本中是否包含"非水下方案"
                            if (containsIP68 && cpxxBox.Text.Contains("非水下方案"))
                            {
                                // 如果包含"非水下方案"，则设置labelFormat的"SX"子字符串为"Not suitable for underwater use"
                                labelFormat.SubStrings.SetSubString("SX", "Not suitable for underwater use");
                            }
                            else if (!containsIP68)
                            {
                                //MessageBox.Show("没有IP68");
                                labelFormat.SubStrings.SetSubString("SX", " ");
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
                                //MessageBox.Show(bq_ipdj + "\n" + containsIP68 + "\n" + containsNonUnderwater.ToString());

                                if (containsIP68 && containsNonUnderwater)
                                {
                                    labelFormat.SubStrings.SetSubString("SX", "非水下.png");
                                }
                                else
                                {
                                    labelFormat.SubStrings.SetSubString("SX", "正常.png");
                                }

                            }

                        }
                    }

                    catch (Exception ex)
                    {
                        MessageBox.Show("修改内容出错 " + ex.Message, "操作提示");
                    }






                    //执行输出判断
                    switch (actionType)
                    {
                        //生成预览图
                        case biaoqian.yulan:

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



                            break;


                        //另存为
                        case biaoqian.lingcun:
                            SaveFileDialog dialog = new SaveFileDialog();
                            dialog.Title = "请选择要保存的文件地址";
                            dialog.Filter = "bwt文件(*.btw)|*.btw";
                            dialog.InitialDirectory = Application.StartupPath + @"\输出文件";
                            dialog.DefaultExt = "btw"; // 设置默认文件扩展名
                            dialog.AddExtension = true; // 确保即使用户未指定扩展名也会添加扩展

                            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                // 获取用户指定的文件路径
                                string saveFilePath = dialog.FileName;
                                //MessageBox.Show(saveFilePath, "操作提示");
                                labelFormat.SaveAs(saveFilePath, true);
                                提示框.AppendText("文件另存完成" + Environment.NewLine);
                            }



                            break;


                        //打印标签
                        case biaoqian.dayin:

                            if (_PrinterName == string.Empty)
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
                                                    double.TryParse(hData, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString("F2"));
                                                    double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString("F3") + "A");
                                                }

                                                labelFormat.SubStrings.SetSubString("CPCD", " ");

                                                //客户型号被选择时
                                                if (checkBox_客户型号.Checked)
                                                {
                                                    //output_灯带型号 = "ART. No.: " + iData;
                                                    if (标签种类_comboBox.Text == "工字标")
                                                    {
                                                        int ai = iData.Length;
                                                        if (ai <= 19) { output_灯带型号 = "ART. No.: " + iData; }
                                                        else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + iData; }
                                                        else { output_灯带型号 = "ART. No.: " + iData.Substring(0, 19) + Environment.NewLine + iData.Substring(19); }
                                                    }
                                                    else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + iData; }

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
                                                    double.TryParse(hData, out double length); double result = length * 3.28; labelFormat.SubStrings.SetSubString("CPCD-2", result.ToString("F2"));
                                                    double washu = length * 10; double anpai = length * 0.093; labelFormat.SubStrings.SetSubString("CPXX-2", washu.ToString() + "W," + anpai.ToString("F3") + "A");
                                                }

                                                labelFormat.SubStrings.SetSubString("CPCD", " ");

                                                //客户型号被选择时
                                                if (checkBox_客户型号.Checked)
                                                {
                                                    //output_灯带型号 = "ART. No.: " + iData;
                                                    if (标签种类_comboBox.Text == "工字标")
                                                    {
                                                        int ai = iData.Length;
                                                        if (ai <= 19) { output_灯带型号 = "ART. No.: " + iData; }
                                                        else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + iData; }
                                                        else { output_灯带型号 = "ART. No.: " + iData.Substring(0, 19) + Environment.NewLine + iData.Substring(19); }
                                                    }
                                                    else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + iData; }

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
                                                    if (标签种类_comboBox.Text == "工字标")
                                                    {
                                                        int ai = iData.Length;
                                                        if (ai <= 19) { output_灯带型号 = "ART. No.: " + iData; }
                                                        else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + iData; }
                                                        else { output_灯带型号 = "ART. No.: " + iData.Substring(0, 19) + Environment.NewLine + iData.Substring(19); }
                                                    }
                                                    else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + iData; }


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
                                                    if (标签种类_comboBox.Text == "工字标")
                                                    {
                                                        int ai = iData.Length;
                                                        if (ai <= 19) { output_灯带型号 = "ART. No.: " + iData; }
                                                        else if (ai > 19 && ai <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + iData; }
                                                        else { output_灯带型号 = "ART. No.: " + iData.Substring(0, 19) + Environment.NewLine + iData.Substring(19); }
                                                    }
                                                    else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + iData; }


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





                            break;


                        default:
                            // 如果传入的 actionType 不是预期值，抛出异常或处理错误
                            throw new ArgumentOutOfRangeException(nameof(actionType), $"未知的操作类型: {actionType}");
                    }


                }




            }






        }

        public void shengcheng_maitou(biaoqian actionType) // 方法名称建议使用大写开头，例如：Test
        {


            using (Engine btEngine = new Engine(true))
            {


                if (!_btw_path.Contains(comboBox_标签规格.Text))
                {
                    string basePath = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\";
                    string modelTypePath = checkBox_常规型号.Checked ? "常规型号" : checkBox_客制型号.Checked ? "客制型号" : null;

                    if (modelTypePath != null)
                    {
                        _btw_path = basePath + modelTypePath + @"\" + comboBox_标签规格.Text;
                    }
                }


                _wjm_ = comboBox_唛头规格.Text + ".btw";

                //寻找文件名_单字匹配("正弯", "侧弯");
                string 复选框 = 判断复选框内容(cpxxBox.Text, comboBox_标签规格.Text);

                模板地址 = _btw_path + @"\" + _wjm_;
                //MessageBox.Show(模板地址, "操作提示");
                模板地址 = 模板地址.Replace("\n", string.Empty).Replace("\r", string.Empty);  //去除换行符，否则下面会报错


                if (_wjm_.Length > 2)
                {
                    //先判断文本内容
                    if (checkBox_唛头型号自动.Checked) { 唛头_寻找灯带型号(); }
                    if (checkBox_唛头电压自动.Checked) { 唛头_寻找灯带电压(); }
                    if (checkBox_唛头色温自动.Checked) { 唛头_寻找色温(); }
                    唛头_寻找订单编号();

                    //加载btw
                    LabelFormatDocument labelFormat = btEngine.Documents.Open(模板地址);

                    //name_CPXXBox.Text = mt.正常型号判断(cpxxBox.Text, checkBox_客户Name.Checked, checkBox_客户型号.Checked, textBox_客户资料.Text, comboBox_标签规格.Text, 标签种类_comboBox.Text, textBox_唛头数量.Text, textBox_唛头尺寸.Text);
                    //labelFormat.SubStrings.SetSubString("CPXX-02", name_CPXXBox.Text );
                    //labelFormat.SubStrings.SetSubString("CPCD", textBox_剪切长度.Text);


                    //PO号
                    if (checkBox_po号.Checked) { labelFormat.SubStrings.SetSubString("PO-01", "PO :    " + textBox_po号.Text); }
                    else { labelFormat.SubStrings.SetSubString("PO-01", "  "); }

                    //判断是否增加标识码
                    if (checkBox_标识码01.Checked) { labelFormat.SubStrings.SetSubString("BSM-01", textBox_标识码01.Text); }
                    else { labelFormat.SubStrings.SetSubString("BSM-01", " "); }
                    if (checkBox_标识码02.Checked) { labelFormat.SubStrings.SetSubString("BSM-02", textBox_标识码02.Text); }
                    else { labelFormat.SubStrings.SetSubString("BSM-02", " "); }


                    //唛头_产品信息
                    StringBuilder 唛头_产品信息 = new StringBuilder();
                    唛头_产品信息.AppendLine("      NAME:    " + textBox_唛头名称.Text);

                    //唛头_产品型号
                    if (textBox_唛头灯带型号.Text.Length <= 27)
                    {
                        唛头_产品信息.AppendLine("ART. NO.:    " + textBox_唛头灯带型号.Text);
                    }
                    else if (textBox_唛头灯带型号.Text.Length > 27 && textBox_唛头灯带型号.Text.Length <= 54)
                    {
                        string firstPart = textBox_唛头灯带型号.Text.Substring(0, 27);
                        int commaIndex = firstPart.LastIndexOf(',');
                        if (commaIndex != -1) // 如果找到了逗号
                        {
                            唛头_产品信息.AppendLine("ART. NO.:    " + textBox_唛头灯带型号.Text.Substring(0, commaIndex + 1).Trim());
                            唛头_产品信息.AppendLine("                     " + textBox_唛头灯带型号.Text.Substring(commaIndex + 1));
                        }
                    }
                    else if (textBox_唛头灯带型号.Text.Length > 54 && textBox_唛头灯带型号.Text.Length <= 81)
                    {
                        int currentPosition = 0;
                        int chunkSize = 28; // 每次处理的字符长度
                        bool isFirstLine = true; // 标记是否是第一行

                        while (currentPosition < textBox_唛头灯带型号.Text.Length)
                        {
                            // 截取当前位置开始的chunkSize个字符或剩余的所有字符
                            string chunk = textBox_唛头灯带型号.Text.Substring(currentPosition, Math.Min(chunkSize, textBox_唛头灯带型号.Text.Length - currentPosition));

                            // 查找chunk中最后一个逗号的位置
                            int commaIndex = chunk.LastIndexOf(',');

                            if (commaIndex != -1)
                            {
                                // 截取从当前位置到逗号加1的文本
                                string lineToAppend = textBox_唛头灯带型号.Text.Substring(currentPosition, commaIndex + 1).Trim();

                                // 根据是否是第一行选择前缀
                                string prefix = isFirstLine ? "ART. NO.:    " : "                     ";

                                // 添加到产品信息
                                唛头_产品信息.AppendLine(prefix + lineToAppend);

                                // 更新当前位置为找到的逗号之后的位置
                                currentPosition += commaIndex + 1;

                                // 从第二行开始，isFirstLine设置为false
                                isFirstLine = false;
                            }
                            else
                            {
                                // 如果没有找到逗号，输出当前chunk，并结束循环
                                if (isFirstLine)
                                {
                                    唛头_产品信息.AppendLine("ART. NO.:    " + chunk.Trim());
                                }
                                else
                                {
                                    唛头_产品信息.AppendLine("                     " + chunk.Trim());
                                }
                                break;
                            }
                        }
                    }

                    唛头_产品信息.AppendLine("VOLTAGE:   " + textBox_唛头电压.Text);


                    //唛头_产品信息.AppendLine("         QTY.:   "+textBox_唛头数量.Text);
                    //唛头_QTY
                    if (textBox_唛头数量.Text.Length <= 27)
                    {
                        唛头_产品信息.AppendLine("         QTY.:   " + textBox_唛头数量.Text);
                    }
                    else if (textBox_唛头数量.Text.Length > 27 && textBox_唛头数量.Text.Length <= 54)
                    {
                        string firstPart = textBox_唛头数量.Text.Substring(0, 27);
                        int commaIndex = firstPart.LastIndexOf(',');
                        if (commaIndex != -1) // 如果找到了逗号
                        {
                            唛头_产品信息.AppendLine("         QTY.:   " + textBox_唛头数量.Text.Substring(0, commaIndex + 1).Trim());
                            唛头_产品信息.AppendLine("                     " + textBox_唛头数量.Text.Substring(commaIndex + 1));
                        }
                    }
                    else if (textBox_唛头数量.Text.Length > 54 && textBox_唛头数量.Text.Length <= 81)
                    {
                        int currentPosition = 0;
                        int chunkSize = 28; // 每次处理的字符长度
                        bool isFirstLine = true; // 标记是否是第一行

                        while (currentPosition < textBox_唛头数量.Text.Length)
                        {
                            // 截取当前位置开始的chunkSize个字符或剩余的所有字符
                            string chunk = textBox_唛头数量.Text.Substring(currentPosition, Math.Min(chunkSize, textBox_唛头数量.Text.Length - currentPosition));

                            // 查找chunk中最后一个逗号的位置
                            int commaIndex = chunk.LastIndexOf(',');

                            if (commaIndex != -1)
                            {
                                // 截取从当前位置到逗号加1的文本
                                string lineToAppend = textBox_唛头数量.Text.Substring(currentPosition, commaIndex + 1).Trim();

                                // 根据是否是第一行选择前缀
                                string prefix = isFirstLine ? "         QTY.:   " : "                     ";

                                // 添加到产品信息
                                唛头_产品信息.AppendLine(prefix + lineToAppend);

                                // 更新当前位置为找到的逗号之后的位置
                                currentPosition += commaIndex + 1;

                                // 从第二行开始，isFirstLine设置为false
                                isFirstLine = false;
                            }
                            else
                            {
                                // 如果没有找到逗号，输出当前chunk，并结束循环
                                if (isFirstLine)
                                {
                                    唛头_产品信息.AppendLine("         QTY.:   " + chunk.Trim());
                                }
                                else
                                {
                                    唛头_产品信息.AppendLine("                     " + chunk.Trim());
                                }
                                break;
                            }
                        }
                    }


                    唛头_产品信息.AppendLine("MEASURE:   " + textBox_唛头尺寸.Text);
                    唛头_产品信息.AppendLine("    COLOR :   " + textBox_唛头色温.Text);
                    labelFormat.SubStrings.SetSubString("CPXX-01", 唛头_产品信息.ToString());

                    labelFormat.SubStrings.SetSubString("DDBH", textBox_订单编号.Text);
                    //labelFormat.SubStrings.SetSubString("CPXX-01", " ");

                    //判断尾巴Made in China
                    if (checkBox_结尾.Checked) { labelFormat.SubStrings.SetSubString("ZGZZ", textBox_结尾.Text); }
                    else { output_尾巴 = " "; }


                    //显指判断
                    if (checkBox_FXK01.Checked)
                    {
                        labelFormat.SubStrings.SetSubString("FXK-01", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-01", textBox_XZ01.Text);
                    }
                    else
                    {
                        labelFormat.SubStrings.SetSubString("FXK-01", "空2.png");
                        labelFormat.SubStrings.SetSubString("XZ-01", textBox_XZ01.Text);
                    }
                    if (checkBox_FXK02.Checked)
                    {
                        labelFormat.SubStrings.SetSubString("FXK-02", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-02", textBox_XZ02.Text);
                    }
                    else
                    {
                        labelFormat.SubStrings.SetSubString("FXK-02", "空2.png");
                        labelFormat.SubStrings.SetSubString("XZ-02", textBox_XZ02.Text);
                    }
                    if (checkBox_FXK03.Checked)
                    {
                        labelFormat.SubStrings.SetSubString("FXK-03", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", textBox_XZ03.Text);
                    }
                    else
                    {
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", " ");
                    }
                    if (checkBox_FXK04.Checked)
                    {
                        labelFormat.SubStrings.SetSubString("FXK-04", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", textBox_XZ04.Text);
                    }
                    else
                    {
                        labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", " ");
                    }
                    if (checkBox_FXK05.Checked)
                    {
                        labelFormat.SubStrings.SetSubString("FXK-05", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", textBox_XZ05.Text);
                    }
                    else
                    {
                        labelFormat.SubStrings.SetSubString("FXK-05", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", " ");
                    }







                    switch (actionType)
                    {
                        //生成预览图
                        case biaoqian.yulan:

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

                            break;


                        //另存为
                        case biaoqian.lingcun:
                            SaveFileDialog dialog = new SaveFileDialog();
                            dialog.Title = "请选择要保存的文件地址";
                            dialog.Filter = "bwt文件(*.btw)|*.btw";
                            dialog.InitialDirectory = Application.StartupPath + @"\输出文件";
                            dialog.DefaultExt = "btw"; // 设置默认文件扩展名
                            dialog.AddExtension = true; // 确保即使用户未指定扩展名也会添加扩展

                            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                // 获取用户指定的文件路径
                                string saveFilePath = dialog.FileName;
                                //MessageBox.Show(saveFilePath, "操作提示");
                                labelFormat.SaveAs(saveFilePath, true);
                                提示框.AppendText("文件另存完成" + Environment.NewLine);
                            }



                            break;


                        //打印标签
                        case biaoqian.dayin:
                            // 确保打印机已选择
                            if (_PrinterName == string.Empty)
                            {
                                MessageBox.Show("请先选择打印机。");
                                return;
                            }

                            // 检查labelFormat是否已初始化
                            if (labelFormat != null)
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
                            break;


                        default:
                            // 如果传入的 actionType 不是预期值，抛出异常或处理错误
                            throw new ArgumentOutOfRangeException(nameof(actionType), $"未知的操作类型: {actionType}");
                    }


                }


            }




        }


        //判断复选框内容
        private static string 判断复选框内容(string input, string 标签规格)
        {

            bool hasConstantCurrent = input.Contains("恒流");
            bool hasConstantVoltage = !hasConstantCurrent; // 如果没有恒流，则默认为恒压

            string firstField = hasConstantCurrent ? "恒流" : "恒压";

            // 检查是否有正弯或侧弯
            bool hasPositiveBend = input.Contains("正弯");
            bool hasSideBend = input.Contains("侧弯");

            string secondField = hasPositiveBend ? "正弯" : (hasSideBend ? "侧弯" : string.Empty);
            //MessageBox.Show(secondField, "操作提示");



            // 构建结果
            if (!string.IsNullOrEmpty(secondField))
            {
                return $"{firstField}-{secondField}.png";
            }
            else if (标签规格.Contains("RCM") || 标签规格.Contains("13013"))
            {
                // 如果包含"RCM"，检查output_灯带型号的文本
                var models = new[] { "F10", "F11", "F15", "F21", "F2222" }; // 侧弯型号
                var models1 = new[] { "F16", "F2219" }; // 正弯型号
                bool isSideBend = models.Any(model => input.Contains(model));
                bool isSideBend1 = models1.Any(model => input.Contains(model));
                //bool isSideBend1 = models1.Any(model => input.Contains(model));
                // 如果是特定的型号，则默认为侧弯
                if (isSideBend)
                {
                    secondField = "侧弯";
                }
                else if (isSideBend1)
                {
                    secondField = "正弯"; // 其他的都是正弯
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
                return $"{firstField}.png";
            }



        }




        //判断产品信息
        private void 判断产品信息(string aa)
        {
            string model = string.Empty;
            string powerValue = string.Empty;
            string voltageValue = string.Empty;
            string ZZ = string.Empty;
            string ledQtyValu= string.Empty;

            // 定义数据
            var data_bis = new List<(string Model, string Power, string Voltage, string LEDsPerMeter, string MaxPower)>
            {
              ("F22S", "22W", "DC 24V", "84LEDs/m", "Max.240W"),
              ("F15A", "12W", "DC 24V", "60LEDs/m", "Max.180W"),
              ("F21A", "12W", "DC 24V", "60LEDs/m", "Max.180W"),
              ("F22A", "12W", "DC 24V", "84LEDs/m", "Max.180W"),
              ("F10B", "4.5W", "DC 24V", "72LEDs/m", "Max.135W"),
              ("F23B", "4.5W", "DC 24V", "144LEDs/m", "Max.135W"),
              ("F15B", "12W", "DC 24V", "72LEDs/m", "Max.180W"),
              ("F21B", "12W", "DC 24V", "72LEDs/m", "Max.180W"),
              ("F22B", "12W", "DC 24V", "108LEDs/m", "Max.180W"),
              ("F22E", "15W", "DC 24V", "84LEDs/m", "Max.120W"),
              ("F15E", "15W", "DC 24V", "60LEDs/m", "Max.240W"),
              ("F22D", "12W", "DC 24V", "144LEDs/m", "Max.180W"),
              ("F22B", "15W", "DC 24V", "108LEDs/m", "Max.150W"),
              ("F22B", "6W", "DC 24V", "108LEDs/m", "Max.180W"),
              ("F15B", "6W", "DC 24V", "72LEDs/m", "Max.360W"),
              ("F16B", "12W", "DC 24V", "72LEDs/m", "Max.360W"),
              ("F23B", "6W", "DC 24V", "144LEDs/m", "Max.120W"),
              ("F15S", "15W", "DC 24V", "56LEDs/m", "Max.450W"),
              ("F16E", "15W", "DC 24V", "60LEDs/m", "Max.240W"),
              ("F16A", "12W", "DC 24V", "60LEDs/m", "Max.240W")
            };







            // 正则表达式模式，
            string pattern1 = @"^(\w+-\w+-\w+)";
            string pattern2 = @"D(\d+)V";
            string pattern21 = @"AC(\d+)V";
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
            Match match21 = Regex.Match(aa, pattern21);
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
                    if (标签种类_comboBox.Text == "工字标") {
                        if (textLength <= 19) { output_name = "Name: " + textBox_客户资料.Text; }
                        else if (textLength > 19 && textLength <= 27) { output_name = "Name: " + Environment.NewLine + textBox_客户资料.Text; }
                        else { output_name = "Name: " + textBox_客户资料.Text.Substring(0, 19) + Environment.NewLine + textBox_客户资料.Text.Substring(19); }
                    }
                    else if (标签种类_comboBox.Text == "品名标") {output_name = "Name:" + textBox_客户资料.Text;}


                }
                else if (!isCustomerNameChecked && isCustomerModelChecked)
                {
                    // 如果只有 checkBox_客户型号 被选中，则输出 2
                    //MessageBox.Show("2");
                    //output_灯带型号 = "ART. No.: " + textBox_客户资料.Text;   
                    if (标签种类_comboBox.Text == "工字标")
                    {
                        if (comboBox_标签规格.Text.Contains("高压") && comboBox_标签规格.Text.Contains("短剪")) { output_灯带型号 = "ART. No.: " + "\n" + artNo; }
                        else if (textLength <= 18) { output_灯带型号 = "ART. No.: " + textBox_客户资料.Text; }
                        else if (textLength > 18 && textLength <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + textBox_客户资料.Text; }
                        else { output_灯带型号 = "ART. No.: " + textBox_客户资料.Text.Substring(0, 18) + Environment.NewLine + textBox_客户资料.Text.Substring(18); }

                        
                        
                    }
                    else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + textBox_客户资料.Text; }



                    //客制标签规格
                    if (cpxxBox.Text.Contains("过温保护")) { output_name = "Name: LED Flex Linear Light with" + Environment.NewLine + " Overheat Protection"; }
                    else if (comboBox_标签规格.Text.Contains("17034")) { output_name = "Name: " + "\n" + "Architectural Outdoor Led Flex"; }
                    else if (comboBox_标签规格.Text.Contains("高温高湿")) { output_name = "Name: " + "LED Flex Linear Light" + "\n" + "for Sauna & Steam Rooms"; }
                    else if (cpxxBox.Text.Contains("3D") && cpxxBox.Text.Contains("地埋")) { output_name = "Name: Free Bend In-ground Light"; }
                    else if (cpxxBox.Text.Contains("地埋") && cpxxBox.Text.Contains("铝合金版")) { output_name = "Name: " + "In-ground Light （Aluminum" + "\n" + "alloy version）"; }
                    else if (cpxxBox.Text.Contains("地埋") && cpxxBox.Text.Contains("不锈钢版") && cpxxBox.Text.Contains("2mm厚")) { output_name = "Name: " + "In-ground Light （" + "\n" + "Strainless steel version, T2）"; }
                    else if (cpxxBox.Text.Contains("地埋") && cpxxBox.Text.Contains("不锈钢版") && cpxxBox.Text.Contains("3mm厚")) { output_name = "Name: " + "In-ground Light （" + "\n" + "Strainless steel version, T3）"; }
                    else if (cpxxBox.Text.Contains("3D") && cpxxBox.Text.Contains("洗墙灯") && cpxxBox.Text.Contains("W3525")) { output_name = "Name: " + "Free Bend Wall Washer"; }
                    else if (cpxxBox.Text.Contains("3D") && cpxxBox.Text.Contains("A1617") && cpxxBox.Text.Contains("A2012")) { output_name = "Name: " + "Free Bend Linear Light"; }
                    else if (cpxxBox.Text.Contains("超长灯")) { output_name = "Name: " + "Ultra-long LED Light "; }
                    else if (cpxxBox.Text.Contains("R30（360°灯）")) { output_name = "Name: " + "360 Neon Flex"; }
                    else if (cpxxBox.Text.Contains("D2230悬吊灯")) { output_name = "Name: " + "Suspended LED Flex Linear " + "\n" + "Light"; }
                    else if (comboBox_标签规格.Text.Contains("14098")) { output_name = "Name: LED Flex Linear Light" + "\n" + "Part code: " + textBox_客户资料.Text; }
                    else { output_name = "Name: LED Flex Linear Light"; }

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
                        if (标签种类_comboBox.Text == "工字标")
                        {
                            if (a1 <= 19) { output_name = "Name: " + part1; }
                            else if (a1 > 19 && a1 <= 27) { output_name = "Name: " + Environment.NewLine + part1; }
                            else { output_name = "Name: " + part1.Substring(0, 19) + Environment.NewLine + part1.Substring(19); }

                            if (comboBox_标签规格.Text.Contains("高压") && comboBox_标签规格.Text.Contains("短剪")) { output_灯带型号 = "ART. No.: " + "\n" + artNo; }
                            else if (a2 <= 18) { output_灯带型号 = "ART. No.: " + part2; }
                            else if (a2 > 18 && a2 <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + part2; }
                            else { output_灯带型号 = "ART. No.: " + part2.Substring(0, 18) + Environment.NewLine + part2.Substring(18); }
                        }
                        else if (标签种类_comboBox.Text == "品名标") {
                            output_name = "Name: " + part1;
                            output_灯带型号 = "ART. No.: " + part2;
                        }



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
                    //客制标签规格
                    if (cpxxBox.Text.Contains("过温保护")) { output_name = "Name: LED Flex Linear Light with" + Environment.NewLine + " Overheat Protection"; }
                    else if (comboBox_标签规格.Text.Contains("17034")) { output_name = "Name: " + "\n" + "Architectural Outdoor Led Flex"; }
                    else if (comboBox_标签规格.Text.Contains("高温高湿")) { output_name = "Name: " + "LED Flex Linear Light" + "\n" + "for Sauna & Steam Rooms"; }
                    else if (cpxxBox.Text.Contains("3D") && cpxxBox.Text.Contains("地埋")) { output_name = "Name: Free Bend In-ground Light"; }
                    else if (cpxxBox.Text.Contains("地埋") && cpxxBox.Text.Contains("铝合金版")) { output_name = "Name: " + "In-ground Light （Aluminum" + "\n" + "alloy version）"; }
                    else if (cpxxBox.Text.Contains("地埋") && cpxxBox.Text.Contains("不锈钢版") && cpxxBox.Text.Contains("2mm厚")) { output_name = "Name: " + "In-ground Light （" + "\n" + "Strainless steel version, T2）"; }
                    else if (cpxxBox.Text.Contains("地埋") && cpxxBox.Text.Contains("不锈钢版") && cpxxBox.Text.Contains("3mm厚")) { output_name = "Name: " + "In-ground Light （" + "\n" + "Strainless steel version, T3）"; }
                    else if (cpxxBox.Text.Contains("3D") && cpxxBox.Text.Contains("洗墙灯") && cpxxBox.Text.Contains("W3525")) { output_name = "Name: " + "Free Bend Wall Washer"; }
                    else if (cpxxBox.Text.Contains("3D") && cpxxBox.Text.Contains("A1617") && cpxxBox.Text.Contains("A2012")) { output_name = "Name: " + "Free Bend Linear Light"; }
                    else if (cpxxBox.Text.Contains("超长灯")) { output_name = "Name: " + "Ultra-long LED Light "; }
                    else if (cpxxBox.Text.Contains("R30（360°灯）")) { output_name = "Name: " + "360 Neon Flex"; }
                    else if (cpxxBox.Text.Contains("D2230悬吊灯")) { output_name = "Name: " + "Suspended LED Flex Linear " + "\n" + "Light"; }
                    else { output_name = "Name: LED Flex Linear Light"; }

                    //output_灯带型号 = $"ART. No.: {artNo}";
                    

                    if (标签种类_comboBox.Text == "工字标")
                    {
                        int a4 = artNo.Length;
                        if (comboBox_标签规格.Text.Contains("高压") && comboBox_标签规格.Text.Contains("短剪")) { output_灯带型号 = "ART. No.: " + "\n" + artNo; }
                        else if (a4 <= 19) { output_灯带型号 = "ART. No.: " + artNo; }
                        else if (a4 > 19 && a4 <= 27) { output_灯带型号 = "ART. No.: " + Environment.NewLine + artNo; }
                        else { output_灯带型号 = "ART. No.: " + artNo.Substring(0, 19) + Environment.NewLine + artNo.Substring(19); }
                    }
                    else if (标签种类_comboBox.Text == "品名标") { output_灯带型号 = "ART. No.: " + artNo; }

                }

                if (comboBox_标签规格.Text.Contains("14098")) { output_灯带型号 = "ART. No.: " + artNo; }


                string lightModel = artNo;
                // 获取最后一个字符
                char lastChar = lightModel[lightModel.Length - 1];
                // 是否为字母
                if (char.IsLetter(lastChar))
                {
                    灯带系列 = lastChar.ToString();
                    //MessageBox.Show(灯带系列);
                }
                //获取最后4个字符
                model = lightModel.Length >= 4 ? lightModel.Substring(lightModel.Length - 4) : lightModel;





            }
            else
            {
                // 如果没有找到匹配项，则输出错误信息
                //MessageBox.Show("未找到灯带型号匹配项。", "错误");
            }

            //电压
            if (match2.Success)
            {
                // 从匹配结果中提取电压值
                voltageValue = match2.Groups[1].Value; // 第一个捕获组匹配的内容

                    // 如果不包含“高压”，则设置 output_电压 为 DC
                    output_电压 = $"Rated Voltage: DC {voltageValue}V";

            }
            else if (match21.Success)
            {
                // 从匹配结果中提取电压值
                voltageValue = match21.Groups[1].Value; // 第一个捕获组匹配的内容

                    // 如果包含“高压”，则设置 output_电压 为 AC
                    output_电压 = $"Rated Voltage: AC {voltageValue}V";

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
                powerValue = match3.Groups[1].Value; // 第一个捕获组匹配的内容

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
                ledQtyValu = match4.Groups[1].Value; // 第一个捕获组匹配的内容

                // 构造输出字符串
                output_灯数 = $"LED Qty.: {ledQtyValu}LEDs/m";
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
                MessageBox.Show("未找到剪切单元信息匹配项", "错误");

            }

            // 色温
            if (parts.Length >= 6)
            {
                string numericValue;
                // 第五个"-"和第六个"-"之间的内容是parts[5]，因为数组索引是从0开始的
                string contentBetweenFifthAndSixth = parts[5];

                char firstLetter = contentBetweenFifthAndSixth[0];
                if (char.IsLetter(firstLetter))
                {
                    ZZ = firstLetter.ToString();

                }


                // 检查 "全彩" 是否存在于 cpxxBox.Text 中
                bool containsFullColor = cpxxBox.Text.Contains("全彩");

                if (!string.IsNullOrEmpty(灯带系列))
                {
                    if (cpxxBox.Text.Contains("全彩"))
                    {

                        if (contentBetweenFifthAndSixth == "R" && containsFullColor) { output_色温 = $"Color: Red(Full color jacket)"; }
                        else if (contentBetweenFifthAndSixth == "B" && containsFullColor) { output_色温 = $"Color: Blue(Full color jacket)"; }
                        else if (contentBetweenFifthAndSixth == "G" && containsFullColor) { output_色温 = $"Color: Green(Full color jacket)"; }
                        else if (contentBetweenFifthAndSixth == "O" && containsFullColor) { output_色温 = $"Color: Orange(Full color jacket)"; }
                        else if (contentBetweenFifthAndSixth == "Y" && containsFullColor) { output_色温 = $"Color: Yellow(Full color jacket)"; }
                        else if (contentBetweenFifthAndSixth == "Y578") { output_色温 = $"Color: Yellow (Full color jacket) (Y578nm)"; }
                        else if (contentBetweenFifthAndSixth == "Y580") { output_色温 = $"Color: Yellow (Full color jacket) (Y580nm)"; }
                        else if (contentBetweenFifthAndSixth == "Y582") { output_色温 = $"Color: Yellow (Full color jacket) (Y582nm)"; }
                        else
                        {
                            // 检查内容是否为纯字母
                            if (Regex.IsMatch(contentBetweenFifthAndSixth, @"^[a-zA-Z]*$"))
                            {
                                output_色温 = $"Color: {contentBetweenFifthAndSixth}";
                            }
                            else
                            {
                                // 如果包含数字，则提取数字部分
                                numericValue = Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty);
                                output_色温 = $"Color: {numericValue}K";
                            }
                        }

                    }
                    else if (cpxxBox.Text.Contains("黑色遮光+雾状发光")) { output_色温 = $"Color: {Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty)}K(Black jacket)"; }
                    else if (cpxxBox.Text.Contains("黑色全彩")) { output_色温 = $"Color: {Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty)}K(Full Black jacket)"; }

                    else
                    {
                        //MessageBox.Show(contentBetweenFifthAndSixth + 灯带系列);
                        if (灯带系列 == "S")
                        {
                            //output_色温 = $"Color: {contentBetweenFifthAndSixth}";// 示例，表示输出所有色温内容

                            // 检查并处理色温范围
                            string pattern = @"W(\d{4})~(\d{4})";
                            Match match = Regex.Match(contentBetweenFifthAndSixth, pattern);
                            if (match.Success)
                            {
                                // 从匹配结果中提取数字部分，并添加 "K"
                                string output = match.Groups[1].Value + "K~" + match.Groups[2].Value + "K";
                                output_色温 = $"Color: {output}";

                            }
                            else
                            {
                                if (contentBetweenFifthAndSixth == "R") { output_色温 = $"Color: Red"; }
                                else if (contentBetweenFifthAndSixth == "B") { output_色温 = $"Color: Blue"; }
                                else if (contentBetweenFifthAndSixth == "G") { output_色温 = $"Color: Green"; }
                                else if (contentBetweenFifthAndSixth == "O") { output_色温 = $"Color: Orange"; }
                                else if (contentBetweenFifthAndSixth == "Y") { output_色温 = $"Color: Yellow"; }
                                else if (contentBetweenFifthAndSixth == "A") { output_色温 = $"Color: Amber"; }
                                else if (contentBetweenFifthAndSixth == "P") { output_色温 = $"Color: Pink"; }
                                else
                                {
                                    // 检查内容是否为纯字母
                                    if (Regex.IsMatch(contentBetweenFifthAndSixth, @"^[a-zA-Z]*$"))
                                    {
                                        output_色温 = $"Color: {contentBetweenFifthAndSixth}";
                                    }
                                    else if (contentBetweenFifthAndSixth.Contains("RGB"))
                                    {
                                        output_色温 = $"Color: {contentBetweenFifthAndSixth}";
                                    }
                                    else
                                    {
                                        // 如果包含数字，则提取数字部分
                                        numericValue = Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty);
                                        output_色温 = $"Color: {numericValue}K";
                                    }
                                }
                            }

                        }
                        else if (灯带系列 == "E")
                        {
                            output_色温 = $"Color: {contentBetweenFifthAndSixth}";// 示例，表示输出所有色温内容
                        }
                        else if (灯带系列 == "D")
                        {
                            // 检查并处理色温范围
                            string pattern = @"W(\d{4})~(\d{4})";
                            Match match = Regex.Match(contentBetweenFifthAndSixth, pattern);
                            if (match.Success)
                            {
                                // 从匹配结果中提取数字部分，并添加 "K"
                                string output = match.Groups[1].Value + "K~" + match.Groups[2].Value + "K";
                                output_色温 = $"Color: {output}";

                            }
                            else
                            {
                                // 检查内容是否为纯字母
                                if (Regex.IsMatch(contentBetweenFifthAndSixth, @"^[a-zA-Z]*$"))
                                {
                                    output_色温 = $"Color: {contentBetweenFifthAndSixth}";
                                }
                                else
                                {
                                    // 如果包含数字，则提取数字部分
                                    numericValue = Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty);
                                    output_色温 = $"Color: {numericValue}K";
                                }
                            }

                        }
                        else if (灯带系列 == "B")
                        {

                            if (contentBetweenFifthAndSixth == "R") { output_色温 = $"Color: Red"; }
                            else if (contentBetweenFifthAndSixth == "B") { output_色温 = $"Color: Blue"; }
                            else if (contentBetweenFifthAndSixth == "G") { output_色温 = $"Color: Green"; }
                            else if (contentBetweenFifthAndSixth == "O") { output_色温 = $"Color: Orange"; }
                            else if (contentBetweenFifthAndSixth == "Y") { output_色温 = $"Color: Yellow"; }
                            else if (contentBetweenFifthAndSixth == "A") { output_色温 = $"Color: Amber"; }
                            else if (contentBetweenFifthAndSixth == "P") { output_色温 = $"Color: Pink"; }
                            else
                            {
                                // 检查内容是否为纯字母
                                if (Regex.IsMatch(contentBetweenFifthAndSixth, @"^[a-zA-Z]*$"))
                                {
                                    output_色温 = $"Color: {contentBetweenFifthAndSixth}";
                                }
                                else
                                {
                                    // 如果包含数字，则提取数字部分
                                    numericValue = Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty);
                                    output_色温 = $"Color: {numericValue}K";
                                }
                            }

                        }
                        else
                        {

                            // 检查内容是否为纯字母
                            if (Regex.IsMatch(contentBetweenFifthAndSixth, @"^[a-zA-Z]*$"))
                            {
                                output_色温 = $"Color: {contentBetweenFifthAndSixth}";
                            }
                            else
                            {
                                // 如果包含数字，则提取数字部分
                                numericValue = Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty);
                                output_色温 = $"Color: {numericValue}K";
                            }
                        }

                    }


                }
                else
                { MessageBox.Show("没有找到灯带系列。", "错误"); }



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
                    最小ip等级 = result;
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
                output_尾巴 = textBox_结尾.Text;
            }
            else
            {
                output_尾巴 = " ";
            }

            //标签种类是BIS标的时候
            if (comboBox_标签规格.Text.Contains("BIS"))
            {
                string power = powerValue + "W";
                string voltage = "DC " + voltageValue + "V";

                // 查找匹配的条目
                var result1 = data_bis.Find(d => d.Model == model && d.Power == power && d.Voltage == voltage);

                // 显示结果在消息框中
                if (result1 != default)
                {
                    //MessageBox.Show($"LEDs/m: {result1.LEDsPerMeter}, Max Power: {result1.MaxPower}", "Result");
                    output_灯带型号 = "ART. No.: " + model + "-" + powerValue + "W";
                    output_功率 = output_功率 + "," + result1.MaxPower;
                    output_灯数 = "LED Qty.: " + result1.LEDsPerMeter;

                    name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;

                    //MessageBox.Show( name_CPXXBox.Text);
                }

            }
            else if (comboBox_标签规格.Text.Contains("18395"))
            {
                output_灯带型号 = output_灯带型号.Replace("ART. No.:", "Product Code:");
                name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
            }
            else if (comboBox_标签规格.Text.Contains("13013"))
            {
                if (cpxxBox.Text.Contains("Ra95")) { output_电压 += "\nCRI: " + "95"; }
                else if (cpxxBox.Text.Contains("Ra90")) { output_电压 += "\nCRI: " + "90"; }
                else { output_电压 += "\nCRI: " + "80"; }

                name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
            }
            //高压短剪的时候
            else if (comboBox_标签规格.Text.Contains("高压") && comboBox_标签规格.Text.Contains("短剪"))
            {
                if (cpxxBox.Text.Contains("-可延长")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "-" + "AC" + voltageValue + "V" + "-" + ledQtyValu + "-" + ZZ + "-Plug" + "\n" + "Rating: 220V-240V~, 50/60Hz," + "\n" + "10W/m,0.042A/m, Max. 3.44A," + "\n" + "Max. 80m" + "\n" + output_长度 + "\n" + output_色温 + "\n" + " " + "\n" + output_尾巴; }
                else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "-" + "AC" + voltageValue + "V" + "-" + ledQtyValu + "-" + ZZ  + "\n" + "Rating: 220V-240V~, 50/60Hz," + "\n" + "10W/m,0.042A/m, Max. 3.44A," + "\n" + "Max. 80m" + "\n" + output_长度 + "\n" + output_色温 + "\n" + " " + "\n" + output_尾巴; }

            }
            else
            {
                name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
            }



            //MessageBox.Show(output_name +"\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴, "提取结果");

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
        private async void button_数据库_Click(object sender, EventArgs e)
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

        //测试按钮
        private void button_test_Click(object sender, EventArgs e)
        {





        }





        //另存为
        private void button_另存为_Click(object sender, EventArgs e)
        {
            //执行工字标和品名标
            //if (标签种类_comboBox.Text == "工字标" || 标签种类_comboBox.Text == "品名标")
            //{
            //    另存工字标和品名标();
            //}
            if (标签种类_comboBox.Text == "工字标" || 标签种类_comboBox.Text == "品名标") { shengcheng_biaoqian(biaoqian.lingcun); }
            else
            {
                //另存唛头();
                shengcheng_maitou(biaoqian.lingcun);
            }
        }


        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void 唛头_寻找灯带型号()
        {
            string input = cpxxBox.Text; // 这里替换成你的实际输入字符串

            // 正则表达式匹配以 "C-F" 开头的型号
            string pattern = @"C-([^-\s]*-[^-\s]*)";

            MatchCollection matches = Regex.Matches(input, pattern);

            int count = matches.Count; // 统计匹配到的型号数量
            //提示框.AppendText($"找到 {count} 个型号" + Environment.NewLine);
            StringBuilder output = new StringBuilder(); // 创建 StringBuilder 对象

            // 输出所有匹配到的型号
            foreach (Match match in matches)
            {
                output.Append("C-" + match.Groups[1].Value); // 追加整个匹配的字符串，包含 "C-"
                output.Append(","); // 追加逗号
            }
            // 从 StringBuilder 中移除最后一个逗号（如果有）
            if (output.Length > 0 && output[output.Length - 1] == ',')
            {
                output.Remove(output.Length - 1, 1);
            }
            // 将 StringBuilder 的内容赋值给一个字符串变量
            string finalOutput = output.ToString();

            // 使用逗号分割字符串，并去除重复内容
            var uniqueModels = new HashSet<string>(finalOutput.Split(','));

            // 将HashSet中的元素连接成一个字符串，元素之间用逗号和空格分隔
            string finalOutput1 = string.Join(", ", uniqueModels);

            // 显示最终的输出
            textBox_唛头灯带型号.Text = finalOutput1;

        }
        private void 唛头_寻找灯带电压()
        {
            string input = cpxxBox.Text; // 这里替换成你的实际输入字符串

            // 正则表达式匹配以 "C-F" 开头的型号
            string pattern = @"-D(\d+)V";

            MatchCollection matches = Regex.Matches(input, pattern);

            int count = matches.Count; // 统计匹配到的型号数量
                                       //提示框.AppendText($"找到 {count} 个型号" + Environment.NewLine);
            StringBuilder output = new StringBuilder(); // 创建 StringBuilder 对象

            // 输出所有匹配到的型号
            foreach (Match match in matches)
            {
                output.Append("DC" + match.Groups[1].Value + "V"); // 追加整个匹配的字符串，包含 "C-"
                output.Append(","); // 追加逗号
            }
            // 从 StringBuilder 中移除最后一个逗号（如果有）
            if (output.Length > 0 && output[output.Length - 1] == ',')
            {
                output.Remove(output.Length - 1, 1);
            }
            // 将 StringBuilder 的内容赋值给一个字符串变量
            string finalOutput = output.ToString();

            // 使用逗号分割字符串，并去除重复内容
            var uniqueModels = new HashSet<string>(finalOutput.Split(','));

            // 将HashSet中的元素连接成一个字符串，元素之间用逗号和空格分隔
            string finalOutput1 = string.Join(", ", uniqueModels);

            // 显示最终的输出
            textBox_唛头电压.Text = finalOutput1;

        }
        private void 唛头_寻找色温()
        {
            string text = cpxxBox.Text;

            // 用于存储输出的StringBuilder
            StringBuilder output = new StringBuilder();

            // 按行分割文本
            string[] lines = text.Split('\n');


            foreach (string line in lines)
            {
                // 检查是否以 "C-" 开头
                if (line.StartsWith("C-"))
                {
                    // 分割行
                    string[] parts = line.Split('-');

                    // 检查是否存在parts[5]并添加到output
                    if (parts.Length > 5)
                    {
                        output.Append(parts[5].Trim()); // 使用Trim()去除可能的前后空白
                        output.Append(","); // 添加逗号分隔符
                    }
                }
            }

            // 去除最后一个逗号（如果有）
            if (output.Length > 0 && output[output.Length - 1] == ',')
            {
                output.Remove(output.Length - 1, 1);
            }

            // 显示最终的输出
            Console.WriteLine(output.ToString());

            // 假设 textBox_唛头色温 是一个TextBox控件
            TextBox textBox_唛头色温 = new TextBox();
            textBox_唛头色温.Text = output.ToString();

        }

        private void 唛头_寻找订单编号()
        {


            string input = cpxxBox.Text; // 这里替换成你的实际输入字符串
            string pattern = @"XS(\d{2})(\d{2})(\d{1})(\d{3})";
            string pattern1 = @"PC(\d{2})(\d{2})(\d{2})(\d{2})";

            Match match = Regex.Match(input, pattern);
            Match match1 = Regex.Match(input, pattern1);

            if (match.Success)
            {
                textBox_订单编号.Clear();
                string specificDigits = match.Groups[2].Value + match.Groups[4].Value;
                Console.WriteLine("匹配的型号标识符: " + match.Value);
                //MessageBox.Show(specificDigits);
                textBox_订单编号.Text += specificDigits.ToString();
                Console.WriteLine("提取的最后五位数字: " + specificDigits);
            }
            else if (match1.Success)
            {
                textBox_订单编号.Clear();
                string specificDigits1 = match1.Groups[2].Value + match1.Groups[4].Value;
                //MessageBox.Show(specificDigits1);
                textBox_订单编号.Text += specificDigits1.ToString();
            }
            else
            {
                Console.WriteLine("没有找到匹配的标识符。");
            }

        }



        private void PrintForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            // 在这里编写窗口关闭时要执行的逻辑
            DialogResult result = MessageBox.Show("你要保存当前配置内容吗？", "确认", MessageBoxButtons.YesNo);

            if (result == DialogResult.No)
            {
                // 如果用户选择 "No"，则直接退出

            }
            else if (result == DialogResult.Yes)
            {
                SaveConfiguration();
            }
        }

        private void comboBox_标签规格_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_标签规格.Text.Contains("14098"))
            {
                checkBox_客户型号.Checked = true;
            }
        }
    }

}
