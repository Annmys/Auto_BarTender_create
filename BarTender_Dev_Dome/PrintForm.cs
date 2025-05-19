using IniParser;
using IniParser.Model;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml; // EPPlus的命名空间
using Seagull.BarTender.Print;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;  // 添加这行运行整合EXE
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Resolution = Seagull.BarTender.Print.Resolution;

using System.Net.Http;

// 注意：这需要引用BarTender的COM对象

namespace BarTender_Dev_Dome

{
    public partial class PrintForm : Form
    {
        private string bq_ipdj = null;
        private string _bmp_path = Application.StartupPath + @"\exp.jpg";
        private string _btw_path = string.Empty;
        private string 模板地址 = string.Empty;
        private string 软件版本 = string.Empty;
        private string 灯带材质 = string.Empty;
        private string 灯带系列 = string.Empty;
        private string _sjk_path = string.Empty;
        private string _PrinterName = string.Empty;
        private string _wjm_ = string.Empty;
        private string output_name = "Name:LED Flex Linear Light";
        private string output_灯带型号 = string.Empty;
        private string output_电压 = string.Empty;
        private string output_功率 = string.Empty;
        private string output_总功率 = string.Empty;
        private string output_光源型号 = string.Empty;
        private string output_灯数 = string.Empty;
        private string output_剪切单元 = string.Empty;
        private string output_长度 = "Length:";
        private string output_灯带长度 = string.Empty;
        private string output_线材长度 = string.Empty;
        private string output_色温 = string.Empty;
        private string output_尾巴 = string.Empty;
        private string output_透镜角度 = string.Empty;
        private string output_name1 = string.Empty;
        private string output_灯带型号1 = string.Empty;
        private string output_13009名称 = string.Empty;
        private string output_13009颜色 = string.Empty;
        private string output_13009色温 = string.Empty;
        private string output_13009流明 = string.Empty;
        private string output_13009功率 = string.Empty;
        private string output_13009条形码 = string.Empty;
        private string 最小ip等级 = string.Empty;
        private string configFilePath;
        private FileIniDataParser parser;
        private IniData data;
        private System.Windows.Forms.Timer memoryTimer;
        private Process currentProcess;
        private bool isGeneratingPreview = false;
        private Thread previewThread;

        public enum biaoqian // 枚举名称建议使用大写，以符合C#的命名规范，例如：ActionType
        {
            dayin,
            lingcun,
            yulan
        }

        // 用于存储Tab页数据的辅助类
        public class TabPageData
        {
            public 结果数据 Data { get; set; }
            public int BoxIndex { get; set; }
        }

        // 结果数据类
        public class 结果数据
        {
            public string 产品型号 { get; set; }
            public string 销售数量 { get; set; }
            public string 备注 { get; set; }
            public List<string> 纸箱规格列表 { get; set; } = new List<string>();
            public List<List<盒子内容>> 盒子列表 { get; set; } = new List<List<盒子内容>>();

            public class 盒子内容
            {
                public string 序号 { get; set; }
                public string 条数 { get; set; }
                public string 米数 { get; set; }
                public string 标签码1 { get; set; }
                public string 标签码2 { get; set; }
                public string 标签码3 { get; set; }
                public string 标签码4 { get; set; }
                public string 线长 { get; set; }

                public string 包装编码 { get; set; }

                public string 纸箱规格 { get; set; }
                // 如果您使用的是列表，则改为：
                // public List<string> 纸箱规格列表 { get; set; } = new List<string>();

                public int 盒装标准 { get; set; }
                public string 客户型号 { get; set; }
                public string 标签显示长度 { get; set; }

            }
        }

        //程序开始运行
        public PrintForm()
        {
            InitializeComponent();

            tabControl_唛头.SelectedIndex = 2;
            //MessageBox.Show("程序打开", "操作提示");

            //读取excel文件必须增加声明才能运行正常
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            LoadConfiguration();

            //更新软件标题栏为10秒一次软件占用内存
            // 获取当前进程
            currentProcess = Process.GetCurrentProcess();
            // 初始化定时器
            memoryTimer = new System.Windows.Forms.Timer();
            memoryTimer.Interval = 10000; // 10秒
            memoryTimer.Tick += UpdateMemoryUsage;
            memoryTimer.Start();
            // 初始更新一次
            UpdateMemoryUsage(null, null);

            // 添加快捷键
            this.KeyPreview = true;  // 确保窗体可以接收键盘事件
            this.KeyDown += (s, e) =>
            {
                switch (e.KeyCode)
                {
                    case Keys.F1:
                        preview_btn.PerformClick();
                        e.Handled = true;
                        break;

                    case Keys.F2:
                        button_另存为.PerformClick();
                        e.Handled = true;
                        break;

                    case Keys.Enter:
                        // 检查当前焦点是否在textBox1上
                        if (ActiveControl == textBox1)
                        {
                            print_btn.PerformClick();
                            e.Handled = true;
                        }
                        break;
                }
            };
        }

        //软件主题标题栏
        private void UpdateMemoryUsage(object sender, EventArgs e)
        {
            try
            {
                // 刷新进程信息
                currentProcess.Refresh();

                // 获取程序内存使用量（MB）
                double appMemoryUsageMB = currentProcess.WorkingSet64 / (1024.0 * 1024.0);

                // 获取系统内存信息
                var ramAvailable = new PerformanceCounter("Memory", "Available MBytes", true);
                var ramCommitted = new PerformanceCounter("Memory", "Committed Bytes", true);

                float availableMemoryMB = ramAvailable.NextValue();
                float committedMemoryBytes = ramCommitted.NextValue();
                float totalMemoryMB = availableMemoryMB + (committedMemoryBytes / (1024 * 1024));

                // 计算剩余百分比
                float memoryFreePercent = (availableMemoryMB / totalMemoryMB) * 100;

                // 更新窗体标题
                this.Text = $"自动_BarTender_标签生成_  程序使用内存: {appMemoryUsageMB:F1}MB | 系统剩余内存: {memoryFreePercent:F1}%         By:Annmy";

                // 释放资源
                ramAvailable.Dispose();
                ramCommitted.Dispose();
            }
            catch (Exception ex)
            {
                this.Text = "自动_BarTender_标签生成_                By:Annmy";
            }
        }

        //释放定时器资源
        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (memoryTimer != null)
            {
                memoryTimer.Stop();    // 停止定时器
                memoryTimer.Dispose(); // 释放定时器资源
            }

            if (currentProcess != null)
            {
                currentProcess.Dispose(); // 释放进程资源
            }

            base.OnFormClosing(e); // 调用基类的关闭方法
        }

        //读取操作记录2
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

        private void printer_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            _PrinterName = printer_comboBox.Text;
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
            _btw_path = "";

            // 添加 ping 测试
            try
            {
                System.Net.NetworkInformation.Ping ping = new System.Net.NetworkInformation.Ping();
                System.Net.NetworkInformation.PingReply reply = ping.Send("192.168.1.33");

                if (reply.Status == System.Net.NetworkInformation.IPStatus.Success)
                {
                    if (reply.RoundtripTime > 2)
                    {
                        MessageBox.Show($"服务器连接较慢，响应时间：{reply.RoundtripTime}ms，可能会影响标签生成速度", "连接测试");
                    }
                }
                else
                {
                    MessageBox.Show($"无法连接到服务器，状态：{reply.Status}", "连接测试");
                    return; // 如果连接失败，可以选择直接返回不执行后续操作
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"测试连接时出错：{ex.Message}", "连接测试");
                return; // 如果发生异常，可以选择直接返回不执行后续操作
            }

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
            else if (checkBox_简化型号.Checked)
            {
                // 如果选中，设置 comboBox_客户编号 为可见
                comboBox_标签规格.Visible = true;

                // 重置 comboBox_标签规格
                comboBox_标签规格.Items.Clear();

                // 重新加载标签规格
                获取标签规格("简化型号");
            }
            else
            {
                // 如果未选中，设置 comboBox_客户编号 为不可见
                comboBox_标签规格.Visible = false;
            }

            if (标签种类_comboBox.Text.Contains("唛头"))
            {
                foreach (TabPage page in tabControl_唛头.TabPages)
                {
                    if (page.Name == "tabPage3")
                    {
                        // 选择该TabPage
                        tabControl_唛头.SelectedTab = page;
                        break;
                    }
                }
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
            if (aa == "简化型号")
            {
                templatesDirectory = Path.Combine(baseDirectory, @"\简化型号" + @"\");

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

        //private void printer_comboBox_SelectedIndexChanged(object sender, EventArgs e)
        //{
        //    _PrinterName = printer_comboBox.Text;
        //}

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
                comboBox_唛头规格.Visible = false;
                groupBox_唛头显指.Visible = false;
            }
            else
            {
                tabControl_唛头.Visible = true;
                comboBox_唛头规格.Visible = false;
                groupBox_唛头显指.Visible = false;
            }
        }

        private string 重构产品信息_工字标(string name_CPXXBox, string textBox_剪切长度)
        {
            // 如果是12098规格，直接返回原始文本
            if (comboBox_标签规格.Text.Contains("12098"))
            {
                return name_CPXXBox;
            }
            // 如果是16008规格，直接返回原始文本
            if (comboBox_标签规格.Text.Contains("16008"))
            {
                return name_CPXXBox;
            }
            if (comboBox_标签规格.Text.Contains("12251"))
            {
                return name_CPXXBox;
            }
            if (comboBox_标签规格.Text.Contains("12090"))
            {
                return name_CPXXBox;
            }
            if (comboBox_标签规格.Text.Contains("12141"))
            {
                return name_CPXXBox;
            }
            if (comboBox_标签规格.Text.Contains("12291") && comboBox_标签规格.Text.Contains("标签型号"))
            {
                return name_CPXXBox;
            }
            if (comboBox_标签规格.Text.Contains("17021") && comboBox_标签规格.Text.Contains("标签型号"))
            {
                return name_CPXXBox;
            }

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
                    return "找不到BIS长度分界点：m\nLength:";
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
                    return "找不到高压短剪长度分界点：m\nLength:";
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
                    return "找不到长度分界点：)\nLength:";
                }
            }
            // 检查 name_CPXXBox 是否包含特定的标识 ")\nLength:"，并找到它的位置

            // 使用 StringBuilder 创建新的文本
            StringBuilder newText = new StringBuilder(name_CPXXBox);

            // 移除从找到的位置到字符串末尾的所有内容
            newText.Remove(lengthIndex, newText.Length - lengthIndex);

            // 在找到的位置后面插入 textBox_剪切长度 的内容
            newText.Insert(lengthIndex, textBox_剪切长度);

            bool contains15019 = comboBox_标签规格.Text.Contains("15019");
            bool doesNotContainUL = !comboBox_标签规格.Text.Contains("UL");
            bool isCorrectSpec = contains15019 && doesNotContainUL;

            if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
            {
                output_尾巴 = "Oneilluminates.com";
                newText.Append("\n" + output_色温);
                newText.Append("\n" + output_尾巴);
            }
            else if (isCorrectSpec)
            {
                output_尾巴 = "Led3.com";
                newText.Append("\n" + output_色温);
                newText.Append("\n" + output_尾巴);
            }
            else if (comboBox_标签规格.Text.Contains("高压") && comboBox_标签规格.Text.Contains("短剪"))
            {
                // 在文本末尾追加其他信息
                newText.Append("\n" + output_色温);
                newText.Append("\n" + " ");
                newText.Append("\n" + output_尾巴);
            }
            else if (comboBox_标签规格.Text.Contains("3525") ||
                          (comboBox_标签规格.Text.Contains("17034") && cpxxBox.Text.Contains("W3525")) ||
                          (comboBox_标签规格.Text.Contains("12058") && cpxxBox.Text.Contains("W3525")) || cpxxBox.Text.Contains("W3525"))
            {
                newText.Append("\n" + output_色温);
                newText.Append("\n" + output_透镜角度);
                newText.Append("\n" + output_尾巴);
            }
            else if (comboBox_标签规格.Text.Contains("12120"))
            {
                newText.Append("\n" + "Caution: Do not overload.");
                newText.Append("\n" + output_尾巴);
            }
            else
            {
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
            _btw_path = "";
            // 假设这是从某个文本框获取的字符串
            string cpxx_text = cpxxBox.Text;
            判断产品信息(cpxx_text);

            using (Engine btEngine = new Engine(true))
            {
                //if (!_btw_path.Contains(comboBox_标签规格.Text))
                //{
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
                else if (checkBox_简化型号.Checked)
                {
                    _btw_path = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\";
                    _btw_path = _btw_path + @"简化型号\" + comboBox_标签规格.Text;
                }
                //}

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

                    if (comboBox_标签规格.Text.Contains("12141"))
                    {
                        if (columnHasContent[2])
                        {
                            _wjm_ = "2.btw";
                        }
                        else { _wjm_ = "1.btw"; }
                    }
                }
                else
                {
                    _wjm_ = "1.btw";
                }

                if (checkBox_标识码01.Checked)
                {
                    _wjm_ = "1.btw";
                    if (comboBox_标签规格.Text.Contains("12141"))
                    {
                        _wjm_ = "2.btw";
                    }
                }
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
                        if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                        {
                            name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, textBox_剪切长度.Text);
                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                        }
                        else if (comboBox_标签规格.Text.Contains("BIS")) { } //2024.9.10发现没有大括号也没有报错，才增加了，不知道后面有没有BUG
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

                        if (comboBox_标签规格.Text.Contains("12098"))
                        {
                            if (灯带系列 == "A")
                            {
                                labelFormat.SubStrings.SetSubString("BSM-05", "Connect to" + "\n" + "MODA RGB SUPER NEON" + "\n" + "products only!");
                            }
                            else
                            {
                                labelFormat.SubStrings.SetSubString("BSM-05", "Connect to" + "\n" + "MODA SUPER NEON" + "\n" + "products only!");
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
                        if (checkBox_标识码03.Checked)
                        {
                            labelFormat.SubStrings.SetSubString("BSM-03", textBox_标识码03.Text);
                        }

                        if (checkBox_标识码04.Checked)
                        {
                            labelFormat.SubStrings.SetSubString("BSM-04", textBox_标识码04.Text);
                        }

                        if (comboBox_标签规格.Text.Contains("12251") && 标签种类_comboBox.Text.Contains("工字标"))
                        {
                            labelFormat.SubStrings.SetSubString("CPCD", textBox_剪切长度.Text);
                            labelFormat.SubStrings.SetSubString("CPXX-2", textBox_客户资料.Text);
                            if (灯带材质 == "FR") { labelFormat.SubStrings.SetSubString("IPDJ", " "); }
                            else { labelFormat.SubStrings.SetSubString("IPDJ", "Not suitable for underwater use"); }
                        }
                        else if (comboBox_标签规格.Text.Contains("12251") && 标签种类_comboBox.Text.Contains("品名标"))
                        {
                            labelFormat.SubStrings.SetSubString("CPXX-2", textBox_客户资料.Text);
                            string 处理后色温 = output_色温.Replace("Color: ", "");
                            labelFormat.SubStrings.SetSubString("SW", 处理后色温);
                            string 处理后功率 = output_功率.Replace("Rated Power: ", ""); // 删除前缀
                            int 斜杠位置 = 处理后功率.IndexOf("/");
                            if (斜杠位置 != -1)
                            {
                                处理后功率 = 处理后功率.Substring(0, 斜杠位置); // 只保留斜杠前的部分
                            }
                            labelFormat.SubStrings.SetSubString("WS", 处理后功率);
                            labelFormat.SubStrings.SetSubString("CPCD", textBox_剪切长度.Text);
                            labelFormat.SubStrings.SetSubString("PO", textBox_po号2.Text);
                        }

                        if (comboBox_标签规格.Text.Contains("12115"))
                        {
                            string 处理后条形码 = textBox_条形码.Text.Replace("EAN code: ", "");
                            labelFormat.SubStrings.SetSubString("TXM", 处理后条形码);
                        }
                        else if (comboBox_标签规格.Text.Contains("12090"))
                        {
                            string originalString = textBox_客户资料.Text;
                            int tabIndex = originalString.LastIndexOf('\t');
                            if (tabIndex != -1)
                            {
                                string part1 = originalString.Substring(0, tabIndex);
                                string part2 = originalString.Substring(tabIndex + 1);
                                labelFormat.SubStrings.SetSubString("KHXH", part1);
                                labelFormat.SubStrings.SetSubString("KHMC", part2);
                            }
                        }

                        if (comboBox_标签规格.Text.Contains("12141"))
                        {
                            string 处理后名称 = output_name.Replace("Name: ", "");
                            string 处理后名称1 = 处理后名称.Replace(Environment.NewLine, "");
                            labelFormat.SubStrings.SetSubString("BSM-02", 处理后名称1);
                            labelFormat.SubStrings.SetSubString("PO", textBox_po号2.Text);

                            string 处理后色温 = output_色温.Replace("Color: ", "");
                            if (处理后色温 == "2700K") { labelFormat.SubStrings.SetSubString("FXK-2700K", "实.png"); }
                            else if (处理后色温 == "3000K") { labelFormat.SubStrings.SetSubString("FXK-3000K", "实.png"); }
                            else if (处理后色温 == "3500K") { labelFormat.SubStrings.SetSubString("FXK-3500K", "实.png"); }
                            else if (处理后色温 == "4000K") { labelFormat.SubStrings.SetSubString("FXK-4000K", "实.png"); }
                            else if (处理后色温 == "4500K") { labelFormat.SubStrings.SetSubString("FXK-4500K", "实.png"); }
                            else if (处理后色温 == "5700K") { labelFormat.SubStrings.SetSubString("FXK-5700K", "实.png"); }
                            else if (处理后色温 == "6500K") { labelFormat.SubStrings.SetSubString("FXK-6500K", "实.png"); }
                            else if (处理后色温 == "Red") { labelFormat.SubStrings.SetSubString("FXK-Red", "实.png"); }
                            else if (处理后色温 == "Blue") { labelFormat.SubStrings.SetSubString("FXK-Blue", "实.png"); }
                            else if (处理后色温 == "RGB") { labelFormat.SubStrings.SetSubString("FXK-RGB", "实.png"); }
                            else if (处理后色温 == "Green") { labelFormat.SubStrings.SetSubString("FXK-Green", "实.png"); }
                            else if (处理后色温 == "Yellow") { labelFormat.SubStrings.SetSubString("FXK-Yellow", "实.png"); }
                            else { labelFormat.SubStrings.SetSubString("Other", 处理后色温); }

                            string pattern1 = @"^(\w+-\w+-\w+)";
                            Match match1 = Regex.Match(cpxxBox.Text, pattern1);
                            if (match1.Success)
                            {
                                string artNo = match1.Groups[1].Value;
                                if (artNo.Contains("F23")) { labelFormat.SubStrings.SetSubString("IPDJ", " "); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                else if (artNo.Contains("F16")) { labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                                else if (artNo.Contains("3525")) { labelFormat.SubStrings.SetSubString("IPDJ", "IP67"); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                else { labelFormat.SubStrings.SetSubString("IPDJ", "IP68"); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                            }
                        }

                        if (comboBox_标签规格.Text.Contains("12120"))
                        {
                            if (灯带系列 == "A") { labelFormat.SubStrings.SetSubString("FXK-A", "实.png"); }
                            else if (灯带系列 == "B") { labelFormat.SubStrings.SetSubString("FXK-B", "实.png"); }
                            else if (灯带系列 == "E") { labelFormat.SubStrings.SetSubString("FXK-E", "实.png"); }
                            else if (灯带系列 == "S") { labelFormat.SubStrings.SetSubString("FXK-S", "实.png"); }

                            string 处理后色温 = output_色温.Replace("Color: ", "");
                            if (处理后色温 == "2700K") { labelFormat.SubStrings.SetSubString("FXK-2700K", "实.png"); }
                            else if (处理后色温 == "3000K") { labelFormat.SubStrings.SetSubString("FXK-3000K", "实.png"); }
                            else if (处理后色温 == "3500K") { labelFormat.SubStrings.SetSubString("FXK-3500K", "实.png"); }
                            else if (处理后色温 == "4000K") { labelFormat.SubStrings.SetSubString("FXK-4000K", "实.png"); }
                            else if (处理后色温 == "4500K") { labelFormat.SubStrings.SetSubString("FXK-4500K", "实.png"); }
                            else if (处理后色温 == "5700K") { labelFormat.SubStrings.SetSubString("FXK-5700K", "实.png"); }
                            else if (处理后色温 == "6500K") { labelFormat.SubStrings.SetSubString("FXK-6500K", "实.png"); }
                            else if (处理后色温 == "Red") { labelFormat.SubStrings.SetSubString("FXK-Red", "实.png"); }
                            else if (处理后色温 == "Blue") { labelFormat.SubStrings.SetSubString("FXK-Blue", "实.png"); }
                            else if (处理后色温 == "RGB") { labelFormat.SubStrings.SetSubString("FXK-RGB", "实.png"); }
                            else if (处理后色温 == "Green") { labelFormat.SubStrings.SetSubString("FXK-Green", "实.png"); }
                            else if (处理后色温 == "Yellow") { labelFormat.SubStrings.SetSubString("FXK-Yellow", "实.png"); }
                            else { labelFormat.SubStrings.SetSubString("Other", 处理后色温); }
                        }

                        if (comboBox_标签规格.Text.Contains("13009"))
                        {
                            output_长度 = "Quantity:   " + textBox_剪切长度.Text;
                            name_CPXXBox.Text = output_13009色温 + "\n" + output_13009流明 + "\n" + output_13009功率 + "\n" + output_电压 + "\n" + output_长度;
                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                            labelFormat.SubStrings.SetSubString("KHXH", textBox_客户资料.Text);
                            labelFormat.SubStrings.SetSubString("KHMC", output_13009名称);
                            labelFormat.SubStrings.SetSubString("KHYS", output_13009颜色);
                            if (textBox_客户资料.Text == "L65XT2004" || textBox_客户资料.Text == "L65XT2006" || textBox_客户资料.Text == "L65XT2008") { labelFormat.SubStrings.SetSubString("CRI80", "CRI80.png"); }
                            else { labelFormat.SubStrings.SetSubString("CRI80", "空.png"); }
                            if (标签种类_comboBox.Text.Contains("品名标")) { labelFormat.SubStrings.SetSubString("TXM", output_13009条形码); }
                            else { labelFormat.SubStrings.SetSubString("TXM", "4251158486185"); }
                        }

                        if (comboBox_标签规格.Text.Contains("标签型号"))
                        {
                            //string cpxx_text = cpxxBox.Text;
                            //判断产品信息(cpxxBox.Text);

                            string cz型号 = output_灯带型号.Replace("ART. No.: ", "");
                            string cz电压 = output_电压.Replace("Rated Voltage: DC ", "");
                            string cz色温 = output_色温.Replace("Color: ", "");
                            string cz色温1 = string.Empty;      //只保留数字色温
                            string cz色温2 = string.Empty;      //转换色温
                            string cz灯数 = new string(output_灯数.Replace("LED Qty.: ", "").Where(char.IsDigit).ToArray());
                            string cz功率 = string.Empty;

                            string patternX1 = @"^(\w+-\w+-\w+)";
                            Match matchX1 = Regex.Match(cpxxBox.Text, patternX1);
                            if (matchX1.Success) { cz型号 = matchX1.Groups[1].Value; }

                            cz色温1 = cz色温.Replace("K", ""); ;

                            if (cz色温.Contains("RGBW")) { cz色温2 = "RGBW"; }
                            else if (cz色温.Contains("K") || cz色温.Contains("k") && !cz色温.Contains("RGBW"))
                            {
                                if (cz色温.Contains("~"))
                                {
                                    cz色温2 = cz色温.Replace("K", "");
                                }
                                else
                                {
                                    // 提取数字部分
                                    string 数字部分 = new string(cz色温.Where(char.IsDigit).ToArray());
                                    if (int.TryParse(数字部分, out int 色温值))
                                    {
                                        if (色温值 >= 1100 && 色温值 < 11111111) { cz色温2 = "白光"; }
                                        else if (色温值 > 11111111)
                                        {
                                            // 将数字转换为字符串
                                            string 色温字符串 = 色温值.ToString();

                                            // 检查长度是否足够
                                            if (色温字符串.Length >= 8)  // 确保有足够的数字
                                            {
                                                // 在第4位后插入~
                                                string 前半部分 = 色温字符串.Substring(0, 4);  // 取前4位
                                                string 后半部分 = 色温字符串.Substring(4);     // 取剩余部分
                                                cz色温2 = $"{前半部分}~{后半部分}";  // 组合结果
                                            }
                                        }
                                    }
                                }
                            }
                            else if (cz色温.Contains("RGB") && !cz色温.Contains("RGBW")) { cz色温2 = "RGB"; }
                            else if (cz色温.Contains("Red")) { cz色温2 = "红"; }
                            else if (cz色温.Contains("Blue")) { cz色温2 = "蓝"; }
                            else if (cz色温.Contains("Green")) { cz色温2 = "绿"; }
                            else if (cz色温.Contains("Orange")) { cz色温2 = "橙"; }
                            else if (cz色温.Contains("Yellow")) { cz色温2 = "黄"; }
                            else if (cz色温.Contains("Amber")) { cz色温2 = "琥珀"; }

                            //UL标12291
                            if (comboBox_标签规格.Text.Contains("12291"))
                            {
                                string matchingCriteria = $"匹配条件：\n型号: {cz型号}\n电压: {cz电压}\n色温: {cz色温2}\n灯数: {cz灯数}";
                                MessageBox.Show(matchingCriteria, "UL标签12291匹配条件");

                                string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12291 UL资料.xlsx";

                                using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                {
                                    var worksheet1 = package1.Workbook.Worksheets[0];
                                    int rowCount1 = worksheet1.Dimension.Rows;
                                    bool found = false;

                                    // 遍历每一行，只检查B列
                                    for (int row1 = 2; row1 <= rowCount1; row1++)
                                    {
                                        string cellValue = worksheet1.Cells[row1, 2].Text; // 只读取B列的内容
                                        if (cpxxBox.Text.Contains("【DMX】"))
                                        {
                                            if (cellValue.Contains(cz型号) && cellValue.Contains(cz电压) && cellValue.Contains(cz色温2) && cellValue.Contains(cz灯数) && cellValue.Contains("DMX"))
                                            {
                                                // 找到匹配项，获取同行的内容
                                                string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                                string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                                output_name1 = "Name: " + DColumnContent;
                                                output_灯带型号1 = "ART. No.: " + CColumnContent;

                                                found = true;
                                                break; // 找到后立即退出循环
                                            }
                                        }
                                        else
                                        {
                                            if (cellValue.Contains(cz型号) && cellValue.Contains(cz电压) && cellValue.Contains(cz色温2) && cellValue.Contains(cz灯数))
                                            {
                                                // 找到匹配项，获取同行的内容
                                                string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                                string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                                output_name1 = "Name: " + DColumnContent;
                                                output_灯带型号1 = "ART. No.: " + CColumnContent;

                                                found = true;
                                                break; // 找到后立即退出循环
                                            }
                                        }
                                    }
                                }
                            }

                            //UL标17021，只需要核对灯带型号
                            else if (comboBox_标签规格.Text.Contains("17021"))
                            {
                                string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\17021 UL资料.xlsx";

                                using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                {
                                    var worksheet1 = package1.Workbook.Worksheets[0];
                                    int rowCount1 = worksheet1.Dimension.Rows;
                                    bool found = false;

                                    // 遍历每一行，只检查B列
                                    for (int row1 = 2; row1 <= rowCount1; row1++)
                                    {
                                        string cellValue = worksheet1.Cells[row1, 2].Text; // 只读取B列的内容

                                        if (cellValue.Contains(cz型号))
                                        {
                                            // 找到匹配项，获取同行的内容
                                            string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                            string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                            output_name1 = "Name: " + DColumnContent;
                                            output_灯带型号1 = "ART. No.: " + CColumnContent;

                                            found = true;
                                            break; // 找到后立即退出循环
                                        }
                                    }
                                }
                            }
                            output_长度 = "Length:" + textBox_剪切长度.Text;
                            name_CPXXBox.Text = output_name1 + "\n" + output_灯带型号1 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                        }

                        //2025.1.23增加简化型号的无附件生成
                        if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && !comboBox_标签规格.Text.Contains("直发"))
                        {
                            output_长度 = "Length: " + textBox_剪切长度.Text;
                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                            else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                        }
                        else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && comboBox_标签规格.Text.Contains("直发"))
                        {
                            output_长度 = "Length: " + textBox_剪切长度.Text;
                            if (comboBox_标签规格.Text.Contains("3525"))
                            {
                                name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度;
                            }
                            else
                            {
                                name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温;
                            }
                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                        }
                        else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                        {
                            output_长度 = "Length: " + textBox_剪切长度.Text;
                            if (cpxxBox.Text.Contains("Ra90"))
                            {
                                output_灯数 = "CRI: " + "90";
                                if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                            }
                            else if (cpxxBox.Text.Contains("Ra95"))
                            {
                                output_灯数 = "CRI: " + "95";
                                if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                            }
                            else if (cpxxBox.Text.Contains("Ra85"))
                            {
                                output_灯数 = "CRI: " + "85";
                                if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                            }
                            else
                            {
                                output_灯数 = "CRI: " + "80";
                                if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                            }
                        }

                        //开始处理带数据库的
                        ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                        // 检查数据库地址不为空时
                        if (!string.IsNullOrEmpty(Box_数据库.Text))
                        {
                            string b2Data, c2Data, g2Data, h2Data, h1Data, a2Data, d2Data, e2Data, f2Data, i2Data, j2Data, k2Data, l2Data;

                            // 使用EPPlus打开Excel文件
                            using (var package = new ExcelPackage(new FileInfo(Box_数据库.Text)))
                            {
                                // 假设Excel工作表名为"Sheet1"
                                var worksheet = package.Workbook.Worksheets["Sheet1"];

                                // 读取B2和C2单元格的数据
                                a2Data = worksheet.Cells["A2"].Value?.ToString() ?? string.Empty; //序号
                                b2Data = worksheet.Cells["B2"].Value?.ToString() ?? string.Empty; //标签码
                                c2Data = worksheet.Cells["C2"].Value?.ToString() ?? string.Empty; //标签码
                                d2Data = worksheet.Cells["D2"].Value?.ToString() ?? string.Empty; //标签码
                                e2Data = worksheet.Cells["E2"].Value?.ToString() ?? string.Empty; //标签码
                                f2Data = worksheet.Cells["F2"].Value?.ToString() ?? string.Empty; //标签码
                                g2Data = worksheet.Cells["G2"].Value?.ToString() ?? string.Empty; //条数
                                h2Data = worksheet.Cells["H2"].Value?.ToString() ?? string.Empty; //长度
                                i2Data = worksheet.Cells["I2"].Value?.ToString() ?? string.Empty; //客户型号，如果同时处理客户名称和客户附件的时候，客户名称在textBox_客户资料中，客户型号才在附件
                                h1Data = worksheet.Cells["H1"].Value?.ToString() ?? string.Empty;
                                j2Data = worksheet.Cells["J2"].Value?.ToString() ?? string.Empty; //PO号
                                k2Data = worksheet.Cells["K2"].Value?.ToString() ?? string.Empty; //条形码
                                l2Data = worksheet.Cells["L2"].Value?.ToString() ?? string.Empty; //线长
                            }

                            switch (_wjm_)
                            {
                                case "1.btw":
                                    labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                    textBox1.Text = g2Data;
                                    //labelFormat.SubStrings.SetSubString("CPCD", h2Data);

                                    // 处理12098规格的情况
                                    if (comboBox_标签规格.Text.Contains("12098"))
                                    {
                                        string inputText = i2Data;
                                        string convertedName = "";
                                        // 1. 处理A系列
                                        if (灯带系列 == "A")
                                        {
                                            // 去掉RGB-前缀
                                            string baseText = inputText.StartsWith("RGB-") ? inputText.Substring(4) : inputText;

                                            // 处理SNX开头的情况
                                            if (baseText.StartsWith("SNX-"))
                                            {
                                                var nameParts = baseText.Split('-');
                                                convertedName = "SUPER-NEON-";

                                                // 处理X-FLAT或X-DOME
                                                if (nameParts.Length >= 2)
                                                {
                                                    switch (nameParts[1])
                                                    {
                                                        case "F":
                                                            convertedName += "X-FLAT";
                                                            break;

                                                        case "D":
                                                            convertedName += "X-DOME";
                                                            break;

                                                        default:
                                                            convertedName += "X-" + nameParts[1];
                                                            break;
                                                    }

                                                    // 添加中间部分，处理倒数第二部分的颜色缩写
                                                    for (int i = 2; i < nameParts.Length; i++)
                                                    {
                                                        if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                                                        {
                                                            string color = ConvertColorAbbreviation(nameParts[i]);
                                                            convertedName += "-" + color;
                                                        }
                                                        else
                                                        {
                                                            convertedName += "-" + nameParts[i];
                                                        }
                                                    }
                                                }

                                                // 最后加上RGB-前缀
                                                convertedName = "RGB-" + convertedName;
                                            }
                                            // 处理SNE开头的情况
                                            else if (baseText.StartsWith("SNE-"))
                                            {
                                                var nameParts = baseText.Split('-');
                                                convertedName = "SUPER-NEON-EDGE";

                                                // 从第二个部分开始添加，处理倒数第二部分的颜色缩写
                                                for (int i = 1; i < nameParts.Length; i++)
                                                {
                                                    if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                                                    {
                                                        string color = ConvertColorAbbreviation(nameParts[i]);
                                                        convertedName += "-" + color;
                                                    }
                                                    else
                                                    {
                                                        convertedName += "-" + nameParts[i];
                                                    }
                                                }

                                                // 最后加上RGB-前缀
                                                convertedName = "RGB-" + convertedName;
                                            }
                                            else
                                            {
                                                convertedName = inputText; // 如果既不是SNX也不是SNE开头，保持原样
                                            }
                                        }
                                        // 2. 处理SNX开头的情况
                                        else if (inputText.StartsWith("SNX-"))
                                        {
                                            var nameParts = inputText.Split('-');
                                            convertedName = "SUPER-NEON-";

                                            // 处理X-FLAT或X-DOME
                                            if (nameParts.Length >= 2)
                                            {
                                                switch (nameParts[1])
                                                {
                                                    case "F":
                                                        convertedName += "X-FLAT";
                                                        break;

                                                    case "D":
                                                        convertedName += "X-DOME";
                                                        break;

                                                    default:
                                                        convertedName += nameParts[1];
                                                        break;
                                                }

                                                // 添加剩余部分，处理颜色缩写
                                                for (int i = 2; i < nameParts.Length; i++)
                                                {
                                                    if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                                                    {
                                                        string color = ConvertColorAbbreviation(nameParts[i]);
                                                        convertedName += "-" + color;
                                                    }
                                                    else
                                                    {
                                                        convertedName += "-" + nameParts[i];
                                                    }
                                                }
                                            }
                                        }
                                        // 3. 处理SNE开头的情况
                                        else if (inputText.StartsWith("SNE-"))
                                        {
                                            var nameParts = inputText.Split('-');
                                            convertedName = "SUPER-NEON-EDGE";

                                            // 从第二个部分开始添加，处理颜色缩写
                                            for (int i = 1; i < nameParts.Length; i++)
                                            {
                                                if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                                                {
                                                    string color = ConvertColorAbbreviation(nameParts[i]);
                                                    convertedName += "-" + color;
                                                }
                                                else
                                                {
                                                    convertedName += "-" + nameParts[i];
                                                }
                                            }
                                        }
                                        else
                                        {
                                            convertedName = inputText; // 默认情况
                                        }

                                        output_name = convertedName;
                                        output_灯带型号 = "Short SKU:" + i2Data;
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + "BATCH " + b2Data;
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("16008"))
                                    {
                                        output_灯带型号 = i2Data;
                                        string 长度1 = "Length: " + h2Data;
                                        string po1 = "Lot #:  " + b2Data;
                                        name_CPXXBox.Text = "www.SGiLighting.com" + "\n" + "LED NEON FLEX LIGHT" + "\n" + output_灯带型号 + "\n" + 长度1 + "\n" + po1;
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("12251") && 标签种类_comboBox.Text.Contains("工字标"))
                                    {
                                        string chazhaoziliao = i2Data;
                                        string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12251资料.xlsx";

                                        try
                                        {
                                            using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                            {
                                                var worksheet1 = package1.Workbook.Worksheets[0];
                                                int rowCount1 = worksheet1.Dimension.Rows;
                                                string fColumnContent = "";
                                                bool found = false;

                                                // 遍历每一行
                                                for (int row1 = 1; row1 <= rowCount1; row1++)
                                                {
                                                    // 检查A到D列
                                                    for (int col = 1; col <= 4; col++) // 1=A, 2=B, 3=C, 4=D
                                                    {
                                                        string cellValue = worksheet1.Cells[row1, col].Text;

                                                        if (cellValue == chazhaoziliao)
                                                        {
                                                            // 找到匹配项，获取同行F列的内容
                                                            fColumnContent = worksheet1.Cells[row1, 6].Text; // F列的内容
                                                            name_CPXXBox.Text = fColumnContent;
                                                            found = true;

                                                            // 可以添加一个消息框显示在哪里找到的（如果需要）
                                                            //MessageBox.Show($"在第{row1}行，第{(char)(col + 64)}列找到匹配项", "查找结果");

                                                            break;
                                                        }
                                                    }

                                                    if (found) break; // 如果找到了就退出外层循环
                                                }

                                                if (string.IsNullOrEmpty(fColumnContent))
                                                {
                                                    MessageBox.Show($"在A、B、C、D列中均未找到匹配内容: {chazhaoziliao}", "查找结果");
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                                        }

                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX-2", i2Data);
                                        if (灯带材质 == "FR") { labelFormat.SubStrings.SetSubString("IPDJ", " "); }
                                        else { labelFormat.SubStrings.SetSubString("IPDJ", "Not suitable for underwater use"); }
                                    }
                                    else if (comboBox_标签规格.Text.Contains("12251") && 标签种类_comboBox.Text.Contains("品名标"))
                                    {
                                        string chazhaoziliao = i2Data;
                                        string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12251资料.xlsx";

                                        try
                                        {
                                            using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                            {
                                                var worksheet1 = package1.Workbook.Worksheets[0];
                                                int rowCount1 = worksheet1.Dimension.Rows;
                                                string fColumnContent = "";
                                                bool found = false;

                                                // 遍历每一行
                                                for (int row1 = 1; row1 <= rowCount1; row1++)
                                                {
                                                    // 检查A到D列
                                                    for (int col = 1; col <= 4; col++) // 1=A, 2=B, 3=C, 4=D
                                                    {
                                                        string cellValue = worksheet1.Cells[row1, col].Text;

                                                        if (cellValue == chazhaoziliao)
                                                        {
                                                            fColumnContent = worksheet1.Cells[row1, 5].Text; //
                                                            name_CPXXBox.Text = fColumnContent;
                                                            found = true;

                                                            // 可以添加一个消息框显示在哪里找到的（如果需要）
                                                            //MessageBox.Show($"在第{row1}行，第{(char)(col + 64)}列找到匹配项", "查找结果");

                                                            break;
                                                        }
                                                    }

                                                    if (found) break; // 如果找到了就退出外层循环
                                                }

                                                if (string.IsNullOrEmpty(fColumnContent))
                                                {
                                                    MessageBox.Show($"在A、B、C、D列中均未找到匹配内容: {chazhaoziliao}", "查找结果");
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                                        }

                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        labelFormat.SubStrings.SetSubString("CPCD", h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX-2", i2Data);
                                        string 处理后色温 = output_色温.Replace("Color: ", "");
                                        string 处理后功率 = output_功率.Replace("Rated Power: ", ""); // 删除前缀
                                        int 斜杠位置 = 处理后功率.IndexOf("/");
                                        if (斜杠位置 != -1)
                                        {
                                            处理后功率 = 处理后功率.Substring(0, 斜杠位置); // 只保留斜杠前的部分
                                        }
                                        labelFormat.SubStrings.SetSubString("WS", 处理后功率);
                                        labelFormat.SubStrings.SetSubString("PO", j2Data);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("12115"))
                                    {
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
                                        //if (checkBox_客户型号.Checked)
                                        if (!string.IsNullOrEmpty(i2Data))
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

                                            if (comboBox_标签规格.Text.Contains("13013"))
                                            {
                                                name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                            //name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);

                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        //MessageBox.Show(name_CPXXBox.Text);
                                        else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                        {
                                            name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }

                                        string 处理后条形码 = k2Data.Replace("EAN code: ", "");
                                        labelFormat.SubStrings.SetSubString("TXM", 处理后条形码);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("12090"))
                                    {
                                        name_CPXXBox.Text = output_功率 + "\n" + output_色温 + "\n" + output_剪切单元 + "\n" + "Rollengte:" + h2Data;
                                        labelFormat.SubStrings.SetSubString("KHXH", i2Data);
                                        labelFormat.SubStrings.SetSubString("KHMC", textBox_客户资料.Text);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("12141"))
                                    {
                                        // 获取灯带长度
                                        output_灯带型号 = i2Data;
                                        //string pattern12 = @"(\d+,\d+)m";
                                        string pattern12 = @"(\d+,\d+)[Mm]";
                                        Regex regex12 = new Regex(pattern12);
                                        Match match = regex12.Match(output_灯带型号);
                                        if (match.Success)
                                        {
                                            string lengthStr = match.Groups[1].Value; // 获取匹配到的长度值（如：3,58319）
                                            output_灯带长度 = lengthStr + "m";
                                            //MessageBox.Show($"提取到的长度：{output_灯带长度}");
                                        }
                                        output_灯带型号 = "ART. No.: " + i2Data;
                                        // 处理功率：去除"Rated Power: "和"W/m"，只保留数字
                                        string powerStr = output_功率.Replace("Rated Power: ", "")
                                                                   .Replace("W/m", "")
                                                                   .Trim();

                                        // 处理长度：去除"m"单位，将逗号替换为小数点
                                        string lengthStr1 = output_灯带长度.Replace("m", "")
                                                                         .Replace(",", ".")
                                                                         .Trim();

                                        // 转换为double进行计算
                                        if (double.TryParse(powerStr, out double power) &&
                                            double.TryParse(lengthStr1, out double length))
                                        {
                                            // 计算总功率
                                            double totalPower = power * length;

                                            // 保留2位小数
                                            output_总功率 = $"Total Power:{totalPower:F2}W";
                                            //MessageBox.Show(output_总功率);
                                        }
                                        output_长度 = "Length of Light:" + output_灯带长度;
                                        output_线材长度 = "Length of Cable: " + l2Data;
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_总功率 + "\n" + output_光源型号 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_线材长度;
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                        string 处理后名称 = output_name.Replace("Name: ", "");
                                        string 处理后名称1 = 处理后名称.Replace(Environment.NewLine, "");
                                        labelFormat.SubStrings.SetSubString("BSM-02", 处理后名称1);
                                        labelFormat.SubStrings.SetSubString("PO", j2Data);

                                        string 处理后色温 = output_色温.Replace("Color: ", "");
                                        if (处理后色温 == "2700K") { labelFormat.SubStrings.SetSubString("FXK-2700K", "实.png"); }
                                        else if (处理后色温 == "3000K") { labelFormat.SubStrings.SetSubString("FXK-3000K", "实.png"); }
                                        else if (处理后色温 == "3500K") { labelFormat.SubStrings.SetSubString("FXK-3500K", "实.png"); }
                                        else if (处理后色温 == "4000K") { labelFormat.SubStrings.SetSubString("FXK-4000K", "实.png"); }
                                        else if (处理后色温 == "4500K") { labelFormat.SubStrings.SetSubString("FXK-4500K", "实.png"); }
                                        else if (处理后色温 == "5700K") { labelFormat.SubStrings.SetSubString("FXK-5700K", "实.png"); }
                                        else if (处理后色温 == "6500K") { labelFormat.SubStrings.SetSubString("FXK-6500K", "实.png"); }
                                        else if (处理后色温 == "Red") { labelFormat.SubStrings.SetSubString("FXK-Red", "实.png"); }
                                        else if (处理后色温 == "Blue") { labelFormat.SubStrings.SetSubString("FXK-Blue", "实.png"); }
                                        else if (处理后色温 == "RGB") { labelFormat.SubStrings.SetSubString("FXK-RGB", "实.png"); }
                                        else if (处理后色温 == "Green") { labelFormat.SubStrings.SetSubString("FXK-Green", "实.png"); }
                                        else if (处理后色温 == "Yellow") { labelFormat.SubStrings.SetSubString("FXK-Yellow", "实.png"); }
                                        else { labelFormat.SubStrings.SetSubString("Other", 处理后色温); }

                                        string pattern1 = @"^(\w+-\w+-\w+)";
                                        Match match1 = Regex.Match(cpxxBox.Text, pattern1);
                                        if (match1.Success)
                                        {
                                            string artNo = match1.Groups[1].Value;
                                            if (artNo.Contains("F23")) { labelFormat.SubStrings.SetSubString("IPDJ", " "); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                            else if (artNo.Contains("F16")) { labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                                            else if (artNo.Contains("3525")) { labelFormat.SubStrings.SetSubString("IPDJ", "IP67"); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                            else { labelFormat.SubStrings.SetSubString("IPDJ", "IP68"); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                                        }
                                    }
                                    else if (comboBox_标签规格.Text.Contains("12120"))
                                    {
                                        output_灯带型号 = "ART. No.: " + i2Data;
                                        name_CPXXBox.Text = "AMBIANCE LUMIERE" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + "Caution: Do not overload." + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);

                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                        if (灯带系列 == "A") { labelFormat.SubStrings.SetSubString("FXK-A", "实.png"); }
                                        else if (灯带系列 == "B") { labelFormat.SubStrings.SetSubString("FXK-B", "实.png"); }
                                        else if (灯带系列 == "E") { labelFormat.SubStrings.SetSubString("FXK-E", "实.png"); }
                                        else if (灯带系列 == "S") { labelFormat.SubStrings.SetSubString("FXK-S", "实.png"); }

                                        string 处理后色温 = output_色温.Replace("Color: ", "");
                                        if (处理后色温 == "2700K") { labelFormat.SubStrings.SetSubString("FXK-2700K", "实.png"); }
                                        else if (处理后色温 == "3000K") { labelFormat.SubStrings.SetSubString("FXK-3000K", "实.png"); }
                                        else if (处理后色温 == "3500K") { labelFormat.SubStrings.SetSubString("FXK-3500K", "实.png"); }
                                        else if (处理后色温 == "4000K") { labelFormat.SubStrings.SetSubString("FXK-4000K", "实.png"); }
                                        else if (处理后色温 == "4500K") { labelFormat.SubStrings.SetSubString("FXK-4500K", "实.png"); }
                                        else if (处理后色温 == "5700K") { labelFormat.SubStrings.SetSubString("FXK-5700K", "实.png"); }
                                        else if (处理后色温 == "6500K") { labelFormat.SubStrings.SetSubString("FXK-6500K", "实.png"); }
                                        else if (处理后色温 == "Red") { labelFormat.SubStrings.SetSubString("FXK-Red", "实.png"); }
                                        else if (处理后色温 == "Blue") { labelFormat.SubStrings.SetSubString("FXK-Blue", "实.png"); }
                                        else if (处理后色温 == "RGB") { labelFormat.SubStrings.SetSubString("FXK-RGB", "实.png"); }
                                        else if (处理后色温 == "Green") { labelFormat.SubStrings.SetSubString("FXK-Green", "实.png"); }
                                        else if (处理后色温 == "Yellow") { labelFormat.SubStrings.SetSubString("FXK-Yellow", "实.png"); }
                                        else { labelFormat.SubStrings.SetSubString("Other", 处理后色温); }
                                    }
                                    else if (comboBox_标签规格.Text.Contains("13009"))
                                    {
                                        string chazhaoziliao = i2Data;
                                        string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\13009资料.xlsx";
                                        try
                                        {
                                            using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                            {
                                                var worksheet1 = package1.Workbook.Worksheets[0];
                                                int rowCount1 = worksheet1.Dimension.Rows;
                                                bool found = false;

                                                // 遍历每一行，只检查A列
                                                for (int row1 = 1; row1 <= rowCount1; row1++)
                                                {
                                                    string cellValue = worksheet1.Cells[row1, 1].Text; // 只读取A列的内容

                                                    if (cellValue == chazhaoziliao)
                                                    {
                                                        // 找到匹配项，获取同行F列的内容
                                                        string BColumnContent = worksheet1.Cells[row1, 2].Text; // B列的内容,名称
                                                        string CColumnContent = worksheet1.Cells[row1, 3].Text; // C列的内容,颜色（名称颜色中间）
                                                        string DColumnContent = worksheet1.Cells[row1, 4].Text; // D列的内容,色温（CCT）
                                                        string EColumnContent = worksheet1.Cells[row1, 5].Text; // E列的内容,流明(Lumen)
                                                        string FColumnContent = worksheet1.Cells[row1, 6].Text; // F列的内容,功率(Wattage)
                                                        string GColumnContent = worksheet1.Cells[row1, 7].Text; // G列的内容.条形码
                                                        output_13009名称 = BColumnContent;
                                                        output_13009颜色 = CColumnContent;
                                                        output_13009色温 = "CCT:         " + DColumnContent;
                                                        output_13009流明 = "Lumen:     " + EColumnContent;
                                                        output_13009功率 = "Wattage:   " + FColumnContent;
                                                        output_13009条形码 = GColumnContent;

                                                        found = true;
                                                        break; // 找到后立即退出循环
                                                    }
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                                        }
                                        output_长度 = "Quantity:   " + h2Data;
                                        name_CPXXBox.Text = output_13009色温 + "\n" + output_13009流明 + "\n" + output_13009功率 + "\n" + output_电压 + "\n" + output_长度;
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        labelFormat.SubStrings.SetSubString("KHXH", i2Data);
                                        labelFormat.SubStrings.SetSubString("KHMC", output_13009名称);
                                        labelFormat.SubStrings.SetSubString("KHYS", output_13009颜色);
                                        if (i2Data == "L65XT2004" || i2Data == "L65XT2006" || i2Data == "L65XT2008") { labelFormat.SubStrings.SetSubString("CRI80", "CRI80.png"); }
                                        else { labelFormat.SubStrings.SetSubString("CRI80", "空.png"); }
                                        if (标签种类_comboBox.Text.Contains("品名标")) { labelFormat.SubStrings.SetSubString("TXM", output_13009条形码); }
                                        else { labelFormat.SubStrings.SetSubString("TXM", "4251158486185"); }
                                    }
                                    else if (comboBox_标签规格.Text.Contains("18395"))
                                    {
                                        output_灯带型号 = output_灯带型号.Replace("ART. No.:", "Product Code:");
                                        output_长度 = "Length:" + h2Data;
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("标签型号"))
                                    {
                                        //string cpxx_text = cpxxBox.Text;
                                        //判断产品信息(cpxxBox.Text);

                                        string cz型号 = output_灯带型号.Replace("ART. No.: ", "");
                                        string cz电压 = output_电压.Replace("Rated Voltage: DC ", "");
                                        string cz色温 = output_色温.Replace("Color: ", "");
                                        string cz色温1 = string.Empty;      //只保留数字色温
                                        string cz色温2 = string.Empty;      //转换色温
                                        string cz灯数 = new string(output_灯数.Replace("LED Qty.: ", "").Where(char.IsDigit).ToArray());
                                        string cz功率 = string.Empty;

                                        string patternX1 = @"^(\w+-\w+-\w+)";
                                        Match matchX1 = Regex.Match(cpxxBox.Text, patternX1);
                                        if (matchX1.Success) { cz型号 = matchX1.Groups[1].Value; }

                                        cz色温1 = cz色温.Replace("K", ""); ;

                                        if (cz色温.Contains("RGBW")) { cz色温2 = "RGBW"; }
                                        else if (cz色温.Contains("K") || cz色温.Contains("k") && !cz色温.Contains("RGBW"))
                                        {
                                            if (cz色温.Contains("~"))
                                            {
                                                cz色温2 = cz色温.Replace("K", "");
                                            }
                                            else
                                            {
                                                // 提取数字部分
                                                string 数字部分 = new string(cz色温.Where(char.IsDigit).ToArray());
                                                if (int.TryParse(数字部分, out int 色温值))
                                                {
                                                    if (色温值 >= 1100 && 色温值 < 11111111) { cz色温2 = "白光"; }
                                                    else if (色温值 > 11111111)
                                                    {
                                                        // 将数字转换为字符串
                                                        string 色温字符串 = 色温值.ToString();

                                                        // 检查长度是否足够
                                                        if (色温字符串.Length >= 8)  // 确保有足够的数字
                                                        {
                                                            // 在第4位后插入~
                                                            string 前半部分 = 色温字符串.Substring(0, 4);  // 取前4位
                                                            string 后半部分 = 色温字符串.Substring(4);     // 取剩余部分
                                                            cz色温2 = $"{前半部分}~{后半部分}";  // 组合结果
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        else if (cz色温.Contains("RGB") && !cz色温.Contains("RGBW")) { cz色温2 = "RGB"; }
                                        else if (cz色温.Contains("Red")) { cz色温2 = "红"; }
                                        else if (cz色温.Contains("Blue")) { cz色温2 = "蓝"; }
                                        else if (cz色温.Contains("Green")) { cz色温2 = "绿"; }
                                        else if (cz色温.Contains("Orange")) { cz色温2 = "橙"; }
                                        else if (cz色温.Contains("Yellow")) { cz色温2 = "黄"; }
                                        else if (cz色温.Contains("Amber")) { cz色温2 = "琥珀"; }

                                        //UL标12291
                                        if (comboBox_标签规格.Text.Contains("12291"))
                                        {
                                            //string matchingCriteria = $"匹配条件：\n型号: {cz型号}\n电压: {cz电压}\n色温: {cz色温2}\n灯数: {cz灯数}";
                                            //MessageBox.Show(matchingCriteria, "UL标签12291匹配条件");

                                            string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12291 UL资料.xlsx";

                                            using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                            {
                                                var worksheet1 = package1.Workbook.Worksheets[0];
                                                int rowCount1 = worksheet1.Dimension.Rows;
                                                bool found = false;

                                                // 遍历每一行，只检查B列
                                                for (int row1 = 2; row1 <= rowCount1; row1++)
                                                {
                                                    string cellValue = worksheet1.Cells[row1, 2].Text; // 只读取B列的内容
                                                    if (cpxxBox.Text.Contains("【DMX】"))
                                                    {
                                                        if (cellValue.Contains(cz型号) && cellValue.Contains(cz电压) && cellValue.Contains(cz色温2) && cellValue.Contains(cz灯数) && cellValue.Contains("DMX"))
                                                        {
                                                            // 找到匹配项，获取同行的内容
                                                            string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                                            string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                                            output_name1 = "Name: " + DColumnContent;
                                                            output_灯带型号1 = "ART. No.: " + CColumnContent;

                                                            found = true;
                                                            break; // 找到后立即退出循环
                                                        }
                                                    }
                                                    else
                                                    {
                                                        if (cellValue.Contains(cz型号) && cellValue.Contains(cz电压) && cellValue.Contains(cz色温2) && cellValue.Contains(cz灯数))
                                                        {
                                                            // 找到匹配项，获取同行的内容
                                                            string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                                            string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                                            output_name1 = "Name: " + DColumnContent;
                                                            output_灯带型号1 = "ART. No.: " + CColumnContent;

                                                            found = true;
                                                            break; // 找到后立即退出循环
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        //UL标17021，只需要核对灯带型号
                                        else if (comboBox_标签规格.Text.Contains("17021"))
                                        {
                                            string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\17021 UL资料.xlsx";

                                            using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                            {
                                                var worksheet1 = package1.Workbook.Worksheets[0];
                                                int rowCount1 = worksheet1.Dimension.Rows;
                                                bool found = false;

                                                // 遍历每一行，只检查B列
                                                for (int row1 = 2; row1 <= rowCount1; row1++)
                                                {
                                                    string cellValue = worksheet1.Cells[row1, 2].Text; // 只读取B列的内容

                                                    if (cellValue.Contains(cz型号))
                                                    {
                                                        // 找到匹配项，获取同行的内容
                                                        string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                                        string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                                        output_name1 = "Name: " + DColumnContent;
                                                        output_灯带型号1 = "ART. No.: " + CColumnContent;

                                                        found = true;
                                                        break; // 找到后立即退出循环
                                                    }
                                                }
                                            }
                                        }
                                        output_长度 = "Length:" + h2Data;
                                        name_CPXXBox.Text = output_name1 + "\n" + output_灯带型号1 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("12573") && comboBox_标签规格.Text.Contains("美国"))
                                    {
                                        output_长度 = "Length:" + h2Data;
                                        output_灯带型号 = "Model:" + i2Data;
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else
                                    {
                                        //常规情况
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
                                        //if (checkBox_客户型号.Checked)
                                        if (!string.IsNullOrEmpty(i2Data))
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

                                            if (comboBox_标签规格.Text.Contains("13013"))
                                            {
                                                name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                            //name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);

                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        //MessageBox.Show(name_CPXXBox.Text);
                                        else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                        {
                                            name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                            name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                    }

                                    if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && !comboBox_标签规格.Text.Contains("直发"))
                                    {
                                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                                        output_长度 = "Length: " + h2Data;
                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                        else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && comboBox_标签规格.Text.Contains("直发"))
                                    {
                                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                                        output_长度 = "Length: " + h2Data;
                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                        else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                                    {
                                        output_长度 = "Length: " + h2Data;
                                        if (cpxxBox.Text.Contains("Ra90"))
                                        {
                                            output_灯数 = "CRI: " + "90";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else if (cpxxBox.Text.Contains("Ra95"))
                                        {
                                            output_灯数 = "CRI: " + "95";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else if (cpxxBox.Text.Contains("Ra85"))
                                        {
                                            output_灯数 = "CRI: " + "85";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else
                                        {
                                            output_灯数 = "CRI: " + "80";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                    }

                                    break;

                                //12141才有2号模板
                                case "2.btw":
                                    labelFormat.SubStrings.SetSubString("XLH", a2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-01", b2Data);
                                    labelFormat.SubStrings.SetSubString("BSM-02", c2Data);
                                    textBox1.Text = g2Data;
                                    if (comboBox_标签规格.Text.Contains("12141"))
                                    {
                                        // 获取灯带长度
                                        output_灯带型号 = i2Data;
                                        //string pattern12 = @"(\d+,\d+)m";
                                        string pattern12 = @"(\d+,\d+)[Mm]";
                                        Regex regex12 = new Regex(pattern12);
                                        Match match = regex12.Match(output_灯带型号);
                                        if (match.Success)
                                        {
                                            string lengthStr = match.Groups[1].Value; // 获取匹配到的长度值（如：3,58319）
                                            output_灯带长度 = lengthStr + "m";
                                            //MessageBox.Show($"提取到的长度：{output_灯带长度}");
                                        }
                                        output_灯带型号 = "ART. No.: " + i2Data;
                                        // 处理功率：去除"Rated Power: "和"W/m"，只保留数字
                                        string powerStr = output_功率.Replace("Rated Power: ", "")
                                                                   .Replace("W/m", "")
                                                                   .Trim();

                                        // 处理长度：去除"m"单位，将逗号替换为小数点
                                        string lengthStr1 = output_灯带长度.Replace("m", "")
                                                                         .Replace(",", ".")
                                                                         .Trim();

                                        // 转换为double进行计算
                                        if (double.TryParse(powerStr, out double power) &&
                                            double.TryParse(lengthStr1, out double length))
                                        {
                                            // 计算总功率
                                            double totalPower = power * length;

                                            // 保留2位小数
                                            output_总功率 = $"Total Power:{totalPower:F2}W";
                                            //MessageBox.Show(output_总功率);
                                        }
                                        output_长度 = "Length of Light:" + output_灯带长度;
                                        output_线材长度 = "Length of Cable: " + l2Data;
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_总功率 + "\n" + output_光源型号 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_线材长度;
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                        string 处理后名称 = output_name.Replace("Name: ", "");
                                        string 处理后名称1 = 处理后名称.Replace(Environment.NewLine, "");
                                        labelFormat.SubStrings.SetSubString("BSM-02", 处理后名称1);
                                        labelFormat.SubStrings.SetSubString("PO", j2Data);

                                        string 处理后色温 = output_色温.Replace("Color: ", "");
                                        if (处理后色温 == "2700K") { labelFormat.SubStrings.SetSubString("FXK-2700K", "实.png"); }
                                        else if (处理后色温 == "3000K") { labelFormat.SubStrings.SetSubString("FXK-3000K", "实.png"); }
                                        else if (处理后色温 == "3500K") { labelFormat.SubStrings.SetSubString("FXK-3500K", "实.png"); }
                                        else if (处理后色温 == "4000K") { labelFormat.SubStrings.SetSubString("FXK-4000K", "实.png"); }
                                        else if (处理后色温 == "4500K") { labelFormat.SubStrings.SetSubString("FXK-4500K", "实.png"); }
                                        else if (处理后色温 == "5700K") { labelFormat.SubStrings.SetSubString("FXK-5700K", "实.png"); }
                                        else if (处理后色温 == "6500K") { labelFormat.SubStrings.SetSubString("FXK-6500K", "实.png"); }
                                        else if (处理后色温 == "Red") { labelFormat.SubStrings.SetSubString("FXK-Red", "实.png"); }
                                        else if (处理后色温 == "Blue") { labelFormat.SubStrings.SetSubString("FXK-Blue", "实.png"); }
                                        else if (处理后色温 == "RGB") { labelFormat.SubStrings.SetSubString("FXK-RGB", "实.png"); }
                                        else if (处理后色温 == "Green") { labelFormat.SubStrings.SetSubString("FXK-Green", "实.png"); }
                                        else if (处理后色温 == "Yellow") { labelFormat.SubStrings.SetSubString("FXK-Yellow", "实.png"); }
                                        else { labelFormat.SubStrings.SetSubString("Other", 处理后色温); }

                                        string pattern1 = @"^(\w+-\w+-\w+)";
                                        Match match1 = Regex.Match(cpxxBox.Text, pattern1);
                                        if (match1.Success)
                                        {
                                            string artNo = match1.Groups[1].Value;
                                            if (artNo.Contains("F23")) { labelFormat.SubStrings.SetSubString("IPDJ", " "); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                            else if (artNo.Contains("F16")) { labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                                            else if (artNo.Contains("3525")) { labelFormat.SubStrings.SetSubString("IPDJ", "IP67"); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                            else { labelFormat.SubStrings.SetSubString("IPDJ", "IP68"); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                                        }
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
                                    //if (checkBox_客户型号.Checked)
                                    if (!string.IsNullOrEmpty(i2Data))
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

                                        if (comboBox_标签规格.Text.Contains("13013"))
                                        {
                                            name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        }
                                        //else if (comboBox_标签规格.Text.Contains("UL")) { name_CPXXBox.Text = output_name  + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                    {
                                        name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }

                                    if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && !comboBox_标签规格.Text.Contains("直发"))
                                    {
                                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                                        output_长度 = "Length: " + h2Data;
                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                        else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && comboBox_标签规格.Text.Contains("直发"))
                                    {
                                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                                        output_长度 = "Length: " + h2Data;
                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                        else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                                    {
                                        output_长度 = "Length: " + h2Data;
                                        if (cpxxBox.Text.Contains("Ra90"))
                                        {
                                            output_灯数 = "CRI: " + "90";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else if (cpxxBox.Text.Contains("Ra95"))
                                        {
                                            output_灯数 = "CRI: " + "95";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else if (cpxxBox.Text.Contains("Ra85"))
                                        {
                                            output_灯数 = "CRI: " + "85";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else
                                        {
                                            output_灯数 = "CRI: " + "80";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                    }
                                    else if (comboBox_标签规格.Text.Contains("18395"))
                                    {
                                        output_灯带型号 = output_灯带型号.Replace("ART. No.:", "Product Code:");
                                        output_长度 = "Length:" + h2Data;
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
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
                                    //if (checkBox_客户型号.Checked)
                                    if (!string.IsNullOrEmpty(i2Data))
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

                                        if (comboBox_标签规格.Text.Contains("13013"))
                                        {
                                            name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        }
                                        //else if (comboBox_标签规格.Text.Contains("UL")) { name_CPXXBox.Text = output_name  + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                        //name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                    {
                                        name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }

                                    if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && !comboBox_标签规格.Text.Contains("直发"))
                                    {
                                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                                        output_长度 = "Length: " + h2Data;
                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                        else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && comboBox_标签规格.Text.Contains("直发"))
                                    {
                                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                                        output_长度 = "Length: " + h2Data;
                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                        else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                                    {
                                        output_长度 = "Length: " + h2Data;
                                        if (cpxxBox.Text.Contains("Ra90"))
                                        {
                                            output_灯数 = "CRI: " + "90";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else if (cpxxBox.Text.Contains("Ra95"))
                                        {
                                            output_灯数 = "CRI: " + "95";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else if (cpxxBox.Text.Contains("Ra85"))
                                        {
                                            output_灯数 = "CRI: " + "85";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else
                                        {
                                            output_灯数 = "CRI: " + "80";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                    }
                                    else if (comboBox_标签规格.Text.Contains("18395"))
                                    {
                                        output_灯带型号 = output_灯带型号.Replace("ART. No.:", "Product Code:");
                                        output_长度 = "Length:" + h2Data;
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
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
                                    //if (checkBox_客户型号.Checked)
                                    if (!string.IsNullOrEmpty(i2Data))
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

                                        if (comboBox_标签规格.Text.Contains("13013"))
                                        {
                                            name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        }
                                        //else if (comboBox_标签规格.Text.Contains("UL")) { name_CPXXBox.Text = output_name  + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                        //name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                    {
                                        name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, h2Data);
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && !comboBox_标签规格.Text.Contains("直发"))
                                    {
                                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                                        output_长度 = "Length: " + h2Data;
                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                        else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && comboBox_标签规格.Text.Contains("直发"))
                                    {
                                        labelFormat.SubStrings.SetSubString("FXK", 复选框);
                                        output_长度 = "Length: " + h2Data;
                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                        else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                    }
                                    else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                                    {
                                        output_长度 = "Length: " + h2Data;
                                        if (cpxxBox.Text.Contains("Ra90"))
                                        {
                                            output_灯数 = "CRI: " + "90";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else if (cpxxBox.Text.Contains("Ra95"))
                                        {
                                            output_灯数 = "CRI: " + "95";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else if (cpxxBox.Text.Contains("Ra85"))
                                        {
                                            output_灯数 = "CRI: " + "85";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                        else
                                        {
                                            output_灯数 = "CRI: " + "80";
                                            if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                            else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                            labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                        }
                                    }
                                    else if (comboBox_标签规格.Text.Contains("18395"))
                                    {
                                        output_灯带型号 = output_灯带型号.Replace("ART. No.:", "Product Code:");
                                        output_长度 = "Length:" + h2Data;
                                        name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
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
                        if (comboBox_标签规格.Text.Contains("直发"))
                        {
                            BPrefixContent = cpxxBox.Text;
                        }

                        //2025.1.24增加简化版工字标
                        if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                        {
                            //简化版工字标状态
                            if (BPrefixContent.Contains("三面发光"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", 灯带系列 + @"T");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }
                            else if (BPrefixContent.Contains("高亮"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "BH");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }
                            else if (BPrefixContent.Contains("翻边"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "BF");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }
                            else if (BPrefixContent.Contains("DTW"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "DTW");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }
                            else if (灯带系列 == "D")
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "D");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }
                            else
                            {
                                // 如果没有找到上述任何关键字，则设置为空
                                labelFormat.SubStrings.SetSubString("XZ", string.Empty);
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK", "空.png");
                                labelFormat.SubStrings.SetSubString("FXK-2", "空.png");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }

                            if (comboBox_标签规格.Text.Contains("13013"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }
                        }
                        else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标"))
                        {
                            //常规状态
                            // 检查是否存在 "Ra90" 或 "Ra95"
                            bool containsRa90 = BPrefixContent.Contains("Ra90");
                            bool containsRa95 = BPrefixContent.Contains("Ra95");

                            //简化版品名标状态
                            if (BPrefixContent.Contains("三面发光"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", 灯带系列 + @"T");
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
                            else if (BPrefixContent.Contains("高亮"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "BH");
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
                            else if (BPrefixContent.Contains("翻边"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "BF");
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
                            else if (BPrefixContent.Contains("DTW"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "DTW");
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
                            else if (BPrefixContent.Contains("Ra85"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ", "Ra85");
                                labelFormat.SubStrings.SetSubString("FXK-2", "实.png");
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");

                                if (comboBox_标签规格.Text.Contains("13013"))
                                {
                                    labelFormat.SubStrings.SetSubString("XZ", " ");
                                    labelFormat.SubStrings.SetSubString("FXK-2", "空.png");
                                }
                                else if (comboBox_标签规格.Text.Contains("17100"))
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
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
                                else if (comboBox_标签规格.Text.Contains("17100"))
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
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
                                else if (comboBox_标签规格.Text.Contains("17100"))
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                }
                            }
                            else
                            {
                                // 如果没有找到上述任何关键字，则设置为空
                                labelFormat.SubStrings.SetSubString("XZ", string.Empty);
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                //labelFormat.SubStrings.SetSubString("FXK", "空.png");
                                labelFormat.SubStrings.SetSubString("FXK-2", "空.png");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }

                            if (comboBox_标签规格.Text.Contains("13013"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }
                        }
                        else
                        {
                            //常规状态
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
                                else if (comboBox_标签规格.Text.Contains("17100"))
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
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
                                else if (comboBox_标签规格.Text.Contains("17100"))
                                {
                                    labelFormat.SubStrings.SetSubString("FXK-3", "实.png");
                                }
                            }
                            else
                            {
                                // 如果没有找到上述任何关键字，则设置为空
                                labelFormat.SubStrings.SetSubString("XZ", string.Empty);
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                //labelFormat.SubStrings.SetSubString("FXK", "空.png");
                                labelFormat.SubStrings.SetSubString("FXK-2", "空.png");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }

                            if (comboBox_标签规格.Text.Contains("13013"))
                            {
                                labelFormat.SubStrings.SetSubString("XZ-2", " ");
                                labelFormat.SubStrings.SetSubString("FXK-3", "空.png");
                            }
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

                        if (comboBox_标签规格.Text.Contains("12115"))
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

                            // 创建一个预览窗体
                            Form previewForm = new Form
                            {
                                Text = "标签预览",
                                Size = new Size(500, 400),
                                BackColor = Color.White,
                                FormBorderStyle = FormBorderStyle.Sizable, // 允许调整大小
                                MaximizeBox = true,
                                MinimizeBox = true,
                                TopMost = false  // 不置顶
                            };

                            // 获取主窗体位置和大小
                            Form mainForm = this.FindForm();

                            // 计算预览窗体位置 - 放在主窗体右侧
                            // 如果主窗体靠近屏幕右边缘，则放在左侧
                            Screen currentScreen = Screen.FromControl(mainForm);
                            int mainFormRight = mainForm.Location.X + mainForm.Width;
                            int availableRightSpace = currentScreen.WorkingArea.Right - mainFormRight;

                            int x, y;
                            if (availableRightSpace >= previewForm.Width + 10) // 右侧有足够空间
                            {
                                // 放在主窗体右侧
                                x = mainFormRight + 10; // 留出10像素间距
                                y = mainForm.Location.Y;
                            }
                            else // 右侧空间不足
                            {
                                // 放在屏幕右上角
                                x = currentScreen.WorkingArea.Right - previewForm.Width;
                                y = currentScreen.WorkingArea.Top;
                            }

                            previewForm.Location = new Point(x, y);

                            // 创建一个PictureBox来显示标签预览 - 固定大小和位置
                            PictureBox pictureBox = new PictureBox
                            {
                                Size = new Size(400, 300),
                                Location = new Point(50, 50),
                                BorderStyle = BorderStyle.FixedSingle,
                                SizeMode = PictureBoxSizeMode.Zoom,
                                BackColor = Color.White
                            };

                            // 添加一个关闭按钮 - 放置在左上角，固定位置
                            Button closeButton = new Button
                            {
                                Text = "关闭预览",
                                Location = new Point(10, 10),
                                Size = new Size(100, 30)
                            };
                            closeButton.Click += (s, args) => previewForm.Close();

                            // 添加上一个和下一个按钮
                            Button prevButton = new Button
                            {
                                Text = "上一个",
                                Location = new Point(120, 10),
                                Size = new Size(80, 30)
                            };

                            Button nextButton = new Button
                            {
                                Text = "下一个",
                                Location = new Point(210, 10),
                                Size = new Size(80, 30)
                            };

                            // 添加行号指示器
                            Label rowIndicator = new Label
                            {
                                Text = "行: 2",
                                Location = new Point(300, 15),
                                Size = new Size(100, 20),
                                TextAlign = ContentAlignment.MiddleLeft
                            };

                            // 当前行索引和总行数
                            int currentRowIndex = 2; // 从第2行开始
                            int totalRows = 2; // 默认值，稍后会更新

                            // 组装界面
                            previewForm.Controls.Add(pictureBox);
                            previewForm.Controls.Add(closeButton);
                            previewForm.Controls.Add(prevButton);
                            previewForm.Controls.Add(nextButton);
                            previewForm.Controls.Add(rowIndicator);

                            // 初始化图像为空
                            pictureBox.Image = null;

                            // 获取Excel总行数
                            if (!string.IsNullOrEmpty(Box_数据库.Text))
                            {
                                try
                                {
                                    using (var package = new ExcelPackage(new FileInfo(Box_数据库.Text)))
                                    {
                                        var worksheet = package.Workbook.Worksheets["Sheet1"];
                                        totalRows = worksheet.Dimension.Rows;
                                        rowIndicator.Text = $"行: {currentRowIndex}/{totalRows}";
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"读取Excel文件失败: {ex.Message}", "错误");
                                }
                            }

                            // 加载指定行数据并更新预览
                            Action<int> loadRowData = (rowIndex) =>
                            {
                                try
                                {
                                    if (!string.IsNullOrEmpty(Box_数据库.Text))
                                    {
                                        // 保存当前行索引
                                        currentRowIndex = rowIndex;

                                        // 更新行指示器
                                        rowIndicator.Text = $"行: {currentRowIndex}/{totalRows}";

                                        // 启用/禁用按钮
                                        prevButton.Enabled = currentRowIndex > 2;
                                        nextButton.Enabled = currentRowIndex < totalRows;

                                        // 读取指定行的数据
                                        using (var package = new ExcelPackage(new FileInfo(Box_数据库.Text)))
                                        {
                                            var worksheet = package.Workbook.Worksheets["Sheet1"];

                                            // 读取当前行的数据
                                            string a2Data = worksheet.Cells[$"A{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string b2Data = worksheet.Cells[$"B{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string c2Data = worksheet.Cells[$"C{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string d2Data = worksheet.Cells[$"D{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string e2Data = worksheet.Cells[$"E{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string f2Data = worksheet.Cells[$"F{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string g2Data = worksheet.Cells[$"G{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string h2Data = worksheet.Cells[$"H{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string i2Data = worksheet.Cells[$"I{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string h1Data = worksheet.Cells["H1"].Value?.ToString() ?? string.Empty;
                                            string j2Data = worksheet.Cells[$"J{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string k2Data = worksheet.Cells[$"K{rowIndex}"].Value?.ToString() ?? string.Empty;
                                            string l2Data = worksheet.Cells[$"L{rowIndex}"].Value?.ToString() ?? string.Empty;

                                            // 生成预览图
                                            if (labelFormat != null)
                                            {
                                                labelFormat.ExportImageToFile(_bmp_path, ImageType.BMP, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(407, 407), OverwriteOptions.Overwrite);
                                                if (pictureBox.Image != null)
                                                {
                                                    pictureBox.Image.Dispose();
                                                }
                                                System.Drawing.Image image = System.Drawing.Image.FromFile(_bmp_path);
                                                Bitmap NmpImage = new Bitmap(image);
                                                pictureBox.Image = NmpImage;
                                                image.Dispose();
                                            }
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"加载行数据错误: {ex.Message}", "操作提示");
                                }
                            };

                            // 设置按钮点击事件
                            prevButton.Click += (s, args) =>
                            {
                                if (currentRowIndex > 2)
                                {
                                    loadRowData(currentRowIndex - 1);
                                }
                            };

                            nextButton.Click += (s, args) =>
                            {
                                if (currentRowIndex < totalRows)
                                {
                                    loadRowData(currentRowIndex + 1);
                                }
                            };

                            // 显示当前标签预览
                            if (labelFormat != null)
                            {
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

                            // 显示预览窗体 - 使用Show()而不是ShowDialog()
                            previewForm.Show();

                            // 创建一个定时器用于自动关闭和更新标题
                            System.Windows.Forms.Timer autoCloseTimer = new System.Windows.Forms.Timer();
                            autoCloseTimer.Interval = 1000; // 设置间隔为 1000 毫秒 (1 秒) 以便更新标题
                            int remainingSeconds = 30; // 倒计时总秒数
                            string originalTitle = previewForm.Text; // 保存原始标题

                            // 定义定时器触发时的事件处理程序
                            autoCloseTimer.Tick += (timerSender, timerArgs) =>
                            {
                                remainingSeconds--; // 秒数减 1

                                if (remainingSeconds > 0)
                                {
                                    // 更新标题显示倒计时
                                    if (previewForm != null && !previewForm.IsDisposed)
                                    {
                                        previewForm.Text = $"{originalTitle} ({remainingSeconds} 秒后自动关闭)";
                                    }
                                }
                                else
                                {
                                    // 时间到，停止定时器并关闭窗口
                                    autoCloseTimer.Stop();
                                    if (previewForm != null && !previewForm.IsDisposed)
                                    {
                                        previewForm.Close(); // 关闭窗口会触发 FormClosed 事件
                                    }
                                    else
                                    {
                                        // 如果窗口已关闭或释放，只需释放定时器
                                        autoCloseTimer.Dispose();
                                    }
                                }
                            };

                            // 添加窗体关闭时的资源清理
                            previewForm.FormClosed += (s, args) =>
                            {
                                // 清理 PictureBox 图片资源
                                if (pictureBox.Image != null)
                                {
                                    pictureBox.Image.Dispose();
                                    pictureBox.Image = null;
                                }
                                // 确保定时器被停止和释放
                                if (autoCloseTimer != null)
                                {
                                    if (autoCloseTimer.Enabled)
                                    {
                                        autoCloseTimer.Stop();
                                    }
                                    autoCloseTimer.Dispose();
                                }
                                // (可选) 可以在这里尝试恢复标题，但窗口即将销毁，通常没必要
                                // if (previewForm != null) previewForm.Text = originalTitle;
                            };

                            // 启动定时器前先更新一次标题
                            previewForm.Text = $"{originalTitle} ({remainingSeconds} 秒后自动关闭)";
                            // 启动定时器
                            autoCloseTimer.Start();

                            break;

                        //MessageBox.Show($"执行预览操作时的值：\n\noutput_name1: {output_name1}\noutput_灯带型号1: {output_灯带型号1}", "生成预览图时的值");

                        //pictureBox.Image = null;
                        //if (labelFormat != null)
                        //{
                        //    //MessageBox.Show(_bmp_path, "操作提示");
                        //    //Generate a thumbnail for it.
                        //    labelFormat.ExportImageToFile(_bmp_path, ImageType.BMP, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(407, 407), OverwriteOptions.Overwrite);
                        //    System.Drawing.Image image = System.Drawing.Image.FromFile(_bmp_path);
                        //    Bitmap NmpImage = new Bitmap(image);
                        //    pictureBox.Image = NmpImage;
                        //    image.Dispose();
                        //}
                        //else
                        //{
                        //    MessageBox.Show("生成图片错误", "操作提示");
                        //}

                        //2023.01.24前能正常使用预览功能代码
                        //pictureBox.Image = null;
                        //if (labelFormat != null)
                        //{
                        //    MessageBox.Show(_bmp_path, "操作提示");
                        //    //Generate a thumbnail for it.
                        //    labelFormat.ExportImageToFile(_bmp_path, ImageType.BMP, Seagull.BarTender.Print.ColorDepth.ColorDepth24bit, new Resolution(407, 407), OverwriteOptions.Overwrite);
                        //    System.Drawing.Image image = System.Drawing.Image.FromFile(_bmp_path);
                        //    Bitmap NmpImage = new Bitmap(image);
                        //    pictureBox.Image = NmpImage;
                        //    image.Dispose();
                        //}
                        //else
                        //{
                        //    MessageBox.Show("生成图片错误", "操作提示");
                        //}

                        //另存为
                        case biaoqian.lingcun:

                            //MessageBox.Show($"执行另存为操作时的值：\n\noutput_name1: {output_name1}\noutput_灯带型号1: {output_灯带型号1}", "另存为时的值");

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
                                        var jData = worksheet.Cells[row, 10].Value?.ToString() ?? string.Empty;
                                        var kData = worksheet.Cells[row, 11].Value?.ToString() ?? string.Empty;
                                        var lData = worksheet.Cells[row, 12].Value?.ToString() ?? string.Empty;

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

                                                // 处理12098规格的情况
                                                if (comboBox_标签规格.Text.Contains("12098"))
                                                {
                                                    string inputText = iData;
                                                    string convertedName = "";
                                                    // 1. 处理A系列
                                                    if (灯带系列 == "A")
                                                    {
                                                        // 去掉RGB-前缀
                                                        string baseText = inputText.StartsWith("RGB-") ? inputText.Substring(4) : inputText;

                                                        // 处理SNX开头的情况
                                                        if (baseText.StartsWith("SNX-"))
                                                        {
                                                            var nameParts = baseText.Split('-');
                                                            convertedName = "SUPER-NEON-";

                                                            // 处理X-FLAT或X-DOME
                                                            if (nameParts.Length >= 2)
                                                            {
                                                                switch (nameParts[1])
                                                                {
                                                                    case "F":
                                                                        convertedName += "X-FLAT";
                                                                        break;

                                                                    case "D":
                                                                        convertedName += "X-DOME";
                                                                        break;

                                                                    default:
                                                                        convertedName += "X-" + nameParts[1];
                                                                        break;
                                                                }

                                                                // 添加中间部分，处理倒数第二部分的颜色缩写
                                                                for (int i = 2; i < nameParts.Length; i++)
                                                                {
                                                                    if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                                                                    {
                                                                        string color = ConvertColorAbbreviation(nameParts[i]);
                                                                        convertedName += "-" + color;
                                                                    }
                                                                    else
                                                                    {
                                                                        convertedName += "-" + nameParts[i];
                                                                    }
                                                                }
                                                            }

                                                            // 最后加上RGB-前缀
                                                            convertedName = "RGB-" + convertedName;
                                                        }
                                                        // 处理SNE开头的情况
                                                        else if (baseText.StartsWith("SNE-"))
                                                        {
                                                            var nameParts = baseText.Split('-');
                                                            convertedName = "SUPER-NEON-EDGE";

                                                            // 从第二个部分开始添加，处理倒数第二部分的颜色缩写
                                                            for (int i = 1; i < nameParts.Length; i++)
                                                            {
                                                                if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                                                                {
                                                                    string color = ConvertColorAbbreviation(nameParts[i]);
                                                                    convertedName += "-" + color;
                                                                }
                                                                else
                                                                {
                                                                    convertedName += "-" + nameParts[i];
                                                                }
                                                            }

                                                            // 最后加上RGB-前缀
                                                            convertedName = "RGB-" + convertedName;
                                                        }
                                                        else
                                                        {
                                                            convertedName = inputText; // 如果既不是SNX也不是SNE开头，保持原样
                                                        }
                                                    }
                                                    // 2. 处理SNX开头的情况
                                                    else if (inputText.StartsWith("SNX-"))
                                                    {
                                                        var nameParts = inputText.Split('-');
                                                        convertedName = "SUPER-NEON-";

                                                        // 处理X-FLAT或X-DOME
                                                        if (nameParts.Length >= 2)
                                                        {
                                                            switch (nameParts[1])
                                                            {
                                                                case "F":
                                                                    convertedName += "X-FLAT";
                                                                    break;

                                                                case "D":
                                                                    convertedName += "X-DOME";
                                                                    break;

                                                                default:
                                                                    convertedName += nameParts[1];
                                                                    break;
                                                            }

                                                            // 添加剩余部分，处理颜色缩写
                                                            for (int i = 2; i < nameParts.Length; i++)
                                                            {
                                                                if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                                                                {
                                                                    string color = ConvertColorAbbreviation(nameParts[i]);
                                                                    convertedName += "-" + color;
                                                                }
                                                                else
                                                                {
                                                                    convertedName += "-" + nameParts[i];
                                                                }
                                                            }
                                                        }
                                                    }
                                                    // 3. 处理SNE开头的情况
                                                    else if (inputText.StartsWith("SNE-"))
                                                    {
                                                        var nameParts = inputText.Split('-');
                                                        convertedName = "SUPER-NEON-EDGE";

                                                        // 从第二个部分开始添加，处理颜色缩写
                                                        for (int i = 1; i < nameParts.Length; i++)
                                                        {
                                                            if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                                                            {
                                                                string color = ConvertColorAbbreviation(nameParts[i]);
                                                                convertedName += "-" + color;
                                                            }
                                                            else
                                                            {
                                                                convertedName += "-" + nameParts[i];
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        convertedName = inputText; // 默认情况
                                                    }

                                                    output_name = convertedName;
                                                    output_灯带型号 = "Short SKU:" + iData;
                                                    name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + "BATCH " + bData;
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("16008"))
                                                {
                                                    output_灯带型号 = iData;
                                                    string 长度1 = "Length: " + hData;
                                                    string po1 = "Lot #:  " + bData;
                                                    name_CPXXBox.Text = "www.SGiLighting.com" + "\n" + "LED NEON FLEX LIGHT" + "\n" + output_灯带型号 + "\n" + 长度1 + "\n" + po1;
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("12251") && 标签种类_comboBox.Text.Contains("工字标"))
                                                {
                                                    string chazhaoziliao = iData;
                                                    string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12251资料.xlsx";

                                                    try
                                                    {
                                                        using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                                        {
                                                            var worksheet1 = package1.Workbook.Worksheets[0];
                                                            int rowCount1 = worksheet1.Dimension.Rows;
                                                            string fColumnContent = "";
                                                            bool found = false;

                                                            // 遍历每一行
                                                            for (int row1 = 1; row1 <= rowCount1; row1++)
                                                            {
                                                                // 检查A到D列
                                                                for (int col = 1; col <= 4; col++) // 1=A, 2=B, 3=C, 4=D
                                                                {
                                                                    string cellValue = worksheet1.Cells[row1, col].Text;

                                                                    if (cellValue == chazhaoziliao)
                                                                    {
                                                                        // 找到匹配项，获取同行F列的内容
                                                                        fColumnContent = worksheet1.Cells[row1, 6].Text; // F列的内容
                                                                        name_CPXXBox.Text = fColumnContent;
                                                                        found = true;

                                                                        // 可以添加一个消息框显示在哪里找到的（如果需要）
                                                                        //MessageBox.Show($"在第{row1}行，第{(char)(col + 64)}列找到匹配项", "查找结果");

                                                                        break;
                                                                    }
                                                                }

                                                                if (found) break; // 如果找到了就退出外层循环
                                                            }

                                                            if (string.IsNullOrEmpty(fColumnContent))
                                                            {
                                                                MessageBox.Show($"在A、B、C、D列中均未找到匹配内容: {chazhaoziliao}", "查找结果");
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                                                    }

                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    labelFormat.SubStrings.SetSubString("CPCD", hData);
                                                    labelFormat.SubStrings.SetSubString("CPXX-2", iData);
                                                    if (灯带材质 == "FR") { labelFormat.SubStrings.SetSubString("IPDJ", " "); }
                                                    else { labelFormat.SubStrings.SetSubString("IPDJ", "Not suitable for underwater use"); }
                                                }
                                                else if (comboBox_标签规格.Text.Contains("12251") && 标签种类_comboBox.Text.Contains("品名标"))
                                                {
                                                    string chazhaoziliao = iData;
                                                    string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12251资料.xlsx";

                                                    try
                                                    {
                                                        using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                                        {
                                                            var worksheet1 = package1.Workbook.Worksheets[0];
                                                            int rowCount1 = worksheet1.Dimension.Rows;
                                                            string fColumnContent = "";
                                                            bool found = false;

                                                            // 遍历每一行
                                                            for (int row1 = 1; row1 <= rowCount1; row1++)
                                                            {
                                                                // 检查A到D列
                                                                for (int col = 1; col <= 4; col++) // 1=A, 2=B, 3=C, 4=D
                                                                {
                                                                    string cellValue = worksheet1.Cells[row1, col].Text;

                                                                    if (cellValue == chazhaoziliao)
                                                                    {
                                                                        fColumnContent = worksheet1.Cells[row1, 5].Text; //
                                                                        name_CPXXBox.Text = fColumnContent;
                                                                        found = true;

                                                                        // 可以添加一个消息框显示在哪里找到的（如果需要）
                                                                        //MessageBox.Show($"在第{row1}行，第{(char)(col + 64)}列找到匹配项", "查找结果");

                                                                        break;
                                                                    }
                                                                }

                                                                if (found) break; // 如果找到了就退出外层循环
                                                            }

                                                            if (string.IsNullOrEmpty(fColumnContent))
                                                            {
                                                                MessageBox.Show($"在A、B、C、D列中均未找到匹配内容: {chazhaoziliao}", "查找结果");
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                                                    }

                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    labelFormat.SubStrings.SetSubString("CPCD", hData);
                                                    labelFormat.SubStrings.SetSubString("CPXX-2", iData);
                                                    string 处理后色温 = output_色温.Replace("Color: ", "");
                                                    string 处理后功率 = output_功率.Replace("Rated Power: ", ""); // 删除前缀
                                                    int 斜杠位置 = 处理后功率.IndexOf("/");
                                                    if (斜杠位置 != -1)
                                                    {
                                                        处理后功率 = 处理后功率.Substring(0, 斜杠位置); // 只保留斜杠前的部分
                                                    }
                                                    labelFormat.SubStrings.SetSubString("WS", 处理后功率);
                                                    labelFormat.SubStrings.SetSubString("PO", jData);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("标签型号"))
                                                {
                                                    //string cpxx_text = cpxxBox.Text;
                                                    //判断产品信息(cpxxBox.Text);

                                                    string cz型号 = output_灯带型号.Replace("ART. No.: ", "");
                                                    string cz电压 = output_电压.Replace("Rated Voltage: DC ", "");
                                                    string cz色温 = output_色温.Replace("Color: ", "");
                                                    string cz色温1 = string.Empty;      //只保留数字色温
                                                    string cz色温2 = string.Empty;      //转换色温
                                                    string cz灯数 = new string(output_灯数.Replace("LED Qty.: ", "").Where(char.IsDigit).ToArray());
                                                    string cz功率 = string.Empty;

                                                    string patternX1 = @"^(\w+-\w+-\w+)";
                                                    Match matchX1 = Regex.Match(cpxxBox.Text, patternX1);
                                                    if (matchX1.Success) { cz型号 = matchX1.Groups[1].Value; }

                                                    cz色温1 = cz色温.Replace("K", ""); ;

                                                    if (cz色温.Contains("RGBW")) { cz色温2 = "RGBW"; }
                                                    else if (cz色温.Contains("K") || cz色温.Contains("k") && !cz色温.Contains("RGBW"))
                                                    {
                                                        if (cz色温.Contains("~"))
                                                        {
                                                            cz色温2 = cz色温.Replace("K", "");
                                                        }
                                                        else
                                                        {
                                                            // 提取数字部分
                                                            string 数字部分 = new string(cz色温.Where(char.IsDigit).ToArray());
                                                            if (int.TryParse(数字部分, out int 色温值))
                                                            {
                                                                if (色温值 >= 1100 && 色温值 < 11111111) { cz色温2 = "白光"; }
                                                                else if (色温值 > 11111111)
                                                                {
                                                                    // 将数字转换为字符串
                                                                    string 色温字符串 = 色温值.ToString();

                                                                    // 检查长度是否足够
                                                                    if (色温字符串.Length >= 8)  // 确保有足够的数字
                                                                    {
                                                                        // 在第4位后插入~
                                                                        string 前半部分 = 色温字符串.Substring(0, 4);  // 取前4位
                                                                        string 后半部分 = 色温字符串.Substring(4);     // 取剩余部分
                                                                        cz色温2 = $"{前半部分}~{后半部分}";  // 组合结果
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                    else if (cz色温.Contains("RGB") && !cz色温.Contains("RGBW")) { cz色温2 = "RGB"; }
                                                    else if (cz色温.Contains("Red")) { cz色温2 = "红"; }
                                                    else if (cz色温.Contains("Blue")) { cz色温2 = "蓝"; }
                                                    else if (cz色温.Contains("Green")) { cz色温2 = "绿"; }
                                                    else if (cz色温.Contains("Orange")) { cz色温2 = "橙"; }
                                                    else if (cz色温.Contains("Yellow")) { cz色温2 = "黄"; }
                                                    else if (cz色温.Contains("Amber")) { cz色温2 = "琥珀"; }

                                                    //UL标12291
                                                    if (comboBox_标签规格.Text.Contains("12291"))
                                                    {
                                                        //string matchingCriteria = $"匹配条件：\n型号: {cz型号}\n电压: {cz电压}\n色温: {cz色温2}\n灯数: {cz灯数}";
                                                        //MessageBox.Show(matchingCriteria, "UL标签12291匹配条件");

                                                        string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12291 UL资料.xlsx";

                                                        using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                                        {
                                                            var worksheet1 = package1.Workbook.Worksheets[0];
                                                            int rowCount1 = worksheet1.Dimension.Rows;
                                                            bool found = false;

                                                            // 遍历每一行，只检查B列
                                                            for (int row1 = 2; row1 <= rowCount1; row1++)
                                                            {
                                                                string cellValue = worksheet1.Cells[row1, 2].Text; // 只读取B列的内容
                                                                if (cpxxBox.Text.Contains("【DMX】"))
                                                                {
                                                                    if (cellValue.Contains(cz型号) && cellValue.Contains(cz电压) && cellValue.Contains(cz色温2) && cellValue.Contains(cz灯数) && cellValue.Contains("DMX"))
                                                                    {
                                                                        // 找到匹配项，获取同行的内容
                                                                        string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                                                        string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                                                        output_name1 = "Name: " + DColumnContent;
                                                                        output_灯带型号1 = "ART. No.: " + CColumnContent;

                                                                        found = true;
                                                                        break; // 找到后立即退出循环
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    if (cellValue.Contains(cz型号) && cellValue.Contains(cz电压) && cellValue.Contains(cz色温2) && cellValue.Contains(cz灯数))
                                                                    {
                                                                        // 找到匹配项，获取同行的内容
                                                                        string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                                                        string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                                                        output_name1 = "Name: " + DColumnContent;
                                                                        output_灯带型号1 = "ART. No.: " + CColumnContent;

                                                                        found = true;
                                                                        break; // 找到后立即退出循环
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                    //UL标17021，只需要核对灯带型号
                                                    else if (comboBox_标签规格.Text.Contains("17021"))
                                                    {
                                                        string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\17021 UL资料.xlsx";

                                                        using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                                        {
                                                            var worksheet1 = package1.Workbook.Worksheets[0];
                                                            int rowCount1 = worksheet1.Dimension.Rows;
                                                            bool found = false;

                                                            // 遍历每一行，只检查B列
                                                            for (int row1 = 2; row1 <= rowCount1; row1++)
                                                            {
                                                                string cellValue = worksheet1.Cells[row1, 2].Text; // 只读取B列的内容

                                                                if (cellValue.Contains(cz型号))
                                                                {
                                                                    // 找到匹配项，获取同行的内容
                                                                    string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                                                    string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                                                    output_name1 = "Name: " + DColumnContent;
                                                                    output_灯带型号1 = "ART. No.: " + CColumnContent;

                                                                    found = true;
                                                                    break; // 找到后立即退出循环
                                                                }
                                                            }
                                                        }
                                                    }
                                                    output_长度 = "Length:" + hData;
                                                    name_CPXXBox.Text = output_name1 + "\n" + output_灯带型号1 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("12115"))
                                                {
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
                                                    //if (checkBox_客户型号.Checked)
                                                    if (!string.IsNullOrEmpty(iData))
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

                                                        if (comboBox_标签规格.Text.Contains("13013"))
                                                        {
                                                            name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                        }
                                                        //else if (comboBox_标签规格.Text.Contains("UL")) { name_CPXXBox.Text = output_name  + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                                        //name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                                    {
                                                        name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }

                                                    string 处理后条形码 = kData.Replace("EAN code: ", "");
                                                    labelFormat.SubStrings.SetSubString("TXM", 处理后条形码);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("12090"))
                                                {
                                                    name_CPXXBox.Text = output_功率 + "\n" + output_色温 + "\n" + output_剪切单元 + "\n" + "Rollengte:" + hData;
                                                    labelFormat.SubStrings.SetSubString("KHXH", iData);
                                                    labelFormat.SubStrings.SetSubString("KHMC", textBox_客户资料.Text);
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("12141"))
                                                {
                                                    // 获取灯带长度
                                                    output_灯带型号 = iData;
                                                    //string pattern12 = @"(\d+,\d+)m";
                                                    string pattern12 = @"(\d+,\d+)[Mm]";
                                                    Regex regex12 = new Regex(pattern12);
                                                    Match match = regex12.Match(output_灯带型号);
                                                    if (match.Success)
                                                    {
                                                        string lengthStr = match.Groups[1].Value; // 获取匹配到的长度值（如：3,58319）
                                                        output_灯带长度 = lengthStr + "m";
                                                        //MessageBox.Show($"提取到的长度：{output_灯带长度}");
                                                    }
                                                    output_灯带型号 = "ART. No.: " + iData;
                                                    // 处理功率：去除"Rated Power: "和"W/m"，只保留数字
                                                    string powerStr = output_功率.Replace("Rated Power: ", "")
                                                                               .Replace("W/m", "")
                                                                               .Trim();

                                                    // 处理长度：去除"m"单位，将逗号替换为小数点
                                                    string lengthStr1 = output_灯带长度.Replace("m", "")
                                                                                     .Replace(",", ".")
                                                                                     .Trim();

                                                    // 转换为double进行计算
                                                    if (double.TryParse(powerStr, out double power) &&
                                                        double.TryParse(lengthStr1, out double length))
                                                    {
                                                        // 计算总功率
                                                        double totalPower = power * length;

                                                        // 保留2位小数
                                                        output_总功率 = $"Total Power:{totalPower:F2}W";
                                                        //MessageBox.Show(output_总功率);
                                                    }
                                                    output_长度 = "Length of Light:" + output_灯带长度;
                                                    output_线材长度 = "Length of Cable: " + lData;
                                                    name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_总功率 + "\n" + output_光源型号 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_线材长度;
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                                    string 处理后名称 = output_name.Replace("Name: ", "");
                                                    string 处理后名称1 = 处理后名称.Replace(Environment.NewLine, "");
                                                    labelFormat.SubStrings.SetSubString("BSM-02", 处理后名称1);
                                                    labelFormat.SubStrings.SetSubString("PO", jData);

                                                    string 处理后色温 = output_色温.Replace("Color: ", "");
                                                    if (处理后色温 == "2700K") { labelFormat.SubStrings.SetSubString("FXK-2700K", "实.png"); }
                                                    else if (处理后色温 == "3000K") { labelFormat.SubStrings.SetSubString("FXK-3000K", "实.png"); }
                                                    else if (处理后色温 == "3500K") { labelFormat.SubStrings.SetSubString("FXK-3500K", "实.png"); }
                                                    else if (处理后色温 == "4000K") { labelFormat.SubStrings.SetSubString("FXK-4000K", "实.png"); }
                                                    else if (处理后色温 == "4500K") { labelFormat.SubStrings.SetSubString("FXK-4500K", "实.png"); }
                                                    else if (处理后色温 == "5700K") { labelFormat.SubStrings.SetSubString("FXK-5700K", "实.png"); }
                                                    else if (处理后色温 == "6500K") { labelFormat.SubStrings.SetSubString("FXK-6500K", "实.png"); }
                                                    else if (处理后色温 == "Red") { labelFormat.SubStrings.SetSubString("FXK-Red", "实.png"); }
                                                    else if (处理后色温 == "Blue") { labelFormat.SubStrings.SetSubString("FXK-Blue", "实.png"); }
                                                    else if (处理后色温 == "RGB") { labelFormat.SubStrings.SetSubString("FXK-RGB", "实.png"); }
                                                    else if (处理后色温 == "Green") { labelFormat.SubStrings.SetSubString("FXK-Green", "实.png"); }
                                                    else if (处理后色温 == "Yellow") { labelFormat.SubStrings.SetSubString("FXK-Yellow", "实.png"); }
                                                    else { labelFormat.SubStrings.SetSubString("Other", 处理后色温); }

                                                    string pattern1 = @"^(\w+-\w+-\w+)";
                                                    Match match1 = Regex.Match(cpxxBox.Text, pattern1);
                                                    if (match1.Success)
                                                    {
                                                        string artNo = match1.Groups[1].Value;
                                                        if (artNo.Contains("F23")) { labelFormat.SubStrings.SetSubString("IPDJ", " "); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                                        else if (artNo.Contains("F16")) { labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                                                        else if (artNo.Contains("3525")) { labelFormat.SubStrings.SetSubString("IPDJ", "IP67"); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                                        else { labelFormat.SubStrings.SetSubString("IPDJ", "IP68"); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                                                    }
                                                }
                                                else if (comboBox_标签规格.Text.Contains("12120"))
                                                {
                                                    output_灯带型号 = "ART. No.: " + iData;
                                                    name_CPXXBox.Text = "AMBIANCE LUMIERE" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + "Caution: Do not overload." + "\n" + output_尾巴;
                                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                                    if (灯带系列 == "A") { labelFormat.SubStrings.SetSubString("FXK-A", "实.png"); }
                                                    else if (灯带系列 == "B") { labelFormat.SubStrings.SetSubString("FXK-B", "实.png"); }
                                                    else if (灯带系列 == "E") { labelFormat.SubStrings.SetSubString("FXK-E", "实.png"); }
                                                    else if (灯带系列 == "S") { labelFormat.SubStrings.SetSubString("FXK-S", "实.png"); }

                                                    string 处理后色温 = output_色温.Replace("Color: ", "");
                                                    if (处理后色温 == "2700K") { labelFormat.SubStrings.SetSubString("FXK-2700K", "实.png"); }
                                                    else if (处理后色温 == "3000K") { labelFormat.SubStrings.SetSubString("FXK-3000K", "实.png"); }
                                                    else if (处理后色温 == "3500K") { labelFormat.SubStrings.SetSubString("FXK-3500K", "实.png"); }
                                                    else if (处理后色温 == "4000K") { labelFormat.SubStrings.SetSubString("FXK-4000K", "实.png"); }
                                                    else if (处理后色温 == "4500K") { labelFormat.SubStrings.SetSubString("FXK-4500K", "实.png"); }
                                                    else if (处理后色温 == "5700K") { labelFormat.SubStrings.SetSubString("FXK-5700K", "实.png"); }
                                                    else if (处理后色温 == "6500K") { labelFormat.SubStrings.SetSubString("FXK-6500K", "实.png"); }
                                                    else if (处理后色温 == "Red") { labelFormat.SubStrings.SetSubString("FXK-Red", "实.png"); }
                                                    else if (处理后色温 == "Blue") { labelFormat.SubStrings.SetSubString("FXK-Blue", "实.png"); }
                                                    else if (处理后色温 == "RGB") { labelFormat.SubStrings.SetSubString("FXK-RGB", "实.png"); }
                                                    else if (处理后色温 == "Green") { labelFormat.SubStrings.SetSubString("FXK-Green", "实.png"); }
                                                    else if (处理后色温 == "Yellow") { labelFormat.SubStrings.SetSubString("FXK-Yellow", "实.png"); }
                                                    else { labelFormat.SubStrings.SetSubString("Other", 处理后色温); }
                                                }
                                                else if (comboBox_标签规格.Text.Contains("13009"))
                                                {
                                                    string chazhaoziliao = iData;
                                                    string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\13009资料.xlsx";
                                                    try
                                                    {
                                                        using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                                                        {
                                                            var worksheet1 = package1.Workbook.Worksheets[0];
                                                            int rowCount1 = worksheet1.Dimension.Rows;
                                                            bool found = false;

                                                            // 遍历每一行，只检查A列
                                                            for (int row1 = 1; row1 <= rowCount1; row1++)
                                                            {
                                                                string cellValue = worksheet1.Cells[row1, 1].Text; // 只读取A列的内容

                                                                if (cellValue == chazhaoziliao)
                                                                {
                                                                    // 找到匹配项，获取同行F列的内容
                                                                    string BColumnContent = worksheet1.Cells[row1, 2].Text; // B列的内容,名称
                                                                    string CColumnContent = worksheet1.Cells[row1, 3].Text; // C列的内容,颜色（名称颜色中间）
                                                                    string DColumnContent = worksheet1.Cells[row1, 4].Text; // D列的内容,色温（CCT）
                                                                    string EColumnContent = worksheet1.Cells[row1, 5].Text; // E列的内容,流明(Lumen)
                                                                    string FColumnContent = worksheet1.Cells[row1, 6].Text; // F列的内容,功率(Wattage)
                                                                    string GColumnContent = worksheet1.Cells[row1, 7].Text; // G列的内容.条形码
                                                                    output_13009名称 = BColumnContent;
                                                                    output_13009颜色 = CColumnContent;
                                                                    output_13009色温 = "CCT:         " + DColumnContent;
                                                                    output_13009流明 = "Lumen:     " + EColumnContent;
                                                                    output_13009功率 = "Wattage:   " + FColumnContent;
                                                                    output_13009条形码 = GColumnContent;

                                                                    found = true;
                                                                    break; // 找到后立即退出循环
                                                                }
                                                            }
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {
                                                        MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                                                    }
                                                    output_长度 = "Quantity:   " + hData;
                                                    name_CPXXBox.Text = output_13009色温 + "\n" + output_13009流明 + "\n" + output_13009功率 + "\n" + output_电压 + "\n" + output_长度;
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    labelFormat.SubStrings.SetSubString("KHXH", iData);
                                                    labelFormat.SubStrings.SetSubString("KHMC", output_13009名称);
                                                    labelFormat.SubStrings.SetSubString("KHYS", output_13009颜色);
                                                    if (iData == "L65XT2004" || iData == "L65XT2006" || iData == "L65XT2008") { labelFormat.SubStrings.SetSubString("CRI80", "CRI80.png"); }
                                                    else { labelFormat.SubStrings.SetSubString("CRI80", "空.png"); }

                                                    if (标签种类_comboBox.Text.Contains("品名标")) { labelFormat.SubStrings.SetSubString("TXM", output_13009条形码); }
                                                    else { labelFormat.SubStrings.SetSubString("TXM", "4251158486185"); }
                                                }
                                                else if (comboBox_标签规格.Text.Contains("18395"))
                                                {
                                                    output_灯带型号 = output_灯带型号.Replace("ART. No.:", "Product Code:");
                                                    output_长度 = "Length:" + hData;
                                                    name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("12573") && comboBox_标签规格.Text.Contains("美国"))
                                                {
                                                    output_长度 = "Length:" + hData;
                                                    output_灯带型号 = "Model:" + iData;
                                                    name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else
                                                {
                                                    //原来的常规逻辑
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
                                                    //if (checkBox_客户型号.Checked)
                                                    if (!string.IsNullOrEmpty(iData))
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

                                                        if (comboBox_标签规格.Text.Contains("13013"))
                                                        {
                                                            name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                        }
                                                        //else if (comboBox_标签规格.Text.Contains("UL")) { name_CPXXBox.Text = output_name  + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                                        //name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                                    {
                                                        name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                        name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                }

                                                if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && !comboBox_标签规格.Text.Contains("直发"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && comboBox_标签规格.Text.Contains("直发"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (cpxxBox.Text.Contains("Ra90"))
                                                    {
                                                        output_灯数 = "CRI: " + "90";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (cpxxBox.Text.Contains("Ra95"))
                                                    {
                                                        output_灯数 = "CRI: " + "95";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (cpxxBox.Text.Contains("Ra85"))
                                                    {
                                                        output_灯数 = "CRI: " + "85";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else
                                                    {
                                                        output_灯数 = "CRI: " + "80";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
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

                                        //12141才有2号模板
                                        if (_wjm_ == "2.btw")
                                        {
                                            if (labelFormat != null)
                                            {
                                                labelFormat.SubStrings.SetSubString("XLH", aData);
                                                labelFormat.SubStrings.SetSubString("BSM-01", bData);
                                                labelFormat.SubStrings.SetSubString("BSM-02", cData);
                                                textBox1.Text = gData;
                                                if (comboBox_标签规格.Text.Contains("12141"))
                                                {
                                                    // 获取灯带长度
                                                    output_灯带型号 = iData;
                                                    //string pattern12 = @"(\d+,\d+)m";
                                                    string pattern12 = @"(\d+,\d+)[Mm]";
                                                    Regex regex12 = new Regex(pattern12);
                                                    Match match = regex12.Match(output_灯带型号);
                                                    if (match.Success)
                                                    {
                                                        string lengthStr = match.Groups[1].Value; // 获取匹配到的长度值（如：3,58319）
                                                        output_灯带长度 = lengthStr + "m";
                                                        //MessageBox.Show($"提取到的长度：{output_灯带长度}");
                                                    }
                                                    output_灯带型号 = "ART. No.: " + iData;
                                                    // 处理功率：去除"Rated Power: "和"W/m"，只保留数字
                                                    string powerStr = output_功率.Replace("Rated Power: ", "")
                                                                               .Replace("W/m", "")
                                                                               .Trim();

                                                    // 处理长度：去除"m"单位，将逗号替换为小数点
                                                    string lengthStr1 = output_灯带长度.Replace("m", "")
                                                                                     .Replace(",", ".")
                                                                                     .Trim();

                                                    // 转换为double进行计算
                                                    if (double.TryParse(powerStr, out double power) &&
                                                        double.TryParse(lengthStr1, out double length))
                                                    {
                                                        // 计算总功率
                                                        double totalPower = power * length;

                                                        // 保留2位小数
                                                        output_总功率 = $"Total Power:{totalPower:F2}W";
                                                        //MessageBox.Show(output_总功率);
                                                    }
                                                    output_长度 = "Length of Light:" + output_灯带长度;
                                                    output_线材长度 = "Length of Cable: " + lData;
                                                    name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_总功率 + "\n" + output_光源型号 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_线材长度;
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);

                                                    string 处理后名称 = output_name.Replace("Name: ", "");
                                                    string 处理后名称1 = 处理后名称.Replace(Environment.NewLine, "");
                                                    labelFormat.SubStrings.SetSubString("BSM-02", 处理后名称1);
                                                    labelFormat.SubStrings.SetSubString("PO", jData);

                                                    string 处理后色温 = output_色温.Replace("Color: ", "");
                                                    if (处理后色温 == "2700K") { labelFormat.SubStrings.SetSubString("FXK-2700K", "实.png"); }
                                                    else if (处理后色温 == "3000K") { labelFormat.SubStrings.SetSubString("FXK-3000K", "实.png"); }
                                                    else if (处理后色温 == "3500K") { labelFormat.SubStrings.SetSubString("FXK-3500K", "实.png"); }
                                                    else if (处理后色温 == "4000K") { labelFormat.SubStrings.SetSubString("FXK-4000K", "实.png"); }
                                                    else if (处理后色温 == "4500K") { labelFormat.SubStrings.SetSubString("FXK-4500K", "实.png"); }
                                                    else if (处理后色温 == "5700K") { labelFormat.SubStrings.SetSubString("FXK-5700K", "实.png"); }
                                                    else if (处理后色温 == "6500K") { labelFormat.SubStrings.SetSubString("FXK-6500K", "实.png"); }
                                                    else if (处理后色温 == "Red") { labelFormat.SubStrings.SetSubString("FXK-Red", "实.png"); }
                                                    else if (处理后色温 == "Blue") { labelFormat.SubStrings.SetSubString("FXK-Blue", "实.png"); }
                                                    else if (处理后色温 == "RGB") { labelFormat.SubStrings.SetSubString("FXK-RGB", "实.png"); }
                                                    else if (处理后色温 == "Green") { labelFormat.SubStrings.SetSubString("FXK-Green", "实.png"); }
                                                    else if (处理后色温 == "Yellow") { labelFormat.SubStrings.SetSubString("FXK-Yellow", "实.png"); }
                                                    else { labelFormat.SubStrings.SetSubString("Other", 处理后色温); }

                                                    string pattern1 = @"^(\w+-\w+-\w+)";
                                                    Match match1 = Regex.Match(cpxxBox.Text, pattern1);
                                                    if (match1.Success)
                                                    {
                                                        string artNo = match1.Groups[1].Value;
                                                        if (artNo.Contains("F23")) { labelFormat.SubStrings.SetSubString("IPDJ", " "); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                                        else if (artNo.Contains("F16")) { labelFormat.SubStrings.SetSubString("IPDJ", bq_ipdj); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                                                        else if (artNo.Contains("3525")) { labelFormat.SubStrings.SetSubString("IPDJ", "IP67"); labelFormat.SubStrings.SetSubString("CPCD", " "); }
                                                        else { labelFormat.SubStrings.SetSubString("IPDJ", "IP68"); labelFormat.SubStrings.SetSubString("CPCD", "1M"); }
                                                    }
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
                                                //if (checkBox_客户型号.Checked)
                                                if (!string.IsNullOrEmpty(iData))
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

                                                    if (comboBox_标签规格.Text.Contains("13013"))
                                                    {
                                                        name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    }
                                                    //else if (comboBox_标签规格.Text.Contains("UL")) { name_CPXXBox.Text = output_name  + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                                    //name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                                {
                                                    name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }

                                                if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && !comboBox_标签规格.Text.Contains("直发"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && comboBox_标签规格.Text.Contains("直发"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (cpxxBox.Text.Contains("Ra90"))
                                                    {
                                                        output_灯数 = "CRI: " + "90";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (cpxxBox.Text.Contains("Ra95"))
                                                    {
                                                        output_灯数 = "CRI: " + "95";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (cpxxBox.Text.Contains("Ra85"))
                                                    {
                                                        output_灯数 = "CRI: " + "85";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else
                                                    {
                                                        output_灯数 = "CRI: " + "80";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
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
                                                //if (checkBox_客户型号.Checked)
                                                if (!string.IsNullOrEmpty(iData))
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

                                                    if (comboBox_标签规格.Text.Contains("13013"))
                                                    {
                                                        name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    }
                                                    //else if (comboBox_标签规格.Text.Contains("UL")) { name_CPXXBox.Text = output_name  + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                                    //name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                                {
                                                    name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }

                                                if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && !comboBox_标签规格.Text.Contains("直发"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && comboBox_标签规格.Text.Contains("直发"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (cpxxBox.Text.Contains("Ra90"))
                                                    {
                                                        output_灯数 = "CRI: " + "90";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (cpxxBox.Text.Contains("Ra95"))
                                                    {
                                                        output_灯数 = "CRI: " + "95";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (cpxxBox.Text.Contains("Ra85"))
                                                    {
                                                        output_灯数 = "CRI: " + "85";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else
                                                    {
                                                        output_灯数 = "CRI: " + "80";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
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
                                                //if (checkBox_客户型号.Checked)
                                                if (!string.IsNullOrEmpty(iData))
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

                                                    if (comboBox_标签规格.Text.Contains("13013"))
                                                    {
                                                        name_CPXXBox.Text = "Totallux.nl" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    }
                                                    //else if (comboBox_标签规格.Text.Contains("UL")) { name_CPXXBox.Text = output_name  + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }

                                                    //name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (comboBox_标签规格.Text.Contains("UL") && comboBox_标签规格.Text.Contains("15019"))
                                                {
                                                    name_CPXXBox.Text = output_name + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
                                                    name_CPXXBox.Text = 重构产品信息_工字标(name_CPXXBox.Text, hData);
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }

                                                if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && !comboBox_标签规格.Text.Contains("直发"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("品名标") && comboBox_标签规格.Text.Contains("直发"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温 + "\n" + "\n" + output_透镜角度; }
                                                    else { name_CPXXBox.Text = output_name + "\n" + "\n" + output_灯带型号 + "\n" + "\n" + output_电压 + "\n" + "\n" + output_功率 + "\n" + "\n" + output_剪切单元 + "\n" + "\n" + output_长度 + "\n" + "\n" + output_色温; }
                                                    labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                }
                                                else if (checkBox_简化型号.Checked && 标签种类_comboBox.Text.Contains("工字标"))
                                                {
                                                    output_长度 = "Length: " + hData;
                                                    if (cpxxBox.Text.Contains("Ra90"))
                                                    {
                                                        output_灯数 = "CRI: " + "90";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (cpxxBox.Text.Contains("Ra95"))
                                                    {
                                                        output_灯数 = "CRI: " + "95";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else if (cpxxBox.Text.Contains("Ra85"))
                                                    {
                                                        output_灯数 = "CRI: " + "85";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_灯数 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
                                                    else
                                                    {
                                                        output_灯数 = "CRI: " + "80";
                                                        if (comboBox_标签规格.Text.Contains("3525")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴; }
                                                        else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴; }
                                                        labelFormat.SubStrings.SetSubString("CPXX", name_CPXXBox.Text);
                                                    }
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

        

        //判断复选框内容
        //private static string 判断复选框内容(string input, string 标签规格)
        private string 判断复选框内容(string input, string 标签规格)
        {
            //是简化型号的时候
            if (checkBox_简化型号.Checked)
            {
                // 检查是否有正弯或侧弯
                bool hasPositiveBend = input.Contains("正弯");
                bool hasSideBend = input.Contains("侧弯");

                string secondField = hasPositiveBend ? "正弯" : (hasSideBend ? "侧弯" : string.Empty);
                if (!string.IsNullOrEmpty(secondField))
                {
                    return $"{secondField}.png";
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
                    return $"{secondField}.png";
                }
                else
                {
                    //return $"{secondField}.png";
                    return "空.png";
                }
            } 
            else if (comboBox_标签规格.Text.Contains("13013") && comboBox_标签规格.Text.Contains("无正侧弯"))
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
                    return $"{firstField}.png";
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
                    return $"{firstField}.png";
                }
                else
                {
                    return $"{firstField}.png";
                }
            }//不是简化型号的时候
            else
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
        }

        

        //判断产品信息
        private void 判断产品信息(string aa)
        {
            string model = string.Empty;
            string powerValue = string.Empty;
            string voltageValue = string.Empty;
            string ZZ = string.Empty;
            string ledQtyValu = string.Empty;

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
            string pattern1 = @"^(\w+-\w+-\w+)";    //灯带型号
            string pattern2 = @"D(\d+)V";
            string pattern21 = @"AC(\d+)V";
            //string pattern3 = @"额定功率(\d+)W";
            string pattern3 = @"额定功率(\d+(?:\.\d+)?)W";
            string pattern4 = @"-(\d+)-";
            string pattern5 = @"(\d+)灯\/(\d+\.?\d*)cm";
            string pattern6 = @"-IP(\d{2})";

            // 使用“-”字符分割输入字符串
            string[] parts = aa.Split('-');
            灯带材质 = parts[1];

            // 使用正则表达式匹配输入字符串
            Match match1 = Regex.Match(aa, pattern1);  //灯带型号
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
                    if (标签种类_comboBox.Text == "工字标")
                    {
                        if (textLength <= 19) { output_name = "Name: " + textBox_客户资料.Text; }
                        else if (textLength > 19 && textLength <= 27) { output_name = "Name: " + Environment.NewLine + textBox_客户资料.Text; }
                        else { output_name = "Name: " + textBox_客户资料.Text.Substring(0, 19) + Environment.NewLine + textBox_客户资料.Text.Substring(19); }

                        if (comboBox_标签规格.Text.Contains("12141")) { output_灯带型号 = "Name: " + textBox_客户资料.Text; }
                    }
                    else if (标签种类_comboBox.Text == "品名标") { output_name = "Name:" + textBox_客户资料.Text; }

                    //2025.3.3
                    if (comboBox_标签规格.Text.Contains("12573") && comboBox_标签规格.Text.Contains("美国"))
                    {
                        //MessageBox.Show("只有客户名称被选择");
                        output_灯带型号 = $"ART. No.: {artNo}";
                        string 处理型号 = output_灯带型号.Replace("ART. No.:", "Model:").Trim();
                        output_灯带型号 = 处理型号;
                    }
                    else { output_灯带型号 = $"ART. No.: {artNo}"; }
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

                        if (comboBox_标签规格.Text.Contains("12141")) { output_灯带型号 = "ART. No.: " + textBox_客户资料.Text; }
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
                    else if (cpxxBox.Text.Contains("W3525")) { output_name = "Name: " + "Free Bend Wall Washer"; }
                    else if (cpxxBox.Text.Contains("A1617")) { output_name = "Name: " + "Free Bend Linear Light"; }
                    else if (cpxxBox.Text.Contains("A2012")) { output_name = "Name: " + "Free Bend Linear Light"; }
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

                            if (comboBox_标签规格.Text.Contains("12141"))
                            {
                                output_name = "Name: " + part1;
                                output_灯带型号 = "ART. No.: " + part2;
                            }
                        }
                        else if (标签种类_comboBox.Text == "品名标")
                        {
                            output_name = "Name: " + part1;
                            output_灯带型号 = "ART. No.: " + part2;
                        }

                        //2025.3.3
                        if (comboBox_标签规格.Text.Contains("12573") && comboBox_标签规格.Text.Contains("美国"))
                        {
                            //MessageBox.Show("2个都被选中");
                            string 处理型号 = output_灯带型号.Replace("ART. No.:", "Model:").Trim();
                            output_灯带型号 = 处理型号;
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
                    //else if (cpxxBox.Text.Contains("3D") && cpxxBox.Text.Contains("洗墙灯") && cpxxBox.Text.Contains("W3525")) { output_name = "Name: " + "Free Bend Wall Washer"; }
                    else if (cpxxBox.Text.Contains("W3525")) { output_name = "Name: " + "Free Bend Wall Washer"; }
                    else if (cpxxBox.Text.Contains("A1617")) { output_name = "Name: " + "Free Bend Linear Light"; }
                    else if (cpxxBox.Text.Contains("A2012")) { output_name = "Name: " + "Free Bend Linear Light"; }
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

                    //2025.3.3
                    if (comboBox_标签规格.Text.Contains("12573") && comboBox_标签规格.Text.Contains("美国"))
                    {
                        //MessageBox.Show("2个都没选中");
                        string 处理型号 = output_灯带型号.Replace("ART. No.:", "Model:").Trim();
                        output_灯带型号 = 处理型号;
                    }
                }

                if (comboBox_标签规格.Text.Contains("14098")) { output_灯带型号 = "ART. No.: " + artNo; }

                //2025.2.18更新，如果是3D系列的话，显示型号和普通的不一样，不需要例如C-SFR-这部分内容,CLEAR的要有C-SFR-
                //if (comboBox_标签规格.Text.Contains("Clear"))
                //{
                //}
                //else
                //{
                //    if (artNo.Contains("1617"))
                //    {
                //        string[] 分割结果 = artNo.Split(new string[] { "1617" }, StringSplitOptions.None);
                //        string 独立型号 = "1617" + 分割结果[1];
                //        output_灯带型号 = "ART. No.: " + 独立型号;
                //    }
                //    else if (artNo.Contains("2008"))
                //    {
                //        string[] 分割结果 = artNo.Split(new string[] { "2008" }, StringSplitOptions.None);
                //        string 独立型号 = "2008" + 分割结果[1];
                //        output_灯带型号 = "ART. No.: " + 独立型号;
                //    }
                //    else if (artNo.Contains("2012"))
                //    {
                //        string[] 分割结果 = artNo.Split(new string[] { "2012" }, StringSplitOptions.None);
                //        string 独立型号 = "2012" + 分割结果[1];
                //        output_灯带型号 = "ART. No.: " + 独立型号;
                //    }
                //}

                //3525灯带型号
                if (comboBox_标签规格.Text.Contains("Clear 3525"))
                {
                    string pattern = @"C-SFB-W(3525\w*)"; // 用括号捕获3525和后面的字母
                    Regex regex = new Regex(pattern);
                    Match match = regex.Match(cpxxBox.Text);
                    if (match.Success)
                    {
                        string modelNumber = match.Groups[1].Value; // 获取3525E
                                                                    //MessageBox.Show(modelNumber); // 用于测试
                        output_灯带型号 = "ART. No.: Panoray " + modelNumber;
                    }
                }
                else if (comboBox_标签规格.Text.Contains("中性 3525"))
                {
                    string pattern = @"C-SFB-W(3525\w*)"; // 用括号捕获3525和后面的字母
                    Regex regex = new Regex(pattern);
                    Match match = regex.Match(cpxxBox.Text);
                    if (match.Success)
                    {
                        string modelNumber = match.Groups[1].Value; // 获取3525E
                                                                    //MessageBox.Show(modelNumber); // 用于测试
                        output_灯带型号 = "ART. No.: " + modelNumber;
                    }
                }

                //A1617灯带型号
                if (comboBox_标签规格.Text.Contains("Clear A1617"))
                {
                    string pattern = @"C-SFB-A(1617\w*)"; // 用括号捕获3525和后面的字母
                    Regex regex = new Regex(pattern);
                    Match match = regex.Match(cpxxBox.Text);
                    if (match.Success)
                    {
                        string modelNumber = match.Groups[1].Value; // 获取3525E
                                                                    //MessageBox.Show(modelNumber); // 用于测试
                        output_灯带型号 = "ART. No.: Arcflex " + modelNumber;
                    }
                }
                else if (comboBox_标签规格.Text.Contains("中性 A1617"))
                {
                    string pattern = @"C-SFB-A(1617\w*)"; // 用括号捕获3525和后面的字母
                    Regex regex = new Regex(pattern);
                    Match match = regex.Match(cpxxBox.Text);
                    if (match.Success)
                    {
                        string modelNumber = match.Groups[1].Value; // 获取3525E
                                                                    //MessageBox.Show(modelNumber); // 用于测试
                        output_灯带型号 = "ART. No.: " + modelNumber;
                    }
                }

                //A2012灯带型号
                if (cpxxBox.Text.Contains("A2012"))
                {
                    string pattern = @"C-SFB-A(2012\w*)";
                    Regex regex = new Regex(pattern);
                    Match match = regex.Match(cpxxBox.Text);
                    if (match.Success)
                    {
                        string modelNumber = match.Groups[1].Value;
                        //MessageBox.Show(modelNumber); // 用于测试
                        output_灯带型号 = "ART. No.: Arcflex " + modelNumber;
                    }
                }
                else if (cpxxBox.Text.Contains("A2012"))
                {
                    string pattern = @"C-SFB-A(2012\w*)";
                    Regex regex = new Regex(pattern);
                    Match match = regex.Match(cpxxBox.Text);
                    if (match.Success)
                    {
                        string modelNumber = match.Groups[1].Value;
                        //MessageBox.Show(modelNumber); // 用于测试
                        output_灯带型号 = "ART. No.: " + modelNumber;
                    }
                }

                //12141灯带型号，需要找出灯带长度
                if (comboBox_标签规格.Text.Contains("12141"))
                {
                    // 使用正则表达式匹配数字,数字格式
                    string pattern = @"(\d+,\d+)[Mm]";
                    Regex regex = new Regex(pattern);
                    Match match = regex.Match(output_灯带型号);

                    if (match.Success)
                    {
                        string lengthStr = match.Groups[1].Value; // 获取匹配到的长度值（如：3,58319）
                        output_灯带长度 = lengthStr + "m";
                        //MessageBox.Show($"提取到的长度：{output_灯带长度}");
                    }
                    //else
                    //{
                    //    MessageBox.Show("未找到符合格式的长度值");
                    //}
                }

                //if (comboBox_标签规格.Text.Contains("12573") && comboBox_标签规格.Text.Contains("美国"))
                //{
                //    string 处理型号 = output_灯带型号.Replace("ART. No.:", "Model:") .Trim();
                //    output_灯带型号 = 处理型号;

                //}

                //13009
                if (comboBox_标签规格.Text.Contains("13009"))
                {
                    string chazhaoziliao = textBox_客户资料.Text;
                    string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\13009资料.xlsx";

                    try
                    {
                        using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                        {
                            var worksheet1 = package1.Workbook.Worksheets[0];
                            int rowCount1 = worksheet1.Dimension.Rows;
                            bool found = false;

                            // 遍历每一行，只检查A列
                            for (int row1 = 1; row1 <= rowCount1; row1++)
                            {
                                string cellValue = worksheet1.Cells[row1, 1].Text; // 只读取A列的内容

                                if (cellValue == chazhaoziliao)
                                {
                                    // 找到匹配项，获取同行F列的内容
                                    string BColumnContent = worksheet1.Cells[row1, 2].Text; // B列的内容,名称
                                    string CColumnContent = worksheet1.Cells[row1, 3].Text; // C列的内容,颜色（名称颜色中间）
                                    string DColumnContent = worksheet1.Cells[row1, 4].Text; // D列的内容,色温（CCT）
                                    string EColumnContent = worksheet1.Cells[row1, 5].Text; // E列的内容,流明(Lumen)
                                    string FColumnContent = worksheet1.Cells[row1, 6].Text; // F列的内容,功率(Wattage)
                                    string GColumnContent = worksheet1.Cells[row1, 7].Text; // G列的内容.条形码
                                    output_13009名称 = BColumnContent;
                                    output_13009颜色 = CColumnContent;
                                    output_13009色温 = "CCT:         " + DColumnContent;
                                    output_13009流明 = "Lumen:     " + EColumnContent;
                                    output_13009功率 = "Wattage:   " + FColumnContent;
                                    output_13009条形码 = GColumnContent;

                                    found = true;
                                    break; // 找到后立即退出循环
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                    }
                }

                //UL标
                if (comboBox_标签规格.Text.Contains("标签型号"))
                {
                    //string cpxx_text = cpxxBox.Text;
                    //判断产品信息(cpxxBox.Text);

                    string cz型号 = output_灯带型号.Replace("ART. No.: ", "");
                    string cz电压 = output_电压.Replace("Rated Voltage: DC ", "");
                    string cz色温 = output_色温.Replace("Color: ", "");
                    string cz色温1 = string.Empty;      //只保留数字色温
                    string cz色温2 = string.Empty;      //转换色温
                    string cz灯数 = new string(output_灯数.Replace("LED Qty.: ", "").Where(char.IsDigit).ToArray());
                    string cz功率 = string.Empty;

                    string patternX1 = @"^(\w+-\w+-\w+)";
                    Match matchX1 = Regex.Match(cpxxBox.Text, patternX1);
                    if (matchX1.Success) { cz型号 = matchX1.Groups[1].Value; }

                    cz色温1 = cz色温.Replace("K", ""); ;

                    if (cz色温.Contains("RGBW")) { cz色温2 = "RGBW"; }
                    else if (cz色温.Contains("K") || cz色温.Contains("k") && !cz色温.Contains("RGBW"))
                    {
                        if (cz色温.Contains("~"))
                        {
                            cz色温2 = cz色温.Replace("K", "");
                        }
                        else
                        {
                            // 提取数字部分
                            string 数字部分 = new string(cz色温.Where(char.IsDigit).ToArray());
                            if (int.TryParse(数字部分, out int 色温值))
                            {
                                if (色温值 >= 1100 && 色温值 < 11111111) { cz色温2 = "白光"; }
                                else if (色温值 > 11111111)
                                {
                                    // 将数字转换为字符串
                                    string 色温字符串 = 色温值.ToString();

                                    // 检查长度是否足够
                                    if (色温字符串.Length >= 8)  // 确保有足够的数字
                                    {
                                        // 在第4位后插入~
                                        string 前半部分 = 色温字符串.Substring(0, 4);  // 取前4位
                                        string 后半部分 = 色温字符串.Substring(4);     // 取剩余部分
                                        cz色温2 = $"{前半部分}~{后半部分}";  // 组合结果
                                    }
                                }
                            }
                        }
                    }
                    else if (cz色温.Contains("RGB") && !cz色温.Contains("RGBW")) { cz色温2 = "RGB"; }
                    else if (cz色温.Contains("Red")) { cz色温2 = "红"; }
                    else if (cz色温.Contains("Blue")) { cz色温2 = "蓝"; }
                    else if (cz色温.Contains("Green")) { cz色温2 = "绿"; }
                    else if (cz色温.Contains("Orange")) { cz色温2 = "橙"; }
                    else if (cz色温.Contains("Yellow")) { cz色温2 = "黄"; }
                    else if (cz色温.Contains("Amber")) { cz色温2 = "琥珀"; }

                    //UL标12291
                    if (comboBox_标签规格.Text.Contains("12291"))
                    {
                        string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12291 UL资料.xlsx";
                        try
                        {
                            using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                            {
                                var worksheet1 = package1.Workbook.Worksheets[0];
                                int rowCount1 = worksheet1.Dimension.Rows;
                                bool found = false;

                                // 遍历每一行，只检查B列
                                for (int row1 = 2; row1 <= rowCount1; row1++)
                                {
                                    string cellValue = worksheet1.Cells[row1, 2].Text; // 只读取B列的内容
                                    if (cpxxBox.Text.Contains("【DMX】"))
                                    {
                                        if (cellValue.Contains(cz型号) && cellValue.Contains(cz电压) && cellValue.Contains(cz色温2) && cellValue.Contains(cz灯数) && cellValue.Contains("DMX"))
                                        {
                                            // 找到匹配项，获取同行的内容
                                            string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                            string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                            output_name1 = "Name: " + DColumnContent;
                                            output_灯带型号1 = "ART. No.: " + CColumnContent;

                                            found = true;
                                            break; // 找到后立即退出循环
                                        }
                                    }
                                    else
                                    {
                                        if (cellValue.Contains(cz型号) && cellValue.Contains(cz电压) && cellValue.Contains(cz色温2) && cellValue.Contains(cz灯数))
                                        {
                                            // 找到匹配项，获取同行的内容
                                            string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                            string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                            output_name1 = "Name: " + DColumnContent;
                                            output_灯带型号1 = "ART. No.: " + CColumnContent;

                                            found = true;
                                            break; // 找到后立即退出循环
                                        }
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                        }
                    }
                    //UL标17021，只需要核对灯带型号
                    else if (comboBox_标签规格.Text.Contains("17021"))
                    {
                        string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\17021 UL资料.xlsx";
                        try
                        {
                            using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                            {
                                var worksheet1 = package1.Workbook.Worksheets[0];
                                int rowCount1 = worksheet1.Dimension.Rows;
                                bool found = false;

                                // 遍历每一行，只检查B列
                                for (int row1 = 2; row1 <= rowCount1; row1++)
                                {
                                    string cellValue = worksheet1.Cells[row1, 2].Text; // 只读取B列的内容

                                    if (cellValue.Contains(cz型号))
                                    {
                                        // 找到匹配项，获取同行的内容
                                        string CColumnContent = worksheet1.Cells[row1, 3].Text;  //客户型号
                                        string DColumnContent = worksheet1.Cells[row1, 4].Text; //客户名称
                                        output_name1 = "Name: " + DColumnContent;
                                        output_灯带型号1 = "ART. No.: " + CColumnContent;

                                        found = true;
                                        break; // 找到后立即退出循环
                                    }
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                        }
                    }
                }

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

                if (comboBox_标签规格.Text.Contains("12098"))
                {
                    output_电压 = $"{voltageValue}V DC ";
                }
                else if (comboBox_标签规格.Text.Contains("13009"))
                {
                    output_电压 = $"Voltage:     {voltageValue}VDC";
                }
                else
                {
                    // 如果不包含“高压”，则设置 output_电压 为 DC
                    output_电压 = $"Rated Voltage: DC {voltageValue}V";
                }

                // 如果不包含“高压”，则设置 output_电压 为 DC
                //output_电压 = $"Rated Voltage: DC {voltageValue}V";
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
            string pattern13 = @"恒流(\d+(?:\.\d+)?)W";
            Match match13 = Regex.Match(aa, pattern13);
            if (match3.Success)
            {
                // 从匹配结果中提取功率值
                powerValue = match3.Groups[1].Value; // 第一个捕获组匹配的内容

                // 检查是否包含"12275"规格
                if (comboBox_标签规格.Text.Contains("12275"))
                {
                    // 将W/m转换为W/FT
                    double powerNum = double.Parse(powerValue);
                    double convertedPower = Math.Round(powerNum / 3.28, 2);
                    output_功率 = $"Rated Power: {convertedPower}W/FT";
                }
                else if (comboBox_标签规格.Text.Contains("12090"))
                {
                    output_功率 = $"Vermogen: {powerValue}W/m";
                }
                else
                {
                    // 保持原有格式(W/m和LEDs/m)
                    output_功率 = $"Rated Power: {powerValue}W/m";
                }
            }
            else if (match13.Success)
            {
                // 从匹配结果中提取功率值
                powerValue = match13.Groups[1].Value; // 第一个捕获组匹配的内容

                output_功率 = $"Rated Power: {powerValue}W/m";
            }
            else
            {
                MessageBox.Show("未找到功率匹配项。", "错误");
            }

            //总功率
            if (comboBox_标签规格.Text.Contains("12141"))
            {
                // 处理功率：去除"Rated Power: "和"W/m"，只保留数字
                string powerStr = output_功率.Replace("Rated Power: ", "")
                                           .Replace("W/m", "")
                                           .Trim();

                // 处理长度：去除"m"单位，将逗号替换为小数点
                string lengthStr = output_灯带长度.Replace("m", "")
                                                 .Replace(",", ".")
                                                 .Trim();

                // 转换为double进行计算
                if (double.TryParse(powerStr, out double power) &&
                    double.TryParse(lengthStr, out double length))
                {
                    // 计算总功率
                    double totalPower = power * length;

                    // 保留2位小数
                    output_总功率 = $"Total Power:{totalPower:F2}W";
                    //MessageBox.Show(output_总功率);
                }
            }

            //光源型号
            if (comboBox_标签规格.Text.Contains("12141"))
            {
                // 去除括号的处理
                string ledType = parts[6];
                if (ledType.Contains("(") && ledType.Contains(")"))
                {
                    ledType = ledType.Replace("(", "").Replace(")", "");
                }

                //MessageBox.Show(parts[6]);
                if (cpxxBox.Text.Contains("Pro"))
                {
                    output_光源型号 = "LED: Nichia " + ledType;
                }
                else
                {
                    if (ledType == "50B" || ledType == "50D") { ledType = "5050"; }
                    output_光源型号 = "LED: SMD " + ledType;
                }
                //MessageBox.Show(output_光源型号);
            }

            // 灯数
            if (match4.Success)
            {
                // 从匹配结果中提取数字
                ledQtyValu = match4.Groups[1].Value; // 第一个捕获组匹配的内容

                // 构造输出字符串
                // 检查是否包含"12275"规格
                if (comboBox_标签规格.Text.Contains("12275"))
                {
                    // 将字符串转换为double后进行计算，并取整
                    if (double.TryParse(ledQtyValu, out double ledNum))
                    {
                        double convertedLed = Math.Round(ledNum / 3.28, 0); // 除以3并取整
                        output_灯数 = $"LED Qty.: {convertedLed}LEDs/FT";
                    }
                    else
                    {
                        MessageBox.Show("灯数值转换失败。", "错误");
                    }
                }
                else if (comboBox_标签规格.Text.Contains("12141")) { output_灯数 = $"Qty.: {ledQtyValu}LEDs/m"; }
                else if (comboBox_标签规格.Text.Contains("16066"))
                {
                    // 使用正则表达式匹配"约XX灯/米"的格式
                    string pattern = @"约(\d+)灯/米";
                    Regex regex = new Regex(pattern);
                    Match ledMatch = regex.Match(cpxxBox.Text);

                    if (ledMatch.Success)
                    {
                        // 提取数字部分
                        ledQtyValu = ledMatch.Groups[1].Value;

                        output_灯数 = $"LED Qty.: {ledQtyValu}LEDs/m";
                    }
                    else
                    {
                        output_灯数 = $"LED Qty.: {ledQtyValu}LEDs/m";
                    }
                }
                else
                {
                    // 保持原有格式(LEDs/m)
                    output_灯数 = $"LED Qty.: {ledQtyValu}LEDs/m";
                }
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

                // 检查是否包含"12275"规格
                if (comboBox_标签规格.Text.Contains("12275"))
                {
                    // 将厘米转换为英寸
                    if (double.TryParse(length, out double lengthNum))
                    {
                        double convertedLength = Math.Round(lengthNum / 2.54, 2); // 除以2.54并保留2位小数

                        if (标签种类_comboBox.Text == "品名标")
                        {
                            output_剪切单元 = $"Min. Cutting Length: {ledQuantity}LEDs({convertedLength}IN)";
                        }
                        else
                        {
                            output_剪切单元 = $"Min. Cutting Length: {ledQuantity}LEDs\n({convertedLength}IN)";
                        }
                    }
                    else
                    {
                        MessageBox.Show("剪切长度转换失败。", "错误");
                    }
                }
                else if (comboBox_标签规格.Text.Contains("12090"))
                {
                    if (标签种类_comboBox.Text == "品名标")
                    {
                        output_剪切单元 = $"In te korten per: {ledQuantity}LEDs({length}cm)";
                    }
                    else
                    {
                        output_剪切单元 = $"In te korten per: {ledQuantity}LEDs\n({length}cm)";
                    }
                }
                else if (comboBox_标签规格.Text.Contains("12141")) { output_剪切单元 = $"Min. Cutting Length: {ledQuantity}LEDs({length}cm)"; }
                else
                {
                    // 保持原有格式(cm)
                    if (标签种类_comboBox.Text == "品名标")
                    {
                        output_剪切单元 = $"Min. Cutting Length: {ledQuantity}LEDs({length}cm)";
                    }
                    else
                    {
                        output_剪切单元 = $"Min. Cutting Length: {ledQuantity}LEDs\n({length}cm)";
                    }
                }
            }
            else
            {
                // 如果没有找到匹配项，则输出错误信息
                MessageBox.Show("未找到剪切单元信息匹配项", "错误");
            }

            //色温
            if (parts.Length >= 6)
            {
                string numericValue;
                // 第五个"-"和第六个"-"之间的内容是parts[5]，因为数组索引是从0开始的
                string contentBetweenFifthAndSixth = parts[5];

                if (parts.Length > 6 && contentBetweenFifthAndSixth.StartsWith("("))
                {
                    // 如果以"("开头，则合并parts[5]和parts[6]
                    contentBetweenFifthAndSixth = parts[5] + "-" + parts[6];
                    //MessageBox.Show("合并后的内容: " + contentBetweenFifthAndSixth);

                    // 使用正则表达式提取括号内的内容
                    Regex regex = new Regex(@"\(([^-]+)-([^)]+)\)");
                    Match match = regex.Match(contentBetweenFifthAndSixth);

                    if (match.Success && match.Groups.Count >= 3)
                    {
                        string firstPart = match.Groups[1].Value; // 例如 W3500
                        string secondPart = match.Groups[2].Value; // 例如 R

                        // 单独显示提取的内容
                        //MessageBox.Show("第一部分: " + firstPart);
                        //MessageBox.Show("第二部分: " + secondPart);
                        numericValue = Regex.Replace(firstPart, @"[^0-9]", string.Empty);
                        string 第一部分 = $"{numericValue}K";

                        string 第二部分 = "";
                        if (secondPart == "R") { 第二部分 = $"Red"; }
                        else if (secondPart == "B") { 第二部分 = $"Blue"; }
                        else if (secondPart == "G") { 第二部分 = $"Green"; }
                        else if (secondPart == "O") { 第二部分 = $"Orange"; }
                        else if (secondPart == "Y") { 第二部分 = $"Yellow"; }

                        output_色温 = "Color: " + 第一部分 + "+" + 第二部分;
                    }
                    else
                    {
                        MessageBox.Show("无法从内容中提取所需部分");
                    }
                }
                else
                {
                    //正常情况
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
                            else if (contentBetweenFifthAndSixth == "Y3955C") { output_色温 = $"Color: Yellow (Full color jacket) (Y3955C)"; }
                            else if (cpxxBox.Text.Contains("黑色全彩"))
                            {
                                //output_色温 = $"Color: {Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty)}K(Full Black jacket)";

                                if (contentBetweenFifthAndSixth.Contains("RGB"))
                                {
                                    output_色温 = $"Color: " + contentBetweenFifthAndSixth + "(Full Black jacket)";
                                    //MessageBox.Show(output_色温);
                                }
                                else
                                {
                                    //MessageBox.Show("123");
                                    output_色温 = $"Color: {Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty)}K(Full Black jacket)";
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
                        else if (cpxxBox.Text.Contains("黑色遮光+雾状发光"))
                        {
                            //MessageBox.Show(contentBetweenFifthAndSixth);
                            if (contentBetweenFifthAndSixth.Contains("RGB"))
                            {
                                output_色温 = $"Color: " + contentBetweenFifthAndSixth + "(Black jacket)";
                                //MessageBox.Show(output_色温);
                            }
                            else if (contentBetweenFifthAndSixth == "R") { output_色温 = $"Color: Red(Black jacket)"; }
                            else if (contentBetweenFifthAndSixth == "B") { output_色温 = $"Color: Blue(Black jacket)"; }
                            else if (contentBetweenFifthAndSixth == "G") { output_色温 = $"Color: Green(Black jacket)"; }
                            else if (contentBetweenFifthAndSixth == "O") { output_色温 = $"Color: Orange(Black jacket)"; }
                            else if (contentBetweenFifthAndSixth == "Y") { output_色温 = $"Color: Yellow(Black jacket)"; }
                            else
                            {
                                //MessageBox.Show("123");
                                output_色温 = $"Color: {Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty)}K(Black jacket)";
                            }
                        }
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

                            if (comboBox_标签规格.Text.Contains("12090"))
                            {
                                // 检查内容是否是RGBW
                                if (contentBetweenFifthAndSixth.Contains("RGBW"))
                                {
                                    output_色温 = $"Kleur: {contentBetweenFifthAndSixth}";
                                }
                                // 检查内容是否为纯字母
                                else if (Regex.IsMatch(contentBetweenFifthAndSixth, @"^[a-zA-Z]*$"))
                                {
                                    output_色温 = $"Kleur: {contentBetweenFifthAndSixth}";
                                }
                                else
                                {
                                    // 如果包含数字，则提取数字部分
                                    numericValue = Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", string.Empty);
                                    output_色温 = $"Kleurtemperatuur: {numericValue}°K";
                                }
                            }
                        }
                    }
                    else
                    { MessageBox.Show("没有找到灯带系列。", "错误"); }
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
            //高压非短剪的时候
            else if (comboBox_标签规格.Text.Contains("高压") && comboBox_标签规格.Text.Contains("非短剪"))
            {
                if (cpxxBox.Text.Contains("C-FR-F10")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "X" + "\n" + "Rating: 220V-240V~, 50/60Hz," + "\n" + " 5W/m, Max 100m" + "\n" + output_长度 + "\n" + output_色温 + "\n" + " " + "\n" + output_尾巴; }
                else if (cpxxBox.Text.Contains("C-FR-F11")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "X" + "\n" + "Rating: 220V-240V~, 50/60Hz," + "\n" + " 5W/m, Max 100m" + "\n" + output_长度 + "\n" + output_色温 + "\n" + " " + "\n" + output_尾巴; }
                else if (cpxxBox.Text.Contains("C-FR-F15")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "X" + "\n" + "Rating: 220V-240V~, 50/60Hz," + "\n" + " 12W/m, Max 80m" + "\n" + output_长度 + "\n" + output_色温 + "\n" + " " + "\n" + output_尾巴; }
                else if (cpxxBox.Text.Contains("C-FR-F21")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "X" + "\n" + "Rating: 220V-240V~, 50/60Hz," + "\n" + " 12W/m, Max 80m" + "\n" + output_长度 + "\n" + output_色温 + "\n" + " " + "\n" + output_尾巴; }
                else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "X" + "\n" + "Rating: 220V-240V~, 50/60Hz," + "\n" + " 5W/m, Max 100m" + "\n" + output_长度 + "\n" + output_色温 + "\n" + " " + "\n" + output_尾巴; }
            }
            //高压短剪的时候
            else if (comboBox_标签规格.Text.Contains("高压") && comboBox_标签规格.Text.Contains("短剪") && !comboBox_标签规格.Text.Contains("非短剪"))
            {
                if (cpxxBox.Text.Contains("-可延长")) { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "-" + "AC" + voltageValue + "V" + "-" + ledQtyValu + "-" + ZZ + "-Plug" + "\n" + "Rating: 220V-240V~, 50/60Hz," + "\n" + "10W/m,0.042A/m, Max. 3.44A," + "\n" + "Max. 80m" + "\n" + output_长度 + "\n" + output_色温 + "\n" + " " + "\n" + output_尾巴; }
                else { name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "-" + "AC" + voltageValue + "V" + "-" + ledQtyValu + "-" + ZZ + "\n" + "Rating: 220V-240V~, 50/60Hz," + "\n" + "10W/m,0.042A/m, Max. 3.44A," + "\n" + "Max. 80m" + "\n" + output_长度 + "\n" + output_色温 + "\n" + " " + "\n" + output_尾巴; }
            }

            //12098的时候
            else if (comboBox_标签规格.Text.Contains("12098"))
            {
                string inputText = textBox_客户资料.Text;
                string convertedName = "";

                // 1. 处理A系列
                if (灯带系列 == "A")
                {
                    // 去掉RGB-前缀
                    string baseText = inputText.StartsWith("RGB-") ? inputText.Substring(4) : inputText;

                    // 处理SNX开头的情况
                    if (baseText.StartsWith("SNX-"))
                    {
                        var nameParts = baseText.Split('-');
                        convertedName = "SUPER-NEON-";

                        // 处理X-FLAT或X-DOME
                        if (nameParts.Length >= 2)
                        {
                            switch (nameParts[1])
                            {
                                case "F":
                                    convertedName += "X-FLAT";
                                    break;

                                case "D":
                                    convertedName += "X-DOME";
                                    break;

                                default:
                                    convertedName += "X-" + nameParts[1];
                                    break;
                            }

                            // 添加中间部分，处理倒数第二部分的颜色缩写
                            for (int i = 2; i < nameParts.Length; i++)
                            {
                                if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                                {
                                    string color = ConvertColorAbbreviation(nameParts[i]);
                                    convertedName += "-" + color;
                                }
                                else
                                {
                                    convertedName += "-" + nameParts[i];
                                }
                            }
                        }

                        // 最后加上RGB-前缀
                        convertedName = "RGB-" + convertedName;
                    }
                    // 处理SNE开头的情况
                    else if (baseText.StartsWith("SNE-"))
                    {
                        var nameParts = baseText.Split('-');
                        convertedName = "SUPER-NEON-EDGE";

                        // 从第二个部分开始添加，处理倒数第二部分的颜色缩写
                        for (int i = 1; i < nameParts.Length; i++)
                        {
                            if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                            {
                                string color = ConvertColorAbbreviation(nameParts[i]);
                                convertedName += "-" + color;
                            }
                            else
                            {
                                convertedName += "-" + nameParts[i];
                            }
                        }

                        // 最后加上RGB-前缀
                        convertedName = "RGB-" + convertedName;
                    }
                    else
                    {
                        convertedName = inputText; // 如果既不是SNX也不是SNE开头，保持原样
                    }
                }
                // 2. 处理SNX开头的情况
                else if (inputText.StartsWith("SNX-"))
                {
                    var nameParts = inputText.Split('-');
                    convertedName = "SUPER-NEON-";

                    // 处理X-FLAT或X-DOME
                    if (nameParts.Length >= 2)
                    {
                        switch (nameParts[1])
                        {
                            case "F":
                                convertedName += "X-FLAT";
                                break;

                            case "D":
                                convertedName += "X-DOME";
                                break;

                            default:
                                convertedName += nameParts[1];
                                break;
                        }

                        // 添加剩余部分，处理颜色缩写
                        for (int i = 2; i < nameParts.Length; i++)
                        {
                            if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                            {
                                string color = ConvertColorAbbreviation(nameParts[i]);
                                convertedName += "-" + color;
                            }
                            else
                            {
                                convertedName += "-" + nameParts[i];
                            }
                        }
                    }
                }
                // 3. 处理SNE开头的情况
                else if (inputText.StartsWith("SNE-"))
                {
                    var nameParts = inputText.Split('-');
                    convertedName = "SUPER-NEON-EDGE";

                    // 从第二个部分开始添加，处理颜色缩写
                    for (int i = 1; i < nameParts.Length; i++)
                    {
                        if (i == nameParts.Length - 2) // 倒数第二个部分，处理颜色缩写
                        {
                            string color = ConvertColorAbbreviation(nameParts[i]);
                            convertedName += "-" + color;
                        }
                        else
                        {
                            convertedName += "-" + nameParts[i];
                        }
                    }
                }
                else
                {
                    convertedName = inputText; // 默认情况
                }

                output_name = convertedName;

                // 显示转换结果
                //MessageBox.Show($"转换后的名称: {output_name}", "名称转换结果");

                output_灯带型号 = "Short SKU:" + textBox_客户资料.Text;
                name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + "BATCH " + textBox_标识码01.Text;
            }
            else if (comboBox_标签规格.Text.Contains("16008"))
            {
                output_灯带型号 = output_灯带型号.Replace("ART. No.: ", "");
                name_CPXXBox.Text = "www.SGiLighting.com" + "\n" + "LED NEON FLEX LIGHT" + "\n" + output_灯带型号 + "\n" + "Length: " + textBox_剪切长度.Text + "\n" + "Lot #:  " + textBox_标识码01.Text;
            }
            else if (comboBox_标签规格.Text.Contains("12251") && 标签种类_comboBox.Text.Contains("工字标"))
            {
                string chazhaoziliao = textBox_客户资料.Text;
                string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12251资料.xlsx";

                try
                {
                    using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                    {
                        var worksheet1 = package1.Workbook.Worksheets[0];
                        int rowCount1 = worksheet1.Dimension.Rows;
                        string fColumnContent = "";
                        bool found = false;

                        // 遍历每一行
                        for (int row1 = 1; row1 <= rowCount1; row1++)
                        {
                            // 检查A到D列
                            for (int col = 1; col <= 4; col++) // 1=A, 2=B, 3=C, 4=D
                            {
                                string cellValue = worksheet1.Cells[row1, col].Text;

                                if (cellValue == chazhaoziliao)
                                {
                                    // 找到匹配项，获取同行F列的内容
                                    fColumnContent = worksheet1.Cells[row1, 6].Text; // F列的内容
                                    name_CPXXBox.Text = fColumnContent;
                                    found = true;

                                    // 可以添加一个消息框显示在哪里找到的（如果需要）
                                    //MessageBox.Show($"在第{row1}行，第{(char)(col + 64)}列找到匹配项", "查找结果");

                                    break;
                                }
                            }

                            if (found) break; // 如果找到了就退出外层循环
                        }

                        if (string.IsNullOrEmpty(fColumnContent))
                        {
                            MessageBox.Show($"在A、B、C、D列中均未找到匹配内容: {chazhaoziliao}", "查找结果");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                }
            }
            else if (comboBox_标签规格.Text.Contains("12251") && 标签种类_comboBox.Text.Contains("品名标"))
            {
                string chazhaoziliao = textBox_客户资料.Text;
                string excelPath1 = @"\\192.168.1.33\Annmy\订单标签自动生成软件\sucai\12251资料.xlsx";

                try
                {
                    using (var package1 = new ExcelPackage(new FileInfo(excelPath1)))
                    {
                        var worksheet1 = package1.Workbook.Worksheets[0];
                        int rowCount1 = worksheet1.Dimension.Rows;
                        string fColumnContent = "";
                        bool found = false;

                        // 遍历每一行
                        for (int row1 = 1; row1 <= rowCount1; row1++)
                        {
                            // 检查A到D列
                            for (int col = 1; col <= 4; col++) // 1=A, 2=B, 3=C, 4=D
                            {
                                string cellValue = worksheet1.Cells[row1, col].Text;

                                if (cellValue == chazhaoziliao)
                                {
                                    // 找到匹配项，获取同行E列的内容
                                    fColumnContent = worksheet1.Cells[row1, 5].Text; // E列的内容
                                    name_CPXXBox.Text = fColumnContent;
                                    found = true;

                                    break;
                                }
                            }

                            if (found) break; // 如果找到了就退出外层循环
                        }

                        if (string.IsNullOrEmpty(fColumnContent))
                        {
                            MessageBox.Show($"在A、B、C、D列中均未找到匹配内容: {chazhaoziliao}", "查找结果");
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"发生错误：\n{ex.Message}", "错误");
                }
            }
            else if (comboBox_标签规格.Text.Contains("12090"))
            {
                name_CPXXBox.Text = output_功率 + "\n" + output_色温 + "\n" + output_剪切单元 + "\n" + "Rollengte:" + textBox_剪切长度.Text;
            }
            else if (comboBox_标签规格.Text.Contains("12141"))
            {
                output_长度 = "Length of Light:" + output_灯带长度;
                output_线材长度 = "Length of Cable: " + textBox_线长.Text;
                name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_总功率 + "\n" + output_光源型号 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_线材长度;
            }
            else if (comboBox_标签规格.Text.Contains("12120"))
            {
                name_CPXXBox.Text = "AMBIANCE LUMIERE" + "\n" + output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + "Caution: Do not overload." + "\n" + output_尾巴;
            }
            else
            {
                name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴;
            }

            //MessageBox.Show(output_name +"\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_尾巴, "提取结果");

            //是否要透镜角度
            bool 是否要透镜角度 = comboBox_标签规格.Text.Contains("3525") ||
                          (comboBox_标签规格.Text.Contains("17034") && cpxxBox.Text.Contains("W3525")) ||
                          (comboBox_标签规格.Text.Contains("12058") && cpxxBox.Text.Contains("W3525")) || cpxxBox.Text.Contains("W3525");

            if (是否要透镜角度)
            {
                try
                {
                    // 先尝试匹配 XX*XX 格式
                    string anglePattern1 = @"透镜角度(\d+)\*(\d+)";
                    // 再尝试匹配单一数字格式
                    string anglePattern2 = @"透镜角度(\d+)(?!\*)"; // (?!\*) 确保数字后面不是*号

                    Regex angleRegex1 = new Regex(anglePattern1);
                    Regex angleRegex2 = new Regex(anglePattern2);

                    Match angleMatch1 = angleRegex1.Match(cpxxBox.Text);
                    Match angleMatch2 = angleRegex2.Match(cpxxBox.Text);

                    if (angleMatch1.Success)
                    {
                        // 处理 XX*XX 格式
                        string angle1 = angleMatch1.Groups[1].Value;
                        string angle2 = angleMatch1.Groups[2].Value;
                        output_透镜角度 = $"Beam Angle: {angle1}°×{angle2}°";
                    }
                    else if (angleMatch2.Success)
                    {
                        // 处理单一角度格式
                        string angle = angleMatch2.Groups[1].Value;
                        output_透镜角度 = $"Beam Angle: {angle}°";
                    }
                    else
                    {
                        // 如果都没匹配到，可以设置为空或默认值
                        output_透镜角度 = "";
                    }

                    name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 + "\n" + output_透镜角度 + "\n" + output_尾巴;
                }
                catch (Exception ex)
                {
                    MessageBox.Show("解析透镜角度时出错：" + ex.Message);
                }
            }
        }

        private string ConvertColorAbbreviation(string colorCode)
        {
            switch (colorCode.ToUpper())
            {
                case "B":
                    return "Blue";

                case "R":
                    return "Red";

                case "G":
                    return "Green";

                case "W":
                    return "White";

                case "Y":
                    return "Yellow";

                case "P":
                    return "Purple";

                case "A":
                    return "Amber";

                default:
                    return colorCode;
            }
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
            try
            {
                // 获取应用程序目录
                string appPath = Application.StartupPath;
                string codePath = Path.Combine(appPath, "CODE", "run.py");  // 注意这里改成 .py

                // 创建进程启动信息
                ProcessStartInfo startInfo = new ProcessStartInfo();
                startInfo.FileName = "python.exe";
                startInfo.Arguments = $"\"{codePath}\"";
                startInfo.UseShellExecute = false;
                startInfo.RedirectStandardOutput = true;
                startInfo.RedirectStandardError = true;
                startInfo.CreateNoWindow = true;

                // 启动进程
                using (Process process = Process.Start(startInfo))
                {
                    // 读取输出
                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    // 等待进程完成
                    process.WaitForExit();

                    // 显示输出结果
                    if (!string.IsNullOrEmpty(output))
                    {
                        MessageBox.Show(output, "Python输出");
                    }

                    // 如果有错误则显示
                    if (!string.IsNullOrEmpty(error))
                    {
                        MessageBox.Show($"Python执行错误：{error}", "错误");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"执行出错：{ex.Message}", "错误");
            }
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

        //标签规格选择内容后的事项
        private void comboBox_标签规格_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox_标签规格.Text.Contains("14098"))
            {
                checkBox_客户型号.Checked = true;
            }
            

            if (comboBox_标签规格.Text.Contains("18422")) { checkBox_结尾.Checked = false; }
            else if (comboBox_标签规格.Text.Contains("17100")){ checkBox_结尾.Checked = false; }
            else { checkBox_结尾.Checked = true; }

            //不同型号提示框操作提示！！————————————————————————————————————————————————————————————————

            //12141工字标操作提示
            if (标签种类_comboBox.Text.Contains("工字标") && comboBox_标签规格.Text.Contains("12141"))
            {
                提示框.AppendText("12141:客户名称和客户型号需要手动打钩" + Environment.NewLine);
                提示框.AppendText("12141:并号使用PO号位置,附件使用PO号的J列" + Environment.NewLine);
                提示框.AppendText("12141:线长使用线长位置,附件使用L列" + Environment.NewLine);
            }

            if (comboBox_标签规格.Text.Contains("13009"))
            {
                提示框.AppendText("13009:只需要填入客户型号，无需打钩（无需客户名称）" + Environment.NewLine);
                提示框.AppendText("13009:如果电压没有变动的话，可以使用默认的规格型号，无需复制" + Environment.NewLine);
            }
            if (comboBox_标签规格.Text.Contains("13453"))
            {
                提示框.AppendText("13453:加载PDF所在文件夹" + Environment.NewLine);
                提示框.AppendText("13453:附件打印数量需要导入数据库控制" + Environment.NewLine);
            }

            if (标签种类_comboBox.Text.Contains("唛头"))
            {
                // 清空comboBox_唛头规格的现有项
                comboBox_唛头规格.Items.Clear();

                // 获取选定的标签规格
                string selectedSpec = comboBox_标签规格.SelectedItem?.ToString();

                // 显示选择的标签规格（用于调试）
                //MessageBox.Show($"选择的标签规格: {selectedSpec ?? "未选择"}", "选择确认", MessageBoxButtons.OK, MessageBoxIcon.Information);

                if (!string.IsNullOrEmpty(selectedSpec))
                {
                    try
                    {
                        // 构建基础路径
                        string basePath = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\唛头\";

                        // 确定型号类型路径
                        string modelTypePath = checkBox_常规型号.Checked ? "常规型号" :
                                              checkBox_客制型号.Checked ? "客制型号" :
                                              checkBox_简化型号.Checked ? "简化型号" : "";

                        // 构建完整的目录路径
                        string btwPath = Path.Combine(basePath, modelTypePath, selectedSpec);

                        // 显示构建的路径（用于调试）
                        //MessageBox.Show($"构建的路径: {btwPath}", "路径信息", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        // 检查目录是否存在
                        if (Directory.Exists(btwPath))
                        {
                            // 获取目录中的所有文件
                            string[] files = Directory.GetFiles(btwPath, "*.btw");

                            // 将文件名（不带扩展名）添加到comboBox_唛头规格
                            foreach (string file in files)
                            {
                                string fileName = Path.GetFileNameWithoutExtension(file);
                                comboBox_唛头规格.Items.Add(fileName);
                            }

                            // 如果有项目，选择第一项
                            if (comboBox_唛头规格.Items.Count > 0)
                            {
                                comboBox_唛头规格.SelectedIndex = 0;
                            }
                        }
                        else
                        {
                            MessageBox.Show($"目录不存在: {btwPath}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"读取文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void button_加载PDF_Click(object sender, EventArgs e)
        {
            if (comboBox_标签规格.Text.Contains("13453"))
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                FolderBrowserDialog folderDialog = new FolderBrowserDialog();
                folderDialog.Description = "请选择文件夹";
                // 可以设置初始目录
                folderDialog.SelectedPath = Application.StartupPath;

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    textBox_pdf.Text = folderDialog.SelectedPath;
                    // 确保拆分PDF的目录存在
                    string outputDirectory = Path.Combine(Application.StartupPath + @"\PDF拆分");

                    if (!Directory.Exists(outputDirectory))
                    {
                        Directory.CreateDirectory(outputDirectory);
                    }

                    // 如果目录存在，先删除目录里的所有文件
                    if (Directory.Exists(outputDirectory))
                    {
                        // 获取目录下的所有文件
                        string[] files = Directory.GetFiles(outputDirectory);
                        foreach (string file in files)
                        {
                            // 检查文件属性，如果文件是只读的，先移除只读属性
                            if ((File.GetAttributes(file) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                            {
                                File.SetAttributes(file, File.GetAttributes(file) & ~FileAttributes.ReadOnly);
                            }
                            // 删除文件
                            File.Delete(file);
                        }
                        // 清空目录后，可以重新创建目录，如果需要的话
                        // Directory.Delete(outputDirectory, false); // 如果需要删除目录本身，取消注释这行代码
                        // Directory.CreateDirectory(outputDirectory); // 如果删除目录后需要重新创建，取消注释这行代码
                    }

                    // 读取源文件夹中的所有PDF文件
                    string sourcePath = textBox_pdf.Text;
                    var pdfFiles = Directory.GetFiles(sourcePath, "*.pdf")
                                           .OrderBy(f => File.GetCreationTime(f))
                                           .ToList();

                    // 依次处理每个PDF文件
                    for (int i = 0; i < pdfFiles.Count; i++)
                    {
                        // 构造新的文件名
                        string newFileName = $"page_{i + 1}.pdf";
                        string destinationPath = Path.Combine(outputDirectory, newFileName);

                        // 如果目标文件已存在，先删除
                        if (File.Exists(destinationPath))
                        {
                            File.Delete(destinationPath);
                        }

                        // 复制文件到新位置
                        File.Copy(pdfFiles[i], destinationPath);
                    }

                    MessageBox.Show("所有PDF文件处理完成！", "成功");
                }
            }
            else
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Multiselect = false;
                dialog.Title = "请选择数据库文件";
                dialog.Filter = "pdf文件(*.pdf)|*.pdf|All files (*.*)|*.*";
                dialog.InitialDirectory = Application.StartupPath;

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    textBox_pdf.Text = dialog.FileName;

                    // 确保拆分PDF的目录存在
                    string outputDirectory = Path.Combine(Application.StartupPath + @"\PDF拆分");

                    if (!Directory.Exists(outputDirectory))
                    {
                        Directory.CreateDirectory(outputDirectory);
                    }

                    // 如果目录存在，先删除目录里的所有文件
                    if (Directory.Exists(outputDirectory))
                    {
                        // 获取目录下的所有文件
                        string[] files = Directory.GetFiles(outputDirectory);
                        foreach (string file in files)
                        {
                            // 检查文件属性，如果文件是只读的，先移除只读属性
                            if ((File.GetAttributes(file) & FileAttributes.ReadOnly) == FileAttributes.ReadOnly)
                            {
                                File.SetAttributes(file, File.GetAttributes(file) & ~FileAttributes.ReadOnly);
                            }
                            // 删除文件
                            File.Delete(file);
                        }
                        // 清空目录后，可以重新创建目录，如果需要的话
                        // Directory.Delete(outputDirectory, false); // 如果需要删除目录本身，取消注释这行代码
                        // Directory.CreateDirectory(outputDirectory); // 如果删除目录后需要重新创建，取消注释这行代码
                    }

                    // 读取PDF文件
                    string pdfPath = textBox_pdf.Text;
                    using (PdfReader reader = new PdfReader(pdfPath))
                    {
                        for (int i = 1; i <= reader.NumberOfPages; i++)
                        {
                            // 创建新的PDF文档
                            string outputPdfPath = Path.Combine(outputDirectory, $"page_{i}.pdf");
                            using (Document document = new Document())
                            {
                                using (PdfCopy copy = new PdfCopy(document, new FileStream(outputPdfPath, FileMode.Create)))
                                {
                                    document.Open();
                                    // 添加单页
                                    copy.AddPage(copy.GetImportedPage(reader, i));
                                }
                            }
                            // 显示进度信息
                            //MessageBox.Show($"页面 {i} 已保存为 {outputPdfPath}");
                        }
                    }
                }
            }
        }

        private void button_打印PDF_Click(object sender, EventArgs e)
        {
            if (comboBox_标签规格.Text.Contains("13453"))
            {
                using (Engine btEngine = new Engine(true))
                {
                    // 获取运行目录下的PDF拆分文件夹路径
                    string pdfSplitDirectory = Path.Combine(Application.StartupPath + @"\PDF拆分");
                    //MessageBox.Show(pdfSplitDirectory);

                    // 获取PDF拆分目录下的所有PDF文件
                    var pdfFiles = Directory.GetFiles(pdfSplitDirectory, "*.pdf")
                      .Select(Path.GetFileName) // 直接提取文件名
                      .OrderBy(fileName => int.Parse(fileName.Substring(5, fileName.Length - 9))) // 提取数字并排序
                      .ToList();

                    // 打开标签格式文件
                    if (标签种类_comboBox.Text.Contains("工字标"))
                    {
                        LabelFormatDocument labelFormat = btEngine.Documents.Open("\\\\192.168.1.33\\Annmy\\订单标签自动生成软件\\moban\\工字标\\客制型号\\(13453) ORLIGHT\\1.btw");
                        //LabelFormatDocument labelFormat = btEngine.Documents.Open("\\\\192.168.1.33\\Annmy\\订单标签自动生成软件\\测试\\1.btw");
                        // 设置打印机名称
                        labelFormat.PrintSetup.PrinterName = _PrinterName;

                        // 从文本框获取打印次数
                        int 次数 = Convert.ToInt32(textBox1.Text);

                        // 确保labelFormat和次数是有效的
                        if (labelFormat == null || 次数 <= 0)
                        {
                            MessageBox.Show("无法打开标签文件或打印次数无效。");
                            return;
                        }

                        if (!string.IsNullOrEmpty(Box_数据库.Text))
                        {
                            // 确保filePath是有效的Excel文件路径
                            string filePath = Box_数据库.Text;
                            // 使用EPPlus打开Excel文件
                            using (var package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                // 读取工作表
                                var worksheet = package.Workbook.Worksheets[0];

                                // 遍历所有PDF文件的完整路径
                                for (int i = 0; i < pdfFiles.Count; i++)
                                {
                                    // 从完整路径中提取文件名
                                    string pdfFileName = pdfFiles[i];

                                    // 将当前PDF文件名设置到标签格式中
                                    labelFormat.SubStrings.SetSubString("PDF", pdfFileName);
                                    //MessageBox.Show(pdfFileName); // 显示文件名

                                    // 根据文件名中的序号确定工作表中相应的行号
                                    // 假设文件名格式为 "page_X.pdf"，序号X从1开始
                                    string pageNumber = Regex.Match(pdfFileName, @"page_(\d+)\.pdf").Groups[1].Value;
                                    int excelRow = 2 + int.Parse(pageNumber) - 1; // 将页码转换为工作表行号

                                    // 确保excelRow在工作表行范围内
                                    if (excelRow >= 2 && excelRow <= worksheet.Dimension.End.Row)
                                    {
                                        // 从工作表读取打印次数
                                        var gData = worksheet.Cells[excelRow, 7].Value?.ToString() ?? "0";
                                        int 次数1 = Convert.ToInt32(gData);

                                        //MessageBox.Show($"打印次数: {次数1}"); // 显示打印次数

                                        // 执行打印操作指定的次数
                                        for (int i1 = 0; i1 < 次数1; i1++)
                                        {
                                            labelFormat.Print("BarPrint" + DateTime.Now, 300);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show($"工作表中没有找到对应的行: {excelRow}");
                                    }
                                }
                            }
                        }
                        else
                        {
                            // 遍历所有PDF文件的完整路径
                            for (int i = 0; i < pdfFiles.Count; i++)
                            {
                                // 从完整路径中提取文件名
                                string pdfFileName = pdfFiles[i];

                                // 将当前PDF文件名设置到标签格式中
                                labelFormat.SubStrings.SetSubString("PDF", pdfFileName);
                                //MessageBox.Show(pdfFileName,次数.ToString()); // 显示文件名

                                // 执行打印操作指定的次数
                                for (int j = 0; j < 次数; j++)
                                {
                                    //MessageBox.Show("打印");
                                    labelFormat.Print("BarPrint" + DateTime.Now, 300); // 假设Print方法接受日志文件名和打印质量参数
                                }
                            }
                        }
                    }
                    //如果是品名标的时候
                    else
                    {
                        LabelFormatDocument labelFormat = btEngine.Documents.Open("\\\\192.168.1.33\\Annmy\\订单标签自动生成软件\\moban\\品名标\\客制型号\\(13453) ORLIGHT\\1.btw");
                        //LabelFormatDocument labelFormat = btEngine.Documents.Open("\\\\192.168.1.33\\Annmy\\订单标签自动生成软件\\测试\\1.btw");
                        // 设置打印机名称
                        labelFormat.PrintSetup.PrinterName = _PrinterName;

                        // 从文本框获取打印次数
                        int 次数 = Convert.ToInt32(textBox1.Text);

                        // 确保labelFormat和次数是有效的
                        if (labelFormat == null || 次数 <= 0)
                        {
                            MessageBox.Show("无法打开标签文件或打印次数无效。");
                            return;
                        }

                        if (!string.IsNullOrEmpty(Box_数据库.Text))
                        {
                            // 确保filePath是有效的Excel文件路径
                            string filePath = Box_数据库.Text;
                            // 使用EPPlus打开Excel文件
                            using (var package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                // 读取工作表
                                var worksheet = package.Workbook.Worksheets[0];

                                // 遍历所有PDF文件的完整路径
                                for (int i = 0; i < pdfFiles.Count; i++)
                                {
                                    // 从完整路径中提取文件名
                                    string pdfFileName = pdfFiles[i];

                                    // 将当前PDF文件名设置到标签格式中
                                    labelFormat.SubStrings.SetSubString("PDF", pdfFileName);
                                    //MessageBox.Show(pdfFileName); // 显示文件名

                                    // 根据文件名中的序号确定工作表中相应的行号
                                    // 假设文件名格式为 "page_X.pdf"，序号X从1开始
                                    string pageNumber = Regex.Match(pdfFileName, @"page_(\d+)\.pdf").Groups[1].Value;
                                    int excelRow = 2 + int.Parse(pageNumber) - 1; // 将页码转换为工作表行号

                                    // 确保excelRow在工作表行范围内
                                    if (excelRow >= 2 && excelRow <= worksheet.Dimension.End.Row)
                                    {
                                        // 从工作表读取打印次数
                                        var gData = worksheet.Cells[excelRow, 7].Value?.ToString() ?? "0";
                                        int 次数1 = Convert.ToInt32(gData);

                                        //MessageBox.Show($"打印次数: {次数1}"); // 显示打印次数

                                        // 执行打印操作指定的次数
                                        for (int i1 = 0; i1 < 次数1; i1++)
                                        {
                                            labelFormat.Print("BarPrint" + DateTime.Now, 300);
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show($"工作表中没有找到对应的行: {excelRow}");
                                    }
                                }
                            }
                        }
                        else
                        {
                            // 遍历所有PDF文件的完整路径
                            for (int i = 0; i < pdfFiles.Count; i++)
                            {
                                // 从完整路径中提取文件名
                                string pdfFileName = pdfFiles[i];

                                // 将当前PDF文件名设置到标签格式中
                                labelFormat.SubStrings.SetSubString("PDF", pdfFileName);
                                //MessageBox.Show(pdfFileName,次数.ToString()); // 显示文件名

                                // 执行打印操作指定的次数
                                for (int j = 0; j < 次数; j++)
                                {
                                    //MessageBox.Show("打印");
                                    labelFormat.Print("BarPrint" + DateTime.Now, 300); // 假设Print方法接受日志文件名和打印质量参数
                                }
                            }
                        }
                    }
                }
            }
            else
            {
                using (Engine btEngine = new Engine(true))
                {
                    // 获取运行目录下的PDF拆分文件夹路径
                    string pdfSplitDirectory = Path.Combine(Application.StartupPath + @"\PDF拆分");

                    // 确保PDF拆分目录存在
                    if (!Directory.Exists(pdfSplitDirectory))
                    {
                        MessageBox.Show("PDF拆分目录不存在。");
                        return;
                    }

                    // 获取PDF拆分目录下的所有PDF文件
                    var pdfFiles = Directory.GetFiles(pdfSplitDirectory, "*.pdf")
                      .Select(Path.GetFileName) // 直接提取文件名
                      .OrderBy(fileName => int.Parse(fileName.Substring(5, fileName.Length - 9))) // 提取数字并排序
                      .ToList();

                    // 打开标签格式文件
                    LabelFormatDocument labelFormat = btEngine.Documents.Open("\\\\192.168.1.33\\Annmy\\订单标签自动生成软件\\测试\\1.btw");

                    // 设置打印机名称
                    labelFormat.PrintSetup.PrinterName = _PrinterName;

                    // 从文本框获取打印次数
                    int 次数 = Convert.ToInt32(textBox1.Text);

                    // 确保labelFormat和次数是有效的
                    if (labelFormat == null || 次数 <= 0)
                    {
                        MessageBox.Show("无法打开标签文件或打印次数无效。");
                        return;
                    }

                    if (!string.IsNullOrEmpty(Box_数据库.Text))
                    {
                        // 确保filePath是有效的Excel文件路径
                        string filePath = Box_数据库.Text;
                        // 使用EPPlus打开Excel文件
                        using (var package = new ExcelPackage(new FileInfo(filePath)))
                        {
                            // 读取工作表
                            var worksheet = package.Workbook.Worksheets[0];

                            // 遍历所有PDF文件的完整路径
                            for (int i = 0; i < pdfFiles.Count; i++)
                            {
                                // 从完整路径中提取文件名
                                string pdfFileName = pdfFiles[i];

                                // 将当前PDF文件名设置到标签格式中
                                labelFormat.SubStrings.SetSubString("PDF", pdfFileName);
                                //MessageBox.Show(pdfFileName); // 显示文件名

                                // 根据文件名中的序号确定工作表中相应的行号
                                // 假设文件名格式为 "page_X.pdf"，序号X从1开始
                                string pageNumber = Regex.Match(pdfFileName, @"page_(\d+)\.pdf").Groups[1].Value;
                                int excelRow = 2 + int.Parse(pageNumber) - 1; // 将页码转换为工作表行号

                                // 确保excelRow在工作表行范围内
                                if (excelRow >= 2 && excelRow <= worksheet.Dimension.End.Row)
                                {
                                    // 从工作表读取打印次数
                                    var gData = worksheet.Cells[excelRow, 7].Value?.ToString() ?? "0";
                                    int 次数1 = Convert.ToInt32(gData);

                                    //MessageBox.Show($"打印次数: {次数1}"); // 显示打印次数

                                    // 执行打印操作指定的次数
                                    for (int i1 = 0; i1 < 次数1; i1++)
                                    {
                                        labelFormat.Print("BarPrint" + DateTime.Now, 300);
                                    }
                                }
                                else
                                {
                                    MessageBox.Show($"工作表中没有找到对应的行: {excelRow}");
                                }
                            }
                        }
                    }
                    else
                    {
                        // 遍历所有PDF文件的完整路径
                        for (int i = 0; i < pdfFiles.Count; i++)
                        {
                            // 从完整路径中提取文件名
                            string pdfFileName = pdfFiles[i];

                            // 将当前PDF文件名设置到标签格式中
                            labelFormat.SubStrings.SetSubString("PDF", pdfFileName);
                            //MessageBox.Show(pdfFileName,次数.ToString()); // 显示文件名

                            // 执行打印操作指定的次数
                            for (int j = 0; j < 次数; j++)
                            {
                                //MessageBox.Show("打印");
                                labelFormat.Print("BarPrint" + DateTime.Now, 300); // 假设Print方法接受日志文件名和打印质量参数
                            }
                        }
                    }
                }
            }
        }

        private void button_包装计算_Click_1(object sender, EventArgs e)
        {
            

            try
            {
                string relativePath = Application.StartupPath + @"\包装计算\net8.0-windows\包装计算.exe";
                Process.Start(relativePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"启动程序失败：{ex.Message}", "错误");
            }
        }

        private void 清空数据库_Click(object sender, EventArgs e)
        {
            Box_数据库.Text = "";
            Box_数据库.BackColor = System.Drawing.Color.White;
        }

        private void 剪切板内容到规格型号_Click(object sender, EventArgs e)
        {
            try
            {
                // 检查剪贴板中是否包含文本
                if (Clipboard.ContainsText())
                {
                    // 获取剪贴板文本
                    string clipboardText = Clipboard.GetText();

                    // 设置到cpxxBox
                    cpxxBox.Text = clipboardText;

                    // 可选：让文本框获得焦点
                    cpxxBox.Focus();

                    // 可选：全选文本
                    cpxxBox.SelectAll();

                    // 可选：添加成功提示
                    // MessageBox.Show("已粘贴剪贴板内容", "成功");
                }
                else
                {
                    MessageBox.Show("剪贴板中没有文本内容", "提示");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"粘贴内容时出错：{ex.Message}", "错误");
            }
        }

        

        private void 纯备注自动_Click(object sender, EventArgs e)
        {
            // 假设您的TabControl名称为tabControl1
            // 查找名为"tabPage4"的TabPage
            foreach (TabPage page in tabControl_唛头.TabPages)
            {
                if (page.Name == "tabPage4")
                {
                    // 选择该TabPage
                    tabControl_唛头.SelectedTab = page;
                    break;
                }
            }

            button_订单导入.PerformClick();

            // 或者如果您知道tabPage4的索引，可以直接使用:
            // tabControl1.SelectedIndex = 3; // 假设tabPage4是第四个选项卡(索引为3)
        }

        private async void button_AI_Click(object sender, EventArgs e)
        {
            try
            {
                // 创建输入框窗体
                using (var inputForm = new Form())
                {
                    inputForm.Text = "输入需求";
                    inputForm.Size = new Size(500, 300);
                    inputForm.StartPosition = FormStartPosition.CenterParent;

                    var textBox = new TextBox
                    {
                        Multiline = true,
                        ScrollBars = ScrollBars.Vertical,
                        Size = new Size(460, 180),
                        Location = new Point(10, 10)
                    };

                    var button = new Button
                    {
                        Text = "生成代码",
                        Size = new Size(100, 30),
                        Location = new Point(190, 200),
                        DialogResult = DialogResult.OK
                    };

                    inputForm.Controls.AddRange(new Control[] { textBox, button });
                    inputForm.AcceptButton = button;

                    if (inputForm.ShowDialog() == DialogResult.OK && !string.IsNullOrWhiteSpace(textBox.Text))
                    {
                        // 构造请求数据
                        var requestData = new
                        {
                            model = "claude-3-5-sonnet-20240620",
                            messages = new[]
                            {
                        new
                        {
                            role = "user",
                           content = $@"请生成一个简单的Python代码示例：
                           {textBox.Text}

                           要求：
                           1. 只返回Python代码，不要有其他说明
                           2. 只使用Python标准库（如random, math等），不要使用需要额外安装的库
                           3. 代码必须是完整可运行的
                           4. 使用print输出结果
                           5. 确保代码格式规范
                           6. 代码要简单易懂"
                        }
                    },
                            temperature = 0.7,
                            max_tokens = 2000
                        };

                        using (var client = new HttpClient())
                        {
                            const string apiKey = "sk-9S1Znc8NehWPQflzDqESpNL1xYykJNVcFCAko0RoeOuBhgzu";
                            const string apiUrl = "https://www.dmxapi.com/v1/chat/completions";

                            client.DefaultRequestHeaders.Clear();
                            client.DefaultRequestHeaders.Authorization = new System.Net.Http.Headers.AuthenticationHeaderValue("Bearer", apiKey);
                            client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                            var jsonContent = System.Text.Json.JsonSerializer.Serialize(requestData);
                            var httpContent = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                            var response = await client.PostAsync(apiUrl, httpContent);

                            if (response.IsSuccessStatusCode)
                            {
                                var responseJson = await response.Content.ReadAsStringAsync();
                                var responseObj = System.Text.Json.JsonDocument.Parse(responseJson);

                                if (responseObj.RootElement.TryGetProperty("choices", out var choices) &&
                                    choices.GetArrayLength() > 0)
                                {
                                    string generatedCode = choices[0].GetProperty("message")
                                                                   .GetProperty("content")
                                                                   .GetString();

                                    // 保存生成的代码到文件
                                    string appPath = Application.StartupPath;
                                    string codeDir = Path.Combine(appPath, "CODE");

                                    // 确保目录存在
                                    if (!Directory.Exists(codeDir))
                                    {
                                        Directory.CreateDirectory(codeDir);
                                    }

                                    string codePath = Path.Combine(codeDir, "run.py");
                                    File.WriteAllText(codePath, generatedCode, Encoding.UTF8);

                                    MessageBox.Show("Python代码已生成并保存到run.py", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                            }
                            else
                            {
                                var errorResponse = await response.Content.ReadAsStringAsync();
                                MessageBox.Show($"API调用失败: {response.StatusCode}\n{errorResponse}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"生成代码出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}