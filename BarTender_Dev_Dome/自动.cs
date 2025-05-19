using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml; // EPPlus的命名空间
using OfficeOpenXml.Style;
using System.IO;
using System.Text.RegularExpressions;
using Seagull.BarTender.Print;

// 定义一个类来存储灯带和线材的信息
public class 产品信息
{
    public string 规格型号 { get; set; }
    public string 原始剪切长度 { get; set; }
    public List<(double 长度, int 数量)> 解析长度列表 { get; set; } = new List<(double, int)>();
    public List<线材信息> 相关线材 { get; set; } = new List<线材信息>();

    // 计算属性：获取所有灯带长度的列表

    public override string ToString()
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"{规格型号}");
        

        if (相关线材.Count > 0)
        {
            foreach (var 线材 in 相关线材)
            {
                string 线材信息 = 线材.ToString().Trim();
                if (!string.IsNullOrEmpty(线材信息))
                {
                    sb.AppendLine(线材信息);
                }
            }
        }

        if (解析长度列表.Count > 0)
        {
            foreach (var (长度, 数量) in 解析长度列表)
            {
                sb.AppendLine($"{长度}米");
                sb.AppendLine($"{数量}条");
            }
        }

        return sb.ToString();
    }
}

public class 线材信息
{
    public string 规格型号 { get; set; }
    public string 原始剪切长度 { get; set; }
    public List<(double 长度, int 数量)> 解析长度列表 { get; set; } = new List<(double, int)>();

    public override string ToString()
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine($"{规格型号}");
        if (解析长度列表.Count > 0)
        {
            sb.AppendLine("解析长度:");
            foreach (var (长度, 数量) in 解析长度列表)
            {
                sb.AppendLine($"{长度}米");
                sb.AppendLine($"{数量}条");
            }
        }

        return sb.ToString();
    }
}

public class 结果数据
{
    public string 产品型号 { get; set; }
    public string 销售数量 { get; set; }
    public string 备注 { get; set; }
    public List<(string 序号, string 条数, string 米数, string 标签码1, string 标签码2,string 线长)> 每个规格包装详情 { get; set; } = new List<(string 序号, string 条数, string 米数, string 标签码1, string 标签码2,string 线长)>();

}


namespace BarTender_Dev_Dome
{
    public partial class PrintForm : Form
    {
        private string 订单地址 = string.Empty;
        private string 客户代码 = string.Empty;
        private string 标签 = string.Empty;
        public List<产品信息> 产品信息列表 = new List<产品信息>();
        private string 待处理规格型号 = string.Empty;

        int 物料编码列 = -1;

        private void button_订单导入_Click(object sender, EventArgs e)
        {
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Multiselect = false;
            dialog.Title = "请选择数据库文件";
            dialog.Filter = "excel文件(*.xlsx)|*.xlsx|All files (*.*)|*.*";
            dialog.InitialDirectory = Application.StartupPath + @"\数据库";

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                订单地址 = dialog.FileName;
                textBox_订单地址.Text = dialog.FileName;

                EXCEL订单数据_提取数据(订单地址);

                自动打印_工字标(标签); 

                //自动打印_品名标(标签); 

            }
        }

        public void EXCEL订单数据_提取数据(string excel文件路径)
        {
            using (var package = new ExcelPackage(new FileInfo(excel文件路径)))
            {
                var worksheet = package.Workbook.Worksheets[0];

                // 初始化列索引
                int 客户代码列 = -1;
                int 规格型号列 = -1;
                int 剪切长度列 = -1;
                int 标签列 = -1;

                // 在第1行查找列标题
                for (int col = 1; col <= worksheet.Dimension.End.Column; col++) 
                {
                    string 列标题 = worksheet.Cells[1, col].Text?.Trim();
                    if (string.IsNullOrEmpty(列标题)) continue;
                    switch (列标题) 
                    {
                        case "客户代码":
                            客户代码列 = col;
                            break;

                        case "物料编码":
                            物料编码列 = col;
                            break;

                        case "规格型号":
                            规格型号列 = col;
                            break;

                        case "剪切长度":
                            剪切长度列 = col;
                            break;

                        case "标签":
                            标签列 = col;
                            break;
                    }
                }

                // 验证必要的列是否都找到
                if (客户代码列 == -1 || 物料编码列 == -1 || 规格型号列 == -1)
                {
                    throw new Exception("未找到必要的列标题（单据编号、规格型号或销售数量）");
                }

                客户代码 = worksheet.Cells[2, 客户代码列].Text;
                标签 = worksheet.Cells[2, 标签列].Text;

                int lastStartRow = -1;

                //开始筛选80.
                for (int row = 2; row <= worksheet.Dimension.End.Row; row++) 
                {
                    var materialCell = worksheet.Cells[row, 物料编码列];
                    //MessageBox.Show($"物料编码列：{materialCell.Text}", "错误");

                    if (materialCell.Value != null && materialCell.Text.StartsWith("80.")) 
                    {
                        // 找到一个80.开头的行
                        if (lastStartRow != -1)
                        {
                            // 处理上一个80.到当前80.之间的规格型号和剪切长度
                            ProcessSpecifications(worksheet, lastStartRow, row - 1, 规格型号列, 剪切长度列, 产品信息列表);
                        }
                        // 更新lastStartRow为当前行
                        lastStartRow = row;
                    }
                }

                // 处理最后一个80.到表格末尾的规格型号和剪切长度
                if (lastStartRow != -1)
                {
                    ProcessSpecifications(worksheet, lastStartRow, worksheet.Dimension.End.Row, 规格型号列, 剪切长度列, 产品信息列表);
                }

                

            }
            
        }

        // 辅助方法：处理指定范围内的规格型号和剪切长度
        private void ProcessSpecifications(ExcelWorksheet worksheet, int startRow, int endRow, int specColumn, int cutLengthColumn, List<产品信息> productList)
        {
            // 获取80.行的规格型号和剪切长度
            string mainSpec = worksheet.Cells[startRow, specColumn].Text;
            string mainCutLength = worksheet.Cells[startRow, cutLengthColumn].Text;

            // 创建新的产品信息对象
            产品信息 product = new 产品信息
            {
                规格型号 = mainSpec,
                原始剪切长度 = mainCutLength
            };

            // 解析剪切长度
            if (!string.IsNullOrEmpty(mainCutLength))
            {
                string[] 长度组 = mainCutLength.Split(new[] { ',', '，', ';', '；', '+' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (string 单个长度 in 长度组)
                {
                    var match = Regex.Match(单个长度.Trim(),
                        @"(\d+(?:\.\d+)?)\s*[Mm][^*]*\*\s*(\d+)\s*(?:PC|PCS|pc|pcs)",
                        RegexOptions.IgnoreCase);

                    if (match.Success)
                    {
                        double 长度 = double.Parse(match.Groups[1].Value);
                        int 数量 = int.Parse(match.Groups[2].Value);
                        product.解析长度列表.Add((长度, 数量));
                    }
                }
            }

            // 获取80.行之后的规格型号和剪切长度（非80.开头的行）
            for (int i = startRow + 1; i <= endRow; i++)
            {
                var currentMaterialCell = worksheet.Cells[i, 物料编码列];
                if (currentMaterialCell.Value != null && !currentMaterialCell.Text.StartsWith("80."))
                {
                    string spec = worksheet.Cells[i, specColumn].Text;
                    string cutLength = worksheet.Cells[i, cutLengthColumn].Text;

                    线材信息 wire = new 线材信息
                    {
                        规格型号 = spec,
                        原始剪切长度 = cutLength
                    };

                    // 解析剪切长度
                    if (!string.IsNullOrEmpty(cutLength))
                    {
                        string[] 长度组 = cutLength.Split(new[] { ',', '，', ';', '；', '+' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (string 单个长度 in 长度组)
                        {
                            var match = Regex.Match(单个长度.Trim(),
                                @"(\d+(?:\.\d+)?)\s*[Mm][^*]*\*\s*(\d+)\s*(?:PC|PCS|pc|pcs)",
                                RegexOptions.IgnoreCase);

                            if (match.Success)
                            {
                                double 长度 = double.Parse(match.Groups[1].Value);
                                int 数量 = int.Parse(match.Groups[2].Value);
                                wire.解析长度列表.Add((长度, 数量));
                            }
                        }
                    }

                    product.相关线材.Add(wire);
                }
            }

            // 将产品信息添加到列表中
            productList.Add(product);
        }



        //打印工字标
        public void 自动打印_工字标(string 路径)
        {
            标签种类_comboBox.Text = "工字标";
            if (产品信息列表.Count > 0)
            {
                // 显示找到的产品信息总数
                MessageBox.Show($"共找到 {产品信息列表.Count} 个产品信息项目，准备开始打印", "产品信息数量");

                // 初始化BarTender引擎
                Engine btEngine = new Engine();
                btEngine.Start(); // 确保引擎已启动

                try
                {

                    // 遍历所有产品信息
                    for (int i = 0; i < 产品信息列表.Count; i++)
                    {
                        // 获取当前产品信息
                        var 产品 = 产品信息列表[i];

                        // 如果产品没有解析长度列表，则跳过
                        if (产品.解析长度列表.Count == 0)
                        {
                            MessageBox.Show($"产品 [{i + 1}] {产品.ToString()} 没有解析到长度信息，跳过打印", "提示");
                            continue;
                        }

                        // 设置产品规格型号到cpxxBox
                        cpxxBox.Text = 产品.ToString();

                        // 遍历产品的所有长度规格
                        foreach (var (长度, 数量) in 产品.解析长度列表)
                        {
                            // 设置长度到textBox_剪切长度
                            textBox_剪切长度.Text = 长度.ToString()+"m";

                            // 设置数量到textBox1
                            textBox1.Text = 数量.ToString();

                            // 直接执行预览按钮的点击事件
                            print_btn.PerformClick();

                            
                        }

                        
                    }

                    // 所有产品打印完成
                    //MessageBox.Show("工字标打印完成，接下来打印品名标", "工字标打印完成");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"打印过程中发生错误: {ex.Message}", "错误");
                }
                finally
                {
                    // 关闭BarTender引擎
                    if (btEngine != null)
                    {
                        btEngine.Stop();
                        btEngine.Dispose();
                    }
                }
            }
            else
            {
                MessageBox.Show("未找到任何产品信息，无法打印", "提示");
            }
        }

        //打印工字标
        public void 自动打印_品名标(string 路径)
        {
            标签种类_comboBox.Text = "品名标";
            if (产品信息列表.Count > 0)
            {
                // 显示找到的产品信息总数
                MessageBox.Show($"共找到 {产品信息列表.Count} 个产品信息项目，准备开始打印", "产品信息数量");

                // 初始化BarTender引擎
                Engine btEngine = new Engine();
                btEngine.Start(); // 确保引擎已启动

                try
                {

                    // 遍历所有产品信息
                    for (int i = 0; i < 产品信息列表.Count; i++)
                    {
                        // 获取当前产品信息
                        var 产品 = 产品信息列表[i];

                        // 如果产品没有解析长度列表，则跳过
                        if (产品.解析长度列表.Count == 0)
                        {
                            MessageBox.Show($"产品 [{i + 1}] {产品.ToString()} 没有解析到长度信息，跳过打印", "提示");
                            continue;
                        }

                        // 设置产品规格型号到cpxxBox
                        cpxxBox.Text = 产品.ToString();

                        // 遍历产品的所有长度规格
                        foreach (var (长度, 数量) in 产品.解析长度列表)
                        {
                            // 设置长度到textBox_剪切长度
                            textBox_剪切长度.Text = 长度.ToString();

                            // 设置数量到textBox1
                            textBox1.Text = 数量.ToString();

                            // 直接执行预览按钮的点击事件
                            print_btn.PerformClick();


                        }


                    }

                    // 所有产品打印完成
                    MessageBox.Show("品名标打印完成", "品名标打印完成");
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"打印过程中发生错误: {ex.Message}", "错误");
                }
                finally
                {
                    // 关闭BarTender引擎
                    if (btEngine != null)
                    {
                        btEngine.Stop();
                        btEngine.Dispose();
                    }
                }
            }
            else
            {
                MessageBox.Show("未找到任何产品信息，无法打印", "提示");
            }
        }

    }
}