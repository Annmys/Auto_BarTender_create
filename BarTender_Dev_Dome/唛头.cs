using OfficeOpenXml; // EPPlus的命名空间
using Seagull.BarTender.Print;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Application = System.Windows.Forms.Application;
using Resolution = Seagull.BarTender.Print.Resolution;

namespace BarTender_Dev_Dome

{
    public partial class PrintForm : Form
    {
        private string lastFolderName = "";
        private string folderPath = "";

        public void shengcheng_maitou(biaoqian actionType) // 方法名称建议使用大写开头，例如：Test
        {
            using (Engine btEngine = new Engine(true))
            {
                if (!_btw_path.Contains(comboBox_标签规格.Text))
                {
                    string basePath = @"\\192.168.1.33\Annmy\订单标签自动生成软件\moban\" + 标签种类_comboBox.Text + @"\";
                    string modelTypePath = checkBox_常规型号.Checked ? "常规型号" : checkBox_客制型号.Checked ? "客制型号" : checkBox_简化型号.Checked ? "简化型号" : null;

                    if (modelTypePath != null)
                    {
                        _btw_path = basePath + modelTypePath + @"\" + comboBox_标签规格.Text;
                    }
                }

                // 清空comboBox_唛头规格的现有项
                //comboBox_唛头规格.Items.Clear();
                // 获取选定的标签规格
                //string selectedSpec = comboBox_标签规格.SelectedItem?.ToString();
                //MessageBox.Show($"选择的标签规格: {selectedSpec ?? "未选择"}", "选择确认", MessageBoxButtons.OK, MessageBoxIcon.Information);

                _wjm_ = comboBox_唛头规格.Text + ".btw";

                //寻找文件名_单字匹配("正弯", "侧弯");
                string 复选框 = 判断复选框内容(cpxxBox.Text, comboBox_标签规格.Text);

                模板地址 = _btw_path + @"\" + _wjm_;
                //MessageBox.Show(模板地址, "操作提示");
                模板地址 = 模板地址.Replace("\n", string.Empty).Replace("\r", string.Empty);  //去除换行符，否则下面会报错

                string cpxx_text = cpxxBox.Text;
                判断产品信息(cpxx_text);

                if (_wjm_.Length > 2)
                {
                    //先判断文本内容
                    if (checkBox_唛头型号自动.Checked) { 唛头_寻找灯带型号(); }
                    if (checkBox_唛头电压自动.Checked) { 唛头_寻找灯带电压(); }

                    //string 唛头灯带型号 = output_灯带型号;
                    //textBox_唛头灯带型号.Text = 唛头灯带型号;
                    //string 唛头电压 = output_电压;
                    //textBox_唛头电压.Text = 唛头电压;

                    string 唛头色温 = output_色温.Replace("Color:", "").Trim();
                    textBox_唛头色温.Text = 唛头色温;
                    //唛头_寻找订单编号();

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

                    //labelFormat.SubStrings.SetSubString("CPXX-01", " ");

                    //判断尾巴Made in China
                    if (comboBox_标签规格.Text.Contains("13013")) { 唛头_产品信息.AppendLine("Made in China"); }
                    else if (checkBox_结尾.Checked) { labelFormat.SubStrings.SetSubString("ZGZZ", textBox_结尾.Text); }
                    else { output_尾巴 = " "; }

                    labelFormat.SubStrings.SetSubString("CPXX-01", 唛头_产品信息.ToString());

                    labelFormat.SubStrings.SetSubString("DDBH", textBox_订单编号.Text);

                    //显指判断
                    //if (checkBox_FXK01.Checked)
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-01", "实.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-01", textBox_XZ01.Text);
                    //}
                    //else
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-01", "空2.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-01", textBox_XZ01.Text);
                    //}
                    //if (checkBox_FXK02.Checked)
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-02", "实.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-02", textBox_XZ02.Text);
                    //}
                    //else
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-02", "空2.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-02", textBox_XZ02.Text);
                    //}
                    //if (checkBox_FXK03.Checked)
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-03", "实.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-03", textBox_XZ03.Text);
                    //}
                    //else
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-03", " ");
                    //}
                    //if (checkBox_FXK04.Checked)
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-04", "实.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-04", textBox_XZ04.Text);
                    //}
                    //else
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-04", " ");
                    //}
                    //if (checkBox_FXK05.Checked)
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-05", "实.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-05", textBox_XZ05.Text);
                    //}
                    //else
                    //{
                    //    labelFormat.SubStrings.SetSubString("FXK-05", "空.png");
                    //    labelFormat.SubStrings.SetSubString("XZ-05", " ");
                    //}

                    //判断显指内容
                    // 重置 XZ 字段的内容
                    labelFormat.SubStrings.SetSubString("XZ-01", string.Empty);
                    labelFormat.SubStrings.SetSubString("XZ-02", string.Empty);
                    labelFormat.SubStrings.SetSubString("XZ-03", string.Empty);
                    labelFormat.SubStrings.SetSubString("XZ-04", string.Empty);
                    labelFormat.SubStrings.SetSubString("XZ-05", string.Empty);

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

                    string 复选框1 = 唛头_判断复选框内容(cpxxBox.Text, comboBox_标签规格.Text);
                    labelFormat.SubStrings.SetSubString("FXK-01", "实.png");
                    labelFormat.SubStrings.SetSubString("XZ-01", 复选框1);

                    //常规状态
                    // 检查是否存在 "Ra90" 或 "Ra95"
                    bool containsRa90 = BPrefixContent.Contains("Ra90");
                    bool containsRa95 = BPrefixContent.Contains("Ra95");

                    if (BPrefixContent.Contains("三面发光"))
                    {
                        labelFormat.SubStrings.SetSubString("XZ-02", 灯带系列 + @"T");
                        labelFormat.SubStrings.SetSubString("FXK-02", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", " ");
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", " ");
                        labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", " ");
                        labelFormat.SubStrings.SetSubString("FXK-05", "空.png");

                        // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                        if (containsRa90 || containsRa95)
                        {
                            labelFormat.SubStrings.SetSubString("FXK-03", "实.png");
                            string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
                            labelFormat.SubStrings.SetSubString("XZ-03", raValue);
                        }
                    }
                    else if (BPrefixContent.Contains("高亮"))
                    {
                        labelFormat.SubStrings.SetSubString("XZ-02", "BH");
                        labelFormat.SubStrings.SetSubString("FXK-02", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", " ");
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", " ");
                        labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", " ");
                        labelFormat.SubStrings.SetSubString("FXK-05", "空.png");

                        // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                        if (containsRa90 || containsRa95)
                        {
                            labelFormat.SubStrings.SetSubString("FXK-03", "实.png");
                            string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
                            labelFormat.SubStrings.SetSubString("XZ-03", raValue);
                        }
                    }
                    else if (BPrefixContent.Contains("翻边"))
                    {
                        labelFormat.SubStrings.SetSubString("XZ-02", "BF");
                        labelFormat.SubStrings.SetSubString("FXK-02", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", " ");
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", " ");
                        labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", " ");
                        labelFormat.SubStrings.SetSubString("FXK-05", "空.png");

                        // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                        if (containsRa90 || containsRa95)
                        {
                            labelFormat.SubStrings.SetSubString("FXK-03", "实.png");
                            string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
                            labelFormat.SubStrings.SetSubString("XZ-03", raValue);
                        }
                    }
                    else if (BPrefixContent.Contains("DTW"))
                    {
                        labelFormat.SubStrings.SetSubString("XZ-02", "DTW");
                        labelFormat.SubStrings.SetSubString("FXK-02", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", " ");
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", " ");
                        labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", " ");
                        labelFormat.SubStrings.SetSubString("FXK-05", "空.png");
                        // 如果同时存在DTW和Ra90或Ra95，设置额外的子字符串
                        if (containsRa90 || containsRa95)
                        {
                            labelFormat.SubStrings.SetSubString("FXK-03", "实.png");
                            string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
                            labelFormat.SubStrings.SetSubString("XZ-03", raValue);
                        }
                    }
                    else if (灯带系列 == "D")
                    {
                        labelFormat.SubStrings.SetSubString("XZ-02", "D");
                        labelFormat.SubStrings.SetSubString("FXK-02", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", " ");
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", " ");
                        labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", " ");
                        labelFormat.SubStrings.SetSubString("FXK-05", "空.png");
                        if (containsRa90 || containsRa95)
                        {
                            labelFormat.SubStrings.SetSubString("FXK-03", "实.png");
                            string raValue = containsRa90 ? "Ra90" : (containsRa95 ? "Ra95" : string.Empty);
                            labelFormat.SubStrings.SetSubString("XZ-03", raValue);
                        }
                    }
                    else if (BPrefixContent.Contains("Ra90"))
                    {
                        labelFormat.SubStrings.SetSubString("XZ-02", "Ra90");
                        labelFormat.SubStrings.SetSubString("FXK-02", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", " ");
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", " ");
                        labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", " ");
                        labelFormat.SubStrings.SetSubString("FXK-05", "空.png");

                        if (comboBox_标签规格.Text.Contains("13013"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ-02", " ");
                            labelFormat.SubStrings.SetSubString("FXK-02", "空.png");
                        }
                        else if (comboBox_标签规格.Text.Contains("17100"))
                        {
                            labelFormat.SubStrings.SetSubString("FXK-03", "实.png");
                        }
                    }
                    else if (BPrefixContent.Contains("Ra95"))
                    {
                        labelFormat.SubStrings.SetSubString("XZ-02", "Ra95");
                        labelFormat.SubStrings.SetSubString("FXK-02", "实.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", " ");
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", " ");
                        labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", " ");
                        labelFormat.SubStrings.SetSubString("FXK-05", "空.png");

                        if (comboBox_标签规格.Text.Contains("13013"))
                        {
                            labelFormat.SubStrings.SetSubString("XZ-02", " ");
                            labelFormat.SubStrings.SetSubString("FXK-02", "空.png");
                        }
                        else if (comboBox_标签规格.Text.Contains("17100"))
                        {
                            labelFormat.SubStrings.SetSubString("FXK-03", "实.png");
                        }
                    }
                    else
                    {
                        // 如果没有找到上述任何关键字，则设置为空

                        labelFormat.SubStrings.SetSubString("XZ-02", " ");
                        //labelFormat.SubStrings.SetSubString("FXK", "空.png");

                        labelFormat.SubStrings.SetSubString("FXK-02", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-03", " ");
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-04", " ");
                        labelFormat.SubStrings.SetSubString("FXK-04", "空.png");
                        labelFormat.SubStrings.SetSubString("XZ-05", " ");
                        labelFormat.SubStrings.SetSubString("FXK-05", "空.png");
                    }

                    if (comboBox_标签规格.Text.Contains("13013"))
                    {
                        labelFormat.SubStrings.SetSubString("XZ-02", " ");
                        labelFormat.SubStrings.SetSubString("FXK-03", "空.png");
                    }

                    switch (actionType)
                    {
                        //生成预览图
                        case biaoqian.yulan:
                            // 创建一个预览窗体
                            Form previewForm = new Form
                            {
                                Text = "唛头预览",
                                Size = new Size(500, 400),
                                BackColor = Color.White,
                                FormBorderStyle = FormBorderStyle.Sizable,
                                MaximizeBox = true,
                                MinimizeBox = true,
                                TopMost = false
                            };

                            // 获取主窗体位置和大小
                            Form mainForm = this.FindForm();

                            // 计算预览窗体位置 - 放在主窗体右侧
                            Screen currentScreen = Screen.FromControl(mainForm);
                            int mainFormRight = mainForm.Location.X + mainForm.Width;
                            int availableRightSpace = currentScreen.WorkingArea.Right - mainFormRight;

                            int x, y;
                            if (availableRightSpace >= previewForm.Width + 10)
                            {
                                x = mainFormRight + 10;
                                y = mainForm.Location.Y;
                            }
                            else
                            {
                                x = currentScreen.WorkingArea.Right - previewForm.Width;
                                y = currentScreen.WorkingArea.Top;
                            }

                            previewForm.Location = new Point(x, y);

                            // 创建一个PictureBox来显示唛头预览
                            PictureBox pictureBox = new PictureBox
                            {
                                Size = new Size(400, 300),
                                Location = new Point(50, 50),
                                BorderStyle = BorderStyle.FixedSingle,
                                SizeMode = PictureBoxSizeMode.Zoom,
                                BackColor = Color.White
                            };

                            // 添加一个关闭按钮
                            Button closeButton = new Button
                            {
                                Text = "关闭预览",
                                Location = new Point(10, 10),
                                Size = new Size(100, 30)
                            };
                            closeButton.Click += (s, args) => previewForm.Close();

                            // 组装界面
                            previewForm.Controls.Add(pictureBox);
                            previewForm.Controls.Add(closeButton);

                            // 生成预览图
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

                            // 显示预览窗体
                            previewForm.Show();

                            // 创建一个定时器用于自动关闭和更新标题
                            System.Windows.Forms.Timer autoCloseTimer = new System.Windows.Forms.Timer();
                            autoCloseTimer.Interval = 1000;
                            int remainingSeconds = 30;
                            string originalTitle = previewForm.Text;

                            autoCloseTimer.Tick += (timerSender, timerArgs) =>
                            {
                                remainingSeconds--;
                                if (remainingSeconds > 0)
                                {
                                    if (previewForm != null && !previewForm.IsDisposed)
                                    {
                                        previewForm.Text = $"{originalTitle} ({remainingSeconds} 秒后自动关闭)";
                                    }
                                }
                                else
                                {
                                    autoCloseTimer.Stop();
                                    if (previewForm != null && !previewForm.IsDisposed)
                                    {
                                        previewForm.Close();
                                    }
                                    else
                                    {
                                        autoCloseTimer.Dispose();
                                    }
                                }
                            };

                            // 添加窗体关闭时的资源清理
                            previewForm.FormClosed += (s, args) =>
                            {
                                if (pictureBox.Image != null)
                                {
                                    pictureBox.Image.Dispose();
                                    pictureBox.Image = null;
                                }
                                if (autoCloseTimer != null)
                                {
                                    if (autoCloseTimer.Enabled)
                                    {
                                        autoCloseTimer.Stop();
                                    }
                                    autoCloseTimer.Dispose();
                                }
                            };

                            // 启动定时器前先更新标题
                            previewForm.Text = $"{originalTitle} ({remainingSeconds} 秒后自动关闭)";
                            autoCloseTimer.Start();

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

        private string 唛头_判断复选框内容(string input, string 标签规格)
        {
            // 是简化型号的时候
            if (checkBox_简化型号.Checked)
            {
                return "空.png";
            }
            // 不是简化型号的时候
            else
            {
                // 直接检查是否包含正弯或侧弯关键词
                bool hasPositiveBend = input.Contains("正弯");
                bool hasSideBend = input.Contains("侧弯");

                // 如果明确包含关键词，直接返回对应结果
                if (hasPositiveBend)
                {
                    return "VB"; // 正弯输出VB
                }
                else if (hasSideBend)
                {
                    return "HB"; // 侧弯输出HB
                }

                // 如果没有明确关键词，根据标签规格和型号判断
                if (标签规格.Contains("RCM") || 标签规格.Contains("13013"))
                {
                    // 侧弯型号列表
                    var 侧弯型号 = new[] { "F10", "F11", "F15", "F21", "F2222" };
                    // 正弯型号列表
                    var 正弯型号 = new[] { "F16", "F2219" };

                    // 检查是否匹配侧弯型号
                    if (侧弯型号.Any(model => input.Contains(model)))
                    {
                        return "HB"; // 侧弯输出HB
                    }
                    // 检查是否匹配正弯型号
                    else if (正弯型号.Any(model => input.Contains(model)))
                    {
                        return "VB"; // 正弯输出VB
                    }
                    else
                    {
                        return "VB"; // 默认为正弯
                    }
                }

                // 如果没有匹配任何条件，返回空字符串
                return string.Empty;
            }
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

        private void 唛头_寻找订单编号(string 订单号)
        {
            string input = 订单号; // 这里替换成你的实际输入字符串
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

        private void 工字标汇总_Click(object sender, EventArgs e)
        {
            textBox1.Text = "1";
            string cpxx_text = cpxxBox.Text;
            判断产品信息(cpxx_text);
            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            using (var folderDialog = new System.Windows.Forms.OpenFileDialog())
            {
                folderDialog.ValidateNames = false;
                folderDialog.CheckFileExists = false;
                folderDialog.CheckPathExists = true;
                folderDialog.FileName = "选择文件夹";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    folderPath = Path.GetDirectoryName(folderDialog.FileName);
                    lastFolderName = Path.GetFileName(folderPath);
                    唛头_寻找订单编号(lastFolderName);
                    EXCEL_包装规格回调(folderPath);

                    string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
                    List<结果数据> 结果列表 = new List<结果数据>();

                    // 读取单位
                    string 单位文件路径 = Path.Combine(folderPath, "订单资料", "单位.txt");
                    string 单位内容 = "m";
                    if (File.Exists(单位文件路径))
                    {
                        单位内容 = File.ReadAllText(单位文件路径).Trim();
                    }

                    foreach (string filePath in excelFiles)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(filePath);
                        if (fileName.Contains("包装材料需求流转单") || fileName.Contains("流转单"))
                            continue;
                        string[] parts = fileName.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
                        if (parts.Length >= 2)
                        {
                            string 型号 = parts[0].Trim();
                            string 销售数量 = parts[1].Trim();
                            string 备注 = parts.Length > 2 ? parts[2].Trim() : "";
                            结果数据 数据 = new 结果数据
                            {
                                产品型号 = 型号,
                                销售数量 = 销售数量,
                                备注 = 备注
                            };
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                foreach (var worksheet in package.Workbook.Worksheets)
                                {
                                    if (worksheet.Dimension == null) continue;
                                    int headerRow = 1;
                                    for (int row = 1; row <= Math.Min(10, worksheet.Dimension.End.Row); row++)
                                    {
                                        if (worksheet.Cells[row, 1].Text.Contains("序号") || worksheet.Cells[row, 1].Text.Equals("序号", StringComparison.OrdinalIgnoreCase))
                                        {
                                            headerRow = row;
                                            break;
                                        }
                                    }
                                    int 序号列 = -1, 条数列 = -1, 米数列 = -1, 标签码1列 = -1, 标签码2列 = -1, 标签码3列 = -1, 标签码4列 = -1, 线长列 = -1, 客户型号列 = -1, 标签显示长度列 = -1;
                                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                    {
                                        string headerText = worksheet.Cells[headerRow, col].Text.Trim();
                                        if (headerText.Contains("序号") || headerText.Equals("序号", StringComparison.OrdinalIgnoreCase)) 序号列 = col;
                                        else if (headerText.Contains("条数") || headerText.Equals("条数", StringComparison.OrdinalIgnoreCase)) 条数列 = col;
                                        else if (headerText.Contains("米数") || headerText.Equals("米数", StringComparison.OrdinalIgnoreCase)) 米数列 = col;
                                        else if (headerText.Contains("标签码1") || headerText.Equals("标签码1", StringComparison.OrdinalIgnoreCase)) 标签码1列 = col;
                                        else if (headerText.Contains("标签码2") || headerText.Equals("标签码2", StringComparison.OrdinalIgnoreCase)) 标签码2列 = col;
                                        else if (headerText.Contains("标签码3") || headerText.Equals("标签码3", StringComparison.OrdinalIgnoreCase)) 标签码3列 = col;
                                        else if (headerText.Contains("标签码4") || headerText.Equals("标签码4", StringComparison.OrdinalIgnoreCase)) 标签码4列 = col;
                                        else if (headerText.Contains("线长") || headerText.Equals("线长", StringComparison.OrdinalIgnoreCase)) 线长列 = col;
                                        else if (headerText.Contains("客户型号") || headerText.Equals("客户型号", StringComparison.OrdinalIgnoreCase)) 客户型号列 = col;
                                        else if (headerText.Contains("标签显示长度") || headerText.Equals("标签显示长度", StringComparison.OrdinalIgnoreCase)) 标签显示长度列 = col;
                                    }
                                    for (int row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
                                    {
                                        string 序号 = 序号列 > 0 ? worksheet.Cells[row, 序号列].Text.Trim() : "";
                                        if (string.IsNullOrEmpty(序号)) continue;
                                        数据.盒子列表.Add(new List<结果数据.盒子内容> {
                                    new 结果数据.盒子内容 {
                                        序号 = worksheet.Cells[row, 序号列].Text,
                                        条数 = 条数列 > 0 ? worksheet.Cells[row, 条数列].Text : "",
                                        米数 = 米数列 > 0 ? worksheet.Cells[row, 米数列].Text : "",
                                        标签码1 = 标签码1列 > 0 ? worksheet.Cells[row, 标签码1列].Text : "",
                                        标签码2 = 标签码2列 > 0 ? worksheet.Cells[row, 标签码2列].Text : "",
                                        标签码3 = 标签码3列 > 0 ? worksheet.Cells[row, 标签码3列].Text : "",
                                        标签码4 = 标签码4列 > 0 ? worksheet.Cells[row, 标签码4列].Text : "",
                                        线长 = 线长列 > 0 ? worksheet.Cells[row, 线长列].Text : "",
                                        客户型号 = 客户型号列 > 0 ? worksheet.Cells[row, 客户型号列].Text : "",
                                        标签显示长度 = 标签显示长度列 > 0 ? worksheet.Cells[row, 标签显示长度列].Text : ""
                                    }
                                });
                                    }
                                }
                            }
                            if (数据.盒子列表.Count > 0)
                                结果列表.Add(数据);
                        }
                    }

                    // 按客户型号或型号分组
                    var 分组 = 结果列表.GroupBy(x =>
                        x.盒子列表.SelectMany(b => b).FirstOrDefault()?.客户型号 ?? x.产品型号
                    );

                    using (var saveFileDialog = new SaveFileDialog())
                    {
                        saveFileDialog.Filter = "Excel文件|*.xlsx";
                        saveFileDialog.Title = "保存工字标汇总";
                        saveFileDialog.FileName = "工字标汇总.xlsx";

                        if (saveFileDialog.ShowDialog() == DialogResult.OK)
                        {
                            string savePath = saveFileDialog.FileName;
                            using (ExcelPackage package = new ExcelPackage())
                            {
                                var worksheet = package.Workbook.Worksheets.Add("工字标汇总");
                                // 表头
                                worksheet.Cells[1, 1].Value = "序号";
                                worksheet.Cells[1, 2].Value = "标签码1";
                                worksheet.Cells[1, 3].Value = "标签码2";
                                worksheet.Cells[1, 4].Value = "标签码3";
                                worksheet.Cells[1, 5].Value = "标签码4";
                                worksheet.Cells[1, 6].Value = "标签码5";
                                worksheet.Cells[1, 7].Value = "条数";
                                worksheet.Cells[1, 8].Value = "长度";
                                worksheet.Cells[1, 9].Value = "客户型号";
                                worksheet.Cells[1, 10].Value = "PO号";
                                worksheet.Cells[1, 11].Value = "条形码";
                                worksheet.Cells[1, 12].Value = "线长";
                                using (var range = worksheet.Cells[1, 1, 1, 12])
                                {
                                    range.Style.Font.Bold = true;
                                    range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                                    range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                                }
                                int rowIndex = 2;
                                int seqNo = 1;

                                // 合并所有分组数据到一个文件
                                foreach (var group in 分组)
                                {
                                    foreach (var 数据 in group)
                                    {
                                        foreach (var 盒子 in 数据.盒子列表)
                                        {
                                            foreach (var 内容 in 盒子)
                                            {
                                                // 长度字段单位判断
                                                string 长度字段 = "";
                                                switch (单位内容)
                                                {
                                                    case "m":
                                                        长度字段 = $"{内容.米数}m";
                                                        break;

                                                    case "mm":
                                                        if (double.TryParse(内容.米数, out double mVal))
                                                            长度字段 = $"{(int)Math.Round(mVal * 1000)}mm";
                                                        else
                                                            长度字段 = $"{内容.米数}mm";
                                                        break;

                                                    case "IN":
                                                    case "Ft":
                                                        长度字段 = $"{内容.标签显示长度}{单位内容}";
                                                        break;

                                                    case "m(IN)":
                                                    case "m(Ft)":
                                                        string 括号单位 = 单位内容 == "m(IN)" ? "IN" : "Ft";
                                                        长度字段 = $"{内容.米数}m({内容.标签显示长度}{括号单位})";
                                                        break;

                                                    default:
                                                        长度字段 = $"{内容.米数}m";
                                                        break;
                                                }

                                                worksheet.Cells[rowIndex, 1].Value = seqNo++;
                                                worksheet.Cells[rowIndex, 2].Value = 内容.标签码1;
                                                worksheet.Cells[rowIndex, 3].Value = 内容.标签码2;
                                                worksheet.Cells[rowIndex, 4].Value = 内容.标签码3;
                                                worksheet.Cells[rowIndex, 5].Value = 内容.标签码4;
                                                worksheet.Cells[rowIndex, 6].Value = ""; // 标签码5
                                                worksheet.Cells[rowIndex, 7].Value = 内容.条数;
                                                worksheet.Cells[rowIndex, 8].Value = 长度字段;
                                                worksheet.Cells[rowIndex, 9].Value = string.IsNullOrWhiteSpace(内容.客户型号) ? 数据.产品型号 : 内容.客户型号;
                                                worksheet.Cells[rowIndex, 10].Value = ""; // PO号
                                                worksheet.Cells[rowIndex, 11].Value = ""; // 条形码
                                                worksheet.Cells[rowIndex, 12].Value = 内容.线长;
                                                rowIndex++;
                                            }
                                        }
                                    }
                                }
                                worksheet.Cells.AutoFitColumns();
                                package.SaveAs(new FileInfo(savePath));
                            }
                            MessageBox.Show($"工字标汇总已保存到: {savePath}", "完成", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
            }
        }

        private void 品名汇总_Click(object sender, EventArgs e)
        {
            textBox1.Text = "1";
            string cpxx_text = cpxxBox.Text;
            判断产品信息(cpxx_text);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // 使用更现代的文件夹选择对话框
            using (var folderDialog = new System.Windows.Forms.OpenFileDialog())
            {
                folderDialog.ValidateNames = false;
                folderDialog.CheckFileExists = false;
                folderDialog.CheckPathExists = true;
                folderDialog.FileName = "选择文件夹";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    folderPath = Path.GetDirectoryName(folderDialog.FileName);
                    // 获取最后一个文件夹名称
                    lastFolderName = Path.GetFileName(folderPath);
                    //MessageBox.Show(lastFolderName);

                    唛头_寻找订单编号(lastFolderName);
                    EXCEL_包装规格回调(folderPath);

                    // 获取文件夹中的所有Excel文件
                    string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

                    // 创建一个列表来存储解析后的文件信息
                    List<结果数据> 结果列表 = new List<结果数据>();

                    // 创建一个字典来存储纸箱规格信息 - 键是文件名，值是规格列表
                    Dictionary<string, List<string>> 纸箱规格字典 = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

                    // 创建一个StringBuilder来收集日志信息
                    StringBuilder logMessages = new StringBuilder();

                    // 首先查找并解析包装材料需求流转单
                    string 流转单路径 = null;
                    foreach (string filePath in excelFiles)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(filePath);
                        if (fileName.Contains("包装材料需求流转单") || fileName.Contains("流转单"))
                        {
                            流转单路径 = filePath;
                            break;
                        }
                    }

                    // 如果找到流转单，解析其中的纸箱规格信息
                    if (!string.IsNullOrEmpty(流转单路径))
                    {
                        logMessages.AppendLine($"找到流转单: {流转单路径}");

                        using (ExcelPackage package = new ExcelPackage(new FileInfo(流转单路径)))
                        {
                            var worksheet = package.Workbook.Worksheets[0]; // 假设流转单在第一个工作表

                            if (worksheet.Dimension != null)
                            {
                                // 直接使用固定列索引，根据您提供的截图
                                int 物料列 = 2;     // B列
                                int 规格列 = 4;     // D列
                                int 文件名列 = 11;   // K列

                                // 直接从第7行开始处理数据
                                int startRow = 7;
                                int endRow = worksheet.Dimension.End.Row;

                                // 遍历所有数据行
                                for (int dataRow = startRow; dataRow <= endRow; dataRow++)
                                {
                                    string 物料 = worksheet.Cells[dataRow, 物料列].Text.Trim();
                                    string 规格 = worksheet.Cells[dataRow, 规格列].Text.Trim();
                                    string 文件名 = worksheet.Cells[dataRow, 文件名列].Text.Trim();

                                    // 如果K列不为空，记录日志
                                    if (!string.IsNullOrEmpty(文件名))
                                    {
                                        logMessages.AppendLine($"行 {dataRow}: 物料='{物料}', 规格='{规格}', 文件名='{文件名}'");

                                        // 检查是否是纸箱（物料列包含"纸箱"关键词）
                                        if (!string.IsNullOrEmpty(物料) && 物料.Contains("纸箱"))
                                        {
                                            // 不管文件名是否已经包含.xlsx扩展名，都先去掉扩展名再存储
                                            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(文件名);

                                            // 将文件名（不带扩展名）作为键，规格作为值存储在字典中
                                            if (!string.IsNullOrEmpty(fileNameWithoutExt) && !string.IsNullOrEmpty(规格))
                                            {
                                                // 如果字典中还没有这个文件名，创建一个新的列表
                                                if (!纸箱规格字典.ContainsKey(fileNameWithoutExt))
                                                {
                                                    纸箱规格字典[fileNameWithoutExt] = new List<string>();
                                                }

                                                // 将规格添加到列表中
                                                纸箱规格字典[fileNameWithoutExt].Add(规格);
                                                logMessages.AppendLine($"关联文件 '{fileNameWithoutExt}' 与纸箱规格: '{规格}'");
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // 输出找到的纸箱规格信息
                        int totalSpecs = 纸箱规格字典.Values.Sum(list => list.Count);
                        logMessages.AppendLine($"共找到 {纸箱规格字典.Count} 个文件对应的 {totalSpecs} 个纸箱规格信息:");
                        foreach (var kvp in 纸箱规格字典)
                        {
                            logMessages.AppendLine($"文件: '{kvp.Key}', 规格数量: {kvp.Value.Count}");
                            for (int i = 0; i < kvp.Value.Count; i++)
                            {
                                logMessages.AppendLine($"  规格 {i + 1}: '{kvp.Value[i]}'");
                            }
                        }
                    }
                    else
                    {
                        logMessages.AppendLine("未找到包装材料需求流转单！");
                    }

                    // 显示日志信息
                    //MessageBox.Show(logMessages.ToString(), "处理日志", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 解析其他Excel文件
                    foreach (string filePath in excelFiles)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(filePath);

                        // 排除包装材料需求流转单
                        if (fileName.Contains("包装材料需求流转单") || fileName.Contains("流转单"))
                            continue;

                        // 解析文件名，按照 F10-4.889-附件 的格式
                        string[] parts = fileName.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

                        if (parts.Length >= 2)
                        {
                            string 型号 = parts[0].Trim();
                            string 销售数量 = parts[1].Trim();
                            string 备注 = parts.Length > 2 ? parts[2].Trim() : "";

                            // 查找对应的纸箱规格
                            List<string> 纸箱规格列表 = new List<string>();
                            if (纸箱规格字典.TryGetValue(fileName, out List<string> foundSpecs))
                            {
                                纸箱规格列表 = foundSpecs;
                                logMessages.AppendLine($"为文件 '{fileName}' 找到 {纸箱规格列表.Count} 个纸箱规格");
                            }
                            else
                            {
                                logMessages.AppendLine($"警告: 未找到文件 '{fileName}' 的纸箱规格");
                            }

                            // 解析Excel文件内容
                            结果数据 数据 = new 结果数据
                            {
                                产品型号 = 型号,
                                销售数量 = 销售数量,
                                备注 = 备注,
                                纸箱规格列表 = 纸箱规格列表 // 设置纸箱规格列表
                            };

                            // 读取Excel文件内容
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                // 遍历所有工作表，每个工作表代表一个盒子
                                foreach (var worksheet in package.Workbook.Worksheets)
                                {
                                    if (worksheet.Dimension == null)
                                        continue;

                                    // 创建一个新的盒子列表
                                    List<结果数据.盒子内容> 盒子内容 = new List<结果数据.盒子内容>();

                                    // 确定表头行
                                    int headerRow = 1;
                                    for (int row = 1; row <= Math.Min(10, worksheet.Dimension.End.Row); row++)
                                    {
                                        if (worksheet.Cells[row, 1].Text.Contains("序号") ||
                                            worksheet.Cells[row, 1].Text.Equals("序号", StringComparison.OrdinalIgnoreCase))
                                        {
                                            headerRow = row;
                                            break;
                                        }
                                    }

                                    // 确定列索引
                                    int 序号列 = -1, 条数列 = -1, 米数列 = -1, 标签码1列 = -1, 标签码2列 = -1, 标签码3列 = -1, 标签码4列 = -1, 线长列 = -1, 纸箱规格列 = -1, 包装编码列 = -1, 盒装标准列 = -1, 客户型号列 = -1, 标签显示长度列 = -1;

                                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                    {
                                        string headerText = worksheet.Cells[headerRow, col].Text.Trim();

                                        if (headerText.Contains("序号") || headerText.Equals("序号", StringComparison.OrdinalIgnoreCase))
                                            序号列 = col;
                                        else if (headerText.Contains("条数") || headerText.Equals("条数", StringComparison.OrdinalIgnoreCase))
                                            条数列 = col;
                                        else if (headerText.Contains("米数") || headerText.Equals("米数", StringComparison.OrdinalIgnoreCase))
                                            米数列 = col;
                                        else if (headerText.Contains("标签码1") || headerText.Equals("标签码1", StringComparison.OrdinalIgnoreCase))
                                            标签码1列 = col;
                                        else if (headerText.Contains("标签码2") || headerText.Equals("标签码2", StringComparison.OrdinalIgnoreCase))
                                            标签码2列 = col;
                                        else if (headerText.Contains("标签码3") || headerText.Equals("标签码3", StringComparison.OrdinalIgnoreCase))
                                            标签码3列 = col;
                                        else if (headerText.Contains("标签码4") || headerText.Equals("标签码4", StringComparison.OrdinalIgnoreCase))
                                            标签码4列 = col;
                                        else if (headerText.Contains("线长") || headerText.Equals("线长", StringComparison.OrdinalIgnoreCase))
                                            线长列 = col;
                                        else if (headerText.Contains("纸箱规格") || headerText.Equals("纸箱规格", StringComparison.OrdinalIgnoreCase))
                                            纸箱规格列 = col;
                                        else if (headerText.Contains("包装编码") || headerText.Equals("包装编码", StringComparison.OrdinalIgnoreCase))
                                            包装编码列 = col;
                                        else if (headerText.Contains("盒装标准") || headerText.Equals("盒装标准", StringComparison.OrdinalIgnoreCase))
                                            盒装标准列 = col;
                                        else if (headerText.Contains("客户型号") || headerText.Equals("客户型号", StringComparison.OrdinalIgnoreCase))
                                            客户型号列 = col;
                                        else if (headerText.Contains("标签显示长度") || headerText.Equals("标签显示长度", StringComparison.OrdinalIgnoreCase))
                                            标签显示长度列 = col;
                                    }

                                    // 读取数据行
                                    for (int row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
                                    {
                                        string 序号 = 序号列 > 0 ? worksheet.Cells[row, 序号列].Text.Trim() : "";

                                        // 如果序号为空，则跳过该行
                                        if (string.IsNullOrEmpty(序号))
                                            continue;

                                        盒子内容.Add(new 结果数据.盒子内容
                                        {
                                            序号 = worksheet.Cells[row, 序号列].Text,
                                            条数 = worksheet.Cells[row, 条数列].Text,
                                            米数 = worksheet.Cells[row, 米数列].Text,
                                            标签码1 = 标签码1列 > 0 ? worksheet.Cells[row, 标签码1列].Text : "",
                                            标签码2 = 标签码2列 > 0 ? worksheet.Cells[row, 标签码2列].Text : "",
                                            标签码3 = 标签码3列 > 0 ? worksheet.Cells[row, 标签码3列].Text : "",
                                            标签码4 = 标签码4列 > 0 ? worksheet.Cells[row, 标签码4列].Text : "",
                                            线长 = 线长列 > 0 ? worksheet.Cells[row, 线长列].Text : "",
                                            纸箱规格 = worksheet.Cells[row, 纸箱规格列].Text,
                                            包装编码 = worksheet.Cells[row, 包装编码列].Text,
                                            盒装标准 = 盒装标准列 > 0 ? (int.TryParse(worksheet.Cells[row, 盒装标准列].Text?.Trim(), out int 标准值) ? 标准值 : 1) : 1,
                                            客户型号 = 客户型号列 > 0 ? worksheet.Cells[row, 客户型号列].Text : "",
                                            标签显示长度 = 标签显示长度列 > 0 ? worksheet.Cells[row, 标签显示长度列].Text : ""
                                        });

                                        // 显示添加的内容
                                        //MessageBox.Show($"添加盒子内容:\n" +
                                        //    $"序号: {worksheet.Cells[row, 序号列].Text}\n" +
                                        //    $"条数: {worksheet.Cells[row, 条数列].Text}\n" +
                                        //    $"米数: {worksheet.Cells[row, 米数列].Text}\n" +
                                        //    $"标签码1: {(标签码1列 > 0 ? worksheet.Cells[row, 标签码1列].Text : "")}\n" +
                                        //    $"标签码2: {(标签码2列 > 0 ? worksheet.Cells[row, 标签码2列].Text : "")}\n" +
                                        //    $"标签码3: {(标签码3列 > 0 ? worksheet.Cells[row, 标签码3列].Text : "")}\n" +
                                        //    $"标签码4: {(标签码4列 > 0 ? worksheet.Cells[row, 标签码4列].Text : "")}\n" +
                                        //    $"线长: {(线长列 > 0 ? worksheet.Cells[row, 线长列].Text : "")}\n" +
                                        //    $"纸箱规格: {worksheet.Cells[row, 纸箱规格列].Text}\n" +
                                        //    $"包装编码: {worksheet.Cells[row, 包装编码列].Text}\n" +
                                        //    $"盒装标准: {(盒装标准列 > 0 ? worksheet.Cells[row, 盒装标准列].Text : "1")}",
                                        //    "调试信息");
                                    }

                                    // 如果盒子内容不为空，则添加到盒子列表
                                    if (盒子内容.Count > 0)
                                    {
                                        数据.盒子列表.Add(盒子内容);
                                    }
                                }
                            }

                            // 如果有盒子数据，则添加到结果列表
                            if (数据.盒子列表.Count > 0)
                            {
                                结果列表.Add(数据);
                            }
                        }
                    }

                    // 再次显示日志信息
                    //MessageBox.Show(logMessages.ToString(), "处理日志", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 显示解析结果
                    if (结果列表.Count > 0)
                    {
                        品名_解析结果(结果列表);
                    }
                    else
                    {
                        MessageBox.Show("未找到符合格式的Excel文件或文件内容不符合要求！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void 品名_解析结果(List<结果数据> 结果列表)
        {
            // 创建一个新窗体来显示解析结果
            Form resultForm = new Form();
            resultForm.Text = "品名_解析结果";
            resultForm.Size = new Size(1000, 700);

            // 创建一个分割面板
            SplitContainer splitContainer = new SplitContainer();
            splitContainer.Dock = DockStyle.Fill;
            splitContainer.Orientation = System.Windows.Forms.Orientation.Vertical;

            splitContainer.SplitterDistance = 40;

            // 创建上半部分的DataGridView，显示文件基本信息
            DataGridView fileGridView = new DataGridView();
            fileGridView.Dock = DockStyle.Fill;
            fileGridView.AutoGenerateColumns = false;
            fileGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            fileGridView.MultiSelect = false;

            // 添加列
            fileGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "产品型号",
                DataPropertyName = "产品型号",
                Width = 50
            });
            fileGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "销售数量",
                DataPropertyName = "销售数量",
                Width = 50
            });
            fileGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "备注",
                DataPropertyName = "备注",
                Width = 50
            });
            fileGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "盒数",
                DataPropertyName = "盒数",
                Width = 40
            });

            // 绑定数据
            fileGridView.DataSource = 结果列表.Select(x => new
            {
                产品型号 = x.产品型号,
                销售数量 = x.销售数量,
                备注 = x.备注,
                盒数 = x.盒子列表.Count
            }).ToList();

            // 创建下半部分的TabControl，显示盒子详情
            TabControl tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;

            // 创建右键菜单
            ContextMenuStrip tabContextMenu = new ContextMenuStrip();

            // 添加多个不同尺寸的打印选项
            ToolStripMenuItem exportItem = new ToolStripMenuItem("汇总");
            ToolStripMenuItem yulang1 = new ToolStripMenuItem("预览--80x150");
            ToolStripMenuItem yulang2 = new ToolStripMenuItem("预览--90x60");
            ToolStripMenuItem yulang3 = new ToolStripMenuItem("预览--90x45");
            ToolStripMenuItem lingcun1 = new ToolStripMenuItem("另存--80x150");
            ToolStripMenuItem lingcun2 = new ToolStripMenuItem("另存--90x60");
            ToolStripMenuItem lingcun3 = new ToolStripMenuItem("另存--90x45");
            ToolStripMenuItem viewBoxDetailsItem1 = new ToolStripMenuItem("打印本盒唛头-80x150");
            ToolStripMenuItem viewBoxDetailsItem2 = new ToolStripMenuItem("打印本盒唛头-90x60");
            ToolStripMenuItem viewBoxDetailsItem3 = new ToolStripMenuItem("打印本盒唛头-90x45");

            // 添加到菜单
            tabContextMenu.Items.Add(exportItem);
            //tabContextMenu.Items.Add(yulang1);
            //tabContextMenu.Items.Add(yulang2);
            //tabContextMenu.Items.Add(yulang3);
            //tabContextMenu.Items.Add(lingcun1);
            //tabContextMenu.Items.Add(lingcun2);
            //tabContextMenu.Items.Add(lingcun3);
            //tabContextMenu.Items.Add(viewBoxDetailsItem1);
            //tabContextMenu.Items.Add(viewBoxDetailsItem2);
            //tabContextMenu.Items.Add(viewBoxDetailsItem3);

            // 添加选择变更事件
            fileGridView.SelectionChanged += (s, e) =>
            {
                // 清空TabControl
                tabControl.TabPages.Clear();

                if (fileGridView.SelectedRows.Count > 0)
                {
                    int selectedIndex = fileGridView.SelectedRows[0].Index;
                    if (selectedIndex >= 0 && selectedIndex < 结果列表.Count)
                    {
                        var 选中数据 = 结果列表[selectedIndex];

                        // 为每个盒子创建一个Tab页
                        for (int i = 0; i < 选中数据.盒子列表.Count; i++)
                        {
                            var 盒子内容 = 选中数据.盒子列表[i];

                            // 创建一个新的Tab页
                            TabPage tabPage = new TabPage($"第{i + 1}盒");

                            // 为Tab页添加右键菜单
                            tabPage.ContextMenuStrip = tabContextMenu;

                            // 存储盒子索引和数据，使用自定义类而不是dynamic
                            tabPage.Tag = new TabPageData
                            {
                                Data = 选中数据,
                                BoxIndex = i
                            };

                            // 创建一个DataGridView来显示盒子内容
                            DataGridView boxGridView = new DataGridView();
                            boxGridView.Dock = DockStyle.Fill;
                            boxGridView.AllowUserToAddRows = false;
                            boxGridView.ReadOnly = true;

                            // 不使用自动调整列宽，而是设置固定宽度
                            boxGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                            // 添加列并设置宽度比例
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "序号",
                                Name = "序号",
                                Width = 80
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "条数",
                                Name = "条数",
                                Width = 30
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "米数",
                                Name = "米数",
                                Width = 50
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "线长",
                                Name = "线长",
                                Width = 100
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "纸箱规格",
                                Name = "纸箱规格",
                                Width = 100
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签码1",
                                Name = "标签码1",
                                Width = 100
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签码2",
                                Name = "标签码2",
                                Width = 100
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签码3",
                                Name = "标签码3",
                                Width = 100
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签码4",
                                Name = "标签码4",
                                Width = 100
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "客户型号",
                                Name = "客户型号",
                                Width = 100
                            });
                            boxGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签显示长度",
                                Name = "标签显示长度",
                                Width = 100
                            });

                            // 添加数据行
                            foreach (var 行数据 in 盒子内容)
                            {
                                int rowIndex = boxGridView.Rows.Add();
                                DataGridViewRow dataRow = boxGridView.Rows[rowIndex];

                                dataRow.Cells["序号"].Value = 行数据.序号;
                                dataRow.Cells["条数"].Value = 行数据.条数;
                                dataRow.Cells["米数"].Value = 行数据.米数;
                                dataRow.Cells["标签码1"].Value = 行数据.标签码1;
                                dataRow.Cells["标签码2"].Value = 行数据.标签码2;
                                dataRow.Cells["标签码3"].Value = 行数据.标签码3;
                                dataRow.Cells["标签码4"].Value = 行数据.标签码4;
                                dataRow.Cells["线长"].Value = 行数据.线长;
                                dataRow.Cells["纸箱规格"].Value = 行数据.纸箱规格;
                                dataRow.Cells["客户型号"].Value = 行数据.客户型号;
                                dataRow.Cells["标签显示长度"].Value = 行数据.标签显示长度;
                            }

                            tabPage.Controls.Add(boxGridView);
                            tabControl.TabPages.Add(tabPage);
                        }
                    }
                }
            };

            // 定义打印代码
            Action<TabPage, biaoqian> processTab = (selectedTab, operation) =>
            {
                if (selectedTab != null && selectedTab.Tag is TabPageData tag)
                {
                    var 选中数据 = tag.Data as 结果数据;
                    int 盒子索引 = tag.BoxIndex;

                    if (选中数据 != null && 盒子索引 >= 0 && 盒子索引 < 选中数据.盒子列表.Count)
                    {
                        var 盒子内容 = 选中数据.盒子列表[盒子索引];

                        // 创建一个新的窗体来显示盒子详情统计
                        Form detailsForm = new Form();
                        detailsForm.Text = $"{选中数据.产品型号} - 第{盒子索引 + 1}盒详情统计";
                        detailsForm.Size = new Size(600, 400);
                        detailsForm.StartPosition = FormStartPosition.CenterParent;

                        // 创建一个ListView来显示统计信息
                        ListView listView = new ListView();
                        listView.Dock = DockStyle.Fill;
                        listView.View = View.Details;
                        listView.FullRowSelect = true;
                        listView.GridLines = true;

                        // 添加列
                        listView.Columns.Add("统计项", 150);
                        listView.Columns.Add("值", 400);

                        // 计算总条数和总米数
                        int 总条数 = 0;
                        double 总米数 = 0;
                        Dictionary<string, int> 条数统计 = new Dictionary<string, int>();

                        foreach (var 行数据 in 盒子内容)
                        {
                            // 解析条数
                            if (int.TryParse(行数据.条数, out int 条数))
                            {
                                总条数 += 条数;

                                // 按米数分组统计条数
                                string 米数Key = 行数据.米数;
                                if (!条数统计.ContainsKey(米数Key))
                                {
                                    条数统计[米数Key] = 0;
                                }
                                条数统计[米数Key] += 条数;
                            }

                            // 解析米数
                            if (double.TryParse(行数据.米数, out double 米数))
                            {
                                总米数 += 米数 * (int.TryParse(行数据.条数, out int 条数值) ? 条数值 : 1);
                            }
                        }

                        // 添加统计信息
                        ListViewItem totalCountItem = new ListViewItem("总条数");
                        totalCountItem.SubItems.Add(总条数.ToString());
                        listView.Items.Add(totalCountItem);

                        ListViewItem totalLengthItem = new ListViewItem("总米数");
                        totalLengthItem.SubItems.Add(总米数.ToString("F3"));
                        listView.Items.Add(totalLengthItem);

                        // 创建一个新的格式化字符串来显示按米数分组的条数统计
                        StringBuilder formattedStats = new StringBuilder();
                        bool isFirst = true;

                        // 按米数排序并格式化为 "米数m*条数PC(S)" 的形式
                        foreach (var 统计 in 条数统计.OrderBy(k => double.Parse(k.Key)))
                        {
                            if (!isFirst)
                            {
                                formattedStats.Append(", ");
                            }

                            // 根据条数决定使用PC还是PCS
                            string pcUnit = 统计.Value == 1 ? "PC" : "PCS";

                            // 使用小写m
                            formattedStats.Append($"{统计.Key}m*{统计.Value}{pcUnit}");
                            isFirst = false;
                        }

                        textBox_唛头数量.Text = formattedStats.ToString();
                        // 添加格式化后的条数统计
                        ListViewItem formattedItem = new ListViewItem("规格明细");
                        formattedItem.SubItems.Add(formattedStats.ToString());
                        listView.Items.Add(formattedItem);

                        // 从盒子内容中获取纸箱规格
                        // 假设当前处理的是第一个盒子的第一行数据
                        if (盒子内容 != null && 盒子内容.Count > 0 && 盒子内容[0] != null)
                        {
                            var 第一行 = 盒子内容[0];

                            if (!string.IsNullOrEmpty(第一行.纸箱规格))
                            {
                                // 添加到ListView
                                ListViewItem specItem = new ListViewItem("纸箱规格");
                                specItem.SubItems.Add(第一行.纸箱规格);
                                listView.Items.Add(specItem);

                                // 将纸箱规格写入textBox_唛头尺寸，并在后面添加CM
                                textBox_唛头尺寸.Text = 第一行.纸箱规格 + " CM";
                            }
                            else
                            {
                                ListViewItem noSpecItem = new ListViewItem("纸箱规格");
                                noSpecItem.SubItems.Add("未指定");
                                listView.Items.Add(noSpecItem);

                                // 如果没有纸箱规格，可以清空文本框或设置默认值
                                textBox_唛头尺寸.Text = "";
                            }
                        }
                        else
                        {
                            ListViewItem noSpecItem = new ListViewItem("纸箱规格");
                            noSpecItem.SubItems.Add("未指定");
                            listView.Items.Add(noSpecItem);

                            // 如果没有盒子内容，可以清空文本框或设置默认值
                            textBox_唛头尺寸.Text = "";
                        }

                        detailsForm.Controls.Add(listView);
                        //detailsForm.ShowDialog();

                        // 处理标签码
                        if (盒子内容 != null && 盒子内容.Count > 0 && 盒子内容[0] != null)
                        {
                            var 第一行 = 盒子内容[0];

                            // 处理标签码1
                            if (!string.IsNullOrEmpty(第一行.标签码1))
                            {
                                textBox_标识码01.Text = 第一行.标签码1;
                                checkBox_标识码01.Checked = true;
                            }
                            else
                            {
                                textBox_标识码01.Text = "";
                                checkBox_标识码01.Checked = false;
                            }

                            // 处理标签码2
                            if (!string.IsNullOrEmpty(第一行.标签码2))
                            {
                                textBox_标识码02.Text = 第一行.标签码2;
                                checkBox_标识码02.Checked = true;
                            }
                            else
                            {
                                textBox_标识码02.Text = "";
                                checkBox_标识码02.Checked = false;
                            }
                        }

                        // 最后根据操作类型执行相应操作
                        switch (operation)
                        {
                            case biaoqian.dayin:
                                print_btn_Click(null, EventArgs.Empty);
                                break;

                            case biaoqian.yulan:
                                shengcheng_maitou(biaoqian.yulan);
                                break;

                            case biaoqian.lingcun:
                                shengcheng_maitou(biaoqian.lingcun);
                                break;
                        }
                    }
                }
            };

            // 处理右键菜单点击事件
            // 处理导出按钮点击事件
            exportItem.Click += (s, evt) =>
            {
                // 调用导出方法
                品名导出(结果列表);
            };

            yulang1.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "80x150";
                processTab(tabControl.SelectedTab, biaoqian.yulan);
            };
            yulang2.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x60";
                processTab(tabControl.SelectedTab, biaoqian.yulan);
            };
            yulang3.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x45";
                processTab(tabControl.SelectedTab, biaoqian.yulan);
            };

            // 处理另存按钮点击事件
            lingcun1.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "80x150";
                processTab(tabControl.SelectedTab, biaoqian.lingcun);
            };
            lingcun2.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x60";
                processTab(tabControl.SelectedTab, biaoqian.lingcun);
            };
            lingcun3.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x45";
                processTab(tabControl.SelectedTab, biaoqian.lingcun);
            };

            // 处理打印按钮点击事件
            viewBoxDetailsItem1.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "80x150";
                processTab(tabControl.SelectedTab, biaoqian.dayin);
            };
            viewBoxDetailsItem2.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x60";
                processTab(tabControl.SelectedTab, biaoqian.dayin);
            };
            viewBoxDetailsItem3.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x45";
                processTab(tabControl.SelectedTab, biaoqian.dayin);
            };

            // 如果有数据，默认选中第一行
            if (fileGridView.Rows.Count > 0)
            {
                fileGridView.Rows[0].Selected = true;
            }

            // 添加控件到分割面板
            splitContainer.Panel1.Controls.Add(fileGridView);
            splitContainer.Panel2.Controls.Add(tabControl);

            // 添加分割面板到窗体
            resultForm.Controls.Add(splitContainer);

            // 显示窗体
            resultForm.ShowDialog();
        }

        private void 品名导出(List<结果数据> 结果列表)
        {
            // 使用SaveFileDialog让用户选择保存位置
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel文件|*.xlsx";
                saveDialog.Title = "保存汇总Excel";
                saveDialog.FileName = "品名汇总表.xlsx";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    // 创建Excel包
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        // 添加工作表
                        var worksheet = package.Workbook.Worksheets.Add("品名汇总");

                        // 设置表头
                        worksheet.Cells[1, 1].Value = "序号";
                        worksheet.Cells[1, 2].Value = "标签码1";
                        worksheet.Cells[1, 3].Value = "标签码2";
                        worksheet.Cells[1, 4].Value = "标签码3";
                        worksheet.Cells[1, 5].Value = "标签码4";
                        worksheet.Cells[1, 6].Value = "标签码5";
                        worksheet.Cells[1, 7].Value = "条数";
                        worksheet.Cells[1, 8].Value = "长度";
                        worksheet.Cells[1, 9].Value = "客户型号";
                        worksheet.Cells[1, 10].Value = "PO号";
                        worksheet.Cells[1, 11].Value = "条形码";
                        worksheet.Cells[1, 12].Value = "线长";

                        // 设置表头样式
                        using (var range = worksheet.Cells[1, 1, 1, 12])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        // 填充数据
                        int rowIndex = 2;
                        int seqNo = 1;

                        foreach (var 数据 in 结果列表)
                        {
                            // 获取客户型号和PO号
                            string 客户型号 = 数据.产品型号 ?? "";
                            string PO号 = textBox_订单编号.Text ?? "";

                            foreach (var 盒子列表 in 数据.盒子列表)
                            {
                                // 改用List来保持顺序
                                List<(string 米数, int 条数)> 米数统计列表 = new List<(string 米数, int 条数)>();

                                // 计算盒子内总条数
                                int 总条数 = 0;
                                foreach (var 内容 in 盒子列表)
                                {
                                    // 统计米数
                                    if (!string.IsNullOrEmpty(内容.米数))
                                    {
                                        string 米数 = 内容.米数;
                                        int 条数 = 1;
                                        int.TryParse(内容.条数, out 条数);
                                        总条数 += 条数;

                                        // 查找是否已存在该米数
                                        var 现有项 = 米数统计列表.FirstOrDefault(x => x.米数 == 米数);
                                        if (现有项 != default)
                                        {
                                            // 如果存在，更新条数
                                            int 索引 = 米数统计列表.IndexOf(现有项);
                                            米数统计列表[索引] = (米数, 现有项.条数 + 条数);
                                        }
                                        else
                                        {
                                            // 如果不存在，添加新项
                                            米数统计列表.Add((米数, 条数));
                                        }
                                    }
                                }

                                // 格式化长度信息
                                StringBuilder 长度信息 = new StringBuilder();
                                bool isFirst = true;

                                // 判断盒子内是否只有一条灯带
                                bool 只有一条 = 总条数 == 1 && 米数统计列表.Count == 1;

                                string 单位文件路径 = Path.Combine(folderPath, "订单资料", "单位.txt");
                                string 单位内容 = "m";
                                if (File.Exists(单位文件路径))
                                {
                                    单位内容 = File.ReadAllText(单位文件路径).Trim();
                                }

                                foreach (var (米数, 条数) in 米数统计列表)
                                {
                                    if (!isFirst)
                                    {
                                        长度信息.Append(",");
                                    }

                                    // 找到所有该米数的内容
                                    var 对应内容 = 盒子列表.Where(x =>
                                        double.TryParse(x.米数, out double mval) && Math.Abs(mval - double.Parse(米数)) < 0.0001
                                    ).ToList();

                                    string 标签显示长度 = 对应内容.FirstOrDefault()?.标签显示长度 ?? "";
                                    string unit = 条数 == 1 ? "PC" : "PCS";

                                    //MessageBox.Show($"单位内容: {单位内容}\n米数: {米数}\n标签显示长度: {标签显示长度}", "拼接前调试");

                                    switch (单位内容)
                                    {
                                        case "m":
                                            if (只有一条)
                                                长度信息.Append($"{米数}m");
                                            else
                                                长度信息.Append($"{米数}m*{条数}{unit}");
                                            break;

                                        case "mm":
                                            // 把米数转成mm
                                            if (double.TryParse(米数, out double mVal))
                                            {
                                                int mmVal = (int)Math.Round(mVal * 1000); // 四舍五入取整
                                                if (只有一条)
                                                    长度信息.Append($"{mmVal}mm");
                                                else
                                                    长度信息.Append($"{mmVal}mm*{条数}{unit}");
                                            }
                                            else
                                            {
                                                // 解析失败，原样输出
                                                if (只有一条)
                                                    长度信息.Append($"{米数}mm");
                                                else
                                                    长度信息.Append($"{米数}mm*{条数}{unit}");
                                            }
                                            break;

                                        case "Ft":
                                        case "IN":
                                            if (只有一条)
                                                长度信息.Append($"{标签显示长度}{单位内容}");
                                            else
                                                长度信息.Append($"{标签显示长度}{单位内容}*{条数}{unit}");
                                            break;

                                        case "m(Ft)":
                                        case "m(IN)":
                                            string 括号单位 = 单位内容 == "m(Ft)" ? "Ft" : "IN";
                                            if (只有一条)
                                                长度信息.Append($"{米数}m({标签显示长度}{括号单位})");
                                            else
                                                长度信息.Append($"{米数}m({标签显示长度}{括号单位})*{条数}{unit}");
                                            break;

                                        default:
                                            if (只有一条)
                                                长度信息.Append($"{米数}m");
                                            else
                                                长度信息.Append($"{米数}m*{条数}{unit}");
                                            break;
                                    }

                                    isFirst = false;
                                }

                                // 如果这个盒子有内容，则添加一行
                                if (盒子列表.Count > 0)
                                {
                                    var 第一行内容 = 盒子列表[0]; // 使用第一行的数据作为基础

                                    // 序号
                                    worksheet.Cells[rowIndex, 1].Value = seqNo++;

                                    // 合并标签码1~4
                                    string 标签码1 = string.Join(",", 盒子列表.Select(x => x.标签码1).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());
                                    string 标签码2 = string.Join(",", 盒子列表.Select(x => x.标签码2).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());
                                    string 标签码3 = string.Join(",", 盒子列表.Select(x => x.标签码3).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());
                                    string 标签码4 = string.Join(",", 盒子列表.Select(x => x.标签码4).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());
                                    // 标签码1-5
                                    worksheet.Cells[rowIndex, 2].Value = 标签码1;
                                    worksheet.Cells[rowIndex, 3].Value = 标签码2;
                                    worksheet.Cells[rowIndex, 4].Value = 标签码3;
                                    worksheet.Cells[rowIndex, 5].Value = 标签码4;
                                    worksheet.Cells[rowIndex, 6].Value = ""; // 标签码5

                                    // 条数（整个盒子的总条数）
                                    worksheet.Cells[rowIndex, 7].Value = 1;

                                    // 长度（汇总格式）
                                    worksheet.Cells[rowIndex, 8].Value = 长度信息.ToString();

                                    // 客户型号汇总
                                    string 客户型号汇总 = string.Join(",", 盒子列表.Select(x => x.客户型号).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());
                                    worksheet.Cells[rowIndex, 9].Value = 客户型号汇总;

                                    // PO号
                                    //worksheet.Cells[rowIndex, 10].Value = PO号;
                                    worksheet.Cells[rowIndex, 10].Value = "";

                                    // 条形码（可以留空或填入特定逻辑生成的条形码）
                                    worksheet.Cells[rowIndex, 11].Value = "";

                                    // 线长
                                    worksheet.Cells[rowIndex, 12].Value = 第一行内容.线长 ?? "";

                                    rowIndex++;
                                }
                            }
                        }

                        // 自动调整列宽
                        worksheet.Cells.AutoFitColumns();

                        // 保存文件
                        try
                        {
                            package.SaveAs(new FileInfo(saveDialog.FileName));
                            MessageBox.Show("汇总表格已成功保存!", "成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"保存文件时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
            }
        }

        private void 打印唛头_Click(object sender, EventArgs e)
        {
            textBox1.Text = "1";
            string cpxx_text = cpxxBox.Text;
            判断产品信息(cpxx_text);

            ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

            // 使用更现代的文件夹选择对话框
            using (var folderDialog = new System.Windows.Forms.OpenFileDialog())
            {
                folderDialog.ValidateNames = false;
                folderDialog.CheckFileExists = false;
                folderDialog.CheckPathExists = true;
                folderDialog.FileName = "选择文件夹";

                if (folderDialog.ShowDialog() == DialogResult.OK)
                {
                    folderPath = Path.GetDirectoryName(folderDialog.FileName);
                    // 获取最后一个文件夹名称
                    lastFolderName = Path.GetFileName(folderPath);
                    //MessageBox.Show(lastFolderName);

                    唛头_寻找订单编号(lastFolderName);
                    EXCEL_包装规格回调(folderPath);

                    // 获取文件夹中的所有Excel文件
                    string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);

                    // 创建一个列表来存储解析后的文件信息
                    List<结果数据> 结果列表 = new List<结果数据>();

                    // 创建一个字典来存储纸箱规格信息 - 键是文件名，值是规格列表
                    Dictionary<string, List<string>> 纸箱规格字典 = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

                    // 创建一个StringBuilder来收集日志信息
                    StringBuilder logMessages = new StringBuilder();

                    // 首先查找并解析包装材料需求流转单
                    string 流转单路径 = null;
                    foreach (string filePath in excelFiles)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(filePath);
                        if (fileName.Contains("包装材料需求流转单") || fileName.Contains("流转单"))
                        {
                            流转单路径 = filePath;
                            break;
                        }
                    }

                    // 如果找到流转单，解析其中的纸箱规格信息
                    if (!string.IsNullOrEmpty(流转单路径))
                    {
                        logMessages.AppendLine($"找到流转单: {流转单路径}");

                        using (ExcelPackage package = new ExcelPackage(new FileInfo(流转单路径)))
                        {
                            var worksheet = package.Workbook.Worksheets[0]; // 假设流转单在第一个工作表

                            if (worksheet.Dimension != null)
                            {
                                // 直接使用固定列索引，根据您提供的截图
                                int 物料列 = 2;     // B列
                                int 规格列 = 4;     // D列
                                int 备注列 = 9;     // I列 - 添加备注列
                                int 文件名列 = 11;   // K列

                                // 直接从第7行开始处理数据
                                int startRow = 7;
                                int endRow = worksheet.Dimension.End.Row;

                                // 遍历所有数据行
                                for (int dataRow = startRow; dataRow <= endRow; dataRow++)
                                {
                                    string 物料 = worksheet.Cells[dataRow, 物料列].Text.Trim();
                                    string 规格 = worksheet.Cells[dataRow, 规格列].Text.Trim();
                                    string 备注 = worksheet.Cells[dataRow, 备注列].Text.Trim();  // 获取备注列内容
                                    string 文件名 = worksheet.Cells[dataRow, 文件名列].Text.Trim();

                                    // 解析盒装标准
                                    int 盒装标准 = 1; // 默认为1盒装
                                    if (!string.IsNullOrEmpty(备注))
                                    {
                                        var match = System.Text.RegularExpressions.Regex.Match(备注, @"(\d+)\s*盒装标准");
                                        if (match.Success && match.Groups.Count > 1)
                                        {
                                            if (int.TryParse(match.Groups[1].Value, out int 解析结果))
                                            {
                                                盒装标准 = 解析结果;
                                            }
                                        }
                                    }

                                    // 如果K列不为空，记录日志
                                    if (!string.IsNullOrEmpty(文件名))
                                    {
                                        logMessages.AppendLine($"行 {dataRow}: 物料='{物料}', 规格='{规格}', 文件名='{文件名}'");

                                        // 检查是否是纸箱（物料列包含"纸箱"关键词）
                                        if (!string.IsNullOrEmpty(物料) && 物料.Contains("纸箱"))
                                        {
                                            // 不管文件名是否已经包含.xlsx扩展名，都先去掉扩展名再存储
                                            string fileNameWithoutExt = Path.GetFileNameWithoutExtension(文件名);

                                            // 将文件名（不带扩展名）作为键，规格作为值存储在字典中
                                            if (!string.IsNullOrEmpty(fileNameWithoutExt) && !string.IsNullOrEmpty(规格))
                                            {
                                                // 如果字典中还没有这个文件名，创建一个新的列表
                                                if (!纸箱规格字典.ContainsKey(fileNameWithoutExt))
                                                {
                                                    纸箱规格字典[fileNameWithoutExt] = new List<string>();
                                                }

                                                // 将规格和盒装标准一起存储
                                                纸箱规格字典[fileNameWithoutExt].Add($"{规格}_{盒装标准}");  // 修改存储格式，包含盒装标准
                                                logMessages.AppendLine($"关联文件 '{fileNameWithoutExt}' 与纸箱规格: '{规格}', 盒装标准: {盒装标准}");
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        // 输出找到的纸箱规格信息
                        int totalSpecs = 纸箱规格字典.Values.Sum(list => list.Count);
                        logMessages.AppendLine($"共找到 {纸箱规格字典.Count} 个文件对应的 {totalSpecs} 个纸箱规格信息:");
                        foreach (var kvp in 纸箱规格字典)
                        {
                            logMessages.AppendLine($"文件: '{kvp.Key}', 规格数量: {kvp.Value.Count}");
                            for (int i = 0; i < kvp.Value.Count; i++)
                            {
                                logMessages.AppendLine($"  规格 {i + 1}: '{kvp.Value[i]}'");
                            }
                        }
                    }
                    else
                    {
                        logMessages.AppendLine("未找到包装材料需求流转单！");
                    }

                    // 显示日志信息
                    //MessageBox.Show(logMessages.ToString(), "处理日志", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 解析其他Excel文件
                    foreach (string filePath in excelFiles)
                    {
                        string fileName = Path.GetFileNameWithoutExtension(filePath);

                        // 排除包装材料需求流转单
                        if (fileName.Contains("包装材料需求流转单") || fileName.Contains("流转单"))
                            continue;

                        // 解析文件名，按照 F10-4.889-附件 的格式
                        string[] parts = fileName.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

                        if (parts.Length >= 2)
                        {
                            string 型号 = parts[0].Trim();
                            string 销售数量 = parts[1].Trim();
                            string 备注 = parts.Length > 2 ? parts[2].Trim() : "";

                            // 查找对应的纸箱规格
                            List<string> 纸箱规格列表 = new List<string>();
                            if (纸箱规格字典.TryGetValue(fileName, out List<string> foundSpecs))
                            {
                                纸箱规格列表 = foundSpecs;
                                logMessages.AppendLine($"为文件 '{fileName}' 找到 {纸箱规格列表.Count} 个纸箱规格");
                            }
                            else
                            {
                                logMessages.AppendLine($"警告: 未找到文件 '{fileName}' 的纸箱规格");
                            }

                            // 解析Excel文件内容
                            结果数据 数据 = new 结果数据
                            {
                                产品型号 = 型号,
                                销售数量 = 销售数量,
                                备注 = 备注,
                                纸箱规格列表 = 纸箱规格列表 // 设置纸箱规格列表
                            };

                            // 读取Excel文件内容
                            using (ExcelPackage package = new ExcelPackage(new FileInfo(filePath)))
                            {
                                // 遍历所有工作表，每个工作表代表一个盒子
                                foreach (var worksheet in package.Workbook.Worksheets)
                                {
                                    if (worksheet.Dimension == null)
                                        continue;

                                    // 创建一个新的盒子列表
                                    List<结果数据.盒子内容> 盒子内容 = new List<结果数据.盒子内容>();

                                    // 确定表头行
                                    int headerRow = 1;
                                    for (int row = 1; row <= Math.Min(10, worksheet.Dimension.End.Row); row++)
                                    {
                                        if (worksheet.Cells[row, 1].Text.Contains("序号") ||
                                            worksheet.Cells[row, 1].Text.Equals("序号", StringComparison.OrdinalIgnoreCase))
                                        {
                                            headerRow = row;
                                            break;
                                        }
                                    }

                                    // 确定列索引
                                    int 序号列 = -1, 条数列 = -1, 米数列 = -1, 标签码1列 = -1, 标签码2列 = -1, 标签码3列 = -1, 标签码4列 = -1, 线长列 = -1, 纸箱规格列 = -1, 包装编码列 = -1, 盒装标准列 = -1, 客户型号列 = -1, 标签显示长度列 = -1;

                                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                                    {
                                        string headerText = worksheet.Cells[headerRow, col].Text.Trim();

                                        if (headerText.Contains("序号") || headerText.Equals("序号", StringComparison.OrdinalIgnoreCase))
                                            序号列 = col;
                                        else if (headerText.Contains("条数") || headerText.Equals("条数", StringComparison.OrdinalIgnoreCase))
                                            条数列 = col;
                                        else if (headerText.Contains("米数") || headerText.Equals("米数", StringComparison.OrdinalIgnoreCase))
                                            米数列 = col;
                                        else if (headerText.Contains("标签码1") || headerText.Equals("标签码1", StringComparison.OrdinalIgnoreCase))
                                            标签码1列 = col;
                                        else if (headerText.Contains("标签码2") || headerText.Equals("标签码2", StringComparison.OrdinalIgnoreCase))
                                            标签码2列 = col;
                                        else if (headerText.Contains("标签码3") || headerText.Equals("标签码3", StringComparison.OrdinalIgnoreCase))
                                            标签码3列 = col;
                                        else if (headerText.Contains("标签码4") || headerText.Equals("标签码4", StringComparison.OrdinalIgnoreCase))
                                            标签码4列 = col;
                                        else if (headerText.Contains("线长") || headerText.Equals("线长", StringComparison.OrdinalIgnoreCase))
                                            线长列 = col;
                                        else if (headerText.Contains("纸箱规格") || headerText.Equals("纸箱规格", StringComparison.OrdinalIgnoreCase))
                                            纸箱规格列 = col;
                                        else if (headerText.Contains("包装编码") || headerText.Equals("包装编码", StringComparison.OrdinalIgnoreCase))
                                            包装编码列 = col;
                                        else if (headerText.Contains("盒装标准") || headerText.Equals("盒装标准", StringComparison.OrdinalIgnoreCase))
                                            盒装标准列 = col;
                                        else if (headerText.Contains("客户型号") || headerText.Equals("客户型号", StringComparison.OrdinalIgnoreCase))
                                            客户型号列 = col;
                                        else if (headerText.Contains("标签显示长度") || headerText.Equals("标签显示长度", StringComparison.OrdinalIgnoreCase))
                                            标签显示长度列 = col;
                                    }

                                    // 读取数据行
                                    for (int row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
                                    {
                                        string 序号 = 序号列 > 0 ? worksheet.Cells[row, 序号列].Text.Trim() : "";

                                        // 如果序号为空，则跳过该行
                                        if (string.IsNullOrEmpty(序号))
                                            continue;

                                        盒子内容.Add(new 结果数据.盒子内容
                                        {
                                            序号 = worksheet.Cells[row, 序号列].Text,
                                            条数 = worksheet.Cells[row, 条数列].Text,
                                            米数 = worksheet.Cells[row, 米数列].Text,
                                            标签码1 = 标签码1列 > 0 ? worksheet.Cells[row, 标签码1列].Text : "",
                                            标签码2 = 标签码2列 > 0 ? worksheet.Cells[row, 标签码2列].Text : "",
                                            标签码3 = 标签码3列 > 0 ? worksheet.Cells[row, 标签码3列].Text : "",
                                            标签码4 = 标签码4列 > 0 ? worksheet.Cells[row, 标签码4列].Text : "",
                                            线长 = 线长列 > 0 ? worksheet.Cells[row, 线长列].Text : "",
                                            纸箱规格 = worksheet.Cells[row, 纸箱规格列].Text,
                                            包装编码 = worksheet.Cells[row, 包装编码列].Text,
                                            盒装标准 = 盒装标准列 > 0 ? (int.TryParse(worksheet.Cells[row, 盒装标准列].Text?.Trim(), out int 标准值) ? 标准值 : 1) : 1,
                                            客户型号 = 客户型号列 > 0 ? worksheet.Cells[row, 客户型号列].Text : "",
                                            标签显示长度 = worksheet.Cells[row, 标签显示长度列].Text,
                                        });

                                        //// 显示添加的内容
                                        //MessageBox.Show($"添加盒子内容:\n" +
                                        //    $"序号: {worksheet.Cells[row, 序号列].Text}\n" +
                                        //    $"条数: {worksheet.Cells[row, 条数列].Text}\n" +
                                        //    $"米数: {worksheet.Cells[row, 米数列].Text}\n" +
                                        //    $"标签码1: {(标签码1列 > 0 ? worksheet.Cells[row, 标签码1列].Text : "")}\n" +
                                        //    $"标签码2: {(标签码2列 > 0 ? worksheet.Cells[row, 标签码2列].Text : "")}\n" +
                                        //    $"标签码3: {(标签码3列 > 0 ? worksheet.Cells[row, 标签码3列].Text : "")}\n" +
                                        //    $"标签码4: {(标签码4列 > 0 ? worksheet.Cells[row, 标签码4列].Text : "")}\n" +
                                        //    $"线长: {(线长列 > 0 ? worksheet.Cells[row, 线长列].Text : "")}\n" +
                                        //    $"纸箱规格: {worksheet.Cells[row, 纸箱规格列].Text}\n" +
                                        //    $"包装编码: {worksheet.Cells[row, 包装编码列].Text}\n" +
                                        //    $"盒装标准: {(盒装标准列 > 0 ? worksheet.Cells[row, 盒装标准列].Text : "1")}",
                                        //    "调试信息");
                                    }

                                    // 如果盒子内容不为空，则添加到盒子列表
                                    if (盒子内容.Count > 0)
                                    {
                                        数据.盒子列表.Add(盒子内容);
                                    }
                                }
                            }

                            // 如果有盒子数据，则添加到结果列表
                            if (数据.盒子列表.Count > 0)
                            {
                                结果列表.Add(数据);
                            }
                        }
                    }

                    // 再次显示日志信息
                    //MessageBox.Show(logMessages.ToString(), "处理日志", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    // 显示解析结果
                    if (结果列表.Count > 0)
                    {
                        唛头_解析结果(结果列表);
                    }
                    else
                    {
                        MessageBox.Show("未找到符合格式的Excel文件或文件内容不符合要求！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void 唛头_解析结果(List<结果数据> 结果列表)
        {
            // 创建一个新窗体来显示解析结果
            Form resultForm = new Form();
            resultForm.Text = "唛头_解析结果";
            resultForm.Size = new Size(1000, 700);

            // 创建一个分割面板
            SplitContainer splitContainer = new SplitContainer();
            splitContainer.Dock = DockStyle.Fill;
            splitContainer.Orientation = System.Windows.Forms.Orientation.Vertical;

            splitContainer.SplitterDistance = 40;

            // 创建上半部分的DataGridView，显示文件基本信息
            DataGridView fileGridView = new DataGridView();
            fileGridView.Dock = DockStyle.Fill;
            fileGridView.AutoGenerateColumns = false;
            fileGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            fileGridView.MultiSelect = false;

            // 添加列
            fileGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "产品型号",
                DataPropertyName = "产品型号",
                Width = 50
            });
            fileGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "销售数量",
                DataPropertyName = "销售数量",
                Width = 50
            });
            fileGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "备注",
                DataPropertyName = "备注",
                Width = 50
            });
            fileGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = "盒数",
                DataPropertyName = "盒数",
                Width = 40
            });

            // 绑定数据
            fileGridView.DataSource = 结果列表.Select(x => new
            {
                产品型号 = x.产品型号,
                销售数量 = x.销售数量,
                备注 = x.备注,
                盒数 = x.盒子列表.Count
            }).ToList();

            // 创建下半部分的TabControl，显示盒子详情
            TabControl tabControl = new TabControl();
            tabControl.Dock = DockStyle.Fill;

            // 创建右键菜单
            ContextMenuStrip tabContextMenu = new ContextMenuStrip();

            // 添加多个不同尺寸的打印选项
            ToolStripMenuItem exportItem = new ToolStripMenuItem("汇总");
            ToolStripMenuItem yulang1 = new ToolStripMenuItem("预览--80x150");
            ToolStripMenuItem yulang2 = new ToolStripMenuItem("预览--90x60");
            ToolStripMenuItem yulang3 = new ToolStripMenuItem("预览--90x45");
            ToolStripMenuItem lingcun1 = new ToolStripMenuItem("另存--80x150");
            ToolStripMenuItem lingcun2 = new ToolStripMenuItem("另存--90x60");
            ToolStripMenuItem lingcun3 = new ToolStripMenuItem("另存--90x45");
            ToolStripMenuItem viewBoxDetailsItem1 = new ToolStripMenuItem("打印本盒唛头-80x150");
            ToolStripMenuItem viewBoxDetailsItem2 = new ToolStripMenuItem("打印本盒唛头-90x60");
            ToolStripMenuItem viewBoxDetailsItem3 = new ToolStripMenuItem("打印本盒唛头-90x45");

            // 添加到菜单
            tabContextMenu.Items.Add(exportItem);
            tabContextMenu.Items.Add(yulang1);
            tabContextMenu.Items.Add(yulang2);
            tabContextMenu.Items.Add(yulang3);
            tabContextMenu.Items.Add(lingcun1);
            tabContextMenu.Items.Add(lingcun2);
            tabContextMenu.Items.Add(lingcun3);
            tabContextMenu.Items.Add(viewBoxDetailsItem1);
            tabContextMenu.Items.Add(viewBoxDetailsItem2);
            tabContextMenu.Items.Add(viewBoxDetailsItem3);

            // 添加选择变更事件
            // 修改选择变更事件中的代码
            // 修改分组逻辑部分的代码
            fileGridView.SelectionChanged += (s, e) =>
            {
                tabControl.TabPages.Clear();

                if (fileGridView.SelectedRows.Count > 0)
                {
                    int selectedIndex = fileGridView.SelectedRows[0].Index;
                    if (selectedIndex >= 0 && selectedIndex < 结果列表.Count)
                    {
                        var 选中数据 = 结果列表[selectedIndex];

                        // 创建一个字典来存储分组后的数据
                        var 分组数据 = new Dictionary<string, List<(List<结果数据.盒子内容> 内容, int 索引)>>();

                        // 先收集相同规格的盒子
                        Dictionary<string, List<(List<结果数据.盒子内容> 内容, int 索引)>> 临时分组 = new Dictionary<string, List<(List<结果数据.盒子内容>, int)>>();

                        // 按规格先收集所有盒子
                        for (int i = 0; i < 选中数据.盒子列表.Count; i++)
                        {
                            var 盒子内容 = 选中数据.盒子列表[i];
                            if (盒子内容.Count > 0)
                            {
                                string 纸箱规格 = 盒子内容[0].纸箱规格 ?? "";
                                int 盒装标准 = 盒子内容[0].盒装标准;

                                string 临时键 = 纸箱规格;

                                if (!临时分组.ContainsKey(临时键))
                                {
                                    临时分组[临时键] = new List<(List<结果数据.盒子内容>, int)>();
                                }

                                临时分组[临时键].Add((盒子内容, i));
                            }
                        }

                        // 然后按盒装标准将盒子分配到箱子
                        foreach (var 规格组 in 临时分组)
                        {
                            string 纸箱规格 = 规格组.Key;
                            var 该规格盒子列表 = 规格组.Value;

                            // 查找该规格的盒装标准（取第一个盒子的标准）
                            int 盒装标准 = 1;
                            if (该规格盒子列表.Count > 0 && 该规格盒子列表[0].内容.Count > 0)
                            {
                                盒装标准 = 该规格盒子列表[0].内容[0].盒装标准;
                            }

                            // 分箱
                            for (int i = 0; i < 该规格盒子列表.Count; i += 盒装标准)
                            {
                                // 每盒装标准个盒子为一箱
                                int 箱号 = i / 盒装标准 + 1;
                                string 分组键 = $"{纸箱规格}_{箱号}";

                                if (!分组数据.ContainsKey(分组键))
                                {
                                    分组数据[分组键] = new List<(List<结果数据.盒子内容>, int)>();
                                }

                                // 取当前箱应有的盒子数（可能不足盒装标准）
                                int 取盒数 = Math.Min(盒装标准, 该规格盒子列表.Count - i);

                                // 将这些盒子添加到当前箱
                                for (int j = 0; j < 取盒数; j++)
                                {
                                    if (i + j < 该规格盒子列表.Count)
                                    {
                                        分组数据[分组键].Add(该规格盒子列表[i + j]);
                                    }
                                }
                            }
                        }

                        // 为每个分组创建一个Tab页
                        int groupIndex = 1;
                        foreach (var group in 分组数据)
                        {
                            var parts = group.Key.Split('_');
                            string 纸箱规格 = parts[0];
                            int 箱号 = int.Parse(parts[1]);
                            var 组内容 = group.Value;

                            // 计算盒装标准
                            int 盒装标准 = 组内容[0].内容[0].盒装标准;

                            // 修改标题显示箱号、规格和盒子数量
                            TabPage tabPage = new TabPage($"第{groupIndex}箱: {纸箱规格} ({组内容.Count}/{盒装标准}盒)");
                            tabPage.ContextMenuStrip = tabContextMenu;

                            // 存储分组数据
                            tabPage.Tag = new TabPageData
                            {
                                Data = 选中数据,
                                BoxIndex = 组内容[0].索引 // 使用组内第一个盒子的索引
                            };

                            // 创建DataGridView显示合并后的数据
                            DataGridView groupGridView = new DataGridView();
                            groupGridView.Dock = DockStyle.Fill;
                            groupGridView.AllowUserToAddRows = false;
                            groupGridView.ReadOnly = true;
                            groupGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

                            // 添加列
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "盒号",
                                Name = "盒号",
                                Width = 60
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "序号",
                                Name = "序号",
                                Width = 80
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "条数",
                                Name = "条数",
                                Width = 30
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "米数",
                                Name = "米数",
                                Width = 50
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "线长",
                                Name = "线长",
                                Width = 100
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "纸箱规格",
                                Name = "纸箱规格",
                                Width = 100
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "盒装标准",
                                Name = "盒装标准",
                                Width = 50
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签码1",
                                Name = "标签码1",
                                Width = 100
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签码2",
                                Name = "标签码2",
                                Width = 100
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签码3",
                                Name = "标签码3",
                                Width = 100
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签码4",
                                Name = "标签码4",
                                Width = 100
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "客户型号",
                                Name = "客户型号",
                                Width = 100
                            });
                            groupGridView.Columns.Add(new DataGridViewTextBoxColumn
                            {
                                HeaderText = "标签显示长度",
                                Name = "标签显示长度",
                                Width = 100
                            });

                            // 添加数据
                            int boxNumber = 1;
                            foreach (var (内容, _) in 组内容)
                            {
                                foreach (var 行数据 in 内容)
                                {
                                    int rowIndex = groupGridView.Rows.Add();
                                    var row = groupGridView.Rows[rowIndex];

                                    row.Cells["盒号"].Value = $"第{boxNumber}盒";
                                    row.Cells["序号"].Value = 行数据.序号;
                                    row.Cells["条数"].Value = 行数据.条数;
                                    row.Cells["米数"].Value = 行数据.米数;
                                    row.Cells["线长"].Value = 行数据.线长;
                                    row.Cells["纸箱规格"].Value = 行数据.纸箱规格;
                                    row.Cells["盒装标准"].Value = 行数据.盒装标准;
                                    row.Cells["标签码1"].Value = 行数据.标签码1;
                                    row.Cells["标签码2"].Value = 行数据.标签码2;
                                    row.Cells["标签码3"].Value = 行数据.标签码3;
                                    row.Cells["标签码4"].Value = 行数据.标签码4;
                                    row.Cells["客户型号"].Value = 行数据.客户型号;
                                    row.Cells["标签显示长度"].Value = 行数据.标签显示长度;
                                    //MessageBox.Show(行数据.标签显示长度);
                                }
                                boxNumber++;
                            }

                            tabPage.Controls.Add(groupGridView);
                            tabControl.TabPages.Add(tabPage);
                            groupIndex++;
                        }
                    }
                }
            };

            //给标签统计数据
            Action<TabPage, biaoqian> processTab = (selectedTab, operation) =>
            {
                if (selectedTab == null || !(selectedTab.Tag is TabPageData tag) || !(tag.Data is 结果数据 选中数据))
                    return;

                // 获取当前选中的箱子
                var 当前数据视图 = selectedTab.Controls.OfType<DataGridView>().FirstOrDefault();
                if (当前数据视图 == null) return;

                // 直接从DataGridView获取数据
                Dictionary<double, int> 米数条数统计 = new Dictionary<double, int>();
                string 纸箱规格 = "";

                // 遍历当前视图中的所有可见行
                foreach (DataGridViewRow 行 in 当前数据视图.Rows)
                {
                    if (!行.Visible) continue;

                    // 获取米数列
                    var 米数单元格 = 行.Cells["米数"];
                    if (米数单元格 != null && 米数单元格.Value != null && double.TryParse(米数单元格.Value.ToString(), out double 米数))
                    {
                        int 条数 = 1; // 默认为1条

                        // 尝试获取条数列
                        var 条数单元格 = 行.Cells["条数"];
                        if (条数单元格 != null && 条数单元格.Value != null)
                        {
                            if (int.TryParse(条数单元格.Value.ToString(), out int 解析条数) && 解析条数 > 0)
                            {
                                条数 = 解析条数;
                            }
                        }

                        // 累计统计
                        if (!米数条数统计.ContainsKey(米数))
                        {
                            米数条数统计[米数] = 0;
                        }
                        米数条数统计[米数] += 条数;
                    }

                    // 获取纸箱规格
                    var 规格单元格 = 行.Cells["纸箱规格"];
                    if (string.IsNullOrEmpty(纸箱规格) && 规格单元格 != null && 规格单元格.Value != null)
                    {
                        纸箱规格 = 规格单元格.Value.ToString();
                    }
                }

                // 格式化输出
                StringBuilder stats = new StringBuilder();
                bool isFirst = true;

                foreach (var 统计 in 米数条数统计.OrderByDescending(k => k.Key))
                {
                    if (!isFirst) stats.Append(", ");
                    string unit = 统计.Value == 1 ? "PC" : "PCS";
                    stats.Append($"{统计.Key:F3}m*{统计.Value}{unit}");
                    isFirst = false;
                }

                // 更新UI
                textBox_唛头数量.Text = stats.ToString();
                if (!string.IsNullOrEmpty(纸箱规格))
                {
                    textBox_唛头尺寸.Text = 纸箱规格 + " CM";
                }

                // 执行操作
                switch (operation)
                {
                    case biaoqian.dayin:
                        print_btn_Click(null, EventArgs.Empty);
                        break;

                    case biaoqian.yulan:
                        shengcheng_maitou(biaoqian.yulan);
                        break;

                    case biaoqian.lingcun:
                        shengcheng_maitou(biaoqian.lingcun);
                        break;
                }
            };

            // 处理导出按钮点击事件
            exportItem.Click += (s, evt) =>
            {
                // 调用唛头专用的导出方法
                唛头导出(结果列表);
            };
            // 处理预览按钮点击事件
            yulang1.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "80x150";
                processTab(tabControl.SelectedTab, biaoqian.yulan);
            };
            yulang2.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x60";
                processTab(tabControl.SelectedTab, biaoqian.yulan);
            };
            yulang3.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x45";
                processTab(tabControl.SelectedTab, biaoqian.yulan);
            };

            // 处理另存按钮点击事件
            lingcun1.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "80x150";
                processTab(tabControl.SelectedTab, biaoqian.lingcun);
            };
            lingcun2.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x60";
                processTab(tabControl.SelectedTab, biaoqian.lingcun);
            };
            lingcun3.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x45";
                processTab(tabControl.SelectedTab, biaoqian.lingcun);
            };

            // 处理打印按钮点击事件
            viewBoxDetailsItem1.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "80x150";
                processTab(tabControl.SelectedTab, biaoqian.dayin);
            };
            viewBoxDetailsItem2.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x60";
                processTab(tabControl.SelectedTab, biaoqian.dayin);
            };
            viewBoxDetailsItem3.Click += (s, evt) =>
            {
                comboBox_唛头规格.Text = "90x45";
                processTab(tabControl.SelectedTab, biaoqian.dayin);
            };

            // 如果有数据，默认选中第一行
            if (fileGridView.Rows.Count > 0)
            {
                fileGridView.Rows[0].Selected = true;
            }

            // 添加控件到分割面板
            splitContainer.Panel1.Controls.Add(fileGridView);
            splitContainer.Panel2.Controls.Add(tabControl);

            // 添加分割面板到窗体
            resultForm.Controls.Add(splitContainer);

            // 显示窗体
            resultForm.ShowDialog();
        }

        private void 唛头导出(List<结果数据> 结果列表)
        {
            // 使用SaveFileDialog让用户选择保存位置
            using (SaveFileDialog saveDialog = new SaveFileDialog())
            {
                saveDialog.Filter = "Excel文件|*.xlsx";
                saveDialog.Title = "保存唛头Excel";
                saveDialog.FileName = "唛头汇总表.xlsx";

                if (saveDialog.ShowDialog() == DialogResult.OK)
                {
                    using (ExcelPackage package = new ExcelPackage())
                    {
                        var worksheet = package.Workbook.Worksheets.Add("唛头汇总");

                        // 设置表头
                        worksheet.Cells[1, 1].Value = "序号";
                        worksheet.Cells[1, 2].Value = "标签码1";
                        worksheet.Cells[1, 3].Value = "标签码2";
                        worksheet.Cells[1, 4].Value = "标签码3";
                        worksheet.Cells[1, 5].Value = "标签码4";
                        worksheet.Cells[1, 6].Value = "标签码5";
                        worksheet.Cells[1, 7].Value = "条数";
                        worksheet.Cells[1, 8].Value = "长度";
                        worksheet.Cells[1, 9].Value = "客户型号";
                        worksheet.Cells[1, 10].Value = "PO号";
                        worksheet.Cells[1, 11].Value = "条形码";
                        worksheet.Cells[1, 12].Value = "线长";

                        // 设置表头样式
                        using (var range = worksheet.Cells[1, 1, 1, 12])
                        {
                            range.Style.Font.Bold = true;
                            range.Style.Fill.PatternType = OfficeOpenXml.Style.ExcelFillStyle.Solid;
                            range.Style.Fill.BackgroundColor.SetColor(System.Drawing.Color.LightGray);
                        }

                        // 填充数据，按箱子汇总
                        int rowIndex = 2;
                        int seqNo = 1;

                        foreach (var 数据 in 结果列表)
                        {
                            // 获取客户型号和PO号
                            string 客户型号 = 数据.产品型号 ?? "";
                            string PO号 = textBox_订单编号.Text ?? "";

                            // 先按纸箱规格和盒装标准分组
                            var 按规格分组 = new Dictionary<string, List<List<结果数据.盒子内容>>>();

                            foreach (var 盒子列表 in 数据.盒子列表)
                            {
                                if (盒子列表.Count > 0)
                                {
                                    string 纸箱规格 = 盒子列表[0].纸箱规格 ?? "";
                                    int 盒装标准 = 盒子列表[0].盒装标准;

                                    // 计算箱号
                                    int 箱号 = 0;
                                    string 分组键 = "";

                                    // 尝试根据已有的分组找到合适的箱子
                                    foreach (var kvp in 按规格分组)
                                    {
                                        if (kvp.Key.StartsWith(纸箱规格) && kvp.Value.Count < 盒装标准)
                                        {
                                            分组键 = kvp.Key;
                                            break;
                                        }
                                    }

                                    // 如果没找到合适的箱子，创建新箱子
                                    if (string.IsNullOrEmpty(分组键))
                                    {
                                        箱号 = 按规格分组.Count(k => k.Key.StartsWith(纸箱规格)) + 1;
                                        分组键 = $"{纸箱规格}_{箱号}";
                                        按规格分组[分组键] = new List<List<结果数据.盒子内容>>();
                                    }

                                    按规格分组[分组键].Add(盒子列表);
                                }
                            }

                            // 按箱子导出数据
                            foreach (var kvp in 按规格分组)
                            {
                                var 箱内盒子列表 = kvp.Value;

                                if (箱内盒子列表.Count > 0)
                                {
                                    var 第一盒 = 箱内盒子列表[0];
                                    var 第一行内容 = 第一盒[0]; // 使用第一盒的第一行内容作为基本信息

                                    // 使用List<(string 米数, int 条数)>代替Dictionary<string, int>来保持顺序
                                    List<(string 米数, int 条数)> 米数统计列表 = new List<(string 米数, int 条数)>();

                                    int 总条数 = 0;

                                    // 遍历箱内所有盒子，按顺序记录
                                    foreach (var 盒子内容 in 箱内盒子列表)
                                    {
                                        foreach (var 内容 in 盒子内容)
                                        {
                                            // 统计米数
                                            if (!string.IsNullOrEmpty(内容.米数))
                                            {
                                                string 米数 = 内容.米数;
                                                int 条数 = 1;
                                                int.TryParse(内容.条数, out 条数);

                                                // 查找是否已存在该米数
                                                var 现有项 = 米数统计列表.FirstOrDefault(x => x.米数 == 米数);
                                                if (现有项 != default)
                                                {
                                                    // 如果存在，更新条数
                                                    int 索引 = 米数统计列表.IndexOf(现有项);
                                                    米数统计列表[索引] = (米数, 现有项.条数 + 条数);
                                                }
                                                else
                                                {
                                                    // 如果不存在，添加新项
                                                    米数统计列表.Add((米数, 条数));
                                                }

                                                总条数 += 条数;
                                            }
                                        }
                                    }

                                    // 格式化长度信息
                                    StringBuilder 长度信息 = new StringBuilder();
                                    bool isFirst = true;

                                    string 单位文件路径 = Path.Combine(folderPath, "订单资料", "单位.txt");
                                    string 单位内容 = "m";
                                    if (File.Exists(单位文件路径))
                                    {
                                        单位内容 = File.ReadAllText(单位文件路径).Trim();
                                    }

                                    foreach (var (米数, 条数) in 米数统计列表)
                                    {
                                        if (!isFirst)
                                            长度信息.Append(",");

                                        string unit = 条数 == 1 ? "PC" : "PCS";
                                        // 找到所有该米数的内容
                                        var 对应内容 = 箱内盒子列表.SelectMany(盒 => 盒)
                                            .Where(x => double.TryParse(x.米数, out double mval) && Math.Abs(mval - double.Parse(米数)) < 0.0001)
                                            .ToList();
                                        string 标签显示长度 = 对应内容.FirstOrDefault()?.标签显示长度 ?? "";

                                        bool 当前项只有一条 = 条数 == 1 && 米数统计列表.Count == 1;

                                        switch (单位内容)
                                        {
                                            case "m":
                                                if (当前项只有一条)
                                                    长度信息.Append($"{米数}m");
                                                else
                                                    长度信息.Append($"{米数}m*{条数}{unit}");
                                                break;

                                            case "mm":
                                                if (double.TryParse(米数, out double mVal))
                                                {
                                                    int mmVal = (int)Math.Round(mVal * 1000);
                                                    if (当前项只有一条)
                                                        长度信息.Append($"{mmVal}mm");
                                                    else
                                                        长度信息.Append($"{mmVal}mm*{条数}{unit}");
                                                }
                                                else
                                                {
                                                    if (当前项只有一条)
                                                        长度信息.Append($"{米数}mm");
                                                    else
                                                        长度信息.Append($"{米数}mm*{条数}{unit}");
                                                }
                                                break;

                                            case "Ft":
                                            case "IN":
                                                if (当前项只有一条)
                                                    长度信息.Append($"{标签显示长度}{单位内容}");
                                                else
                                                    长度信息.Append($"{标签显示长度}{单位内容}*{条数}{unit}");
                                                break;

                                            case "m(Ft)":
                                            case "m(IN)":
                                                string 括号单位 = 单位内容 == "m(Ft)" ? "Ft" : "IN";
                                                if (当前项只有一条)
                                                    长度信息.Append($"{米数}m({标签显示长度}{括号单位})");
                                                else
                                                    长度信息.Append($"{米数}m({标签显示长度}{括号单位})*{条数}{unit}");
                                                break;

                                            default:
                                                if (当前项只有一条)
                                                    长度信息.Append($"{米数}m");
                                                else
                                                    长度信息.Append($"{米数}m*{条数}{unit}");
                                                break;
                                        }

                                        isFirst = false;
                                    }

                                    // 序号
                                    worksheet.Cells[rowIndex, 1].Value = seqNo++;

                                    // 合并箱内所有盒子的标签码1~4（去重，非空，用逗号分隔）
                                    string 标签码1 = string.Join(",", 箱内盒子列表.SelectMany(盒 => 盒.Select(x => x.标签码1)).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());
                                    string 标签码2 = string.Join(",", 箱内盒子列表.SelectMany(盒 => 盒.Select(x => x.标签码2)).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());
                                    string 标签码3 = string.Join(",", 箱内盒子列表.SelectMany(盒 => 盒.Select(x => x.标签码3)).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());
                                    string 标签码4 = string.Join(",", 箱内盒子列表.SelectMany(盒 => 盒.Select(x => x.标签码4)).Where(x => !string.IsNullOrWhiteSpace(x)).Distinct());

                                    worksheet.Cells[rowIndex, 2].Value = 标签码1;
                                    worksheet.Cells[rowIndex, 3].Value = 标签码2;
                                    worksheet.Cells[rowIndex, 4].Value = 标签码3;
                                    worksheet.Cells[rowIndex, 5].Value = 标签码4;
                                    worksheet.Cells[rowIndex, 6].Value = ""; // 标签码5

                                    // 条数（整个箱子的总条数）
                                    worksheet.Cells[rowIndex, 7].Value = 1;

                                    // 长度信息（整个箱子的汇总）
                                    worksheet.Cells[rowIndex, 8].Value = 长度信息.ToString();

                                    // 客户型号汇总（合并所有盒子的所有内容的客户型号）
                                    string 客户型号汇总 = string.Join(",",
                                        箱内盒子列表
                                            .SelectMany(盒 => 盒.Select(x => x.客户型号))
                                            .Where(x => !string.IsNullOrWhiteSpace(x))
                                            .Distinct()
                                    );
                                    worksheet.Cells[rowIndex, 9].Value = 客户型号汇总;

                                    // 其他信息
                                    worksheet.Cells[rowIndex, 10].Value = PO号;
                                    worksheet.Cells[rowIndex, 11].Value = "";
                                    worksheet.Cells[rowIndex, 12].Value = 第一行内容.线长 ?? "";

                                    rowIndex++;
                                }
                            }
                        }

                        // 调整列宽
                        worksheet.Cells.AutoFitColumns();

                        // 保存文件
                        package.SaveAs(new FileInfo(saveDialog.FileName));
                        MessageBox.Show("导出成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
        }

        private void EXCEL_包装规格回调(string folderPath)
        {
            // 查找包装材料需求流转单
            string[] excelFiles = Directory.GetFiles(folderPath, "*.xlsx");
            string 流转单路径 = null;

            foreach (string filePath in excelFiles)
            {
                string fileName = Path.GetFileName(filePath);
                if (fileName.Contains("包装材料需求流转单") || fileName.Contains("流转单"))
                {
                    流转单路径 = filePath;
                    break;
                }
            }

            // 创建一个字典来存储文件名与纸箱规格的关系
            Dictionary<string, List<(string 规格, int 数量, string 包装编码, int 盒装标准)>> 纸箱规格字典 =
                new Dictionary<string, List<(string, int, string, int)>>(StringComparer.OrdinalIgnoreCase);

            // 创建一个字典来存储BOM物料编码与纸箱规格的关系
            Dictionary<string, List<(string 规格, int 数量, string 包装编码, int 盒装标准)>> BOM物料纸箱规格字典 =
                new Dictionary<string, List<(string, int, string, int)>>(StringComparer.OrdinalIgnoreCase);

            StringBuilder logMessages = new StringBuilder();
            // 如果找到流转单，解析其中的纸箱规格信息
            if (!string.IsNullOrEmpty(流转单路径))
            {
                logMessages.AppendLine($"找到流转单: {流转单路径}");

                try
                {
                    // 确保文件可写
                    FileInfo fileInfo = new FileInfo(流转单路径);
                    if (fileInfo.IsReadOnly)
                    {
                        fileInfo.IsReadOnly = false;
                    }

                    // 使用正确的访问模式打开文件
                    using (var package = new ExcelPackage(fileInfo))
                    {
                        var worksheet = package.Workbook.Worksheets[0]; // 假设流转单在第一个工作表

                        // 查找表头行
                        int startRow = 7;
                        int endRow = worksheet.Dimension.End.Row;

                        // 直接指定列的位置，不进行检查
                        int 物料列 = 2;       // 假设物料在第2列
                        int 物料编码列 = 3;    // 假设物料编码在第3列
                        int 规格列 = 4;       // 假设规格在第4列
                        int 需求数量列 = 5;    // 假设需求数量在第5列
                        int 备注列 = 9;
                        int 文件名列 = 11;      // 假设文件名在第7列

                        logMessages.AppendLine($"使用固定列位置: 物料列: {物料列}, 物料编码列: {物料编码列}, 规格列: {规格列}, 需求数量列: {需求数量列}, 备注列: {备注列}, 文件名列: {文件名列}");

                        // 当前处理的BOM物料编码
                        string 当前BOM物料编码 = "";

                        // 遍历所有数据行
                        for (int dataRow = startRow; dataRow <= endRow; dataRow++)
                        {
                            string 物料 = worksheet.Cells[dataRow, 物料列].Text.Trim();
                            string 物料编码 = 物料编码列 > 0 ? worksheet.Cells[dataRow, 物料编码列].Text.Trim() : "";
                            string 规格 = worksheet.Cells[dataRow, 规格列].Text.Trim();
                            string 需求数量文本 = worksheet.Cells[dataRow, 需求数量列].Text.Trim();
                            string 备注 = 备注列 > 0 ? worksheet.Cells[dataRow, 备注列].Text.Trim() : "";
                            string 文件名 = worksheet.Cells[dataRow, 文件名列].Text.Trim();

                            // 如果是BOM物料，记录编码
                            if (!string.IsNullOrEmpty(物料) && 物料.Contains("BOM物料"))
                            {
                                当前BOM物料编码 = 物料编码;
                                logMessages.AppendLine($"行 {dataRow}: 发现BOM物料: '{物料}', 编码: '{物料编码}'");
                            }

                            // 解析盒装标准
                            int 盒装标准 = 1; // 默认为1盒装
                            if (!string.IsNullOrEmpty(备注))
                            {
                                // 使用正则表达式提取数字 + "盒装标准"
                                var match = System.Text.RegularExpressions.Regex.Match(备注, @"(\d+)\s*盒装标准");
                                if (match.Success && match.Groups.Count > 1)
                                {
                                    if (int.TryParse(match.Groups[1].Value, out int 解析结果))
                                    {
                                        盒装标准 = 解析结果;
                                    }
                                }
                                else
                                {
                                    // 如果正则表达式没有匹配到，尝试其他方法
                                    if (备注.Contains("2盒装标准"))
                                        盒装标准 = 2;
                                    else if (备注.Contains("3盒装标准"))
                                        盒装标准 = 3;
                                    else if (备注.Contains("5盒装标准"))
                                        盒装标准 = 5;
                                }
                            }
                            // 尝试解析需求数量
                            int 需求数量 = 0;
                            int.TryParse(需求数量文本, out 需求数量);

                            // 如果是纸箱，同时关联到当前BOM物料编码
                            if (!string.IsNullOrEmpty(物料) && 物料.Contains("纸箱") && !string.IsNullOrEmpty(当前BOM物料编码))
                            {
                                if (需求数量 > 0)
                                {
                                    if (!BOM物料纸箱规格字典.ContainsKey(当前BOM物料编码))
                                    {
                                        BOM物料纸箱规格字典[当前BOM物料编码] = new List<(string, int, string, int)>();
                                    }

                                    BOM物料纸箱规格字典[当前BOM物料编码].Add((规格, 需求数量, 物料编码, 盒装标准));
                                    logMessages.AppendLine($"关联BOM物料 '{当前BOM物料编码}' 与纸箱规格: '{规格}', 数量: {需求数量}, 物料编码: '{物料编码}', 盒装标准: {盒装标准}");
                                }
                            }

                            // 同时也保持原来的文件名关联
                            if (!string.IsNullOrEmpty(物料) && 物料.Contains("纸箱") && !string.IsNullOrEmpty(文件名))
                            {
                                if (需求数量 > 0)
                                {
                                    string fileNameWithoutExt = Path.GetFileNameWithoutExtension(文件名);
                                    if (!string.IsNullOrEmpty(fileNameWithoutExt) && !string.IsNullOrEmpty(规格))
                                    {
                                        if (!纸箱规格字典.ContainsKey(fileNameWithoutExt))
                                        {
                                            纸箱规格字典[fileNameWithoutExt] = new List<(string, int, string, int)>();
                                        }
                                        纸箱规格字典[fileNameWithoutExt].Add((规格, 需求数量, 物料编码, 盒装标准));
                                        logMessages.AppendLine($"关联文件 '{fileNameWithoutExt}' 与纸箱规格: '{规格}', 数量: {需求数量}, 物料编码: '{物料编码}', 盒装标准: {盒装标准}");
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"解析流转单时出错: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("未找到包装材料需求流转单", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            // 打印BOM物料与纸箱规格的关联信息
            logMessages.AppendLine("\nBOM物料与纸箱规格的关联信息:");
            foreach (var entry in BOM物料纸箱规格字典)
            {
                logMessages.AppendLine($"BOM物料编码: {entry.Key}");
                foreach (var item in entry.Value)
                {
                    logMessages.AppendLine($"  规格: {item.规格}, 数量: {item.数量}, 包装编码: {item.包装编码}, 盒装标准: {item.盒装标准}");
                }
            }

            // 打印文件名与纸箱规格的关联信息
            logMessages.AppendLine("\n文件名与纸箱规格的关联信息:");
            foreach (var entry in 纸箱规格字典)
            {
                logMessages.AppendLine($"文件名: {entry.Key}");
                foreach (var item in entry.Value)
                {
                    logMessages.AppendLine($"  规格: {item.规格}, 数量: {item.数量}, 包装编码: {item.包装编码}, 盒装标准: {item.盒装标准}");
                }
            }

            // 显示解析日志
            //MessageBox.Show(logMessages.ToString(), "流转单解析日志", MessageBoxButtons.OK, MessageBoxIcon.Information);

            // 创建一个字典来跟踪每个文件中每个规格已分配的数量
            Dictionary<string, Dictionary<string, int>> 文件规格已分配数量 = new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);

            // 初始化已分配数量字典
            foreach (var kvp in 纸箱规格字典)
            {
                string fileName = kvp.Key;
                文件规格已分配数量[fileName] = new Dictionary<string, int>();

                foreach (var item in kvp.Value)
                {
                    文件规格已分配数量[fileName][item.规格] = 0;
                }
            }
            // 创建一个字典来跟踪每个规格已分配的数量（用于BOM物料编码匹配）
            Dictionary<string, int> 规格已分配数量 = new Dictionary<string, int>();

            // 初始化规格已分配数量字典
            foreach (var kvp in BOM物料纸箱规格字典)
            {
                foreach (var item in kvp.Value)
                {
                    if (!规格已分配数量.ContainsKey(item.规格))
                    {
                        规格已分配数量[item.规格] = 0;
                    }
                }
            }

            // 回填纸箱规格数据
            foreach (string filePath in excelFiles)
            {
                string fileName = Path.GetFileNameWithoutExtension(filePath);

                // 排除包装材料需求流转单
                if (fileName.Contains("包装材料需求流转单") || fileName.Contains("流转单"))
                    continue;

                // 显示当前处理的文件
                //MessageBox.Show($"处理文件: '{fileName}'", "处理进度", MessageBoxButtons.OK, MessageBoxIcon.Information);

                // 查找对应的纸箱规格信息
                string matchedFileName = null;
                List<(string 规格, int 数量, string 包装编码, int 盒装标准)> 纸箱规格列表 = null;

                foreach (var kvp in 纸箱规格字典)
                {
                    string fileNameWithoutExt = kvp.Key;
                    // 检查文件名是否匹配
                    if (fileName.Equals(fileNameWithoutExt, StringComparison.OrdinalIgnoreCase))
                    {
                        matchedFileName = fileNameWithoutExt;
                        纸箱规格列表 = kvp.Value;
                        StringBuilder matchInfo = new StringBuilder();
                        matchInfo.AppendLine($"找到匹配的纸箱规格信息: '{fileNameWithoutExt}'");
                        foreach (var spec in 纸箱规格列表)
                        {
                            matchInfo.AppendLine($"  规格: '{spec.规格}', 数量: {spec.数量}, 包装编码: '{spec.包装编码}', 盒装标准: {spec.盒装标准}");
                        }
                        //MessageBox.Show(matchInfo.ToString(), "匹配信息", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        break;
                    }
                }

                if (纸箱规格列表 == null || 纸箱规格列表.Count == 0 || matchedFileName == null)
                {
                    MessageBox.Show($"文件 '{fileName}' 没有匹配的纸箱规格信息", "提示", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    continue;
                }

                // 读取Excel文件内容并回填纸箱规格
                FileInfo fileInfo = new FileInfo(filePath);
                if (fileInfo.IsReadOnly)
                {
                    fileInfo.IsReadOnly = false;
                }

                try
                {
                    using (ExcelPackage package = new ExcelPackage(fileInfo))
                    {
                        bool hasChanges = false;

                        // 遍历所有工作表
                        for (int sheetIndex = 0; sheetIndex < package.Workbook.Worksheets.Count; sheetIndex++)
                        {
                            var worksheet = package.Workbook.Worksheets[sheetIndex];

                            if (worksheet.Dimension == null)
                                continue;

                            StringBuilder sheetLog = new StringBuilder();
                            sheetLog.AppendLine($"处理工作表: '{worksheet.Name}'");

                            // 修改确定表头行的逻辑
                            int headerRow = 1; // 直接使用第1行作为表头行，因为从截图看表头就在第1行

                            // 确定列索引
                            int 包装编码列 = -1, 纸箱规格列 = -1, 盒装标准列 = -1;

                            // 遍历所有列以找到需要的列
                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                string headerText = worksheet.Cells[headerRow, col].Text.Trim();

                                switch (headerText)
                                {
                                    case "包装编码":
                                        包装编码列 = col;
                                        break;

                                    case "纸箱规格":
                                        纸箱规格列 = col;
                                        break;

                                    case "盒装标准":
                                        盒装标准列 = col;
                                        break;
                                }
                            }

                            // 如果没有找到包装编码列，使用列I（第9列）
                            if (包装编码列 <= 0)
                            {
                                包装编码列 = 9; // 从截图看，包装编码在第I列
                            }

                            // 如果没有找到纸箱规格列，在包装编码列后添加
                            if (纸箱规格列 <= 0)
                            {
                                纸箱规格列 = 包装编码列 + 1;
                                // 添加列标题
                                worksheet.Cells[headerRow, 纸箱规格列].Value = "纸箱规格";
                                worksheet.Cells[headerRow, 纸箱规格列].Style.Font.Bold = true;
                            }

                            // 如果没有找到盒装标准列，在纸箱规格列后添加
                            if (盒装标准列 <= 0)
                            {
                                盒装标准列 = 纸箱规格列 + 1;
                                // 添加列标题
                                worksheet.Cells[headerRow, 盒装标准列].Value = "盒装标准";
                                worksheet.Cells[headerRow, 盒装标准列].Style.Font.Bold = true;
                            }

                            // 添加安全检查
                            if (worksheet.Dimension == null)
                            {
                                MessageBox.Show($"工作表 '{worksheet.Name}' 为空或格式不正确", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                continue;
                            }

                            int maxColumn = worksheet.Dimension.End.Column;
                            if (包装编码列 > maxColumn || 纸箱规格列 > maxColumn || 盒装标准列 > maxColumn)
                            {
                                // 扩展工作表以容纳新列
                                if (盒装标准列 > maxColumn)
                                {
                                    // 确保工作表有足够的列
                                    for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                                    {
                                        worksheet.Cells[row, 盒装标准列].Value = row == 1 ? "盒装标准" : "";
                                    }
                                }
                            }

                            // 确保在保存前设置hasChanges为true
                            hasChanges = true;

                            // 获取工作表中的第一个包装编码
                            string 工作表包装编码 = "";
                            string 工作表BOM物料编码 = "";

                            // 尝试在工作表中查找BOM物料编码
                            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                            {
                                for (int row = 1; row <= Math.Min(20, worksheet.Dimension.End.Row); row++)
                                {
                                    string cellValue = worksheet.Cells[row, col].Text.Trim();
                                    // 检查是否是BOM物料编码格式
                                    if (cellValue.Contains("30.23.") || (cellValue.Contains("BOM") && cellValue.Contains("物料")))
                                    {
                                        // 提取完整的编码部分
                                        if (cellValue.Contains("30.23."))
                                        {
                                            // 使用正则表达式提取完整的编码
                                            var match = System.Text.RegularExpressions.Regex.Match(cellValue, @"30\.23\.\d+");
                                            if (match.Success)
                                            {
                                                工作表BOM物料编码 = match.Value;
                                            }
                                        }

                                        sheetLog.AppendLine($"  找到BOM物料编码: '{工作表BOM物料编码}'");
                                        break;
                                    }
                                }
                                if (!string.IsNullOrEmpty(工作表BOM物料编码))
                                    break;
                            }

                            // 获取工作表中的第一个包装编码
                            for (int row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
                            {
                                string 包装编码 = worksheet.Cells[row, 包装编码列].Text.Trim();
                                if (!string.IsNullOrEmpty(包装编码))
                                {
                                    工作表包装编码 = 包装编码;
                                    break;
                                }
                            }
                            if (string.IsNullOrEmpty(工作表包装编码))
                            {
                                sheetLog.AppendLine($"  工作表 '{worksheet.Name}' 中未找到有效的包装编码，跳过处理");
                                //MessageBox.Show(sheetLog.ToString(), "工作表处理", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                continue;
                            }

                            sheetLog.AppendLine($"  工作表 '{worksheet.Name}' 中包装编码: '{工作表包装编码}'");

                            // 查找匹配的纸箱规格
                            string 分配规格 = "";
                            int 盒装标准 = 1;
                            bool 匹配成功 = false;

                            // 如果找到了BOM物料编码，尝试通过它匹配纸箱规格
                            if (!string.IsNullOrEmpty(工作表BOM物料编码))
                            {
                                sheetLog.AppendLine($"  尝试通过BOM物料编码 '{工作表BOM物料编码}' 匹配纸箱规格");

                                // 首先尝试完全匹配
                                if (BOM物料纸箱规格字典.ContainsKey(工作表BOM物料编码))
                                {
                                    var BOM匹配规格列表 = BOM物料纸箱规格字典[工作表BOM物料编码];

                                    foreach (var item in BOM匹配规格列表)
                                    {
                                        // 检查是否已经分配了足够的数量
                                        if (规格已分配数量.ContainsKey(item.规格) && 规格已分配数量[item.规格] < item.数量 * item.盒装标准)
                                        {
                                            分配规格 = item.规格;
                                            盒装标准 = item.盒装标准;
                                            规格已分配数量[item.规格]++;
                                            sheetLog.AppendLine($"  为工作表 '{worksheet.Name}' 分配纸箱规格: '{分配规格}' (通过BOM物料编码精确匹配)");
                                            sheetLog.AppendLine($"  规格 '{分配规格}' 已分配数量: {规格已分配数量[item.规格]}/{item.数量 * item.盒装标准} (盒装标准: {item.盒装标准})");
                                            匹配成功 = true;
                                            break;
                                        }
                                    }
                                }

                                // 记录所有可用的BOM物料编码，帮助调试
                                sheetLog.AppendLine("  可用的BOM物料编码:");
                                foreach (var key in BOM物料纸箱规格字典.Keys)
                                {
                                    sheetLog.AppendLine($"    - {key}");
                                }
                            }
                            // 如果BOM物料编码匹配失败，尝试通过包装编码匹配
                            if (!匹配成功)
                            {
                                // 首先尝试直接匹配包装编码
                                sheetLog.AppendLine($"  尝试通过包装编码 '{工作表包装编码}' 匹配");

                                foreach (var item in 纸箱规格列表)
                                {
                                    if (item.包装编码 == 工作表包装编码)
                                    {
                                        // 检查是否已经分配了足够的数量
                                        if (文件规格已分配数量[matchedFileName].ContainsKey(item.规格) &&
                                            文件规格已分配数量[matchedFileName][item.规格] < item.数量 * item.盒装标准)
                                        {
                                            分配规格 = item.规格;
                                            盒装标准 = item.盒装标准;
                                            文件规格已分配数量[matchedFileName][item.规格]++;
                                            sheetLog.AppendLine($"  为工作表 '{worksheet.Name}' 分配纸箱规格: '{分配规格}' (直接匹配包装编码: '{工作表包装编码}')");
                                            sheetLog.AppendLine($"  规格 '{分配规格}' 已分配数量: {文件规格已分配数量[matchedFileName][item.规格]}/{item.数量 * item.盒装标准} (盒装标准: {item.盒装标准})");
                                            匹配成功 = true;
                                            break;
                                        }
                                    }
                                }
                            }

                            // 如果前两种匹配都失败，尝试使用任何可用的规格
                            if (!匹配成功)
                            {
                                // 打印所有可用的纸箱规格和包装编码，帮助调试
                                sheetLog.AppendLine($"  工作表包装编码 '{工作表包装编码}' 没有直接匹配项，尝试使用可用规格");
                                foreach (var item in 纸箱规格列表)
                                {
                                    sheetLog.AppendLine($"  可用规格: '{item.规格}', 数量: {item.数量}, 已分配: {文件规格已分配数量[matchedFileName][item.规格]}/{item.数量 * item.盒装标准}, 包装编码: '{item.包装编码}', 盒装标准: {item.盒装标准}");
                                }
                                // 尝试分配任何可用的纸箱规格
                                foreach (var item in 纸箱规格列表)
                                {
                                    // 检查是否已经分配了足够的数量
                                    if (文件规格已分配数量[matchedFileName].ContainsKey(item.规格) &&
                                        文件规格已分配数量[matchedFileName][item.规格] < item.数量 * item.盒装标准)
                                    {
                                        分配规格 = item.规格;
                                        盒装标准 = item.盒装标准;
                                        文件规格已分配数量[matchedFileName][item.规格]++;
                                        sheetLog.AppendLine($"  为工作表 '{worksheet.Name}' 分配纸箱规格: '{分配规格}' (未匹配包装编码，使用可用规格)");
                                        sheetLog.AppendLine($"  规格 '{分配规格}' 已分配数量: {文件规格已分配数量[matchedFileName][item.规格]}/{item.数量 * item.盒装标准} (盒装标准: {item.盒装标准})");
                                        匹配成功 = true;
                                        break;
                                    }
                                }
                            }

                            // 如果找到了匹配的规格，则回填
                            if (!string.IsNullOrEmpty(分配规格))
                            {
                                int 回填计数 = 0;
                                for (int row = headerRow + 1; row <= worksheet.Dimension.End.Row; row++)
                                {
                                    string 行包装编码 = worksheet.Cells[row, 包装编码列].Text.Trim();
                                    if (行包装编码 == 工作表包装编码)
                                    {
                                        // 直接设置值
                                        worksheet.Cells[row, 纸箱规格列].Value = 分配规格;
                                        worksheet.Cells[row, 盒装标准列].Value = 盒装标准;
                                        回填计数++;
                                        hasChanges = true;
                                    }
                                }

                                sheetLog.AppendLine($"  工作表 '{worksheet.Name}' 回填纸箱规格 '{分配规格}' 到 {回填计数} 行");
                            }

                            // 如果有修改，则保存文件
                            if (hasChanges)
                            {
                                try
                                {
                                    package.Save();
                                    //MessageBox.Show($"已成功回填纸箱规格到文件: '{filePath}'", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show($"保存文件时出错: '{filePath}', 错误: {ex.Message}", "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                }
                            }

                            // 显示当前工作表的处理结果
                            //MessageBox.Show(sheetLog.ToString(), "工作表处理", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        // 如果有修改，则保存文件
                        if (hasChanges)
                        {
                            try
                            {
                                // 在保存前强制刷新
                                package.Workbook.Calculate();

                                // 保存文件
                                package.Save();

                                //MessageBox.Show($"已成功回填纸箱规格到文件: '{filePath}'", "保存成功", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"保存文件时出错: '{filePath}', 错误: {ex.Message}", "保存失败", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"处理文件时出错: '{filePath}', 错误: {ex.Message}", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }   //结束EXCEL 包装规格回调 }

        /// 显示解析结果
    }
}