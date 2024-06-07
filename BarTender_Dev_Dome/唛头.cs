using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Data;
using System.Text.RegularExpressions; // 引入正则表达式命名空间

namespace maitou
{

    public class 唛头
    {

        private string _bmp_path = Application.StartupPath + @"\exp.jpg";
        string 灯带系列 = "";
        string output_name = "Name:LED Flex Linear Light";
        string output_灯带型号 = "";
        string output_电压 = "";
        string output_色温 = "";
        string output_尾巴 = "";
        string output_唛头数量 = "";

        public void 正常型号判断(string aa,
            bool checkBox_客户Name,
            bool checkBox_客户型号,
            string textBox_客户资料,
            string comboBox_标签规格,
            string 标签种类_comboBox,
            string textBox_唛头数量,
            string textBox_唛头尺寸)
        {
            
            // 正则表达式模式，
            string pattern1 = @"^(\w+-\w+-\w+)";
            string pattern2 = @"D(\d+)V";
            //string pattern3 = @"额定功率(\d+)W";
            string pattern3 = @"额定功率(\d+(?:\.\d+)?)W";
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
                output_灯带型号 = $"{artNo}";

                // 使用信息框输出结果
                //MessageBox.Show(output1, "提取结果");

                // 检查复选框是否同时被选中
                bool isCustomerNameChecked = checkBox_客户Name;
                bool isCustomerModelChecked = checkBox_客户型号;

                string originalString = textBox_客户资料;
                int spaceIndex = originalString.LastIndexOf('	');


                if (isCustomerNameChecked && !isCustomerModelChecked)
                {
                    // 如果只有 checkBox_客户Name 被选中，则输出 1
                    //MessageBox.Show("1");
                    output_name = "" + textBox_客户资料;
                }
                else if (!isCustomerNameChecked && isCustomerModelChecked)
                {
                    // 如果只有 checkBox_客户型号 被选中，则输出 2
                    //MessageBox.Show("2");
                    output_灯带型号 = "" + textBox_客户资料;   


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
                        output_name = "" + part1 ;
                        output_灯带型号 = "" + part2 ;

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
                    output_name = "LED Flex Linear Light";
                    output_灯带型号 = $"{artNo}";

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
                if (comboBox_标签规格.Contains("高压"))
                {
                    // 如果包含“高压”，则设置 output_电压 为 AC
                    output_电压 = $"AC {voltageValue}V";
                }
                else
                {
                    // 如果不包含“高压”，则设置 output_电压 为 DC
                    output_电压 = $"DC {voltageValue}V";
                }

            }
            else
            {
                // 如果没有找到匹配项，则输出错误信息
                MessageBox.Show("未找到电压匹配项。", "错误");
            }



            // 色温
            if (parts.Length >= 6)
            {
                // 第五个"-"和第六个"-"之间的内容是parts[5]，因为数组索引是从0开始的
                string contentBetweenFifthAndSixth = parts[5];

                // 检查灯带系列是否为"S"或"D"
                string series = 灯带系列; 
                if (series.Equals("S") || series.Equals("D"))
                {

                    // 如果灯带系列是"S"或"D"，执行新的逻辑

                    output_色温 = $"{contentBetweenFifthAndSixth}";// 示例，表示输出所有色温内容

                }
                else
                {
                    // 如果灯带系列是"S"或"D"，按照现有逻辑判断色温
                    // 检查内容是否为纯字母
                    if (Regex.IsMatch(contentBetweenFifthAndSixth, @"^[a-zA-Z]*$"))
                    {
                        output_色温 = $"{contentBetweenFifthAndSixth}";
                    }
                    else
                    {
                        // 如果包含数字，则提取数字部分
                        string numericValue = Regex.Replace(contentBetweenFifthAndSixth, @"[^0-9]", "");
                        output_色温 = $"{numericValue}K";
                    }
                }

                 


            }
            else
            {
                MessageBox.Show("未找到色温匹配项。", "错误");
            }

            // 计算逗号的数量
            int commaCount = textBox_唛头数量.Count(c => c == ',');

            // 如果逗号数量大于3，在第三个逗号后面添加换行符
            if (commaCount > 3)
            {
                // 找到第三个逗号的位置
                int thirdCommaIndex = textBox_唛头数量.IndexOf(',', textBox_唛头数量.IndexOf(',', textBox_唛头数量.IndexOf(',') + 1) + 1);

                // 确保第三个逗号的位置是有效的
                if (thirdCommaIndex != -1 && thirdCommaIndex < textBox_唛头数量.Length)
                {
                    // 在第三个逗号后面添加换行符
                    output_唛头数量 =textBox_唛头数量.Insert(thirdCommaIndex + 1, Environment.NewLine);
                }
                else
                {
                    // 如果由于某种原因找不到第三个逗号，使用原始文本
                    output_唛头数量 = textBox_唛头数量;
                }
            }
            else {
                output_唛头数量 = textBox_唛头数量;
            }
            



            MessageBox.Show(output_name +"\n" + output_灯带型号 + "\n" + output_电压 + "\n" +  output_唛头数量 + "\n" +  textBox_唛头尺寸  + "\n" + output_色温 , "提取结果");
            // name_CPXXBox.Text = output_name + "\n" + output_灯带型号 + "\n" + output_电压 + "\n" + output_功率 + "\n" + output_灯数 + "\n" + output_剪切单元 + "\n" + output_长度 + "\n" + output_色温 ;

        }

        public void 中国制造(bool checkBox_取消中国制造)
        {
            //判断尾巴Made in China
            if (checkBox_取消中国制造)
            {
                output_尾巴 = " ";
            }
            else
            {
                output_尾巴 = "Made in China";
            }

        }
    }
}