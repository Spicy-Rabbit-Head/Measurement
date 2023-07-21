using System;
using System.Linq;
using System.Text.RegularExpressions;

namespace Measurement.FinsTcp
{
    internal class ConvertClass
    {
        /// <summary>
        /// 得到枚举值
        /// </summary>
        /// <param name="txt">如：D100,W100.1</param>
        /// <param name="txtq"></param>
        /// <returns></returns>
        internal static PlcMemory GetPlcMemory(string txt, out string txtq)
        {
            PlcMemory pm;

            // 去除文本两端的空白字符，并将文本转换为大写后获取第一个字符
            var da = txt.Trim().ToUpper().FirstOrDefault();

            // 根据第一个字符的不同，将枚举类型 PlcMemory 赋值给变量 pm
            switch (da)
            {
                case 'D':
                    pm = PlcMemory.DM;
                    break;
                case 'W':
                    pm = PlcMemory.WR;
                    break;
                case 'H':
                    pm = PlcMemory.HR;
                    break;
                case 'A':
                    pm = PlcMemory.AR;
                    break;
                case 'C':
                    pm = PlcMemory.CNT;
                    break;
                case 'I':
                    pm = PlcMemory.CIO;
                    break;
                case 'T':
                    pm = PlcMemory.TIM;
                    break;
                default:
                    // 如果前缀不匹配任何已知类型，则抛出异常，显示无效的前缀和寄存器
                    throw new Exception($"寄存器【{txt}】无效的前缀[{da}]");
            }

            // 通过正则表达式保留 txt 中的数字和小数点，去除其他字符
            txtq = Regex.Replace(txt, "[^0-9.]", "");
            return pm;
        }
    }
}