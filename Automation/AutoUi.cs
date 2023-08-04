using System;
using System.Threading;
using System.Windows.Automation;
using System.Windows.Forms;

namespace Measurement.Automation
{
    public class AutoUi
    {
        // 设置其他应用程序的进程名称
        private const string TargetAppName = ".qcc - 250B Network Analyzer [SERVER 1]";

        // 取消令牌
        private static readonly CancellationTokenSource CancellationTokenSource = new CancellationTokenSource();

        // 目标应用程序的主窗口
        private static AutomationElement targetWindow;

        // 验证目标应用程序是否存在
        public bool VerificationState(string url)
        {
            try
            {
                // 查找目标应用程序的主窗口
                targetWindow = FindWindowByTitle(url + TargetAppName);
                return targetWindow != null;
            }
            catch
            {
                return false;
            }
        }

        // 根据窗口标题查找窗口
        private static AutomationElement FindWindowByTitle(string title)
        {
            try
            {
                Condition condition = new PropertyCondition(AutomationElement.NameProperty, title);
                return AutomationElement.RootElement.FindFirst(TreeScope.Children, condition);
            }
            catch
            {
                return null;
            }
        }

        // 等待主线程超时
        public void WaitMainThreadTimeout()
        {
            try
            {
                Thread.Sleep(30000);
                if (!CancellationTokenSource.Token.IsCancellationRequested)
                {
                    ExecutiveExit(targetWindow);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        // 执行退出
        private static bool ExecutiveExit(AutomationElement element)
        {
            try
            {
                // 获取所有子元素
                var elementAll = element.FindAll(TreeScope.Children, Condition.TrueCondition);

                // 如果没有子元素，返回false
                if (elementAll.Count == 0) return false;

                // 遍历子元素 
                foreach (AutomationElement childElement in elementAll)
                {
                    // 如果子元素是对话框并且可用，执行退出
                    if (childElement.Current.ClassName == "#32770" && childElement.Current.IsEnabled)
                    {
                        childElement.SetFocus();
                        SendKeys.SendWait("{ESC}");
                        return true;
                    }

                    return ExecutiveExit(childElement);
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        // 保存数据
        public bool SaveData()
        {
            try
            {
                // 目标窗口是否可用
                if (!targetWindow.Current.IsEnabled)
                {
                    // 执行退出
                    var executiveExit = ExecutiveExit(targetWindow);

                    // 如果退出失败，返回false
                    if (!executiveExit) return false;
                }

                // 执行保存
                targetWindow.SetFocus();
                Thread.Sleep(10);
                SendKeys.SendWait("^s");
                return true;
            }
            catch
            {
                return false;
            }
        }

        // 停止令牌
        public void StopToken()
        {
            CancellationTokenSource.Cancel();
        }
    }
}