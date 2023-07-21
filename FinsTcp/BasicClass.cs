using System.Net.NetworkInformation;

namespace Measurement.FinsTcp
{
    internal class BasicClass
    {
        // 检查PLC链接状况
        internal static bool PingCheck(string ip, int timeOut)
        {
            // 创建一个 Ping 实例
            var ping = new Ping();

            // 发送 ICMP 回显请求到指定的 IP 地址，并等待响应
            var pingReply = ping.Send(ip, timeOut);

            // 如果 pingReply 为 null，表示没有收到响应
            if (pingReply == null) return false;

            // 判断响应的状态是否为 Success，是否成功收到 ICMP 回显响应
            // 返回 true 表示 Ping 成功，否则返回 false 表示 Ping 失败
            return pingReply.Status == IPStatus.Success;
        }
    }
}