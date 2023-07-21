namespace Measurement.FinsTcp
{
    internal class ErrorCode
    {
        /// <summary>
        /// （若返回的头指令为3）检查命令头中的错误代码
        /// </summary>
        /// <param name="code">错误代码</param>
        /// <returns>指示程序是否可以继续进行</returns>
        internal static bool CheckHeadError(byte code)
        {
            switch (code)
            {
                // 成功的情况
                case 0x00: return true;

                // 抛出异常 : 头指令不是FINS
                case 0x01: return false;

                // 抛出异常 : 数据长度太长
                case 0x02: return false;

                // 抛出异常 : 命令不支持
                case 0x03: return false;
            }

            // 没有匹配的情况
            return false; // 抛出异常 : 未知错误
        }

        /// <summary>
        /// 检查命令帧中的EndCode
        /// </summary>
        /// <param name="main">主码</param>
        /// <param name="sub">副码</param>
        /// <returns>指示程序是否可以继续进行</returns>
        internal static bool CheckEndCode(byte main, byte sub)
        {
            switch (main)
            {
                case 0x00:
                    switch (sub)
                    {
                        // 成功的情况
                        case 0x00: return true;

                        // 错误码64，是因为PLC中产生了报警，但是数据还是能正常得到的，屏蔽64报警或清除plc错误可解决
                        case 0x40: return true;

                        // 抛出异常 : 服务已取消
                        case 0x01: return false;
                    }

                    break;

                case 0x01:
                    switch (sub)
                    {
                        // 抛出异常 : 本地节点不在网络中
                        case 0x01: return false;

                        // 抛出异常 : 令牌超时
                        case 0x02: return false;

                        // 抛出异常 : 重试失败
                        case 0x03: return false;

                        // 抛出异常 : 发送帧太多
                        case 0x04: return false;

                        // 抛出异常 : 节点地址范围错误
                        case 0x05: return false;

                        // 抛出异常 : 节点地址重复
                        case 0x06: return false;
                    }

                    break;

                case 0x02:
                    switch (sub)
                    {
                        // 抛出异常 : 目标节点不在网络中
                        case 0x01: return false;

                        // 抛出异常 : 单元缺失
                        case 0x02: return false;

                        // 抛出异常 : 第三节点缺失
                        case 0x03: return false;

                        // 抛出异常 : 目标节点忙
                        case 0x04: return false;

                        // 抛出异常 : 响应超时
                        case 0x05: return false;
                    }

                    break;

                case 0x03:
                    switch (sub)
                    {
                        // 抛出异常 : 通信控制器错误
                        case 0x01: return false;

                        // 抛出异常 : CPU 单元错误
                        case 0x02: return false;

                        // 抛出异常 : 控制器错误
                        case 0x03: return false;

                        // 抛出异常 : 单元编号错误
                        case 0x04: return false;
                    }

                    break;

                case 0x04:
                    switch (sub)
                    {
                        // 抛出异常 : 未定义的命令
                        case 0x01: return false;

                        // 抛出异常 : 不被当前模型/版本支持
                        case 0x02: return false;
                    }

                    break;

                case 0x05:
                    switch (sub)
                    {
                        // 抛出异常 : 目标地址设置错误
                        case 0x01: return false;

                        // 抛出异常 : 没有路由表
                        case 0x02: return false;

                        // 抛出异常 : 路由表错误
                        case 0x03: return false;

                        // 抛出异常 : 继电器数量过多
                        case 0x04: return false;
                    }

                    break;

                case 0x10:
                    switch (sub)
                    {
                        // 抛出异常 : 命令太长
                        case 0x01: return false;

                        // 抛出异常 : 命令太短
                        case 0x02: return false;

                        // 抛出异常 : 元素/数据不匹配
                        case 0x03: return false;

                        // 抛出异常 : 命令格式错误
                        case 0x04: return false;

                        // 抛出异常 : 头部错误
                        case 0x05: return false;
                    }

                    break;

                case 0x11:
                    switch (sub)
                    {
                        // 抛出异常 : 区域分类缺失
                        case 0x01: return false;

                        // 抛出异常 : 访问大小错误
                        case 0x02: return false;

                        // 抛出异常 : 地址范围错误
                        case 0x03: return false;

                        // 抛出异常 : 地址范围超出
                        case 0x04: return false;

                        // 抛出异常 : 程序缺失
                        case 0x06: return false;

                        // 抛出异常 : 关系错误
                        case 0x09: return false;

                        // 抛出异常 : 重复数据访问
                        case 0x0a: return false;

                        // 抛出异常 : 响应过长
                        case 0x0b: return false;

                        // 抛出异常 : 参数错误
                        case 0x0c: return false;
                    }

                    break;

                case 0x20:
                    switch (sub)
                    {
                        // 抛出异常 : 受保护的
                        case 0x02: return false;

                        // 抛出异常 : 表缺失
                        case 0x03: return false;

                        // 抛出异常 : 数据缺失
                        case 0x04: return false;

                        // 抛出异常 : 程序缺失
                        case 0x05: return false;

                        // 抛出异常 : 文件缺失
                        case 0x06: return false;

                        // 抛出异常 : 数据不匹配
                        case 0x07: return false;
                    }

                    break;

                case 0x21:
                    switch (sub)
                    {
                        // 抛出异常 : 只读
                        case 0x01: return false;

                        // 抛出异常 : 受保护，无法写入数据链路表
                        case 0x02: return false;

                        // 抛出异常 : 无法注册
                        case 0x03: return false;

                        // 抛出异常 : 程序缺失
                        case 0x05: return false;

                        // 抛出异常 : 文件缺失
                        case 0x06: return false;

                        // 抛出异常 : 文件名已存在
                        case 0x07: return false;

                        // 抛出异常 : 无法更改
                        case 0x08: return false;
                    }

                    break;

                case 0x22:
                    switch (sub)
                    {
                        // 抛出异常 : 执行期间不可行
                        case 0x01: return false;

                        // 抛出异常 : 运行时不可行
                        case 0x02: return false;

                        // 抛出异常 : 错误的PLC模式
                        case 0x03: return false;

                        // 抛出异常 : 错误的PLC模式
                        case 0x04: return false;

                        // 抛出异常 : 错误的PLC模式
                        case 0x05: return false;

                        // 抛出异常 : 错误的PLC模式
                        case 0x06: return false;

                        // 抛出异常 : 指定的节点不是轮询节点
                        case 0x07: return false;

                        // 抛出异常 : 无法执行步骤
                        case 0x08: return false;
                    }

                    break;

                case 0x23:
                    switch (sub)
                    {
                        // 抛出异常 : 文件设备缺失
                        case 0x01: return false;

                        // 抛出异常 : 内存缺失
                        case 0x02: return false;

                        // 抛出异常 : 时钟缺失
                        case 0x03: return false;
                    }

                    break;

                case 0x24:
                    switch (sub)
                    {
                        // 抛出异常 : 表缺失
                        case 0x01: return false;
                    }

                    break;

                case 0x25:
                    switch (sub)
                    {
                        // 抛出异常 : 内存错误
                        case 0x02: return false;

                        // 抛出异常 : I/O设置错误
                        case 0x03: return false;

                        // 抛出异常 : I/O点过多
                        case 0x04: return false;

                        // 抛出异常 : CPU总线错误
                        case 0x05: return false;

                        // 抛出异常 : I/O重复
                        case 0x06: return false;

                        // 抛出异常 : CPU总线错误
                        case 0x07: return false;

                        // 抛出异常 : SYSMAC BUS/2 错误
                        case 0x09: return false;

                        // 抛出异常 : CPU总线单元错误
                        case 0x0a: return false;

                        // 抛出异常 : SYSMAC BUS编号重复
                        case 0x0d: return false;

                        // 抛出异常 : 内存错误
                        case 0x0f: return false;

                        // 抛出异常 : SYSMAC BUS终端缺失
                        case 0x10: return false;
                    }

                    break;

                case 0x26:
                    switch (sub)
                    {
                        // 抛出异常 : 无保护
                        case 0x01: return false;

                        // 抛出异常 : 密码错误
                        case 0x02: return false;

                        // 抛出异常 : 受保护
                        case 0x04: return false;

                        // 抛出异常 : 服务已在执行
                        case 0x05: return false;

                        // 抛出异常 : 服务已停止
                        case 0x06: return false;

                        // 抛出异常 : 无执行权限
                        case 0x07: return false;

                        // 抛出异常 : 执行前需要进行设置
                        case 0x08: return false;

                        // 抛出异常 : 必要项未设置
                        case 0x09: return false;

                        // 抛出异常 : 编号已定义
                        case 0x0a: return false;

                        // 抛出异常 : 错误无法清除
                        case 0x0b: return false;
                    }

                    break;

                case 0x30:
                    switch (sub)
                    {
                        // 抛出异常 : 无访问权限
                        case 0x01: return false;
                    }

                    break;

                case 0x40:
                    switch (sub)
                    {
                        // 抛出异常 : 服务终止
                        case 0x01: return false;
                    }

                    break;
            }

            // 没有找到对应的异常
            // 抛出异常 : 未知异常
            return false;
        }
    }
}