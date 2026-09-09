using System.Runtime.InteropServices;

namespace DocuFormatPro.Services
{
    /// <summary>
    /// COM 消息过滤器：Word 内部忙（重排/更新域）时拒绝调用（RPC_E_CALL_REJECTED），
    /// 指示 COM 等待后自动重试而不是立刻抛异常。须在 STA 线程上注册。
    /// </summary>
    internal class ComCallRetryFilter : IOleMessageFilter
    {
        public int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo)
            => 0; // SERVERCALL_ISHANDLED

        // dwTickCount 为该调用已等待的毫秒数；超过 30 秒放弃，否则建议等 100ms 后重试
        public int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType)
            => dwTickCount > 30000 ? -1 : 100;

        public int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType)
            => 2; // PENDINGTYPE_WAITDEFHANDLE

        public static void Register()
        {
            CoRegisterMessageFilter(new ComCallRetryFilter(), out _);
        }

        public static void Revoke()
        {
            CoRegisterMessageFilter(null, out _);
        }

#pragma warning disable SYSLIB0001
        [DllImport("ole32.dll")]
        private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter? oldFilter);
#pragma warning restore SYSLIB0001
    }

    [ComImport, Guid("00000016-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    internal interface IOleMessageFilter
    {
        [PreserveSig]
        int HandleInComingCall(int dwCallType, IntPtr hTaskCaller, int dwTickCount, IntPtr lpInterfaceInfo);
        [PreserveSig]
        int RetryRejectedCall(IntPtr hTaskCallee, int dwTickCount, int dwRejectType);
        [PreserveSig]
        int MessagePending(IntPtr hTaskCallee, int dwTickCount, int dwPendingType);
    }
}
