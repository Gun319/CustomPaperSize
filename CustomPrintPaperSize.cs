using System;
using System.Runtime.InteropServices;
using System.Security;
using System.Text;

namespace CustomPrintPaperSize
{
    public class CustomPrintPaperSize
    {
        #region WindowsAPI

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct StructPrinterDefaults
        {
            [MarshalAs(UnmanagedType.LPTStr)]
            public String pDatatype;

            public IntPtr pDevMode;

            [MarshalAs(UnmanagedType.I4)]
            public int DesiredAccess;
        };

        [DllImport("winspool.Drv", EntryPoint = "OpenPrinter", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = false,
            CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool OpenPrinter(
            [MarshalAs(UnmanagedType.LPTStr)] string printerName,
            out IntPtr phPrinter,
            ref StructPrinterDefaults pd);

        [DllImport("winspool.Drv", EntryPoint = "ClosePrinter", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = false,
            CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool ClosePrinter(IntPtr phPrinter);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct StructSize
        {
            public Int32 width;
            public Int32 height;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct StructRect
        {
            public Int32 left;
            public Int32 top;
            public Int32 right;
            public Int32 bottom;
        }

        [StructLayout(LayoutKind.Explicit, CharSet = CharSet.Unicode)]
        internal struct FormInfo1
        {
            [FieldOffset(0), MarshalAs(UnmanagedType.I4)] public uint Flags;
            [FieldOffset(4), MarshalAs(UnmanagedType.LPWStr)] public String pName;
            [FieldOffset(8)] public StructSize Size;
            [FieldOffset(16)] public StructRect ImageableArea;
        };

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Ansi/* changed from CharSet=CharSet.Auto */)]
        internal struct StructDevMode
        {
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public String dmDeviceName;
            [MarshalAs(UnmanagedType.U2)] public short dmSpecVersion;
            [MarshalAs(UnmanagedType.U2)] public short dmDriverVersion;
            [MarshalAs(UnmanagedType.U2)] public short dmSize;
            [MarshalAs(UnmanagedType.U2)] public short dmDriverExtra;
            [MarshalAs(UnmanagedType.U4)] public int dmFields;
            [MarshalAs(UnmanagedType.I2)] public short dmOrientation;
            [MarshalAs(UnmanagedType.I2)] public short dmPaperSize;
            [MarshalAs(UnmanagedType.I2)] public short dmPaperLength;
            [MarshalAs(UnmanagedType.I2)] public short dmPaperWidth;
            [MarshalAs(UnmanagedType.I2)] public short dmScale;
            [MarshalAs(UnmanagedType.I2)] public short dmCopies;
            [MarshalAs(UnmanagedType.I2)] public short dmDefaultSource;
            [MarshalAs(UnmanagedType.I2)] public short dmPrintQuality;
            [MarshalAs(UnmanagedType.I2)] public short dmColor;
            [MarshalAs(UnmanagedType.I2)] public short dmDuplex;
            [MarshalAs(UnmanagedType.I2)] public short dmYResolution;
            [MarshalAs(UnmanagedType.I2)] public short dmTTOption;
            [MarshalAs(UnmanagedType.I2)] public short dmCollate;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 32)] public String dmFormName;
            [MarshalAs(UnmanagedType.U2)] public short dmLogPixels;
            [MarshalAs(UnmanagedType.U4)] public int dmBitsPerPel;
            [MarshalAs(UnmanagedType.U4)] public int dmPelsWidth;
            [MarshalAs(UnmanagedType.U4)] public int dmPelsHeight;
            [MarshalAs(UnmanagedType.U4)] public int dmNup;
            [MarshalAs(UnmanagedType.U4)] public int dmDisplayFrequency;
            [MarshalAs(UnmanagedType.U4)] public int dmICMMethod;
            [MarshalAs(UnmanagedType.U4)] public int dmICMIntent;
            [MarshalAs(UnmanagedType.U4)] public int dmMediaType;
            [MarshalAs(UnmanagedType.U4)] public int dmDitherType;
            [MarshalAs(UnmanagedType.U4)] public int dmReserved1;
            [MarshalAs(UnmanagedType.U4)] public int dmReserved2;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct PRINTER_INFO_9
        {
            public IntPtr pDevMode;
        }

        [DllImport("winspool.Drv", EntryPoint = "AddFormW", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = true,
          CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool AddForm(IntPtr phPrinter, [MarshalAs(UnmanagedType.I4)] int level, ref FormInfo1 form);

        /*    This method is not used
          [DllImport("winspool.Drv", EntryPoint="SetForm", SetLastError=true,
            CharSet=CharSet.Unicode, ExactSpelling=false,
            CallingConvention=CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
          internal static extern bool SetForm(IntPtr phPrinter, string paperName,
           [MarshalAs(UnmanagedType.I4)] int level, ref FormInfo1 form);
        */
        [DllImport("winspool.Drv", EntryPoint = "DeleteForm", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = false,
            CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool DeleteForm(IntPtr phPrinter, [MarshalAs(UnmanagedType.LPTStr)] string pName);

        [DllImport("kernel32.dll", EntryPoint = "GetLastError", SetLastError = false, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern Int32 GetLastError();

        [DllImport("GDI32.dll", EntryPoint = "CreateDC", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = false,
            CallingConvention = CallingConvention.StdCall),
        SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern IntPtr CreateDC([MarshalAs(UnmanagedType.LPTStr)] string pDrive, [MarshalAs(UnmanagedType.LPTStr)]
        string pName, [MarshalAs(UnmanagedType.LPTStr)] string pOutput, ref StructDevMode pDevMode);

        [DllImport("GDI32.dll", EntryPoint = "ResetDC", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = false,
          CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern IntPtr ResetDC(IntPtr hDC, ref StructDevMode pDevMode);

        [DllImport("GDI32.dll", EntryPoint = "DeleteDC", SetLastError = true, CharSet = CharSet.Unicode, ExactSpelling = false,
          CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool DeleteDC(IntPtr hDC);

        [DllImport("winspool.Drv", EntryPoint = "SetPrinterA", SetLastError = true, CharSet = CharSet.Auto, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall), SuppressUnmanagedCodeSecurityAttribute()]
        internal static extern bool SetPrinter(
            IntPtr hPrinter,
            [MarshalAs(UnmanagedType.I4)] int level,
            IntPtr pPrinter, [MarshalAs(UnmanagedType.I4)]
            int command);

        /*
         LONG DocumentProperties(
           HWND hWnd,               // handle to parent window
           HANDLE hPrinter,         // handle to printer object
           LPTSTR pDeviceName,      // device name
           PDEVMODE pDevModeOutput, // modified device mode
           PDEVMODE pDevModeInput,  // original device mode
           DWORD fMode              // mode options
           );
         */
        [DllImport("winspool.Drv", EntryPoint = "DocumentPropertiesA", SetLastError = true, ExactSpelling = true,
            CallingConvention = CallingConvention.StdCall)]
        public static extern int DocumentProperties(
           IntPtr hwnd,
           IntPtr hPrinter,
           [MarshalAs(UnmanagedType.LPStr)] string pDeviceName /* changed from String to string */,
           IntPtr pDevModeOutput,
           IntPtr pDevModeInput,
           int fMode
           );

        [DllImport("winspool.Drv", EntryPoint = "GetPrinterA", SetLastError = true,
        ExactSpelling = true, CallingConvention = CallingConvention.StdCall)]
        public static extern bool GetPrinter(
           IntPtr hPrinter,
           int dwLevel /* changed type from Int32 */,
           IntPtr pPrinter,
           int dwBuf /* chagned from Int32*/,
           out int dwNeeded /* changed from Int32*/
           );

        // SendMessageTimeout tools
        [Flags]
        public enum SendMessageTimeoutFlags : uint
        {
            SMTO_NORMAL = 0x0000,
            SMTO_BLOCK = 0x0001,
            SMTO_ABORTIFHUNG = 0x0002,
            SMTO_NOTIMEOUTIFNOTHUNG = 0x0008
        }
        const int WM_SETTINGCHANGE = 0x001A;
        const int HWND_BROADCAST = 0xffff;

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessageTimeout(
           IntPtr windowHandle,
           uint Msg,
           IntPtr wParam,
           IntPtr lParam,
           SendMessageTimeoutFlags flags,
           uint timeout,
           out IntPtr result
           );

        #endregion

        public CustomPrintPaperSize() { }

        /// <summary>
        /// 将自定义纸张大小添加到指定打印机
        /// </summary>
        /// <param name="printerName">打印机名称</param>
        /// <param name="paperName">纸张名称</param>
        /// <param name="widthMm">纸张宽度，单位mm</param>
        /// <param name="heightMm">纸张高度，单位mm</param>
        /// <param name="leftMm">左边距，单位mm</param>
        /// <param name="topMm">顶边距，单位mm</param>
        /// <returns></returns>
        public static string AddCustomPaperSize(string printerName, string paperName, float widthMm, float heightMm, float leftMm = 0, float topMm = 0)
        {
            // 添加自定义纸张大小的代码对于Windows NT是不同的
            if (PlatformID.Win32NT == Environment.OSVersion.Platform)
            {
                const int PRINTER_ACCESS_USE = 0x00000008;
                const int PRINTER_ACCESS_ADMINISTER = 0x00000004;

                const int FORM_PRINTER = 0x00000002;

                StructPrinterDefaults defaults = new StructPrinterDefaults
                {
                    pDatatype = null,
                    pDevMode = IntPtr.Zero,
                    DesiredAccess = PRINTER_ACCESS_ADMINISTER | PRINTER_ACCESS_USE
                };

                // 打开打印机
                if (OpenPrinter(printerName, out IntPtr hPrinter, ref defaults))
                {
                    try
                    {
                        // 如果表单已存在，请删除该表单
                        DeleteForm(hPrinter, paperName);

                        // 创建并初始化FORM_INFO_1结构
                        FormInfo1 formInfo = new FormInfo1
                        {
                            Flags = 0,
                            pName = paperName
                        };

                        // 所有尺寸均为毫米的千分之一
                        formInfo.Size.width = (int)(widthMm * 1000.0);
                        formInfo.Size.height = (int)(heightMm * 1000.0);

                        formInfo.ImageableArea.left = (int)(leftMm * 1000.0);
                        formInfo.ImageableArea.top = (int)(topMm * 1000.0);
                        formInfo.ImageableArea.right = formInfo.Size.width;
                        formInfo.ImageableArea.bottom = formInfo.Size.height;

                        if (!AddForm(hPrinter, 1, ref formInfo))
                        {
                            StringBuilder strBuilder = new StringBuilder();
                            strBuilder.AppendFormat($"无法添加自定义纸张大小 {paperName} 到打印机 {printerName}, 系统错误号: {GetLastError()}");
                            return strBuilder.ToString();
                        }

                        // INIT
                        const int DM_OUT_BUFFER = 2;
                        const int DM_IN_BUFFER = 8;
                        StructDevMode devMode = new StructDevMode();
                        IntPtr hPrinterInfo;
                        PRINTER_INFO_9 printerInfo;
                        printerInfo.pDevMode = IntPtr.Zero;


                        // 获取DEV_MODE缓冲区的大小
                        int iDevModeSize = DocumentProperties(IntPtr.Zero, hPrinter, printerName, IntPtr.Zero, IntPtr.Zero, 0);

                        if (iDevModeSize < 0)
                            return "无法获取DEVMODE结构的大小.";

                        // 分配缓冲区
                        IntPtr hDevMode = Marshal.AllocCoTaskMem(iDevModeSize + 100);

                        // 获取指向DEV_MODE缓冲区的指针
                        int iRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, hDevMode, IntPtr.Zero, DM_OUT_BUFFER);

                        if (iRet < 0)
                            return "无法获取DEVMODE结构.";

                        // 填充DEV_MODE结构
                        devMode = (StructDevMode)Marshal.PtrToStructure(hDevMode, devMode.GetType());

                        // 设置表单名称字段以指示将修改此字段
                        devMode.dmFields = 0x10000; // DM_FORMNAME

                        // 设置表单名称
                        devMode.dmFormName = paperName;

                        // 将DEV_MODE结构放回指针中
                        Marshal.StructureToPtr(devMode, hDevMode, true);

                        // 将新的变更与旧的变更合并
                        iRet = DocumentProperties(IntPtr.Zero, hPrinter, printerName, printerInfo.pDevMode, printerInfo.pDevMode, DM_IN_BUFFER | DM_OUT_BUFFER);

                        if (iRet < 0)
                            return "无法设置此打印机的方向设置.";

                        // 获取打印机信息大小
                        GetPrinter(hPrinter, 9, IntPtr.Zero, 0, out int iPrinterInfoSize);
                        if (iPrinterInfoSize == 0)
                            return "Get Printer 失败。无法获取共享打印机_INFO_9结构所需的字节";

                        // 分配缓冲区
                        hPrinterInfo = Marshal.AllocCoTaskMem(iPrinterInfoSize + 100);

                        // 获取指向打印机信息缓冲区的指针
                        bool bSuccess = GetPrinter(hPrinter, 9, hPrinterInfo, iPrinterInfoSize, out _);

                        if (!bSuccess)
                            return "Get Printer 失败。无法获取共享的PRINTER_INFO_9结构";

                        // 填写打印机信息结构
                        printerInfo = (PRINTER_INFO_9)Marshal.PtrToStructure(hPrinterInfo, printerInfo.GetType());
                        printerInfo.pDevMode = hDevMode;

                        // 获取指向打印机信息结构的指针
                        Marshal.StructureToPtr(printerInfo, hPrinterInfo, true);

                        // 设置打印机设置
                        bSuccess = SetPrinter(hPrinter, 9, hPrinterInfo, 0);

                        if (!bSuccess)
                            return $"{Marshal.GetLastWin32Error()}，SetPrinter() 失败。无法设置打印机设置";

                        // 告诉所有打开的程序发生了此更改
                        _ = SendMessageTimeout(
                           new IntPtr(HWND_BROADCAST),
                           WM_SETTINGCHANGE,
                           IntPtr.Zero,
                           IntPtr.Zero,
                           SendMessageTimeoutFlags.SMTO_NORMAL,
                           1000,
                           out _);
                    }
                    finally
                    {
                        ClosePrinter(hPrinter);
                    }
                    return "设置成功";
                }
                else
                {
                    StringBuilder strBuilder = new StringBuilder();
                    strBuilder.AppendFormat($"无法打开 {printerName} 打印机，系统错误号: {GetLastError()}");
                    return strBuilder.ToString();
                }
            }
            else
            {
                StructDevMode pDevMode = new StructDevMode();
                IntPtr hDC = CreateDC(null, printerName, null, ref pDevMode);
                if (hDC != IntPtr.Zero)
                {
                    const long DM_PAPERSIZE = 0x00000002L;
                    const long DM_PAPERLENGTH = 0x00000004L;
                    const long DM_PAPERWIDTH = 0x00000008L;
                    pDevMode.dmFields = (int)(DM_PAPERSIZE | DM_PAPERWIDTH | DM_PAPERLENGTH);
                    pDevMode.dmPaperSize = 256;
                    pDevMode.dmPaperWidth = (short)(widthMm * 1000.0);
                    pDevMode.dmPaperLength = (short)(heightMm * 1000.0);
                    ResetDC(hDC, ref pDevMode);
                    DeleteDC(hDC);
                    return "";
                }
                return "";
            }
        }
    }
}
