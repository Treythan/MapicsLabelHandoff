using System.Runtime.InteropServices;
using System.Runtime.Versioning;
using System.Security;
using System.Text;

namespace MapicsLabelHandoff
{
    internal static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            ApplicationConfiguration.Initialize();

            StringBuilder sb = new StringBuilder();
            foreach (string arg in args) { sb.Append(arg); }
            string[] parsedArgs = sb.ToString().Split("||");

            switch (parsedArgs[0])
            {
                case "1":
                    Application.Run(new Form1(parsedArgs));
                    break;
                case "2":
                    Application.Run(new Form2(parsedArgs));
                    break;
            }
        }
    }

    public static class Marshall
    {
        internal const String OLEAUT32 = "oleaut32.dll";
        internal const String OLE32 = "ole32.dll";

        [System.Security.SecurityCritical]
        public static Object GetActiveObject(String progID)
        {
            Object obj = null;
            Guid clsid;
            try
            {
                CLSIDFromProgIDEx(progID, out clsid);
            }
            catch (Exception)
            {
                CLSIDFromProgID(progID, out clsid);
            }

            GetActiveObject(ref clsid, IntPtr.Zero, out obj);
            return obj;
        }

        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]
        private static extern void CLSIDFromProgIDEx([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        [DllImport(OLE32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]
        private static extern void CLSIDFromProgID([MarshalAs(UnmanagedType.LPWStr)] String progId, out Guid clsid);

        [DllImport(OLEAUT32, PreserveSig = false)]
        [ResourceExposure(ResourceScope.None)]
        [SuppressUnmanagedCodeSecurity]
        [System.Security.SecurityCritical]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr reserved, [MarshalAs(UnmanagedType.Interface)] out Object ppunk);

    }
}