using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

namespace libExcelRegex
{
    public static class MXHelper
    {
        public static IList<string> GetMXRecords(string domain)
        {
            IntPtr ptr1 = IntPtr.Zero;
            IntPtr ptr2 = IntPtr.Zero;

            if (Environment.OSVersion.Platform != PlatformID.Win32NT)
            {
                throw new NotSupportedException();
            }

            var ret = new List<string>();
            
            int hr = Win32Interop.DnsQuery(
                    ref domain, 
                    Win32Interop.DnsQueryTypes.DNS_TYPE_MX,
                    Win32Interop.DnsQueryOptions.DNS_QUERY_STANDARD, 
                    0, 
                    ref ptr1, 
                    0);
            
            if (hr != 0)
            {
                throw new Win32Exception(hr);
            }

            for (ptr2 = ptr1; !ptr2.Equals(IntPtr.Zero); )
            {
                var mx = (Win32Interop.MXRecord) Marshal.PtrToStructure(ptr2, typeof(Win32Interop.MXRecord));

                if (mx.wType == 15)
                {
                    ret.Add(Marshal.PtrToStringAuto(mx.pNameExchange));
                }

                ptr2 = mx.pNext;
            }

            Win32Interop.DnsRecordListFree(ptr1, 0);

            return ret;
        }
    }
}
