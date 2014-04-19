using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System;
using System.Collections;
using System.ComponentModel;
using System.Runtime.InteropServices;

namespace libExcelRegex
{
    public static class Win32Interop
    {
        [DllImport("dnsapi", EntryPoint = "DnsQuery_W", CharSet = CharSet.Unicode, SetLastError = true, ExactSpelling = true)]
        public static extern int DnsQuery([MarshalAs(UnmanagedType.VBByRefStr)] ref string pszName, DnsQueryTypes wType, DnsQueryOptions options, int aipServers, ref IntPtr ppQueryResults, int pReserved);

        [DllImport("dnsapi", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern void DnsRecordListFree(IntPtr pRecordList, int FreeType);

        public enum DnsQueryOptions
        {
            DNS_QUERY_ACCEPT_TRUNCATED_RESPONSE = 1,
            DNS_QUERY_BYPASS_CACHE = 8,
            DNS_QUERY_DONT_RESET_TTL_VALUES = 0x100000,
            DNS_QUERY_NO_HOSTS_FILE = 0x40,
            DNS_QUERY_NO_LOCAL_NAME = 0x20,
            DNS_QUERY_NO_NETBT = 0x80,
            DNS_QUERY_NO_RECURSION = 4,
            DNS_QUERY_NO_WIRE_QUERY = 0x10,
            DNS_QUERY_RESERVED = -16777216,
            DNS_QUERY_RETURN_MESSAGE = 0x200,
            DNS_QUERY_STANDARD = 0,
            DNS_QUERY_TREAT_AS_FQDN = 0x1000,
            DNS_QUERY_USE_TCP_ONLY = 2,
            DNS_QUERY_WIRE_ONLY = 0x100
        }

        public enum DnsQueryTypes
        {
            DNS_TYPE_A = 1,
            DNS_TYPE_NS = 2,
            DNS_TYPE_CNAME = 5,
            DNS_TYPE_SOA = 6,
            DNS_TYPE_PTR = 12,
            DNS_TYPE_HINFO = 13,
            DNS_TYPE_MX = 15,
            DNS_TYPE_TXT = 16,
            DNS_TYPE_AAAA = 28
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MXRecord
        {
            public IntPtr pNext;
            public string pName;
            public short wType;
            public short wDataLength;
            public int flags;
            public int dwTtl;
            public int dwReserved;
            public IntPtr pNameExchange;
            public short wPreference;
            public short Pad;
        }
    }
}
