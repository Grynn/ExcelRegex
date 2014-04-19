using System;
using ExcelDna.Integration;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
using System.IO;
using System.Text;
using System.Linq;
using System.Web;
using System.Net;
using System.Diagnostics;
using System.ComponentModel;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace libExcelRegex
{
    public class MyFunctions
    {
        //http://regexlib.com/Search.aspx?k=email&c=-1&m=5&ps=20 seems to be a good source of email validation regexes
        private static readonly Regex ReIsEmail = new Regex(
            @"^([a-zA-Z0-9_\-\.]+)@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.)|(([a-zA-Z0-9\-]+\.)+))([a-zA-Z]{2,4}|[0-9]{1,3})(\]?)$",
            RegexOptions.Compiled);

        [ExcelCommand(MenuName = "&TestOnRecalc", MenuText = "&Enable Beep")]
        public static void EnableBeep()
        {
            XlCall.Excel(XlCall.xlcOnRecalc, null, "Beep");
        }

        [ExcelCommand(MenuText = "Show Log Window")]
        public static void ShowLog()
        {
            ExcelDna.Logging.LogDisplay.Show();
        }

        [ExcelFunction(Name="DNS.Resolve")]
        public static string DNSResolve(
            [ExcelArgument(Description="Hostname or IP address to resolve")] string Hostname, 
            [ExcelArgument(Description="If a hostname resolves to multiple addresses, which address to return")] int EntryId = 0)
        {
            var ipEntry = Dns.GetHostEntry(Hostname);
            if (EntryId > ipEntry.AddressList.Length) 
            {
                throw new ArgumentOutOfRangeException("EntryId");
            }
            
            return ipEntry.AddressList[EntryId].ToString();   
        }

        //TODO: MX records should be sorted by weight
        [ExcelFunction(Name = "DNS.GetMX")]
        public static string DNSGetMX(
            [ExcelArgument(Description = "Domain for which to lookup MX records")] string Domain,
            [ExcelArgument(Description = "Return nth MX record")] int EntryId = 0)
        {
            var entries = MXHelper.GetMXRecords(Domain);
            if (EntryId > entries.Count())
            {
                throw new ArgumentOutOfRangeException("EntryId");
            }

            return entries[EntryId].ToString();
        }

        [ExcelFunction(Name = "DNS.ResolveAll")]
        public static object[] DNSResolveAll([ExcelArgument(Description = "Hostname or IP address to resolve")] string Hostname)
        {
            var ipEntry = Dns.GetHostEntry(Hostname);
            return ipEntry.AddressList.Select(x => x.ToString()).Cast<object>().ToArray();
        }

        [ExcelFunction(Name = "RegexExtract")]
        public static string RegexExtract(string Input, string Pattern, double GroupNum)
        {
            var m = Regex.Match(Input, Pattern);
            string ret = string.Empty;

            if (m.Success)
                ret = m.Groups[Convert.ToInt32(GroupNum)].Value;

            return ret;
        }

        [ExcelFunction(Name = "IsEmail")]
        public static bool IsEmail(string email)
        {
            return ReIsEmail.IsMatch(email);
        }

        [ExcelFunction(Name = "FileHash")]
        private string FileHash(string Filename)
        {
            MD5 md5 = MD5.Create();
            using (var fs = File.OpenRead(Filename))
            {
                return md5
                    .ComputeHash(fs)
                    .Aggregate(
                        new StringBuilder(), 
                        (sb, b) => sb.AppendFormat("{0:X2}", b),
                        x => x.ToString()
                    );
            }
        }

        [ExcelFunction(Name = "TimespanToMinutes")]
        public static double TimespanToMinutes(string Timespan)
        {
            var re = new Regex(@"(\d+)\s+(year|day|week|month|hour|minute)s?",
                                 RegexOptions.Singleline | RegexOptions.IgnoreCase);
            MatchCollection ms = re.Matches(Timespan);
            double sum = 0;

            foreach (Match m in ms)
            {
                if (m.Success)
                {
                    string unit = m.Groups[2].Value.ToLower();
                    double val = Convert.ToDouble(m.Groups[1].Value);

                    if (unit.StartsWith("year"))
                        sum += 365*24*60*val;
                    else if (unit.StartsWith("day"))
                        sum += 24*60*val;
                    else if (unit.StartsWith("week"))
                        sum += 7*24*60*val;
                    else if (unit.StartsWith("month"))
                        sum += 30*24*60*val;
                    else if (unit.StartsWith("hour"))
                        sum += 60*val;
                    else if (unit.StartsWith("minute"))
                        sum += val;
                    else
                        throw new Exception("unexpected " + unit);
                }
            }
            return sum;
        }

        [ExcelFunction(Name = "Shorten")]
        public static string Shorten(string url)
        {
            var u = new Uri(url);
            if (u.Host.ToLower() == "bit.ly")
            {
                return u.ToString();
            }
            const string a = "http://api.bitly.com/v3/shorten?login=grynn&apiKey=R_4fccca4db3839b3a2bf2bdfde1e3fc22&longUrl={0}&format=xml";
            var b = string.Format(a, HttpUtility.UrlEncode(url));
            dynamic xhr = GetXHR();

            xhr.open("GET", b, false);
            xhr.send();
            string res = xhr.responseXml.SelectSingleNode("response/status_code").text;
            if (res.Trim() == "200")
            {
                res = xhr.responseXml.SelectSingleNode("response/data/url").text;
                return res.Trim();
            }
            throw new Exception("Error: " + res);
        }

        public static object GetXHR()
        {
            return Activator.CreateInstance(Type.GetTypeFromProgID("MSXML2.XMLHTTP"));
        }

        
    }
}