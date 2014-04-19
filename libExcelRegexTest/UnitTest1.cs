using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using MyFunctions = libExcelRegex.MyFunctions;

namespace libExcelRegexTest
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestDnsResolve() 
        {
            var x = MyFunctions.DNSResolve("abc.lvh.me");
            Assert.AreEqual("127.0.0.1", x);

            try 
            { 
                x = MyFunctions.DNSResolve("abc.lvh.me", 2);
                Assert.Fail("Expected ArgumentOutOfRangeException");
            }
            catch (ArgumentOutOfRangeException) 
            {
                Assert.IsTrue(true);
            }
        }

        [TestMethod]
        public void TestRegexExtract()
        {
            var actual = MyFunctions.RegexExtract("Hello World", @"^\w+", 0);
            Assert.AreEqual("Hello", actual);
        }

        [TestMethod]
        public void TestRegexExtract1()
        {
            var actual = MyFunctions.RegexExtract("Hello World", @"\s+(\w+)", 1);
            Assert.AreEqual("World", actual);
        }

        [TestMethod]
        public void TestIsEmail()
        {
            var actual = MyFunctions.IsEmail("someone@somewhere.com");
            Assert.AreEqual(true, actual);
        }

        [TestMethod]
        public void TestIsEmail1()
        {
            var actual = MyFunctions.IsEmail("someone@somewhere..com");
            Assert.AreEqual(false, actual);
        }

        [TestMethod]
        public void TestIsEmail2()
        {
            var good = new string[]
                           {
                               "abattini_by@sancharnet.in",
                               "zoemg@yahoo.co.in",
                               "mirza_ca@yahoo.co.in",
                               "axvicky@yahoo.co.uk",
                               "ABCINDIA@VSNL.COM",
                               "abesde_79@yahoo.co.in",
                               "abs3hat@nse.co.in",
                               "nirmal.pandey@itc.in"
                           };
            
            foreach (var email in good)
                Assert.AreEqual(true, MyFunctions.IsEmail(email));
        }

        [TestMethod]
        public void TestIsEmail3()
        {
            var bad = new string[]
                           {
                               "ABEPA4585D",
                               "abhi_s_joshi",
                               "123@456.78910"
                           };
            
            foreach (var email in bad)
                Assert.AreEqual(false, MyFunctions.IsEmail(email));
        }

        [TestMethod]
        public void TestTimespanToMinutes()
        {
            var a = MyFunctions.TimespanToMinutes("3 minutes");
            Assert.AreEqual(3, a);

            a = MyFunctions.TimespanToMinutes("3 hours 2 minutes");
            Assert.AreEqual(182, a);
        }

    }
}
