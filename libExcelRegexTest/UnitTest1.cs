using Microsoft.VisualStudio.TestTools.UnitTesting;

using MyFunctions = libExcelRegex.MyFunctions;

namespace libExcelRegexTest
{
    [TestClass]
    public class UnitTest1
    {
        
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
    }
}
