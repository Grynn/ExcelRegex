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
    }
}
