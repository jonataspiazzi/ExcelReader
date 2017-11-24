using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace CoeCall.Test
{
    [TestClass]
    public class ExcelReaderTest
    {
        [TestMethod]
        public void ShouldGetValues()
        {
            var stream = GetType().Assembly.GetManifestResourceStream("ER.Test.allTypes.xlsx");

            using (var excel = new ExcelReader(stream))
            {
                Assert.AreEqual("geral", excel.GetValue("B", 2));
                Assert.AreEqual("12.4568", excel.GetValue("B", 3));
                Assert.AreEqual("45.25", excel.GetValue("B", 4));
                Assert.AreEqual("18.5599", excel.GetValue("B", 5));
                Assert.AreEqual("32408", excel.GetValue("B", 6));
                Assert.AreEqual("42952", excel.GetValue("B", 7));
                Assert.AreEqual("0.489594", excel.GetValue("B", 8));
                Assert.AreEqual("0.1845", excel.GetValue("B", 9));
                Assert.AreEqual("0.2", excel.GetValue("B", 10));
                Assert.AreEqual("10500000", excel.GetValue("B", 11));
                Assert.AreEqual("texto", excel.GetValue("B", 12));
                Assert.AreEqual("texto2", excel.GetValue("B", 13));
                Assert.AreEqual("texto", excel.GetValue("B", 14));
                Assert.AreEqual("texto2", excel.GetValue("B", 15));
                Assert.AreEqual("a", excel.GetValue("B", 16));
                Assert.AreEqual("1", excel.GetValue("B", 17));
            }
        }

        [TestMethod]
        public void ShouldGetValuesTyped()
        {
            var stream = GetType().Assembly.GetManifestResourceStream("ER.Test.allTypes.xlsx");

            using (var excel = new ExcelReader(stream))
            {
                Assert.AreEqual(null, excel.GetValue<int?>("B", 1));
                Assert.AreEqual("geral", excel.GetValue<string>("B", 2));
                Assert.AreEqual(12.4568f, excel.GetValue<float?>("B", 3));
                Assert.AreEqual(45.25, excel.GetValue<double>("B", 4));
                Assert.AreEqual(18.5599m, excel.GetValue<decimal>("B", 5));
                Assert.AreEqual(new DateTime(1988, 9, 22), excel.GetValue<DateTime?>("B", 6));
                Assert.AreEqual(new DateTime(2017, 8, 5), excel.GetValue<DateTime>("B", 7));
                Assert.AreEqual(new TimeSpan(0, 11, 45, 0, 922), excel.GetValue<TimeSpan>("B", 8));
                Assert.AreEqual(0.1845, excel.GetValue<double>("B", 9));
                Assert.AreEqual(0.2, excel.GetValue<double>("B", 10));
                Assert.AreEqual(10500000, excel.GetValue<int>("B", 11));
                Assert.AreEqual("texto", excel.GetValue<string>("B", 12));
                Assert.AreEqual("texto2", excel.GetValue<string>("B", 13));
                Assert.AreEqual("texto", excel.GetValue<string>("B", 14));
                Assert.AreEqual("texto2", excel.GetValue<string>("B", 15));
                Assert.AreEqual('a', excel.GetValue<char?>("B", 16));
                Assert.AreEqual(true, excel.GetValue<bool?>("B", 17));
            }
        }
    }
}
