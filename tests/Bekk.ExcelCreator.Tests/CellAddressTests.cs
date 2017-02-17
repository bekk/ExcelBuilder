using Bekk.ExcelCreator.Entities;
using NUnit.Framework;

namespace Bekk.ExcelCreator.Tests
{
    public class CellAddressTests
    {
        [TestCase(0u,0u, ExpectedResult = "A1")]
        [TestCase(1u,0u, ExpectedResult = "A2")]
        [TestCase(0u,1u, ExpectedResult = "B1")]
        [TestCase(1u,25u, ExpectedResult = "Z2")]
        [TestCase(1u,26u, ExpectedResult = "AA2")]
        [TestCase(1u,27u, ExpectedResult = "AB2")]
        [TestCase(1u,82u, ExpectedResult = "CE2")]
        [TestCase(13u, 16383u, ExpectedResult = "XFD14")]
        public string ToString_WithValidValues_ReturnsCorrectAddress(uint row, uint col)
        {
            var target = new CellAddress(row, col);
            return target.ToString();
        }
    }
}