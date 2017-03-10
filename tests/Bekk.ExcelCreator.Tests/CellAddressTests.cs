using Bekk.ExcelBuilder.Contracts;
using Xunit;

namespace Bekk.ExcelBuilder.Tests
{
    public class CellAddressTests
    {
		[Theory]
		[InlineData(0u,0u, "A1")]
		[InlineData(1u, 0u, "A2")]
		[InlineData(0u, 1u, "B1")]
		[InlineData(1u, 25u, "Z2")]
		[InlineData(1u, 26u, "AA2")]
		[InlineData(1u, 27u, "AB2")]
		[InlineData(1u, 82u, "CE2")]
		[InlineData(13u, 16383u, "XFD14")]
        public void ToString_WithValidValues_ReturnsCorrectAddress(uint row, uint col, string expected)
        {
            var target = new CellAddress(row, col);
			Assert.Equal(expected, target.ToString());
        }

		[Theory]
		[InlineData("A1", 0u, 0u)]
		[InlineData("z134", 133u, 25u)]
		[InlineData("AA2", 1u, 26u)]
		[InlineData("XFD14", 13u, 16383u)]
        public void Parse_WithValidValues_ReturnsAddress(string input, uint expectedRow, uint expectedCol)
        {
            var result = CellAddress.Parse(input);
            Assert.Equal(expectedCol,result.Col);
            Assert.Equal(expectedRow, result.Row);
        }
    }
}