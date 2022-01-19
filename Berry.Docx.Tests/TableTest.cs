using Microsoft.VisualStudio.TestTools.UnitTesting;
using Berry.Docx;
using Berry.Docx.Documents;

namespace Berry.Docx.Tests
{
    [TestClass]
    public class TableTest
    {
        [TestMethod]
        [DataRow(1)]
        [DataRow(2)]
        [DataRow(3)]
        public void Add_TableRow_ReturnSameRowCount(int row)
        {
            var doc = new Document("test.docx");
            Table table = new Table(doc, row, 1);
            Assert.AreEqual(row, table.RowCount);
            doc.Close();
        }
    }
}
