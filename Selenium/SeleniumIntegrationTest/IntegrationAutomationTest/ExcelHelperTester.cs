using IntegrationAutomation;
using IntegrationAutomation.Excel;
using Moq;
using System.Data;
using Xunit;

namespace IntegrationAutomationTest
{
    public class ExcelHelperTester
    {
        [Fact]
        public void NoColumnsAndNoRowsReturnTrue()
        {
            DataTable originalDt = new DataTable();
            DataTable compareDt = new DataTable();

            Mock<ILog> logService = new Mock<ILog>();
            ExcelHelper operation = new ExcelHelper(logService.Object);
            var result = operation.CompareDataTable(originalDt, compareDt);
            Assert.True(result == true);
        }

        [Fact]
        public void SameColumnsAndDifferentRowsReturnFalse()
        {
            DataTable originalDt = new DataTable();
            AddColumns(originalDt, new string[] { "1", "2" });
            AddRows(originalDt, new string[] { "1,1", "2,2" });
            DataTable compareDt = new DataTable();
            AddColumns(compareDt, new string[] { "1", "2" });
            AddRows(compareDt, new string[] { "1,1", "2,2", "3,3" });

            Mock<ILog> logService = new Mock<ILog>();
            ExcelHelper operation = new ExcelHelper(logService.Object);
            var result = operation.CompareDataTable(originalDt, compareDt);
            Assert.True(result == false);
        }

        [Fact]
        public void DifferentColumnsAndSameRowsReturnFalse()
        {
            DataTable originalDt = new DataTable();
            AddColumns(originalDt, new string[] { "1", "2", "3" });
            AddRows(originalDt, new string[] { "1,1,1", "2,2,2" });
            DataTable compareDt = new DataTable();
            AddColumns(compareDt, new string[] { "1", "2" });
            AddRows(compareDt, new string[] { "1,1", "2,2" });

            Mock<ILog> logService = new Mock<ILog>();
            ExcelHelper operation = new ExcelHelper(logService.Object);
            var result = operation.CompareDataTable(originalDt, compareDt);
            Assert.True(result == false);
        }

        [Fact]
        public void DifferentColumnsAndDifferentRowsReturnFalse()
        {
            DataTable originalDt = new DataTable();
            AddColumns(originalDt, new string[] { "1", "2", "3" });
            AddRows(originalDt, new string[] { "1,1,1", "2,2,2" });
            DataTable compareDt = new DataTable();
            AddColumns(compareDt, new string[] { "1", "2" });
            AddRows(compareDt, new string[] { "1,1", "2,2","3,3" });

            Mock<ILog> logService = new Mock<ILog>();
            ExcelHelper operation = new ExcelHelper(logService.Object);
            var result = operation.CompareDataTable(originalDt, compareDt);
            Assert.True(result == false);
        }

        [Fact]
        public void SameColumnsAndSameRowsReturnTrue()
        {
            DataTable originalDt = new DataTable();
            AddColumns(originalDt, new string[] { "1", "2", "3" });
            AddRows(originalDt, new string[] { "1,1,1", "2,2,2" });
            DataTable compareDt = new DataTable();
            AddColumns(compareDt, new string[] { "1", "2", "3" });
            AddRows(compareDt, new string[] { "1,1,1", "2,2,2" });

            Mock<ILog> logService = new Mock<ILog>();
            ExcelHelper operation = new ExcelHelper(logService.Object);
            var result = operation.CompareDataTable(originalDt, compareDt);
            Assert.True(result == true);
        }

        private void AddColumns(DataTable dt, string[] columnNames)
        {
            if (columnNames != null)
            {
                foreach (var columnName in columnNames)
                {
                    dt.Columns.Add(columnName);
                }
            }
        }

        private void AddRows(DataTable dt, string[] rowDatas, char splitChar = ',')
        {
            if (rowDatas != null)
            {
                foreach (var rowData in rowDatas)
                {
                    dt.Rows.Add(rowData.Split(splitChar));
                }
            }
        }
    }
}
