using IntegrationAutomation;
using Moq;
using System.Collections.Generic;
using System.Linq;
using Xunit;

namespace IntegrationAutomationTest
{
    public class IntegrationAutomationTester
    {
        [Fact]
        public void NoScenesWithSuccess()
        {
            List<ActionRowEntity> rows = new List<ActionRowEntity>();
            Mock<OperationInterface> seleniumOperator = new Mock<OperationInterface>();
            Mock<OperationInterface> excelOperator = new Mock<OperationInterface>();
            IntegrationAutomationOperation operation = new IntegrationAutomationOperation(seleniumOperator.Object, excelOperator.Object);
            var result = operation.PerformActions(rows);
            Assert.True(result.Where(c => c.IsSuccess).GroupBy(d => d.ScenarioName).Select(e => e.Key).ToList<string>().Count == 0);
        }

        [Fact]
        public void OneSceneWithAllSuccess()
        {
            List<ActionRowEntity> rows = new List<ActionRowEntity>();
            rows.Add(new ActionRowEntity() { ScenarioName = "1", FrameworkName = "Selenium", CommandComponent = "test" });

            Mock<OperationInterface> seleniumOperator = new Mock<OperationInterface>();
            Mock<OperationInterface> excelOperator = new Mock<OperationInterface>();
            IntegrationAutomationOperation operation = new IntegrationAutomationOperation(seleniumOperator.Object, excelOperator.Object);
            var result = operation.PerformActions(rows);
            Assert.True(result.Where(c=> c.IsSuccess).GroupBy(d => d.ScenarioName).Select(e=>e.Key).ToList<string>().Count == 1);
        }

        [Fact]
        public void TwoScenesWithAllSuccess()
        {
            List<ActionRowEntity> rows = new List<ActionRowEntity>();
            rows.Add(new ActionRowEntity() { ScenarioName = "1", FrameworkName = "Selenium", CommandComponent = "test" });
            rows.Add(new ActionRowEntity() { ScenarioName = "2", CommandComponent = "test" });

            Mock<OperationInterface> seleniumOperator = new Mock<OperationInterface>();
            Mock<OperationInterface> excelOperator = new Mock<OperationInterface>();
            IntegrationAutomationOperation operation = new IntegrationAutomationOperation(seleniumOperator.Object, excelOperator.Object);
            var result = operation.PerformActions(rows);
            Assert.True(result.Where(c => c.IsSuccess).GroupBy(d => d.ScenarioName).Select(e => e.Key).ToList<string>().Count == 2);
        }

        [Fact]
        public void TwoScenesWithOneSuccess()
        {
            List<ActionRowEntity> rows = new List<ActionRowEntity>();
            rows.Add(new ActionRowEntity() { ScenarioName = "1", FrameworkName = "Selenium", CommandComponent = "test" });
            rows.Add(new ActionRowEntity() { ScenarioName = "2", FrameworkName = "Selenium", CommandComponent = "test" });

            Mock<OperationInterface> seleniumOperator = new Mock<OperationInterface>();
            seleniumOperator.Setup(c => c.PerforAction(It.Is<ActionRowEntity>(d => d.ScenarioName == "2"))).Throws(new System.Exception());
            Mock<OperationInterface> excelOperator = new Mock<OperationInterface>();
            IntegrationAutomationOperation operation = new IntegrationAutomationOperation(seleniumOperator.Object, excelOperator.Object);
            var result = operation.PerformActions(rows);
            Assert.True(result.Where(c => c.IsSuccess).GroupBy(d => d.ScenarioName).Select(e => e.Key).ToList<string>().Count == 1);
        }
    }
}