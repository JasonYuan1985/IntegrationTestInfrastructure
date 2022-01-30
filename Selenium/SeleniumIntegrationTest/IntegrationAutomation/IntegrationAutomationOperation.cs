using IntegrationAutomation.Selenium;
using System.Data;

namespace IntegrationAutomation
{
    public class IntegrationAutomationOperation
    {
        OperationInterface SeleniumOperator {get;set;}
        public IntegrationAutomationOperation(OperationInterface seleniumOperator)
        {
            SeleniumOperator = seleniumOperator;
        }
        public List<ActionRowEntity> GetActionRowEntities(DataTable dateTable)
        {
            var result = new List<ActionRowEntity>();
            foreach (DataRow rowEntity in dateTable.Rows)
            {
                var actionRowEntity = new ActionRowEntity();
                actionRowEntity.ScenarioName = rowEntity["Scenario"].ToString();
                actionRowEntity.FrameworkName = rowEntity["Framework"].ToString();
                actionRowEntity.CommandName = rowEntity["Command Name"].ToString();
                actionRowEntity.CommandComponent = rowEntity["Command Component"].ToString();
                actionRowEntity.CommandValue = rowEntity["Command Value"].ToString();
                actionRowEntity.WaitTime = rowEntity["Wait Time"].ToString();
                result.Add(actionRowEntity);
            }

            return result;
        }

        private List<string> TemplateColumnNames
        {
            get
            {
                return new List<string>()
                {
                    "Scenario",
                    "Framework",
                    "Command Name",
                    "Command Component",
                    "Command Value",
                    "Wait Time"
                };
            }
        }

        public bool CheckTemplate(DataTable dataTable)
        {
            foreach (var title in TemplateColumnNames)
            {
                if (!dataTable.Columns.Contains(title))
                {
                    return false;
                }
            }
            return true;
        }

        public List<ActionRowEntity> PerformActions(List<ActionRowEntity> entities)
        {
            List<string> scenarioNames = entities.Select(c => c.ScenarioName).Distinct().ToList();
            foreach(var scenarioName in scenarioNames)
            {
                var scenarioEntities = entities.Where(c => c.ScenarioName == scenarioName);
                try
                {
                    foreach (ActionRowEntity entity in scenarioEntities)
                    {
                        if (entity.FrameworkName == "Selenium")
                        {
                            if (SeleniumOperator == null)
                            {
                                SeleniumOperator = new SeleniumHelper();
                            }
                            SeleniumOperator.PerforAction(entity);
                        }

                        entity.IsSuccess = true;
                    }
                }
                catch (Exception ex)
                {
                    scenarioEntities.ToList().ForEach(c=>c.IsSuccess = false);
                }
            }

            return entities;
        }
    }
}
