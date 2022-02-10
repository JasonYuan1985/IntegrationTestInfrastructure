using IntegrationAutomation.Excel;
using IntegrationAutomation.Selenium;
using System.Data;

namespace IntegrationAutomation
{
    public class IntegrationAutomationOperation
    {
        OperationInterface _seleniumOperator { get; set; }
        OperationInterface _excelOperator { get; set; }
        OperationInterface _systemOperator { get; set; }
        public IntegrationAutomationOperation(OperationInterface seleniumOperator, OperationInterface excelOperator, OperationInterface systemOperator)
        {
            _seleniumOperator = seleniumOperator;
            _excelOperator = excelOperator;
            _systemOperator = systemOperator;
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
                actionRowEntity.CommandType = rowEntity["Command Type"].ToString();
                actionRowEntity.CommandComponent = rowEntity["Command Component"].ToString();
                actionRowEntity.CommandValue = rowEntity["Command Value"].ToString();
                actionRowEntity.WaitTime = rowEntity["Wait Time"].ToString();
                actionRowEntity.ActionDisabled = rowEntity["Action Disabled"].ToString() == "Y";
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
                    "Command Type",
                    "Command Component",
                    "Command Value",
                    "Wait Time",
                    "Action Disabled"
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
            foreach (var scenarioName in scenarioNames)
            {
                var scenarioEntities = entities.Where(c => c.ScenarioName == scenarioName && !c.ActionDisabled);
                try
                {
                    foreach (ActionRowEntity entity in scenarioEntities)
                    {
                        if (entity.FrameworkName == "Selenium")
                        {
                            if (_seleniumOperator == null)
                            {
                                _seleniumOperator = new SeleniumHelper(new TxtLogWriter());
                            }
                            _seleniumOperator.PerforAction(entity);
                        }
                        else if(entity.FrameworkName == "Excel")
                        {
                            if (_excelOperator == null)
                            {
                                _excelOperator = new ExcelHelper(new TxtLogWriter());
                            }
                            _excelOperator.PerforAction(entity);
                        }
                        else if(entity.FrameworkName == "System")
                        {
                            if (_systemOperator == null)
                            {
                                _systemOperator = new SystemHelper(new TxtLogWriter());
                            }
                            _systemOperator.PerforAction(entity);
                        }
                        entity.IsSuccess = true;
                    }
                }
                catch (Exception ex)
                {
                    scenarioEntities.ToList().ForEach(c => c.IsSuccess = false);
                }
            }

            return entities;
        }
    }
}
