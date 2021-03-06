namespace IntegrationAutomation
{
    public class ActionRowEntity
    {
        public string ScenarioName { get; set; }
        public string FrameworkName { get; set; }
        public string CommandName { get; set; }
        public string CommandType { get; set; }
        public string CommandComponent { get; set; }
        public string CommandValue { get; set; }
        public string WaitTime { get; set; }
        public bool ActionDisabled { get; set; }
        public bool IsSuccess { get; set; }

        public override string ToString()
        {
            return $"{ScenarioName}, {FrameworkName}, {CommandName}, {CommandType}, {CommandComponent}, {CommandValue}, {WaitTime}";
        }
    }
}
