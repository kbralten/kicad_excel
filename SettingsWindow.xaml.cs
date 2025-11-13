using System.Windows;

namespace KiCadExcelBridge
{
    public partial class SettingsWindow : Window
    {
        public string SymbolPrefix { get; set; } = "symbol";
        public string FootprintPrefix { get; set; } = "footprint";
        public int ServerPort { get; set; } = 8088;

        public SettingsWindow()
        {
            InitializeComponent();
            LoadSettings();
        }

        private void LoadSettings()
        {
            var config = ConfigurationManager.Load();
            SymbolPrefix = config.SymbolPrefix;
            FootprintPrefix = config.FootprintPrefix;
            ServerPort = config.ServerPort;

            SymbolPrefixTextBox.Text = SymbolPrefix;
            FootprintPrefixTextBox.Text = FootprintPrefix;
            ServerPortTextBox.Text = ServerPort.ToString();
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            if (!ValidateAndSave())
            {
                return;
            }

            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private bool ValidateAndSave()
        {
            // Validate port
            if (!int.TryParse(ServerPortTextBox.Text, out var port) || port < 1 || port > 65535)
            {
                MessageBox.Show("Server port must be a number between 1 and 65535.", "Invalid Port", MessageBoxButton.OK, MessageBoxImage.Warning);
                return false;
            }

            SymbolPrefix = SymbolPrefixTextBox.Text?.Trim() ?? "symbol";
            FootprintPrefix = FootprintPrefixTextBox.Text?.Trim() ?? "footprint";
            ServerPort = port;

            var config = ConfigurationManager.Load();
            config.SymbolPrefix = SymbolPrefix;
            config.FootprintPrefix = FootprintPrefix;
            config.ServerPort = ServerPort;
            ConfigurationManager.Save(config);

            return true;
        }
    }
}
