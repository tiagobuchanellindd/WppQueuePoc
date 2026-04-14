using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Windows;
using WppQueuePoc.Abstractions;
using WppQueuePoc.Models;
using WppQueuePoc.Services;

namespace WppQueuePoc.App
{
    public partial class MainWindow : Window
    {
        private readonly IWppStatusProvider _wppStatusProvider;
        private readonly IPrintSpoolerService _printSpoolerService;
        private readonly IPrintTicketService _printTicketService;
        private string? _currentQueueName;
        public MainWindow()
        {
            InitializeComponent();
            _wppStatusProvider = new WppRegistryService();
            _printSpoolerService = new PrintSpoolerService(_wppStatusProvider);
            _printTicketService = new PrintTicketService();
            PrintProcessorTextBox.Text = "WinPrint";
            DataTypeTextBox.Text = "RAW";
            SetCurrentQueue(null);
        }

        private void ClearOutput()
        {
            OutputTextBox.Clear();
        }

        private void AppendOutput(string text)
        {
            OutputTextBox.AppendText($"{text}\n");
            OutputTextBox.ScrollToEnd();
        }

        private bool ShowValidationWarning(string commandName, string message)
        {
            AppendOutput(message);
            StatusTextBlock.Text = $"Validation warning ({commandName})";
            MessageBox.Show(this, message, "Validation", MessageBoxButton.OK, MessageBoxImage.Warning);
            return false;
        }

        private bool ValidateRequiredFields(string commandName, params (string FieldName, string? Value)[] fields)
        {
            var missingFields = fields
                .Where(field => string.IsNullOrWhiteSpace(field.Value))
                .Select(field => field.FieldName)
                .ToArray();

            if (missingFields.Length == 0)
                return true;

            return ShowValidationWarning(commandName, $"Please fill in required fields: {string.Join(", ", missingFields)}.");
        }

        private bool TryGetQueueNameOrWarn(string queueNameFromField, string? currentQueueName, string commandName, out string queueName)
        {
            queueName = ResolveQueueName(queueNameFromField, currentQueueName);
            if (string.IsNullOrWhiteSpace(queueName))
                return ShowValidationWarning(commandName, "Please specify the queue name in 'Queue Name' or select a current queue.");

            queueName = queueName.Trim();
            return true;
        }

        private bool ValidateTicketUpdateInput(string commandName, string duplexing, string outputColor, string orientation)
        {
            var hasAnyValue =
                !string.IsNullOrWhiteSpace(duplexing) ||
                !string.IsNullOrWhiteSpace(outputColor) ||
                !string.IsNullOrWhiteSpace(orientation);

            if (hasAnyValue)
                return true;

            return ShowValidationWarning(commandName, "Please provide at least one ticket field: Duplexing, OutputColor, or Orientation.");
        }

        private async Task ExecuteAsync(string commandName, Func<string> action)
        {
            ClearOutput();
            StatusTextBlock.Text = $"Running: {commandName}";
            try
            {
                var result = await Task.Run(action);
                AppendOutput($"> {commandName}");
                AppendOutput(result);
                StatusTextBlock.Text = $"Ready ({commandName})";
            }
            catch (Exception ex)
            {
                AppendOutput($"> {commandName} [ERROR]");
                AppendOutput(ex.ToString());
                StatusTextBlock.Text = $"Error ({commandName})";
            }
        }

        private static string NormalizeOptional(string s) => string.IsNullOrWhiteSpace(s) ? "" : s.Trim();

        private static string ResolveQueueName(string nomeDoCampo, string? nomeAtual)
        {
            return !string.IsNullOrWhiteSpace(nomeDoCampo) ? nomeDoCampo.Trim() : (nomeAtual ?? "");
        }

        private void SetCurrentQueue(string? queueName)
        {
            _currentQueueName = queueName;
            if (string.IsNullOrWhiteSpace(queueName))
                CurrentQueueTextBlock.Text = "Current queue: (none)";
            else
                CurrentQueueTextBlock.Text = $"Current queue: {queueName}";
        }

        private void OnNewQueueClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            QueueNameTextBox.Text = "";
            DriverNameTextBox.Text = "";
            PortNameTextBox.Text = "";
            PrintProcessorTextBox.Text = "WinPrint";
            DataTypeTextBox.Text = "RAW";
            CommentTextBox.Text = "";
            LocationTextBox.Text = "";
            TicketDuplexingTextBox.Text = "";
            TicketOutputColorTextBox.Text = "";
            TicketOrientationTextBox.Text = "";
            SetCurrentQueue(null);
            AppendOutput("Ready to create new print queue.");
        }

        private void OnWppStatusClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var status = _wppStatusProvider.GetWppStatus();
            AppendOutput(status.ToString());
        }

        private void OnListQueuesClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var filas = _printSpoolerService.ListQueues();
            AppendOutput("Installed queues:");
            foreach (var fila in filas)
            {
                AppendOutput($"- {fila.Name} ({fila.DriverName} @ {fila.PortName})");
            }
        }

        private void OnListPortsClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var ports = _printSpoolerService.ListPorts();
            AppendOutput("Ports:");
            foreach (var port in ports)
            {
                AppendOutput($"- {port}");
            }
        }

        private void OnListDriversClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var drivers = _printSpoolerService.ListDrivers();
            AppendOutput("Drivers:");
            foreach (var driver in drivers)
            {
                AppendOutput($"- {driver}");
            }
        }

        private void OnListProcessorsClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var processors = _printSpoolerService.ListPrintProcessors();
            AppendOutput("Print processors:");
            foreach (var proc in processors)
            {
                AppendOutput($"- {proc}");
            }
        }

        private void OnListDataTypesClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var procName = PrintProcessorTextBox.Text;
            if (string.IsNullOrWhiteSpace(procName))
            {
                ShowValidationWarning("list-data-types", "Please specify the print processor name to see data types.");
                return;
            }
            var tipos = _printSpoolerService.ListDataTypes(procName);
            AppendOutput($"Data types for {procName}:");
            foreach (var t in tipos)
            {
                AppendOutput($"- {t}");
            }
        }

        private void OnAddWsdPortClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var port = PortNameTextBox.Text;
            if (string.IsNullOrWhiteSpace(port))
            {
                ShowValidationWarning("add-wsd-port", "Please enter the WSD port name.");
                return;
            }
            try
            {
                _printSpoolerService.AddWsdPort(port);
                AppendOutput($"WSD port '{port}' added.");
            }
            catch (Exception ex)
            {
                AppendOutput($"Error adding WSD port:\n{ex.Message}");
                StatusTextBlock.Text = $"Error (Add WSD Port)";
            }
        }

        private void OnCreateQueueClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var queueName = QueueNameTextBox.Text;
            var driver = DriverNameTextBox.Text;
            var port = PortNameTextBox.Text;
            var processor = PrintProcessorTextBox.Text;
            var dataType = DataTypeTextBox.Text;
            var comment = CommentTextBox.Text;
            var location = LocationTextBox.Text;
            if (!ValidateRequiredFields("create-queue",
                    ("Queue Name", queueName),
                    ("Driver Name", driver),
                    ("Port Name", port),
                    ("Print Processor", processor),
                    ("Data Type", dataType)))
                return;

            _printSpoolerService.CreateQueue(queueName, driver, port, processor, dataType, comment, location);
            AppendOutput($"Queue '{queueName}' created.");
            SetCurrentQueue(queueName);
        }

        private void OnUpdateQueueClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var queueNameRaw = QueueNameTextBox.Text;
            var currQueue = _currentQueueName;
            var newQueueName = NormalizeOptional(QueueNameTextBox.Text);
            var newDriverName = NormalizeOptional(DriverNameTextBox.Text);
            var newPortName = NormalizeOptional(PortNameTextBox.Text);
            var comment = NormalizeOptional(CommentTextBox.Text);
            var location = NormalizeOptional(LocationTextBox.Text);
            string queueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currQueue, "update-queue", out queueName))
                return;

            var hasAnyUpdate =
                !string.IsNullOrWhiteSpace(newQueueName) ||
                !string.IsNullOrWhiteSpace(newDriverName) ||
                !string.IsNullOrWhiteSpace(newPortName) ||
                !string.IsNullOrWhiteSpace(comment) ||
                !string.IsNullOrWhiteSpace(location);

            if (!hasAnyUpdate)
            {
                ShowValidationWarning("update-queue", "Please provide at least one field to update: Queue Name, Driver Name, Port Name, Comment, or Location.");
                return;
            }

            _printSpoolerService.UpdateQueue(queueName, newQueueName, newDriverName, newPortName, comment, location);
            AppendOutput($"Queue '{queueName}' updated.");
            SetCurrentQueue(string.IsNullOrWhiteSpace(newQueueName) ? queueName : newQueueName);
        }

        private void OnDeleteQueueClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var queueNameRaw = QueueNameTextBox.Text;
            var currQueue = _currentQueueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currQueue, "delete-queue", out var queueName))
                return;

            _printSpoolerService.DeleteQueue(queueName);
            AppendOutput($"Queue '{queueName}' deleted.");
            SetCurrentQueue(null);
        }

        private async void OnInspectQueueClick(object sender, RoutedEventArgs e)
        {
            var queueNameRaw = QueueNameTextBox.Text;
            var currQueue = _currentQueueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currQueue, "inspect", out var queueName))
                return;

            await ExecuteAsync("inspect", () =>
            {
                var result = _printSpoolerService.InspectQueue(queueName);
                var sb = new StringBuilder();
                sb.AppendLine($"[Inspect] Queue: {queueName}");
                sb.AppendLine($"  - Port: {result.PortName}");
                sb.AppendLine($"  - GlobalWpp: {result.GlobalWppStatus}");
                sb.AppendLine($"  - Classification: {result.Classification}");
                sb.AppendLine($"  - Details: {result.Details}");
                return sb.ToString();
            });
        }

        private async void OnTicketInfoClick(object sender, RoutedEventArgs e)
        {
            var queueNameRaw = QueueNameTextBox.Text;
            var currQueue = _currentQueueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currQueue, "ticket-info-default", out var queueName))
                return;

            await ExecuteAsync("ticket-info (default)", () =>
            {
                var result = _printTicketService.GetDefaultTicketInfo(queueName);
                return FormatTicketInfoResult(result);
            });
        }

        private async void OnTicketUserInfoClick(object sender, RoutedEventArgs e)
        {
            var queueNameRaw = QueueNameTextBox.Text;
            var currQueue = _currentQueueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currQueue, "ticket-info-user", out var queueName))
                return;

            await ExecuteAsync("ticket-info (user)", () =>
            {
                var result = _printTicketService.GetUserTicketInfo(queueName);
                return FormatTicketInfoResult(result);
            });
        }

        private async void OnTicketDefaultUpdateClick(object sender, RoutedEventArgs e)
        {
            var queueNameRaw = QueueNameTextBox.Text;
            var currQueue = _currentQueueName;
            var ticketDuplexing = TicketDuplexingTextBox.Text;
            var ticketOutputColor = TicketOutputColorTextBox.Text;
            var ticketOrientation = TicketOrientationTextBox.Text;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currQueue, "ticket-update-default", out var queueName))
                return;

            if (!ValidateTicketUpdateInput("ticket-update-default", ticketDuplexing, ticketOutputColor, ticketOrientation))
                return;

            await ExecuteAsync("ticket-update-default", () =>
            {
                var request = new PrintTicketUpdateRequest(
                    NormalizeOptional(ticketDuplexing),
                    NormalizeOptional(ticketOutputColor),
                    NormalizeOptional(ticketOrientation)
                );
                var result = _printTicketService.UpdateDefaultTicket(queueName, request);
                return FormatTicketUpdateResult(result);
            });
        }

        private async void OnTicketUserUpdateClick(object sender, RoutedEventArgs e)
        {
            var queueNameRaw = QueueNameTextBox.Text;
            var currQueue = _currentQueueName;
            var ticketDuplexing = TicketDuplexingTextBox.Text;
            var ticketOutputColor = TicketOutputColorTextBox.Text;
            var ticketOrientation = TicketOrientationTextBox.Text;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currQueue, "ticket-update-user", out var queueName))
                return;

            if (!ValidateTicketUpdateInput("ticket-update-user", ticketDuplexing, ticketOutputColor, ticketOrientation))
                return;

            await ExecuteAsync("ticket-update-user", () =>
            {
                var request = new PrintTicketUpdateRequest(
                    NormalizeOptional(ticketDuplexing),
                    NormalizeOptional(ticketOutputColor),
                    NormalizeOptional(ticketOrientation)
                );
                var result = _printTicketService.UpdateUserTicket(queueName, request);
                return FormatTicketUpdateResult(result);
            });
        }

        private static string FormatTicketInfoResult(PrintTicketInfoResult result)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Ticket:");
            sb.AppendLine($"  Queue: {result.QueueName}");
            sb.AppendLine($"  Available: {result.Available}");
            sb.AppendLine($"  Details: {result.Details}");
            if (result.Attributes != null)
                foreach (var kv in result.Attributes)
                    sb.AppendLine($"    {kv.Key}: {kv.Value}");
            return sb.ToString();
        }

        private string FormatTicketUpdateResult(PrintTicketUpdateResult result)
        {
            var sb = new StringBuilder();
            if (result.Applied)
            {
                sb.AppendLine("Update completed successfully!");
                sb.AppendLine(FormatTicketInfoResult(new PrintTicketInfoResult(result.QueueName, result.Applied, result.Details, result.AppliedValues)));
            }
            else
            {
                if (result.Details.Contains("No changes needed"))
                    sb.AppendLine("No update required: values already set.");
                else
                    sb.AppendLine("Error updating ticket: " + result.Details);
            }
            return sb.ToString();
        }
    }
}
