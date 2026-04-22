using System.Text;
using System.Windows;
using WppQueuePoc.Abstractions;
using WppQueuePoc.Models;
using WppQueuePoc.Services;

namespace WppQueuePoc.App
{
    public partial class MainWindow : System.Windows.Window
    {
        // PrinterPolicyEnforcer instance and state
        private PrinterPolicyEnforcer? _policyEnforcer;
        private bool _isPolicyMonitorRunning = false;

        private readonly IWppStatusProvider _wppStatusProvider;
        private readonly IPrintSpoolerService _printSpoolerService;
        private readonly IPrintTicketService _printTicketService;
        private string? _currentQueueName;
        private QueueInfo? _currentQueueInfo = null;

        public MainWindow()
        {
            InitializeComponent();
            _wppStatusProvider = new WppRegistryService();
            _printSpoolerService = new PrintSpoolerService(_wppStatusProvider);
            _printTicketService = new PrintTicketService();
            PrintProcessorTextBox.Text = "WinPrint";
            DataTypeTextBox.Text = "RAW";
            SetCurrentQueue(null);

            InitializePolicyEnforcer();
        }

        private void InitializePolicyEnforcer()
        {
            var policy = new PrinterPolicyEnforcer.Policy
            {
                EnforceDuplex = true, RequiredDuplexValue = "TwoSidedLongEdge",
                EnforceColor = true, RequiredColorValue = "Monochrome",
                EnforceOrientation = true, RequiredOrientationValue = "Portrait"
            };
            _policyEnforcer = new PrinterPolicyEnforcer(policy, _printTicketService);
            _policyEnforcer.StatusChanged += (s, msg) => Dispatcher.Invoke(() => AppendOutput($"[Policy] {msg}"));
            _policyEnforcer.EnforcementLog += (s, log) => Dispatcher.Invoke(() => AppendOutput($"[Policy] {log}"));
            _policyEnforcer.Error += (s, ex) => Dispatcher.Invoke(() => AppendOutput($"[Policy][ERROR] {ex.Message}"));
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

        /// <summary>
        /// Runs a command on a background thread to keep UI responsive, while standardizing status/output updates.
        /// On success, appends result and invokes an optional callback; on failure, logs exception details.
        /// </summary>
        private async Task ExecuteAsync(string commandName, Func<string> action, Action? onSuccess = null)
        {
            ClearOutput();
            StatusTextBlock.Text = $"Running: {commandName}";
            try
            {
                var result = await Task.Run(action);
                onSuccess?.Invoke();
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

        private static string? NormalizeOptionalToNull(string s) => string.IsNullOrWhiteSpace(s) ? null : s.Trim();

        private static string ResolveQueueName(string queueNameFromField, string? currentQueueName)
        {
            return !string.IsNullOrWhiteSpace(queueNameFromField) ? queueNameFromField.Trim() : (currentQueueName ?? "");
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
            _currentQueueInfo = null;
            SetCurrentQueue(null);
            AppendOutput("Ready to create new print queue.");
        }

        private async void OnWppStatusClick(object sender, RoutedEventArgs e)
        {
            await ExecuteAsync("wpp-status", () => _wppStatusProvider.GetWppStatus().ToString());
        }

        private async void OnListQueuesClick(object sender, RoutedEventArgs e)
        {
            await ExecuteAsync("list-queues", () =>
            {
                var queues = _printSpoolerService.ListQueues();
                var output = new StringBuilder();
                output.AppendLine("Installed queues:");
                foreach (var queue in queues)
                {
                    output.AppendLine($"- Printer: {queue.Name} (Driver: {queue.DriverName} / Port: {queue.PortName})");
                }

                return output.ToString();
            });
        }

        private async void OnListPortsClick(object sender, RoutedEventArgs e)
        {
            await ExecuteAsync("list-ports", () =>
            {
                var ports = _printSpoolerService.ListPorts();
                var output = new StringBuilder();
                output.AppendLine("Ports:");
                foreach (var port in ports)
                {
                    output.AppendLine($"- {port}");
                }

                return output.ToString();
            });
        }

        private async void OnListDriversClick(object sender, RoutedEventArgs e)
        {
            await ExecuteAsync("list-drivers", () =>
            {
                var drivers = _printSpoolerService.ListDrivers();
                var output = new StringBuilder();
                output.AppendLine("Drivers:");
                foreach (var driver in drivers)
                {
                    output.AppendLine($"- {driver}");
                }

                return output.ToString();
            });
        }

        private async void OnListProcessorsClick(object sender, RoutedEventArgs e)
        {
            await ExecuteAsync("list-processors", () =>
            {
                var processors = _printSpoolerService.ListPrintProcessors();
                var output = new StringBuilder();
                output.AppendLine("Print processors:");
                foreach (var processor in processors)
                {
                    output.AppendLine($"- {processor}");
                }

                return output.ToString();
            });
        }

        private async void OnListDataTypesClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var printProcessorName = PrintProcessorTextBox.Text;
            if (string.IsNullOrWhiteSpace(printProcessorName))
            {
                ShowValidationWarning("list-data-types", "Please specify the print processor name to see data types.");
                return;
            }

            await ExecuteAsync("list-data-types", () =>
            {
                var dataTypes = _printSpoolerService.ListDataTypes(printProcessorName);
                var output = new StringBuilder();
                output.AppendLine($"Data types for {printProcessorName}:");
                foreach (var dataType in dataTypes)
                {
                    output.AppendLine($"- {dataType}");
                }

                return output.ToString();
            });
        }

        private async void OnAddWsdPortClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var port = PortNameTextBox.Text;
            if (string.IsNullOrWhiteSpace(port))
            {
                ShowValidationWarning("add-wsd-port", "Please enter the WSD port name.");
                return;
            }

            await ExecuteAsync("add-wsd-port", () =>
            {
                _printSpoolerService.AddWsdPort(port);
                return $"WSD port '{port}' added.";
            });
        }

        private async void OnCreateQueueClick(object sender, RoutedEventArgs e)
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

            await ExecuteAsync(
                "create-queue",
                () =>
                {
                _printSpoolerService.CreateQueue(queueName, driver, port, processor, dataType, comment, location);
                    return $"Queue '{queueName}' created.";
                },
                () => SetCurrentQueue(queueName));
        }

        private async void OnUpdateQueueClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var queueNameRaw = QueueNameTextBox.Text;
            var currentQueueName = _currentQueueName;
            var newQueueName = NormalizeOptionalToNull(QueueNameTextBox.Text);
            var newDriverName = NormalizeOptionalToNull(DriverNameTextBox.Text);
            var newPortName = NormalizeOptionalToNull(PortNameTextBox.Text);
            var comment = NormalizeOptionalToNull(CommentTextBox.Text);
            var location = NormalizeOptionalToNull(LocationTextBox.Text);
            string queueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currentQueueName, "update-queue", out queueName))
                return;

            var hasAnyUpdate =
                newQueueName is not null ||
                newDriverName is not null ||
                newPortName is not null ||
                comment is not null ||
                location is not null;

            if (!hasAnyUpdate)
            {
                ShowValidationWarning("update-queue", "Please provide at least one field to update: Queue Name, Driver Name, Port Name, Comment, or Location.");
                return;
            }

            await ExecuteAsync(
                "update-queue",
                () =>
                {
                _printSpoolerService.UpdateQueue(queueName, newQueueName, newDriverName, newPortName, comment, location);
                    return $"Queue '{queueName}' updated.";
                },
                () => SetCurrentQueue(string.IsNullOrWhiteSpace(newQueueName) ? queueName : newQueueName));
        }

        private async void OnDeleteQueueClick(object sender, RoutedEventArgs e)
        {
            ClearOutput();
            var queueNameRaw = QueueNameTextBox.Text;
            var currentQueueName = _currentQueueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currentQueueName, "delete-queue", out var queueName))
                return;

            await ExecuteAsync(
                "delete-queue",
                () =>
                {
                _printSpoolerService.DeleteQueue(queueName);
                    return $"Queue '{queueName}' deleted.";
                },
                () => SetCurrentQueue(null));
        }

        private async void OnInspectQueueClick(object sender, RoutedEventArgs e)
        {
            var queueNameRaw = QueueNameTextBox.Text;
            var currentQueueName = _currentQueueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currentQueueName, "inspect", out var queueName))
                return;

            QueueInfo? queueInfo = null;
            await ExecuteAsync("inspect", () =>
            {
                var result = _printSpoolerService.InspectQueue(queueName);
                queueInfo = _printSpoolerService
                    .ListQueues()
                    .FirstOrDefault(q => string.Equals(q.Name, queueName, StringComparison.OrdinalIgnoreCase));

                var sb = new StringBuilder();
                sb.AppendLine($"[Inspect] Queue: {queueName}");
                sb.AppendLine($"  - Port: {result.PortName}");
                if (queueInfo is not null)
                {
                    sb.AppendLine($"  - Driver: {queueInfo.DriverName}");
                    sb.AppendLine($"  - Comment: {queueInfo.Comment}");
                    sb.AppendLine($"  - Location: {queueInfo.Location}");
                }
                sb.AppendLine($"  - GlobalWpp: {result.GlobalWppStatus}");
                sb.AppendLine($"  - Classification: {result.Classification}");
                sb.AppendLine($"  - Details: {result.Details}");
                _currentQueueInfo = queueInfo;

                return sb.ToString();
            },
            () =>
            {
                if (queueInfo is null)
                    return;

                QueueNameTextBox.Text = queueInfo.Name;
                DriverNameTextBox.Text = queueInfo.DriverName;
                PortNameTextBox.Text = queueInfo.PortName;
                CommentTextBox.Text = queueInfo.Comment;
                LocationTextBox.Text = queueInfo.Location;
                SetCurrentQueue(queueInfo.Name);
            });
        }

        private async void OnTicketInfoClick(object sender, RoutedEventArgs e)
        {
            var queueNameRaw = QueueNameTextBox.Text;
            var currentQueueName = _currentQueueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currentQueueName, "ticket-info-default", out var queueName))
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
            var currentQueueName = _currentQueueName;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currentQueueName, "ticket-info-user", out var queueName))
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
            var currentQueueName = _currentQueueName;
            var ticketDuplexing = TicketDuplexingTextBox.Text;
            var ticketOutputColor = TicketOutputColorTextBox.Text;
            var ticketOrientation = TicketOrientationTextBox.Text;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currentQueueName, "ticket-update-default", out var queueName))
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
            var currentQueueName = _currentQueueName;
            var ticketDuplexing = TicketDuplexingTextBox.Text;
            var ticketOutputColor = TicketOutputColorTextBox.Text;
            var ticketOrientation = TicketOrientationTextBox.Text;
            if (!TryGetQueueNameOrWarn(queueNameRaw, currentQueueName, "ticket-update-user", out var queueName))
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
        private void OnTogglePolicyMonitorClick(object sender, RoutedEventArgs e)
        {
            if (_policyEnforcer == null) return;

            if (_isPolicyMonitorRunning)
            {
                _policyEnforcer.Stop();
                _isPolicyMonitorRunning = false;
                PolicyMonitorButton.Content = "Start Policy Monitor";
                AppendOutput("[Policy] Policy monitor stopped.");
            }
            else
            {
                ClearOutput();
                _policyEnforcer.Start(_currentQueueInfo?.Name);
                _isPolicyMonitorRunning = true;
                PolicyMonitorButton.Content = "Stop Policy Monitor";
                AppendOutput("[Policy] Policy monitor started.");
            }
        }
    }
}
