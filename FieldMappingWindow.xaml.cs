using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Windows;
using System.Windows.Input;

namespace KiCadExcelBridge
{
    public partial class FieldMappingWindow : Window, INotifyPropertyChanged
    {
        private readonly SheetMapping _sheetMapping;
        private readonly ObservableCollection<FieldMappingRowViewModel> _fieldMappings;
        private readonly ObservableCollection<ExcelColumnOption> _columns;
        private readonly List<Dictionary<int, string>> _previewRows;
        private DataTable _previewTable = new();
        private DataView _previewView = new DataTable().DefaultView;
        private FieldMappingRowViewModel? _selectedMapping;

        public event PropertyChangedEventHandler? PropertyChanged;

        public ObservableCollection<FieldMappingRowViewModel> FieldMappings => _fieldMappings;

        public DataView PreviewView
        {
            get => _previewView;
            private set
            {
                if (!ReferenceEquals(_previewView, value))
                {
                    _previewView = value;
                    OnPropertyChanged(nameof(PreviewView));
                }
            }
        }

        public FieldMappingRowViewModel? SelectedMapping
        {
            get => _selectedMapping;
            set
            {
                if (!ReferenceEquals(_selectedMapping, value))
                {
                    _selectedMapping = value;
                    OnPropertyChanged(nameof(SelectedMapping));
                }
            }
        }

        public ICommand RemoveFieldCommand { get; }

        public FieldMappingWindow(SheetMapping sheetMapping)
        {
            InitializeComponent();

            _sheetMapping = sheetMapping ?? throw new ArgumentNullException(nameof(sheetMapping));

            Title = $"Field Mapping - {Path.GetFileName(sheetMapping.SourceFile)} | {sheetMapping.SheetName}";

            EpplusLicenseHelper.EnsureNonCommercialLicense();
            var preview = ExcelPreviewService.LoadPreview(sheetMapping.SourceFile, sheetMapping.SheetName, sheetMapping.IgnoreHeader, maxRows: 5);
            _columns = new ObservableCollection<ExcelColumnOption>(preview.Columns);
            _previewRows = preview.Rows;

            _fieldMappings = new ObservableCollection<FieldMappingRowViewModel>(
                sheetMapping.FieldMappings.Select(m => new FieldMappingRowViewModel(m, _columns, OnMappingChanged))
            );

            RemoveFieldCommand = new RelayCommand<FieldMappingRowViewModel>(RemoveField, vm => vm?.CanRemove == true);

            DataContext = this;
            SelectedMapping = _fieldMappings.FirstOrDefault();

            UpdatePreview();
        }

        private void RemoveField_Click(object sender, RoutedEventArgs e)
        {
            if (SelectedMapping?.CanRemove == true)
            {
                RemoveField(SelectedMapping);
            }
        }

        private void AddField_Click(object sender, RoutedEventArgs e)
        {
            var baseName = "CustomField";
            var index = 1;
            string candidate;
            do
            {
                candidate = $"{baseName}{index}";
                index++;
            } while (_fieldMappings.Any(f => string.Equals(f.FieldName, candidate, StringComparison.OrdinalIgnoreCase)));

            var model = new FieldMapping
            {
                FieldName = candidate,
                Category = "Custom",
                Visible = true,
                IsRequired = false
            };

            var vm = new FieldMappingRowViewModel(model, _columns, OnMappingChanged);
            _fieldMappings.Add(vm);
            SelectedMapping = vm;
            CommandManager.InvalidateRequerySuggested();
        }

        private void AutoMap_Click(object sender, RoutedEventArgs e)
        {
            foreach (var mapping in _fieldMappings)
            {
                if (mapping.SelectedColumn != null)
                {
                    continue;
                }

                var normalizedField = Normalize(mapping.FieldName);
                var match = _columns.FirstOrDefault(col => Normalize(col.Header ?? string.Empty) == normalizedField || Normalize(col.DisplayName).Contains(normalizedField));
                if (match != null)
                {
                    mapping.SelectedColumn = match;
                    mapping.Visible = true;
                }
            }
        }

        private void ClearMapping_Click(object sender, RoutedEventArgs e)
        {
            foreach (var mapping in _fieldMappings)
            {
                mapping.SelectedColumn = null;
                if (!mapping.IsRequired)
                {
                    mapping.Visible = false;
                }
            }
        }

        private void Ok_Click(object sender, RoutedEventArgs e)
        {
            var missingRequired = _fieldMappings
                .Where(m => m.IsRequired && m.SelectedColumn == null)
                .Select(m => m.FieldName)
                .ToList();

            if (missingRequired.Count > 0)
            {
                MessageBox.Show($"Please map the required field(s): {string.Join(", ", missingRequired)}", "Missing Required Fields", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            _sheetMapping.FieldMappings = _fieldMappings.Select(vm => vm.ToModel()).ToList();
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void RemoveField(FieldMappingRowViewModel? vm)
        {
            if (vm == null || !vm.CanRemove)
            {
                return;
            }

            var wasSelected = ReferenceEquals(SelectedMapping, vm);
            _fieldMappings.Remove(vm);
            if (wasSelected)
            {
                SelectedMapping = _fieldMappings.FirstOrDefault();
            }
            UpdatePreview();
            CommandManager.InvalidateRequerySuggested();
        }

        private void OnMappingChanged(FieldMappingRowViewModel _)
        {
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            var visibleMappings = _fieldMappings
                .Where(m => m.SelectedColumn != null && m.Visible)
                .ToList();

            _previewTable = new DataTable();
            foreach (var mapping in visibleMappings)
            {
                var columnName = mapping.FieldName;
                if (_previewTable.Columns.Contains(columnName))
                {
                    columnName = columnName + " (dup)";
                }
                _previewTable.Columns.Add(columnName);
            }

            foreach (var row in _previewRows)
            {
                var dataRow = _previewTable.NewRow();
                for (int i = 0; i < visibleMappings.Count; i++)
                {
                    var mapping = visibleMappings[i];
                    var columnIndex = mapping.SelectedColumn?.Index;
                    if (columnIndex.HasValue && columnIndex.Value >= 1)
                    {
                        dataRow[i] = row.TryGetValue(columnIndex.Value, out var value) ? value : string.Empty;
                    }
                }
                _previewTable.Rows.Add(dataRow);
            }

            PreviewView = _previewTable.DefaultView;
        }

        private static string Normalize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            return new string(value.Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());
        }

        protected void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private sealed class RelayCommand<T> : ICommand
        {
            private readonly Action<T?> _execute;
            private readonly Func<T?, bool>? _canExecute;

            public RelayCommand(Action<T?> execute, Func<T?, bool>? canExecute = null)
            {
                _execute = execute;
                _canExecute = canExecute;
            }

            public event EventHandler? CanExecuteChanged
            {
                add { CommandManager.RequerySuggested += value; }
                remove { CommandManager.RequerySuggested -= value; }
            }

            public bool CanExecute(object? parameter)
            {
                return _canExecute == null || _canExecute((T?)parameter);
            }

            public void Execute(object? parameter)
            {
                _execute((T?)parameter);
            }
        }
    }

    public sealed class ExcelColumnOption
    {
        public int Index { get; init; }
        public string Letter { get; init; } = string.Empty;
        public string? Header { get; init; }
        public string DisplayName => string.IsNullOrWhiteSpace(Header) ? Letter : $"{Header} ({Letter})";
    }

    public sealed class FieldMappingRowViewModel : INotifyPropertyChanged
    {
        private readonly Action<FieldMappingRowViewModel> _onChanged;
        private string _fieldName;
        private ExcelColumnOption? _selectedColumn;
        private bool _visible;
        private readonly ObservableCollection<ExcelColumnSelection> _columnSelections;
        private ExcelColumnSelection? _selectedColumnSelection;

        public FieldMappingRowViewModel(FieldMapping model, ObservableCollection<ExcelColumnOption> columns, Action<FieldMappingRowViewModel> onChanged)
        {
            Model = model;
            Columns = columns;
            _onChanged = onChanged;
            _fieldName = model.FieldName;
            _visible = model.Visible;
            if (model.ColumnIndex.HasValue)
            {
                _selectedColumn = columns.FirstOrDefault(c => c.Index == model.ColumnIndex.Value);
            }

            _columnSelections = new ObservableCollection<ExcelColumnSelection>();
            Columns.CollectionChanged += ColumnsCollectionChanged;
            RebuildColumnSelections();
        }

        public FieldMapping Model { get; }
        public ObservableCollection<ExcelColumnOption> Columns { get; }
        public ObservableCollection<ExcelColumnSelection> ColumnSelections => _columnSelections;
        public string FieldName
        {
            get => _fieldName;
            set
            {
                if (_fieldName != value)
                {
                    _fieldName = value;
                    OnPropertyChanged(nameof(FieldName));
                    _onChanged(this);
                }
            }
        }

        public ExcelColumnOption? SelectedColumn
        {
            get => _selectedColumn;
            set
            {
                if (_selectedColumn != value)
                {
                    _selectedColumn = value;
                    _selectedColumnSelection = _columnSelections.FirstOrDefault(s => Equals(s.Option, value))
                        ?? _columnSelections.FirstOrDefault(s => s.IsNone);
                    OnPropertyChanged(nameof(SelectedColumn));
                    OnPropertyChanged(nameof(SelectedColumnSelection));
                    _onChanged(this);
                }
            }
        }

        public ExcelColumnSelection? SelectedColumnSelection
        {
            get => _selectedColumnSelection;
            set
            {
                if (!ReferenceEquals(_selectedColumnSelection, value))
                {
                    _selectedColumnSelection = value;
                    var option = value?.Option;
                    if (!Equals(_selectedColumn, option))
                    {
                        _selectedColumn = option;
                        OnPropertyChanged(nameof(SelectedColumn));
                        _onChanged(this);
                    }
                    OnPropertyChanged(nameof(SelectedColumnSelection));
                }
            }
        }

        public bool Visible
        {
            get => _visible;
            set
            {
                if (_visible != value)
                {
                    _visible = value;
                    OnPropertyChanged(nameof(Visible));
                    _onChanged(this);
                }
            }
        }

        public bool IsFieldNameEditable => Model.IsCustom;
        public bool CanRemove => Model.IsCustom && !Model.IsRequired;
        public bool IsRequired => Model.IsRequired;

        public FieldMapping ToModel()
        {
            Model.FieldName = FieldName;
            Model.Visible = Visible;
            Model.ColumnIndex = SelectedColumn?.Index;
            Model.ColumnHeader = SelectedColumn?.Header;
            return Model;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged(string propertyName)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        private void ColumnsCollectionChanged(object? sender, NotifyCollectionChangedEventArgs e)
        {
            RebuildColumnSelections();
        }

        private void RebuildColumnSelections()
        {
            var currentColumn = _selectedColumn;
            _columnSelections.Clear();
            _columnSelections.Add(new ExcelColumnSelection(null));
            foreach (var column in Columns)
            {
                _columnSelections.Add(new ExcelColumnSelection(column));
            }

            _selectedColumnSelection = _columnSelections.FirstOrDefault(s => Equals(s.Option, currentColumn))
                ?? _columnSelections.FirstOrDefault(s => s.IsNone);

            OnPropertyChanged(nameof(ColumnSelections));
            OnPropertyChanged(nameof(SelectedColumnSelection));
        }
    }

    public sealed class ExcelColumnSelection
    {
        public ExcelColumnSelection(ExcelColumnOption? option)
        {
            Option = option;
        }

        public ExcelColumnOption? Option { get; }
        public bool IsNone => Option == null;
        public string DisplayName => Option == null ? "(none)" : Option.DisplayName;
    }
}
