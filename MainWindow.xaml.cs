using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Windows;
using Microsoft.Win32;
using OfficeOpenXml;

namespace KiCadExcelBridge;

public partial class MainWindow : Window
{
    public event Action<List<string>, List<SheetMapping>>? OnFileSelectionChanged;
    private readonly ObservableCollection<string> _sourceFiles = new();
    private readonly ObservableCollection<SheetMapping> _sheetMappings = new();
    private readonly Dictionary<string, FileSystemWatcher> _fileWatchers = new();
    private readonly Dictionary<string, System.Windows.Threading.DispatcherTimer> _debounceTimers = new();

    public MainWindow()
    {
        InitializeComponent();
        SourceFilesList.ItemsSource = _sourceFiles;
        SheetMappingsGrid.ItemsSource = _sheetMappings;
        EpplusLicenseHelper.EnsureNonCommercialLicense();
        LoadConfiguration();
        // Start watchers for any files loaded from configuration
        foreach (var f in _sourceFiles)
        {
            WatchFile(f);
        }
    }

    private void LoadConfiguration()
    {
        var config = ConfigurationManager.Load();
        foreach (var file in config.SourceFiles)
        {
            _sourceFiles.Add(file);
        }
        foreach (var mapping in config.SheetMappings)
        {
            EnsureFieldMappingDefaults(mapping);
            _sheetMappings.Add(mapping);
        }
        // Global IgnoreHeader setting deprecated - per-sheet IgnoreHeader is used instead.
    }


    private void AddFile_Click(object sender, RoutedEventArgs e)
    {
        var openFileDialog = new OpenFileDialog
        {
            Filter = "Excel Files (*.xlsx)|*.xlsx|CSV files (*.csv)|*.csv|All files (*.*)|*.*",
            Multiselect = true
        };

        if (openFileDialog.ShowDialog() == true)
        {
            foreach (var file in openFileDialog.FileNames)
            {
                if (!_sourceFiles.Contains(file))
                {
                    _sourceFiles.Add(file);
                    RefreshSheetsForFile(file);
                    WatchFile(file);
                }
            }
        }
    }

    private void RemoveFile_Click(object sender, RoutedEventArgs e)
    {
        var selectedFiles = SourceFilesList.SelectedItems.Cast<string>().ToList();
        foreach (var file in selectedFiles)
        {
            _sourceFiles.Remove(file);
            UnwatchFile(file);
            var mappingsToRemove = _sheetMappings
                .Where(m => string.Equals(m.SourceFile, file, StringComparison.OrdinalIgnoreCase) || string.Equals(Path.GetFileName(m.SourceFile), Path.GetFileName(file), StringComparison.OrdinalIgnoreCase))
                .ToList();
            foreach (var mapping in mappingsToRemove)
            {
                _sheetMappings.Remove(mapping);
            }
        }
    }

    private void RefreshSheetsForFile(string filePath)
    {
        try
        {
            var newSheets = new List<string>();
            if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
            {
                using var package = new ExcelPackage(new FileInfo(filePath));
                foreach (var worksheet in package.Workbook.Worksheets)
                {
                    newSheets.Add(worksheet.Name);
                }
            }
            else if (Path.GetExtension(filePath).Equals(".csv", StringComparison.OrdinalIgnoreCase))
            {
                newSheets.Add(Path.GetFileNameWithoutExtension(filePath));
            }

            // Preserve existing category labels for sheets that remain
            var existingForFile = _sheetMappings.Where(m => string.Equals(m.SourceFile, filePath, StringComparison.OrdinalIgnoreCase) || string.Equals(Path.GetFileName(m.SourceFile), Path.GetFileName(filePath), StringComparison.OrdinalIgnoreCase)).ToList();
            var existingLookup = existingForFile.ToDictionary(m => m.SheetName, m => m);

            // Remove mappings for sheets that no longer exist
            var toRemove = existingForFile.Where(m => !newSheets.Contains(m.SheetName)).ToList();
            foreach (var rem in toRemove)
            {
                _sheetMappings.Remove(rem);
            }

            // Add or update sheets
            foreach (var sheet in newSheets)
            {
                if (existingLookup.TryGetValue(sheet, out var existing))
                {
                    existing.SourceFile = filePath;
                    EnsureFieldMappingDefaults(existing);
                }
                else
                {
                    var mapping = new SheetMapping
                    {
                        SourceFile = filePath,
                        SheetName = sheet,
                        CategoryLabel = string.Empty,
                        IgnoreHeader = true,
                        FieldMappings = FieldMappingDefaults.CreateDefaults()
                    };
                    _sheetMappings.Add(mapping);
                }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error loading file: {ex.Message}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void WatchFile(string filePath)
    {
        try
        {
            var fullPath = Path.GetFullPath(filePath);
            if (_fileWatchers.ContainsKey(fullPath)) return;

            var dir = Path.GetDirectoryName(fullPath);
            var name = Path.GetFileName(fullPath);
            if (string.IsNullOrEmpty(dir) || string.IsNullOrEmpty(name)) return;

            var watcher = new FileSystemWatcher(dir, name)
            {
                NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size
            };

            FileSystemEventHandler onChanged = (s, e) =>
            {
                // We handle Deleted immediately; other events are debounced.
                Application.Current.Dispatcher.Invoke(() =>
                {
                    if (e.ChangeType == WatcherChangeTypes.Deleted)
                    {
                        var mappingsToRemove = _sheetMappings.Where(m => string.Equals(m.SourceFile, fullPath, StringComparison.OrdinalIgnoreCase)).ToList();
                        foreach (var mapping in mappingsToRemove) _sheetMappings.Remove(mapping);
                        UnwatchFile(fullPath);
                        _sourceFiles.Remove(fullPath);
                        // If a debounce timer exists, stop it
                        if (_debounceTimers.TryGetValue(fullPath, out var t))
                        {
                            t.Stop();
                            _debounceTimers.Remove(fullPath);
                        }
                    }
                    else
                    {
                        // Debounce refresh: reset or start a DispatcherTimer for this file
                        if (!_debounceTimers.TryGetValue(fullPath, out var timer))
                        {
                            timer = new System.Windows.Threading.DispatcherTimer
                            {
                                Interval = TimeSpan.FromMilliseconds(500)
                            };
                            timer.Tick += (sender, args) =>
                            {
                                timer.Stop();
                                _debounceTimers.Remove(fullPath);
                                // Attempt to refresh sheets (file may be in-use briefly)
                                try
                                {
                                    RefreshSheetsForFile(fullPath);
                                }
                                catch
                                {
                                    // ignore refresh errors; next file change will retry
                                }
                            };
                            _debounceTimers[fullPath] = timer;
                        }
                        else
                        {
                            // reset interval
                            timer.Stop();
                        }
                        timer.Start();
                    }
                });
            };

            watcher.Changed += onChanged;
            watcher.Renamed += (s, e) =>
            {
                // Treat rename as deletion of the old name and possible add of the new
                onChanged(s, new FileSystemEventArgs(WatcherChangeTypes.Deleted, Path.GetDirectoryName(e.OldFullPath) ?? "", Path.GetFileName(e.OldFullPath)));
                onChanged(s, new FileSystemEventArgs(WatcherChangeTypes.Changed, Path.GetDirectoryName(e.FullPath) ?? "", Path.GetFileName(e.FullPath)));
            };
            watcher.Deleted += onChanged;
            watcher.EnableRaisingEvents = true;

            _fileWatchers[fullPath] = watcher;
        }
        catch
        {
            // ignore watcher failures
        }
    }

    private void UnwatchFile(string filePath)
    {
        try
        {
            var fullPath = Path.GetFullPath(filePath);
            if (_fileWatchers.TryGetValue(fullPath, out var watcher))
            {
                watcher.EnableRaisingEvents = false;
                watcher.Dispose();
                _fileWatchers.Remove(fullPath);
            }
        }
        catch
        {
            // ignore
        }
    }

    private void ConfigureFields_Click(object sender, RoutedEventArgs e)
    {
        if (SheetMappingsGrid.SelectedItem is not SheetMapping selected)
        {
            MessageBox.Show("Select a sheet mapping first.", "Field Mapping", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        EnsureFieldMappingDefaults(selected);

        var resolvedPath = ResolveSourceFilePath(selected.SourceFile);
        if (!File.Exists(resolvedPath))
        {
            MessageBox.Show($"Unable to locate source file:\n{resolvedPath}", "Field Mapping", MessageBoxButton.OK, MessageBoxImage.Warning);
            return;
        }

        selected.SourceFile = resolvedPath;

        var dialog = new FieldMappingWindow(selected)
        {
            Owner = this
        };

        if (dialog.ShowDialog() == true)
        {
            ApplyChanges();
        }
    }

    private void RefreshFiles_Click(object sender, RoutedEventArgs e)
    {
        var selectedFiles = SourceFilesList.SelectedItems.Cast<string>().ToList();
        if (selectedFiles.Count == 0)
        {
            // Refresh all
            foreach (var f in _sourceFiles.ToList())
            {
                try { RefreshSheetsForFile(f); } catch { }
            }
        }
        else
        {
            foreach (var f in selectedFiles)
            {
                try { RefreshSheetsForFile(f); } catch { }
            }
        }
    }

    private void SheetMappingsGrid_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
    {
        // Prevent Delete key from removing rows
        if (e.Key == System.Windows.Input.Key.Delete)
        {
            e.Handled = true;
        }
    }

    private void Ok_Click(object sender, RoutedEventArgs e)
    {
        ApplyChanges();
        Close();
    }

    private void Cancel_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }

    private void Apply_Click(object sender, RoutedEventArgs e)
    {
        ApplyChanges();
    }

    private void ApplyChanges()
    {
            var duplicateCategories = _sheetMappings
                .Select(m => m.CategoryLabel?.Trim())
                .Where(label => !string.IsNullOrWhiteSpace(label))
                .GroupBy(label => label, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .Select(group => group.Key)
                .ToList();

            if (duplicateCategories.Count > 0)
            {
                MessageBox.Show($"Category names must be unique. Please resolve duplicates: {string.Join(", ", duplicateCategories)}", "Duplicate Categories", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

        var filePaths = _sourceFiles.ToList();

        var config = new AppConfiguration
        {
            SourceFiles = filePaths,
            SheetMappings = _sheetMappings.ToList()
        };
        ConfigurationManager.Save(config);

        OnFileSelectionChanged?.Invoke(filePaths, _sheetMappings.ToList());
    }

    private void EnsureFieldMappingDefaults(SheetMapping mapping)
    {
        if (mapping.FieldMappings == null || mapping.FieldMappings.Count == 0)
        {
            mapping.FieldMappings = FieldMappingDefaults.CreateDefaults();
        }
        else
        {
            FieldMappingDefaults.EnsureDefaults(mapping.FieldMappings);
        }
    }

    private string ResolveSourceFilePath(string storedPath)
    {
        if (File.Exists(storedPath))
        {
            return storedPath;
        }

        var directMatch = _sourceFiles.FirstOrDefault(s => string.Equals(s, storedPath, StringComparison.OrdinalIgnoreCase));
        if (directMatch != null)
        {
            return directMatch;
        }

        var fileName = Path.GetFileName(storedPath);
        var nameMatch = _sourceFiles.FirstOrDefault(s => string.Equals(Path.GetFileName(s), fileName, StringComparison.OrdinalIgnoreCase));
        return nameMatch ?? storedPath;
    }
}