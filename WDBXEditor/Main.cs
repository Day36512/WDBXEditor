
using WDBXEditor.Reader;
using WDBXEditor.Storage;
using WDBXEditor.Archives.MPQ;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using static WDBXEditor.Common.Constants;
using static WDBXEditor.Forms.InputBox;
using System.Threading.Tasks;
using WDBXEditor.Forms;
using WDBXEditor.Common;
using System.Text.RegularExpressions;
using System.Net;
using System.Web.Script.Serialization;
using WDBXEditor.Reader.FileTypes;
using System.Runtime.InteropServices;

namespace WDBXEditor
{
    public partial class Main : Form
    {
        protected DBEntry LoadedEntry;

        private BindingSource _bindingsource = new BindingSource();
        private FileSystemWatcher watcher = new FileSystemWatcher();
        private readonly HashSet<string> _changedCellKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _changedRowKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private string _highlightedRowKey = string.Empty;
        private readonly JavaScriptSerializer _serializer = new JavaScriptSerializer();
        private UiPreferences _uiPreferences = new UiPreferences();
        private bool _suppressColumnWidthPersistence;
        private bool _applyingUiPreferences;
        private bool _isLargeDataset;
        private const string NewWindowMenuName = "newWindowTopToolStripMenuItem";

        private const int LargeDatasetRowThreshold = 25000;

        private bool IsLoaded => (LoadedEntry != null && _bindingsource.DataSource != null);
        private DBEntry GetEntry() => Database.Entries.FirstOrDefault(x => x.FileName == txtCurEntry.Text && x.BuildName == txtCurDefinition.Text);

        [DllImport("dwmapi.dll")]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int attr, ref int attrValue, int attrSize);

        [DllImport("uxtheme.dll", CharSet = CharSet.Unicode)]
        private static extern int SetWindowTheme(IntPtr hWnd, string appName, string partList);


        private string UiPreferencesPath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "WDBXEditor",
            "ui-preferences.json");

        private void InitializeUi()
        {
            _bindingsource.DataSource = null;
            advancedDataGridView.DataSource = _bindingsource;

            HookUiBehavior();
            ApplyThemeAndLayout();
        }

        private void HookUiBehavior()
        {
            ConfigureDataGridViewPerformance();

            advancedDataGridView.CellFormatting -= advancedDataGridView_CellFormatting;
            advancedDataGridView.CellFormatting += advancedDataGridView_CellFormatting;

            advancedDataGridView.RowPrePaint -= advancedDataGridView_RowPrePaint;
            advancedDataGridView.RowPrePaint += advancedDataGridView_RowPrePaint;

            advancedDataGridView.RowPostPaint -= advancedDataGridView_RowPostPaint;
            advancedDataGridView.RowPostPaint += advancedDataGridView_RowPostPaint;

            advancedDataGridView.ColumnWidthChanged -= advancedDataGridView_ColumnWidthChanged;
            advancedDataGridView.ColumnWidthChanged += advancedDataGridView_ColumnWidthChanged;

            advancedDataGridView.Sorted -= advancedDataGridView_Sorted;
            advancedDataGridView.Sorted += advancedDataGridView_Sorted;

            advancedDataGridView.DataError -= advancedDataGridView_DataError;
            advancedDataGridView.DataError += advancedDataGridView_DataError;

            advancedDataGridView.RowHeaderMouseClick -= advancedDataGridView_RowHeaderMouseClick;
            advancedDataGridView.RowHeaderMouseClick += advancedDataGridView_RowHeaderMouseClick;

            advancedDataGridView.Scroll -= advancedDataGridView_Scroll;
            advancedDataGridView.Scroll += advancedDataGridView_Scroll;
        }

        private void advancedDataGridView_Scroll(object sender, ScrollEventArgs e)
        {
            if (string.IsNullOrEmpty(_highlightedRowKey))
                return;

            // Horizontal scrolling can leave custom row paint stale until another redraw happens.
            // Force a repaint so the highlighted row stays visually correct.
            if (e.ScrollOrientation == ScrollOrientation.HorizontalScroll)
                advancedDataGridView.Invalidate();
        }

        private void advancedDataGridView_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= advancedDataGridView.Rows.Count)
                return;

            if (!TryCommitPendingGridEdit())
                return;

            DataGridViewRow row = advancedDataGridView.Rows[e.RowIndex];
            string rowKey = GetRowIdentity(row);

            // Toggle off if the same row header is clicked again.
            if (string.Equals(_highlightedRowKey, rowKey, StringComparison.OrdinalIgnoreCase))
            {
                _highlightedRowKey = string.Empty;
                advancedDataGridView.ClearSelection();
                advancedDataGridView.Invalidate();
                return;
            }

            _highlightedRowKey = rowKey;

            advancedDataGridView.ClearSelection();
            advancedDataGridView.SelectRow(e.RowIndex);

            if (advancedDataGridView.Columns.Count > 0)
            {
                for (int i = 0; i < advancedDataGridView.Columns.Count; i++)
                {
                    if (advancedDataGridView.Columns[i].Visible)
                    {
                        try
                        {
                            advancedDataGridView.CurrentCell = advancedDataGridView.Rows[e.RowIndex].Cells[i];
                        }
                        catch
                        {
                        }
                        break;
                    }
                }
            }

            advancedDataGridView.Invalidate();
        }

        private void ConfigureDataGridViewPerformance()
        {
            if (advancedDataGridView == null)
                return;

            advancedDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            advancedDataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            advancedDataGridView.AllowUserToResizeRows = false;

            // This is the important part: clicking into a cell should actually let you edit it.
            advancedDataGridView.EditMode = DataGridViewEditMode.EditOnKeystrokeOrF2;
            advancedDataGridView.SelectionMode = DataGridViewSelectionMode.CellSelect;
            advancedDataGridView.MultiSelect = false;
            advancedDataGridView.StandardTab = true;

            try
            {
                typeof(DataGridView)
                    .GetProperty("DoubleBuffered", System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)
                    ?.SetValue(advancedDataGridView, true, null);
            }
            catch
            {
            }
        }

        private DataGridViewAutoSizeColumnsMode GetEffectiveAutoSizeMode(DataGridViewAutoSizeColumnsMode requestedMode)
        {
            if (!_isLargeDataset)
                return requestedMode;

            switch (requestedMode)
            {
                case DataGridViewAutoSizeColumnsMode.DisplayedCells:
                case DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader:
                    return DataGridViewAutoSizeColumnsMode.None;
                default:
                    return requestedMode;
            }
        }

        private void UpdateDatasetPerformanceMode()
        {
            int rowCount = LoadedEntry?.Data?.Rows?.Count ?? 0;
            _isLargeDataset = rowCount >= LargeDatasetRowThreshold;

            if (advancedDataGridView == null)
                return;

            advancedDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.None;
            advancedDataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            advancedDataGridView.AllowUserToResizeRows = false;

            if (_isLargeDataset)
            {
                advancedDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
                advancedDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            }
        }

        private bool TryCommitPendingGridEdit()
        {
            if (advancedDataGridView == null)
                return true;

            try
            {
                if (advancedDataGridView.IsCurrentCellInEditMode)
                    advancedDataGridView.EndEdit(DataGridViewDataErrorContexts.Commit);

                if (_bindingsource != null)
                    _bindingsource.EndEdit();

                return !advancedDataGridView.IsCurrentCellInEditMode;
            }
            catch
            {
                try
                {
                    advancedDataGridView.CancelEdit();
                    _bindingsource?.CancelEdit();
                }
                catch
                {
                }

                return false;
            }
        }

        private void ApplyThemeAndLayout()
        {
            EnableDarkTitleBar();
            ApplyModernGridTheme();
            ApplyDarkChrome();
            ApplyReadableSizing();
        }

        private string GetSelectedColumnModeKey()
        {
            if (cbColumnMode?.SelectedItem is KeyValuePair<string, DataGridViewAutoSizeColumnsMode> item)
                return item.Key;

            return "None";
        }

        private void SetColumnModeSelection(string key)
        {
            if (cbColumnMode == null)
                return;

            _applyingUiPreferences = true;
            try
            {
                for (int i = 0; i < cbColumnMode.Items.Count; i++)
                {
                    if (cbColumnMode.Items[i] is KeyValuePair<string, DataGridViewAutoSizeColumnsMode> item &&
                        string.Equals(item.Key, key, StringComparison.OrdinalIgnoreCase))
                    {
                        cbColumnMode.SelectedIndex = i;
                        return;
                    }
                }

                if (cbColumnMode.Items.Count > 0)
                    cbColumnMode.SelectedIndex = 0;
            }
            finally
            {
                _applyingUiPreferences = false;
            }
        }

        private string GetCurrentEntryPreferenceKey()
        {
            if (LoadedEntry == null)
                return string.Empty;

            return string.Concat(LoadedEntry.FileName ?? string.Empty, "|", LoadedEntry.BuildName ?? string.Empty);
        }

        private void SaveCurrentColumnWidthsForEntry()
        {
            if (!IsLoaded)
                return;

            string entryKey = GetCurrentEntryPreferenceKey();
            if (string.IsNullOrWhiteSpace(entryKey))
                return;

            Dictionary<string, int> widths = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (DataGridViewColumn col in advancedDataGridView.Columns)
            {
                if (!col.Visible)
                    continue;

                widths[col.Name] = col.Width;
            }

            _uiPreferences.ColumnWidths[entryKey] = widths;
        }

        private void ApplyStoredColumnWidthsForEntry()
        {
            if (!IsLoaded)
                return;

            string entryKey = GetCurrentEntryPreferenceKey();
            if (string.IsNullOrWhiteSpace(entryKey) || !_uiPreferences.ColumnWidths.TryGetValue(entryKey, out Dictionary<string, int> widths))
                return;

            _suppressColumnWidthPersistence = true;
            try
            {
                foreach (DataGridViewColumn col in advancedDataGridView.Columns)
                {
                    if (!col.Visible || !widths.TryGetValue(col.Name, out int width))
                        continue;

                    col.Width = Math.Max(col.MinimumWidth, width);
                }
            }
            finally
            {
                _suppressColumnWidthPersistence = false;
            }
        }

        private void LoadUiPreferences()
        {
            try
            {
                if (!File.Exists(UiPreferencesPath))
                    return;

                string json = File.ReadAllText(UiPreferencesPath);
                if (string.IsNullOrWhiteSpace(json))
                    return;

                UiPreferences loaded = _serializer.Deserialize<UiPreferences>(json);
                if (loaded != null)
                    _uiPreferences = loaded;
            }
            catch
            {
                _uiPreferences = new UiPreferences();
            }
        }

        private void ApplyUiPreferences()
        {
            _applyingUiPreferences = true;
            try
            {
                Rectangle desiredBounds = new Rectangle(
                    _uiPreferences.WindowLeft,
                    _uiPreferences.WindowTop,
                    _uiPreferences.WindowWidth,
                    _uiPreferences.WindowHeight);

                if (_uiPreferences.WindowWidth > 0 &&
                    _uiPreferences.WindowHeight > 0 &&
                    Screen.AllScreens.Any(screen => screen.WorkingArea.IntersectsWith(desiredBounds)))
                {
                    StartPosition = FormStartPosition.Manual;
                    Bounds = desiredBounds;
                }

                SetColumnModeSelection(_uiPreferences.LastColumnModeKey ?? "None");

                if (_uiPreferences.WindowMaximized)
                    WindowState = FormWindowState.Maximized;
            }
            finally
            {
                _applyingUiPreferences = false;
            }
        }

        private void SaveUiPreferences()
        {
            try
            {
                CaptureWindowBounds();

                _uiPreferences.LastColumnModeKey = GetSelectedColumnModeKey();

                string folder = Path.GetDirectoryName(UiPreferencesPath);
                if (!Directory.Exists(folder))
                    Directory.CreateDirectory(folder);

                string json = _serializer.Serialize(_uiPreferences);
                File.WriteAllText(UiPreferencesPath, json);
            }
            catch
            {
            }
        }

        private void CaptureWindowBounds()
        {
            Rectangle bounds = WindowState == FormWindowState.Normal ? Bounds : RestoreBounds;

            _uiPreferences.WindowLeft = bounds.Left;
            _uiPreferences.WindowTop = bounds.Top;
            _uiPreferences.WindowWidth = bounds.Width;
            _uiPreferences.WindowHeight = bounds.Height;
            _uiPreferences.WindowMaximized = WindowState == FormWindowState.Maximized;
        }

        private static string EscapeFilterValue(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            return value
                .Replace("'", "''")
                .Replace("[", "[[]")
                .Replace("]", "[]]")
                .Replace("%", "[%]")
                .Replace("*", "[*]");
        }

        private void ApplyFileFilter()
        {
            if (!(lbFiles.DataSource is BindingSource source))
                return;

            string text = EscapeFilterValue(txtFilter?.Text?.Trim());
            string build = EscapeFilterValue(cbBuild?.Text?.Trim());

            List<string> filters = new List<string>();
            if (!string.IsNullOrEmpty(text))
                filters.Add(string.Format("[Value] LIKE '%{0}%'", text));

            if (!string.IsNullOrEmpty(build))
                filters.Add(string.Format("[Value] LIKE '%{0}%'", build));

            source.Filter = filters.Count > 0 ? string.Join(" AND ", filters) : string.Empty;
        }

        private string GetRowIdentity(DataGridViewRow row)
        {
            if (row == null)
                return string.Empty;

            if (LoadedEntry != null &&
                row.DataBoundItem is DataRowView rowView &&
                rowView.Row != null &&
                rowView.Row.Table != null &&
                rowView.Row.Table.Columns.Contains(LoadedEntry.Key))
            {
                return Convert.ToString(rowView.Row[LoadedEntry.Key]) ?? string.Empty;
            }

            return row.Index.ToString();
        }

        private string GetChangedCellKey(int rowIndex, int columnIndex)
        {
            if (rowIndex < 0 || columnIndex < 0 ||
                rowIndex >= advancedDataGridView.Rows.Count ||
                columnIndex >= advancedDataGridView.Columns.Count)
                return string.Empty;

            string rowKey = GetRowIdentity(advancedDataGridView.Rows[rowIndex]);
            if (string.IsNullOrEmpty(rowKey))
                return string.Empty;

            return rowKey + "|" + advancedDataGridView.Columns[columnIndex].Name;
        }

        private bool IsHighlightedRow(DataGridViewRow row)
        {
            if (row == null || string.IsNullOrEmpty(_highlightedRowKey))
                return false;

            return string.Equals(GetRowIdentity(row), _highlightedRowKey, StringComparison.OrdinalIgnoreCase);
        }

        private void advancedDataGridView_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0 || e.RowIndex >= advancedDataGridView.Rows.Count)
                return;

            if (_changedCellKeys.Count == 0 && LoadedEntry == null)
                return;

            DataGridViewColumn column = advancedDataGridView.Columns[e.ColumnIndex];
            bool isKeyColumn = LoadedEntry != null &&
                               string.Equals(column.Name, LoadedEntry.Key, StringComparison.OrdinalIgnoreCase);

            string changedCellKey = GetChangedCellKey(e.RowIndex, e.ColumnIndex);
            bool isChangedCell = !string.IsNullOrEmpty(changedCellKey) && _changedCellKeys.Contains(changedCellKey);
            bool isHighlightedRow = IsHighlightedRow(advancedDataGridView.Rows[e.RowIndex]);

            if (isHighlightedRow && !isChangedCell)
            {
                e.CellStyle.BackColor = Color.FromArgb(36, 49, 71);
                e.CellStyle.ForeColor = Color.FromArgb(238, 241, 245);
                e.CellStyle.SelectionBackColor = Color.FromArgb(59, 130, 246);
                e.CellStyle.SelectionForeColor = Color.White;
            }

            if (isKeyColumn)
            {
                e.CellStyle.BackColor = isHighlightedRow ? Color.FromArgb(42, 56, 82) : Color.FromArgb(30, 34, 40);
                e.CellStyle.ForeColor = isHighlightedRow ? Color.FromArgb(220, 228, 240) : Color.FromArgb(170, 178, 188);
            }

            if (isChangedCell)
            {
                e.CellStyle.BackColor = Color.FromArgb(86, 60, 18);
                e.CellStyle.ForeColor = Color.FromArgb(255, 241, 214);
                e.CellStyle.SelectionBackColor = Color.FromArgb(191, 145, 39);
                e.CellStyle.SelectionForeColor = Color.White;
            }
        }

        private void EnsureNewWindowMenu()
        {
            if (menuStrip == null)
                return;

            ToolStripMenuItem existing = menuStrip.Items
                .OfType<ToolStripMenuItem>()
                .FirstOrDefault(x => string.Equals(x.Name, NewWindowMenuName, StringComparison.Ordinal));

            if (existing != null)
                return;

            ToolStripMenuItem newWindowItem = new ToolStripMenuItem("New &Window")
            {
                Name = NewWindowMenuName,
                ShortcutKeys = Keys.Control | Keys.Alt | Keys.N,
                ShowShortcutKeys = true
            };

            newWindowItem.Click += (sender, e) => OpenNewWindow();

            int insertIndex = Math.Min(1, menuStrip.Items.Count);
            menuStrip.Items.Insert(insertIndex, newWindowItem);
        }

        private void OpenNewWindow()
        {
            try
            {
                if (!TryCommitPendingGridEdit())
                {
                    MessageBox.Show(
                        this,
                        "Finish or cancel the current cell edit before opening another window.",
                        "Pending Edit",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                    return;
                }

                this.ActiveControl = null;

                if (!InstanceManager.LoadNewInstance(Array.Empty<string>()))
                {
                    MessageBox.Show(
                        this,
                        "Could not open another WDBX Editor window.",
                        "New Window",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    this,
                    "Could not open another WDBX Editor window.\r\n\r\n" + ex.Message,
                    "New Window",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void advancedDataGridView_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= advancedDataGridView.Rows.Count)
                return;

            DataGridViewRow row = advancedDataGridView.Rows[e.RowIndex];
            string rowKey = GetRowIdentity(row);
            bool isHighlightedRow = IsHighlightedRow(row);

            if (_changedRowKeys.Contains(rowKey))
            {
                row.HeaderCell.Style.BackColor = Color.FromArgb(110, 72, 24);
                row.HeaderCell.Style.ForeColor = Color.FromArgb(255, 241, 214);
            }
            else if (isHighlightedRow)
            {
                row.HeaderCell.Style.BackColor = Color.FromArgb(59, 130, 246);
                row.HeaderCell.Style.ForeColor = Color.White;
            }
            else
            {
                row.HeaderCell.Style.BackColor = advancedDataGridView.RowHeadersDefaultCellStyle.BackColor;
                row.HeaderCell.Style.ForeColor = advancedDataGridView.RowHeadersDefaultCellStyle.ForeColor;
            }
        }

        private void advancedDataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= advancedDataGridView.Rows.Count)
                return;

            DataGridViewRow row = advancedDataGridView.Rows[e.RowIndex];
            if (!IsHighlightedRow(row))
                return;

            Rectangle rowBounds = new Rectangle(
                advancedDataGridView.RowHeadersWidth,
                e.RowBounds.Top,
                Math.Max(0, advancedDataGridView.ClientSize.Width - advancedDataGridView.RowHeadersWidth - 1),
                e.RowBounds.Height - 1);

            using (Pen pen = new Pen(Color.FromArgb(59, 130, 246), 2f))
            {
                e.Graphics.DrawRectangle(pen, rowBounds);
            }
        }

        private void advancedDataGridView_ColumnWidthChanged(object sender, DataGridViewColumnEventArgs e)
        {
            if (_suppressColumnWidthPersistence || _applyingUiPreferences || !IsLoaded)
                return;

            if (advancedDataGridView.AutoSizeColumnsMode != DataGridViewAutoSizeColumnsMode.None)
                return;

            SaveCurrentColumnWidthsForEntry();
            SaveUiPreferences();
        }

        private void advancedDataGridView_Sorted(object sender, EventArgs e)
        {
            advancedDataGridView.Invalidate();
        }

        private void advancedDataGridView_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.ThrowException = false;

            try
            {
                advancedDataGridView.CancelEdit();
                _bindingsource?.CancelEdit();
            }
            catch
            {
            }

            string message = "The value could not be committed to this cell.";
            if (e.Exception != null && !string.IsNullOrWhiteSpace(e.Exception.Message))
                message += "\r\n\r\n" + e.Exception.Message;

            MessageBox.Show(
                this,
                message,
                "Cell Edit Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
        }

        public Main()
        {
            InitializeComponent();
            InitializeUi();
        }

        public Main(string[] filenames)
        {
            InitializeComponent();
            InitializeUi();

            Parallel.For(0, filenames.Length, f => InstanceManager.AutoRun.Enqueue(filenames[f]));
        }

        private void EnableDarkTitleBar()
        {
            try
            {
                if (Environment.OSVersion.Version.Major >= 10)
                {
                    int useDark = 1;
                    DwmSetWindowAttribute(this.Handle, 20, ref useDark, sizeof(int));
                    DwmSetWindowAttribute(this.Handle, 19, ref useDark, sizeof(int));
                }
            }
            catch
            {
            }
        }

        private void ApplyModernGridTheme()
        {
            Font gridFont = new Font("Segoe UI", 12.9F, FontStyle.Regular);
            Font headerFont = new Font("Segoe UI Semibold", 12.0F, FontStyle.Bold);

            Color bgMain = Color.FromArgb(24, 28, 34);
            Color bgAlt = Color.FromArgb(30, 35, 42);
            Color bgHeader = Color.FromArgb(38, 44, 53);
            Color bgRowHeader = Color.FromArgb(34, 39, 47);
            Color gridLines = Color.FromArgb(58, 66, 78);

            Color textPrimary = Color.FromArgb(230, 232, 235);

            Color accentBlue = Color.FromArgb(59, 130, 246);
            Color accentBlueSoft = Color.FromArgb(37, 99, 235);

            Color accentGold = Color.FromArgb(245, 182, 66);
            Color accentGoldSoft = Color.FromArgb(191, 145, 39);

            Color accentRedSoft = Color.FromArgb(160, 62, 62);

            advancedDataGridView.Font = gridFont;
            advancedDataGridView.BackgroundColor = bgMain;
            advancedDataGridView.BorderStyle = BorderStyle.None;
            advancedDataGridView.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            advancedDataGridView.GridColor = gridLines;

            advancedDataGridView.EnableHeadersVisualStyles = false;
            advancedDataGridView.AllowUserToResizeRows = false;
            advancedDataGridView.StandardTab = true;

            advancedDataGridView.RowHeadersWidth = 58;
            advancedDataGridView.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;

            advancedDataGridView.DefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = bgMain,
                ForeColor = textPrimary,
                SelectionBackColor = accentBlue,
                SelectionForeColor = Color.White,
                Font = gridFont,
                Padding = new Padding(8, 0, 8, 0),
                Alignment = DataGridViewContentAlignment.MiddleLeft,
                WrapMode = DataGridViewTriState.False,
                NullValue = ""
            };

            advancedDataGridView.AlternatingRowsDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = bgAlt,
                ForeColor = textPrimary,
                SelectionBackColor = accentBlueSoft,
                SelectionForeColor = Color.White,
                Font = gridFont,
                Padding = new Padding(8, 0, 8, 0),
                Alignment = DataGridViewContentAlignment.MiddleLeft,
                WrapMode = DataGridViewTriState.False
            };

            advancedDataGridView.ColumnHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = bgHeader,
                ForeColor = accentGold,
                SelectionBackColor = bgHeader,
                SelectionForeColor = accentGold,
                Font = headerFont,
                Padding = new Padding(8, 0, 8, 0),
                Alignment = DataGridViewContentAlignment.MiddleLeft,
                WrapMode = DataGridViewTriState.True
            };

            advancedDataGridView.RowHeadersDefaultCellStyle = new DataGridViewCellStyle
            {
                BackColor = bgRowHeader,
                ForeColor = accentGoldSoft,
                SelectionBackColor = accentRedSoft,
                SelectionForeColor = Color.White,
                Font = gridFont,
                Alignment = DataGridViewContentAlignment.MiddleCenter,
                Padding = new Padding(0),
                WrapMode = DataGridViewTriState.False
            };

            advancedDataGridView.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;
            advancedDataGridView.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single;

            try
            {
                if (advancedDataGridView.IsHandleCreated)
                    SetWindowTheme(advancedDataGridView.Handle, "DarkMode_Explorer", null);
            }
            catch
            {
            }
        }

        private void ApplyDarkChrome()
        {
            Color formBg = Color.FromArgb(24, 28, 34);
            Color panelBg = Color.FromArgb(32, 37, 44);
            Color controlBg = Color.FromArgb(38, 44, 53);
            Color textPrimary = Color.FromArgb(230, 232, 235);
            Color accentBlue = Color.FromArgb(59, 130, 246);
            Color accentGold = Color.FromArgb(245, 182, 66);
            Color accentRed = Color.FromArgb(220, 88, 88);

            this.BackColor = formBg;
            this.ForeColor = textPrimary;

            if (menuStrip != null)
            {
                menuStrip.BackColor = formBg;
                menuStrip.ForeColor = textPrimary;
                menuStrip.RenderMode = ToolStripRenderMode.Professional;
                menuStrip.Renderer = new DarkToolStripRenderer();
                ApplyToolStripTheme(menuStrip.Items, textPrimary);
                menuStrip.Invalidate();
            }

            if (contextMenuStrip != null)
            {
                contextMenuStrip.RenderMode = ToolStripRenderMode.Professional;
                contextMenuStrip.Renderer = new DarkToolStripRenderer();
                contextMenuStrip.BackColor = Color.FromArgb(30, 34, 40);
                contextMenuStrip.ForeColor = textPrimary;
                ApplyToolStripTheme(contextMenuStrip.Items, textPrimary);
            }

            if (filecontextMenuStrip != null)
            {
                filecontextMenuStrip.RenderMode = ToolStripRenderMode.Professional;
                filecontextMenuStrip.Renderer = new DarkToolStripRenderer();
                filecontextMenuStrip.BackColor = Color.FromArgb(30, 34, 40);
                filecontextMenuStrip.ForeColor = textPrimary;
                ApplyToolStripTheme(filecontextMenuStrip.Items, textPrimary);
            }

            if (lbFiles != null)
            {
                lbFiles.BackColor = panelBg;
                lbFiles.ForeColor = textPrimary;
                lbFiles.BorderStyle = BorderStyle.FixedSingle;

                try
                {
                    if (lbFiles.IsHandleCreated)
                        SetWindowTheme(lbFiles.Handle, "DarkMode_Explorer", null);
                }
                catch
                {
                }
            }

            if (txtFilter != null)
            {
                txtFilter.BackColor = controlBg;
                txtFilter.ForeColor = textPrimary;
                txtFilter.BorderStyle = BorderStyle.FixedSingle;
            }

            if (cbBuild != null)
                SetupDarkComboBox(cbBuild, controlBg, textPrimary);

            if (cbColumnMode != null)
                SetupDarkComboBox(cbColumnMode, controlBg, textPrimary);

            if (btnReset != null)
            {
                btnReset.BackColor = accentGold;
                btnReset.ForeColor = Color.FromArgb(28, 28, 28);
                btnReset.FlatStyle = FlatStyle.Flat;
                btnReset.FlatAppearance.BorderColor = Color.FromArgb(191, 145, 39);
                btnReset.FlatAppearance.MouseOverBackColor = Color.FromArgb(255, 196, 84);
                btnReset.FlatAppearance.MouseDownBackColor = Color.FromArgb(220, 162, 58);
            }

            if (lblCurrentProcess != null)
                lblCurrentProcess.ForeColor = accentBlue;

            if (txtStats != null)
            {
                txtStats.BackColor = formBg;
                txtStats.ForeColor = accentGold;
                txtStats.BorderStyle = BorderStyle.None;
            }

            if (txtCurrentCell != null)
            {
                txtCurrentCell.BackColor = formBg;
                txtCurrentCell.ForeColor = accentRed;
                txtCurrentCell.BorderStyle = BorderStyle.None;
            }

            if (txtCurEntry != null)
            {
                txtCurEntry.BackColor = Color.FromArgb(38, 44, 53);
                txtCurEntry.ForeColor = textPrimary;
                txtCurEntry.BorderStyle = BorderStyle.FixedSingle;
            }

            if (txtCurDefinition != null)
            {
                txtCurDefinition.BackColor = Color.FromArgb(38, 44, 53);
                txtCurDefinition.ForeColor = textPrimary;
                txtCurDefinition.BorderStyle = BorderStyle.FixedSingle;
            }

            if (gbSettings != null)
            {
                gbSettings.BackColor = formBg;
                gbSettings.ForeColor = accentGold;
            }

            if (gbFilter != null)
            {
                gbFilter.BackColor = formBg;
                gbFilter.ForeColor = accentGold;
            }

            if (columnFilter != null)
            {
                columnFilter.BackColor = panelBg;
                columnFilter.ForeColor = textPrimary;
            }
        }

        private void ApplyReadableSizing()
        {
            Font uiFont = new Font("Segoe UI", 12.0F, FontStyle.Regular);
            Font uiBold = new Font("Segoe UI Semibold", 12.0F, FontStyle.Bold);
            Font menuFont = new Font("Segoe UI", 11.25F, FontStyle.Regular);

            this.Font = uiFont;
            this.MinimumSize = new Size(1040, 720);

            if (menuStrip != null)
            {
                menuStrip.Font = menuFont;
                menuStrip.Padding = new Padding(8, 5, 0, 5);
                menuStrip.ImageScalingSize = new Size(16, 16);
                menuStrip.Height = 32;
            }

            if (contextMenuStrip != null)
            {
                contextMenuStrip.Font = menuFont;
                contextMenuStrip.ImageScalingSize = new Size(16, 16);
            }

            if (filecontextMenuStrip != null)
            {
                filecontextMenuStrip.Font = menuFont;
                filecontextMenuStrip.ImageScalingSize = new Size(16, 16);
            }

            if (advancedDataGridView != null)
            {
                advancedDataGridView.Font = uiFont;
                advancedDataGridView.DefaultCellStyle.Font = uiFont;
                advancedDataGridView.AlternatingRowsDefaultCellStyle.Font = uiFont;
                advancedDataGridView.ColumnHeadersDefaultCellStyle.Font = uiBold;
                advancedDataGridView.RowHeadersDefaultCellStyle.Font = uiFont;
            }

            if (lbFiles != null)
                lbFiles.Font = uiFont;

            if (txtFilter != null)
                txtFilter.Font = uiFont;

            if (cbBuild != null)
            {
                cbBuild.Font = uiFont;
                cbBuild.ItemHeight = 26;
                cbBuild.Height = 30;
                cbBuild.IntegralHeight = false;
                cbBuild.DropDownHeight = 300;
            }

            if (cbColumnMode != null)
            {
                cbColumnMode.Font = uiFont;
                cbColumnMode.ItemHeight = 26;
                cbColumnMode.Height = 30;
                cbColumnMode.IntegralHeight = false;
                cbColumnMode.DropDownHeight = 300;
            }

            if (btnReset != null)
            {
                btnReset.Font = uiFont;
                btnReset.Height = 32;
                if (btnReset.Width < 82)
                    btnReset.Width = 82;
            }

            if (txtCurEntry != null)
                txtCurEntry.Font = uiFont;

            if (txtCurDefinition != null)
                txtCurDefinition.Font = uiFont;

            if (txtStats != null)
                txtStats.Font = uiFont;

            if (txtCurrentCell != null)
                txtCurrentCell.Font = uiFont;

            if (lblCurrentProcess != null)
                lblCurrentProcess.Font = uiFont;

            if (gbSettings != null)
                gbSettings.Font = uiBold;

            if (gbFilter != null)
                gbFilter.Font = uiBold;

            if (columnFilter != null)
                columnFilter.Font = uiFont;
        }

        private void ApplyToolStripTheme(ToolStripItemCollection items, Color textColor)
        {
            foreach (ToolStripItem item in items)
            {
                item.ForeColor = textColor;
                item.BackColor = Color.FromArgb(30, 34, 40);

                if (item is ToolStripMenuItem menuItem && menuItem.HasDropDownItems)
                    ApplyToolStripTheme(menuItem.DropDownItems, textColor);
            }
        }

        private void SetupDarkComboBox(ComboBox combo, Color backColor, Color foreColor)
        {
            combo.BackColor = backColor;
            combo.ForeColor = foreColor;
            combo.FlatStyle = FlatStyle.Flat;
            combo.DrawMode = DrawMode.OwnerDrawFixed;
            combo.DropDownStyle = ComboBoxStyle.DropDownList;
            combo.DrawItem -= DarkComboBox_DrawItem;
            combo.DrawItem += DarkComboBox_DrawItem;
        }

        private void DarkComboBox_DrawItem(object sender, DrawItemEventArgs e)
        {
            ComboBox combo = sender as ComboBox;
            if (combo == null)
                return;

            Color back = Color.FromArgb(38, 44, 53);
            Color hover = Color.FromArgb(45, 52, 63);
            Color text = Color.FromArgb(230, 232, 235);
            Color accent = Color.FromArgb(245, 182, 66);

            if (e.Index < 0)
            {
                using (SolidBrush brush = new SolidBrush(back))
                    e.Graphics.FillRectangle(brush, e.Bounds);
                return;
            }

            bool selected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;

            using (SolidBrush brush = new SolidBrush(selected ? hover : back))
                e.Graphics.FillRectangle(brush, e.Bounds);

            Rectangle textBounds = new Rectangle(
                e.Bounds.X + 6,
                e.Bounds.Y,
                e.Bounds.Width - 12,
                e.Bounds.Height);

            string displayText = combo.GetItemText(combo.Items[e.Index]);
            TextRenderer.DrawText(
                e.Graphics,
                displayText,
                combo.Font,
                textBounds,
                selected ? accent : text,
                TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis);

            e.DrawFocusRectangle();
        }

        private void ApplyModernGridSizing()
        {
            advancedDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            advancedDataGridView.ColumnHeadersHeight = 42;
            advancedDataGridView.RowTemplate.Height = 32;

            if (_isLargeDataset)
                return;

            foreach (DataGridViewRow row in advancedDataGridView.Rows)
                row.Height = 32;
        }

        private void ApplyMinimumHeaderWidths()
        {
            if (advancedDataGridView.Columns.Count == 0)
                return;

            Font headerFont = advancedDataGridView.ColumnHeadersDefaultCellStyle.Font ?? advancedDataGridView.Font;
            DataGridViewAutoSizeColumnsMode oldMode = advancedDataGridView.AutoSizeColumnsMode;

            // We temporarily force manual sizing while applying minimums so width writes do not fight
            // autosize logic or custom header/filter painting.
            advancedDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;

            foreach (DataGridViewColumn col in advancedDataGridView.Columns)
            {
                if (!col.Visible)
                    continue;

                string headerText = col.HeaderText ?? string.Empty;

                int textWidth = TextRenderer.MeasureText(headerText + "  ", headerFont).Width;

                // Room for text + padding + sort/filter/header glyphs.
                int minWidth = textWidth + 40;

                if (minWidth < 70)
                    minWidth = 70;

                col.MinimumWidth = minWidth;

                if (col.Width < minWidth)
                    col.Width = minWidth;
            }

            advancedDataGridView.AutoSizeColumnsMode = oldMode;
        }

        private void Main_Load(object sender, EventArgs e)
        {
#if DEBUG
            wdb5ParserToolStripMenuItem.Visible = true;
#endif

            if (!Directory.Exists(TEMP_FOLDER))
                Directory.CreateDirectory(TEMP_FOLDER);

            openFileDialog.Filter = string.Join("|", SupportedFileTypes.Select(x => string.Format("{0} ({1})|{1}", x.Key, x.Value)));

            Parallel.ForEach(this.Controls.Cast<Control>(), c => c.KeyDown += new KeyEventHandler(KeyDownEvent));

            Task.Run(Database.LoadDefinitions)
                .ContinueWith(x =>
                {
                    Task.Run(UpdateManager.CheckForUpdate).ContinueWith(y => Watcher(), TaskScheduler.FromCurrentSynchronizationContext());
                    AutoRun();
                },
                TaskScheduler.FromCurrentSynchronizationContext());

            InstanceManager.AutoRunAdded += delegate
            {
                this.Invoke((MethodInvoker)delegate
                {
                    InstanceManager.FlashWindow(this);
                    AutoRun();
                });
            };

            LoadColumnSizeDropdown();
            LoadRecentList();
            LoadUiPreferences();
            EnsureNewWindowMenu();

            this.Text = string.Format("WDBX Editor ({0})", VERSION);

            ApplyThemeAndLayout();
            ApplyUiPreferences();

            try
            {
                if (advancedDataGridView.IsHandleCreated)
                    SetWindowTheme(advancedDataGridView.Handle, "DarkMode_Explorer", null);
            }
            catch
            {
            }
        }

        private void Main_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (Database.Entries.Count(x => x.Changed) > 0)
                if (MessageBox.Show("You have unsaved changes. Do you wish to exit?", "Unsaved Changes", MessageBoxButtons.YesNo) == DialogResult.No)
                    e.Cancel = true;

            if (!e.Cancel)
            {
                SaveUiPreferences();

                try { Directory.Delete(TEMP_FOLDER, true); } catch { }

                ProgressBarHandle(false, "", false);
                InstanceManager.Stop();
                watcher.EnableRaisingEvents = false;
                FormHandler.Close();
            }
        }

        private void SetSource(DBEntry dt, bool resetcolumns = true)
        {
            if (dt?.Header.IsTypeOf<HTFX>() == true && lbFiles.Items.Count > 1)
            {
                if (new LoadHotfix().ShowDialog(this) != DialogResult.OK)
                    return;
            }

            if (!ReferenceEquals(LoadedEntry, dt))
            {
                _changedCellKeys.Clear();
                _changedRowKeys.Clear();
                _highlightedRowKey = string.Empty;
            }

            advancedDataGridView.RowHeadersVisible = false;
            advancedDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            advancedDataGridView.ColumnHeadersVisible = false;
            advancedDataGridView.SuspendLayout();

            if (_bindingsource.IsSorted)
                _bindingsource.RemoveSort();

            if (!string.IsNullOrWhiteSpace(_bindingsource.Filter))
                _bindingsource.RemoveFilter();

            advancedDataGridView.Columns.Clear();
            _bindingsource.DataSource = null;
            _bindingsource.Clear();

            if (dt != null)
            {
                this.Tag = dt.Tag;
                this.Text = string.Format("WDBX Editor ({0}) - {1} {2}", VERSION, dt.FileName, dt.BuildName);
                LoadedEntry = dt;
                UpdateDatasetPerformanceMode();

                _bindingsource.DataSource = dt.Data;
                _bindingsource.ResetBindings(true);

                columnFilter.Reset(dt.Data.Columns, resetcolumns);
                advancedDataGridView.Columns[LoadedEntry.Key].ReadOnly = true;
                advancedDataGridView.ClearSelection();

                if (advancedDataGridView.Rows.Count > 0 && advancedDataGridView.Columns.Count > 0)
                    advancedDataGridView.CurrentCell = advancedDataGridView.Rows[0].Cells[0];

                txtStats.Text = string.Format("{0} fields, {1} rows{2}", LoadedEntry.Data.Columns.Count, LoadedEntry.Data.Rows.Count, _isLargeDataset ? "  |  Performance mode" : string.Empty);
                wotLKItemFixToolStripMenuItem.Enabled = LoadedEntry.IsFileOf("Item", Expansion.WotLK);
                colourPickerToolStripMenuItem.Enabled = (LoadedEntry.IsFileOf("LightIntBand") || LoadedEntry.IsFileOf("LightData"));
                if (!colourPickerToolStripMenuItem.Enabled)
                    FormHandler.Close<ColourConverter>();
            }
            else
            {
                this.Text = string.Format("WDBX Editor ({0})", VERSION);
                this.Tag = string.Empty;
                LoadedEntry = null;
                _isLargeDataset = false;

                txtStats.Text = txtCurEntry.Text = txtCurDefinition.Text = string.Empty;
                columnFilter.Reset(null, true);
                FormHandler.Close();
            }

            advancedDataGridView.ClearCopyData();
            advancedDataGridView.ClearChanges();
            pasteToolStripMenuItem.Enabled = false;
            undoToolStripMenuItem.Enabled = false;
            redoToolStripMenuItem.Enabled = false;
        }

        #region Data Grid

        private void advancedDataGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            advancedDataGridView.RowHeadersVisible = true;
            advancedDataGridView.ColumnHeadersVisible = true;
            advancedDataGridView.ResumeLayout(false);

            ApplyModernGridSizing();
            ApplyMinimumHeaderWidths();

            DataGridViewAutoSizeColumnsMode requestedMode = DataGridViewAutoSizeColumnsMode.None;
            if (cbColumnMode.SelectedItem is KeyValuePair<string, DataGridViewAutoSizeColumnsMode> selectedItem)
                requestedMode = selectedItem.Value;

            DataGridViewAutoSizeColumnsMode effectiveMode = GetEffectiveAutoSizeMode(requestedMode);
            advancedDataGridView.AutoSizeColumnsMode = effectiveMode;

            if (effectiveMode == DataGridViewAutoSizeColumnsMode.None)
                ApplyStoredColumnWidthsForEntry();

            try
            {
                if (advancedDataGridView.IsHandleCreated)
                    SetWindowTheme(advancedDataGridView.Handle, "DarkMode_Explorer", null);
            }
            catch
            {
            }

            advancedDataGridView.Invalidate();
            ProgressBarHandle(false);
        }

        private void advancedDataGridView_FilterStringChanged(object sender, EventArgs e)
        {
            _bindingsource.Filter = advancedDataGridView.FilterString;
        }

        private void advancedDataGridView_SortStringChanged(object sender, EventArgs e)
        {
            _bindingsource.Sort = advancedDataGridView.SortString;
        }

        private void advancedDataGridView_CurrentCellChanged(object sender, EventArgs e)
        {
            var cell = advancedDataGridView.CurrentCell;
            txtCurrentCell.Text = $"X: {cell?.ColumnIndex ?? 0}, Y: {cell?.RowIndex ?? 0}";
        }

        private void advancedDataGridView_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {
            if (IsLoaded && LoadedEntry.Data != null)
                txtStats.Text = string.Format("{0} fields, {1} rows{2}", LoadedEntry.Data.Columns.Count, LoadedEntry.Data.Rows.Count, _isLargeDataset ? "  |  Performance mode" : string.Empty);
        }

        private void advancedDataGridView_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            if (IsLoaded && LoadedEntry.Data != null)
                txtStats.Text = string.Format("{0} fields, {1} rows{2}", LoadedEntry.Data.Columns.Count, LoadedEntry.Data.Rows.Count, _isLargeDataset ? "  |  Performance mode" : string.Empty);
        }

        private void columnFilter_ItemCheckChanged(object sender, ItemCheckEventArgs e)
        {
            advancedDataGridView.SetVisible(e.Index, (e.NewValue == CheckState.Checked));
        }

        private void columnFilter_HideEmptyPressed(object sender, EventArgs e)
        {
            if (!IsLoaded)
                return;

            foreach (var c in advancedDataGridView.GetEmptyColumns())
                columnFilter.SetItemChecked(c, false);
        }

        private void cbColumnMode_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!(cbColumnMode.SelectedItem is KeyValuePair<string, DataGridViewAutoSizeColumnsMode> item))
                return;

            DataGridViewAutoSizeColumnsMode effectiveMode = GetEffectiveAutoSizeMode(item.Value);
            advancedDataGridView.AutoSizeColumnsMode = effectiveMode;

            if (effectiveMode == DataGridViewAutoSizeColumnsMode.None)
                ApplyStoredColumnWidthsForEntry();

            if (_applyingUiPreferences)
                return;

            _uiPreferences.LastColumnModeKey = item.Key;
            SaveUiPreferences();
        }

        private void advancedDataGridView_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0 && e.ColumnIndex >= 0 && IsLoaded)
            {
                string changedCellKey = GetChangedCellKey(e.RowIndex, e.ColumnIndex);
                string changedRowKey = GetRowIdentity(advancedDataGridView.Rows[e.RowIndex]);

                if (!string.IsNullOrEmpty(changedCellKey))
                    _changedCellKeys.Add(changedCellKey);

                if (!string.IsNullOrEmpty(changedRowKey))
                    _changedRowKeys.Add(changedRowKey);

                advancedDataGridView.InvalidateRow(e.RowIndex);
            }

            if (!LoadedEntry.Changed)
            {
                LoadedEntry.Changed = true;
                UpdateListBox();
            }
        }

        private void advancedDataGridView_UndoRedoChanged(object sender, EventArgs e)
        {
            undoToolStripMenuItem.Enabled = advancedDataGridView.CanUndo;
            redoToolStripMenuItem.Enabled = advancedDataGridView.CanRedo;
        }

        private void advancedDataGridView_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void advancedDataGridView_DragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            foreach (string file in files)
                if (Regex.IsMatch(file, Constants.FileRegexPattern, RegexOptions.Compiled | RegexOptions.IgnoreCase))
                    InstanceManager.AutoRun.Enqueue(file);

            AutoRun();
        }
        #endregion

        #region Data Grid Context
        private void advancedDataGridView_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right)
            {
                if (contextMenuStrip.Visible)
                {
                    contextMenuStrip.Tag = null;
                    contextMenuStrip.Hide();
                }

                return;
            }

            DataGridView.HitTestInfo info = advancedDataGridView.HitTest(e.X, e.Y);

            if (info.Type != DataGridViewHitTestType.RowHeader &&
                info.Type != DataGridViewHitTestType.Cell)
            {
                if (contextMenuStrip.Visible)
                {
                    contextMenuStrip.Tag = null;
                    contextMenuStrip.Hide();
                }

                return;
            }

            if (advancedDataGridView.IsCurrentCellInEditMode)
            {
                bool sameCell =
                    info.Type == DataGridViewHitTestType.Cell &&
                    advancedDataGridView.CurrentCell != null &&
                    info.RowIndex == advancedDataGridView.CurrentCell.RowIndex &&
                    info.ColumnIndex == advancedDataGridView.CurrentCell.ColumnIndex;

                if (!sameCell && !TryCommitPendingGridEdit())
                    return;
            }

            if (info.Type == DataGridViewHitTestType.RowHeader && info.RowIndex >= 0)
            {
                advancedDataGridView.ClearSelection();
                advancedDataGridView.SelectRow(info.RowIndex);
                contextMenuStrip.Tag = null;
                viewInEditorToolStripMenuItem.Enabled = false;
                contextMenuStrip.Show(Cursor.Position);
                return;
            }

            if (info.RowIndex < 0 || info.ColumnIndex < 0)
                return;

            DataGridViewCell targetCell = advancedDataGridView.Rows[info.RowIndex].Cells[info.ColumnIndex];

            contextMenuStrip.Tag = targetCell;

            advancedDataGridView.ClearSelection();
            targetCell.Selected = true;

            viewInEditorToolStripMenuItem.Enabled = true;
            contextMenuStrip.Show(Cursor.Position);
        }

        private void copyToolStripMenuItem_Click(object sender, EventArgs e)
        {
            advancedDataGridView.SetCopyData();
            pasteToolStripMenuItem.Enabled = true;
        }

        private void pasteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataRowView row;
            if (advancedDataGridView.SelectedRows.Count > 0)
                row = ((DataRowView)advancedDataGridView.CurrentRow.DataBoundItem);
            else if (advancedDataGridView.SelectedCells.Count > 0)
                row = ((DataRowView)advancedDataGridView.CurrentCell.OwningRow.DataBoundItem);
            else
                return;

            if (row?.Row != null)
            {
                advancedDataGridView.PasteCopyData(row.Row);
            }
            else
            {
                _bindingsource.EndEdit();
                advancedDataGridView.NotifyCurrentCellDirty(true);
                advancedDataGridView.EndEdit();
                advancedDataGridView.NotifyCurrentCellDirty(false);

                row = ((DataRowView)advancedDataGridView.CurrentRow.DataBoundItem);
                if (row?.Row != null)
                    advancedDataGridView.PasteCopyData(row.Row);

                if (!LoadedEntry.Changed)
                {
                    LoadedEntry.Changed = true;
                    UpdateListBox();
                }
            }
        }

        private void gotoIdToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GotoLine();
        }

        private void insertLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertLine();
        }

        private void clearLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DefaultRowValues();

            if (!LoadedEntry.Changed)
            {
                LoadedEntry.Changed = true;
                UpdateListBox();
            }
        }

        private void deleteLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded) return;

            if (advancedDataGridView.SelectedRows.Count == 0 && advancedDataGridView.SelectedCells.Count == 0)
                return;

            if (advancedDataGridView.SelectedRows.Count == 0)
                advancedDataGridView.SelectRow(advancedDataGridView.CurrentCell.OwningRow.Index);

            SendKeys.Send("{delete}");
            if (!LoadedEntry.Changed)
            {
                LoadedEntry.Changed = true;
                UpdateListBox();
            }
        }

        private void viewInEditorToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DataGridViewCell cell = contextMenuStrip.Tag as DataGridViewCell;
            if (cell == null)
                return;

            if (!TryCommitPendingGridEdit())
                return;

            try
            {
                if (!ReferenceEquals(advancedDataGridView.CurrentCell, cell))
                    advancedDataGridView.CurrentCell = cell;
            }
            catch
            {
                return;
            }

            using (var form = new TextEditor())
            {
                form.CellValue = Convert.ToString(cell.Value) ?? string.Empty;

                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    advancedDataGridView.BeginEdit(false);
                    cell.Value = form.CellValue;
                    advancedDataGridView.EndEdit();
                }
            }
        }
        #endregion

        #region Menu Items

        private void loadFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (openFileDialog.ShowDialog(this) == DialogResult.OK)
            {
                using (var loaddefs = new LoadDefinition())
                {
                    loaddefs.Files = openFileDialog.FileNames;
                    if (loaddefs.ShowDialog(this) != DialogResult.OK)
                        return;
                }

                ProgressBarHandle(true, "Loading files...");
                Task.Run(() => Database.LoadFiles(openFileDialog.FileNames))
                .ContinueWith(x =>
                {
                    if (x.Result.Count > 0)
                        new ErrorReport(x.Result).ShowDialog(this);

                    UpdateRecentList(openFileDialog.FileNames);
                    LoadFiles(openFileDialog.FileNames);
                    ProgressBarHandle(false);

                }, TaskScheduler.FromCurrentSynchronizationContext());
            }
        }

        private void loadRecentFilesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string[] files = new string[] { ((ToolStripMenuItem)sender).Tag.ToString() };

            using (var loaddefs = new LoadDefinition())
            {
                loaddefs.Files = files;
                if (loaddefs.ShowDialog(this) != DialogResult.OK)
                    return;
            }

            ProgressBarHandle(true, "Loading files...");
            Task.Run(() => Database.LoadFiles(files))
            .ContinueWith(x =>
            {
                if (x.Result.Count > 0)
                    new ErrorReport(x.Result).ShowDialog(this);

                UpdateRecentList(files);
                LoadFiles(files);
                ProgressBarHandle(false);

            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void openFromMPQToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var mpq = new LoadMPQ())
            {
                if (mpq.ShowDialog(this) == DialogResult.OK)
                {
                    using (var loaddefs = new LoadDefinition())
                    {
                        loaddefs.Files = mpq.Streams.Keys;
                        if (loaddefs.ShowDialog(this) != DialogResult.OK)
                            return;
                    }

                    ProgressBarHandle(true, "Loading files...");
                    Task.Run(() => Database.LoadFiles(mpq.Streams))
                    .ContinueWith(x =>
                    {
                        if (x.Result.Count > 0)
                            new ErrorReport(x.Result).ShowDialog(this);

                        LoadFiles(mpq.Streams.Keys);
                        ProgressBarHandle(false);
                    }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        private void openFromCASCToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var mpq = new LoadMPQ())
            {
                mpq.IsMPQ = false;

                if (mpq.ShowDialog(this) == DialogResult.OK)
                {
                    using (var loaddefs = new LoadDefinition())
                    {
                        loaddefs.Files = mpq.FileNames.Values;
                        if (loaddefs.Files.Count() == 0)
                            loaddefs.Files = mpq.Streams.Keys;

                        if (loaddefs.ShowDialog(this) != DialogResult.OK)
                            return;
                    }

                    ProgressBarHandle(true, "Loading files...");
                    Task.Run(() => Database.LoadFiles(mpq.Streams))
                    .ContinueWith(x =>
                    {
                        if (x.Result.Count > 0)
                            new ErrorReport(x.Result).ShowDialog(this);

                        LoadFiles(mpq.Streams.Keys);
                        ProgressBarHandle(false);
                    }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        private void saveToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded) return;
            SaveFile(false);
        }

        private void saveAsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded) return;
            SaveFile();
        }

        private void saveAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveAll();
        }

        private void editDefinitionsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new EditDefinition().ShowDialog(this);
        }

        private void findToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Find();
        }

        private void replaceToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Replace();
        }

        private void reloadToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Reload();
        }

        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseFile();
        }

        private void closeAllToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CloseAllFiles();
        }

        private void undoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Undo();
        }

        private void redoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Redo();
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new About().ShowDialog(this);
        }

        private void helpToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Help.ShowHelp(this, Path.Combine(Path.GetDirectoryName(Application.ExecutablePath), "Help.chm"));
        }

        private void insertToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertLine();
        }

        private void newLineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            NewLine();
        }

        private void playerLocationRecorderToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormHandler.Show<PlayerLocation>();
        }

        private void colourPickerToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormHandler.Show<ColourConverter>();
        }

        #endregion

        #region Export Menu Items
        private void toSQLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded) return;

            using (var sql = new LoadSQL() { Entry = LoadedEntry, ConnectionOnly = true })
            {
                if (sql.ShowDialog(this) == DialogResult.OK)
                {
                    ProgressBarHandle(true, "Exporting to SQL...");
                    Task.Factory.StartNew(() => { LoadedEntry.ToSQLTable(sql.ConnectionString); })
                    .ContinueWith(x =>
                    {
                        if (x.IsFaulted)
                            MessageBox.Show("An error occured exporting to SQL.");
                        else
                            MessageBox.Show("Sucessfully exported to SQL.");

                        ProgressBarHandle(false);
                    }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        private void toSQLFileToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded) return;

            using (var sfd = new SaveFileDialog() { FileName = LoadedEntry.TableStructure.Name + ".sql", Filter = "SQL Files|*.sql" })
            {
                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    ProgressBarHandle(true, "Exporting to SQL file...");
                    Task.Factory.StartNew(() =>
                    {
                        using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create))
                        {
                            string sql = LoadedEntry.ToSQL();
                            byte[] data = Encoding.UTF8.GetBytes(sql);
                            fs.Write(data, 0, data.Length);
                        }
                    })
                    .ContinueWith(x =>
                    {
                        ProgressBarHandle(false);

                        if (x.IsFaulted)
                            MessageBox.Show($"Error generating SQL file {x.Exception.Message}");
                        else
                            MessageBox.Show($"File successfully exported to {sfd.FileName}");

                    }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        private void toCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded) return;

            using (var sfd = new SaveFileDialog())
            {
                sfd.FileName = LoadedEntry.TableStructure.Name + ".csv";
                sfd.Filter = "CSV files (*.csv)|*.csv|Text files (*.txt)|*.txt";

                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    ProgressBarHandle(true, "Exporting to CSV...");
                    Task.Factory.StartNew(() =>
                    {
                        using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create))
                        {
                            string sql = LoadedEntry.ToCSV();
                            byte[] data = Encoding.UTF8.GetBytes(sql);
                            fs.Write(data, 0, data.Length);
                        }
                    })
                    .ContinueWith(x =>
                    {
                        ProgressBarHandle(false);

                        if (x.IsFaulted)
                            MessageBox.Show($"Error generating CSV file {x.Exception.Message}");
                        else
                            MessageBox.Show($"File successfully exported to {sfd.FileName}");

                    }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        private void toMPQToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded) return;

            using (var sfd = new SaveFileDialog())
            {
                sfd.InitialDirectory = Path.GetDirectoryName(LoadedEntry.FilePath);
                sfd.OverwritePrompt = false;
                sfd.CheckFileExists = false;

                switch (Path.GetExtension(LoadedEntry.FilePath).ToLower().TrimStart('.'))
                {
                    case "dbc":
                    case "db2":
                        sfd.FileName = LoadedEntry.TableStructure.Name + ".mpq";
                        sfd.Filter = "MPQ Files|*.mpq";
                        break;
                    default:
                        MessageBox.Show("Only DBC and DB2 files can be saved to MPQ.");
                        return;
                }

                if (sfd.ShowDialog(this) == DialogResult.OK)
                    LoadedEntry.ToMPQ(sfd.FileName);
            }
        }

        private void toJSONToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded) return;

            using (var sfd = new SaveFileDialog())
            {
                sfd.FileName = LoadedEntry.TableStructure.Name + ".json";
                sfd.Filter = "JSON files (*.json)|*.json|Text files (*.txt)|*.txt";

                if (sfd.ShowDialog(this) == DialogResult.OK)
                {
                    ProgressBarHandle(true, "Exporting to JSON...");
                    Task.Factory.StartNew(() =>
                    {
                        using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create))
                        {
                            string sql = LoadedEntry.ToJSON();
                            byte[] data = Encoding.UTF8.GetBytes(sql);
                            fs.Write(data, 0, data.Length);
                        }
                    })
                    .ContinueWith(x =>
                    {
                        ProgressBarHandle(false);

                        if (x.IsFaulted)
                            MessageBox.Show($"Error generating JSON file {x.Exception.Message}");
                        else
                            MessageBox.Show($"File successfully exported to {sfd.FileName}");

                    }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        #endregion

        #region Import Menu Items
        private void fromCSVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded) return;

            using (var loadCsv = new LoadCSV() { Entry = LoadedEntry })
            {
                switch (loadCsv.ShowDialog(this))
                {
                    case DialogResult.OK:
                        SetSource(GetEntry(), false);
                        advancedDataGridView.CacheData();
                        MessageBox.Show("CSV import succeeded.");
                        break;
                    case DialogResult.Abort:
                        ProgressBarHandle(false);
                        if (!string.IsNullOrWhiteSpace(loadCsv.ErrorMessage))
                            MessageBox.Show("CSV import failed: " + loadCsv.ErrorMessage);
                        else
                            MessageBox.Show("CSV import failed due to incorrect file format.");
                        break;
                }

                ProgressBarHandle(false);
            }
        }

        private void fromSQLToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (!IsLoaded)
            {
                MessageBox.Show("Open a file first.");
                return;
            }

            using (var importSql = new LoadSQL() { Entry = LoadedEntry })
            {
                switch (importSql.ShowDialog(this))
                {
                    case DialogResult.OK:
                        SetSource(GetEntry(), false);
                        advancedDataGridView.CacheData();
                        MessageBox.Show("SQL import succeeded.");
                        break;
                    case DialogResult.Abort:
                        if (!string.IsNullOrWhiteSpace(importSql.ErrorMessage))
                            MessageBox.Show(importSql.ErrorMessage);
                        else
                            MessageBox.Show("SQL import failed due to incorrect file format.");
                        break;
                }

                ProgressBarHandle(false);
            }
        }

        #endregion

        #region Tool Menu Items
        private void wotLKItemFixToolStripMenuItem_Click(object sender, EventArgs e)
        {
            using (var itemfix = new WotLKItemFix())
            {
                itemfix.Entry = LoadedEntry;
                if (itemfix.ShowDialog(this) == DialogResult.OK)
                    SetSource(LoadedEntry);
            }
        }

        private void legionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            new LegionParser().ShowDialog(this);
        }
        #endregion

        #region File ListView
        private void closeToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DBEntry selection = (DBEntry)((DataRowView)lbFiles.SelectedItem)["Key"];
            if (LoadedEntry == selection)
                CloseFile();
            else
            {
                Database.Entries.Remove(selection);
                Database.Entries.TrimExcess();
                UpdateListBox();
            }
        }

        private void editToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DBEntry selection = (DBEntry)((DataRowView)lbFiles.SelectedItem)["Key"];
            if (LoadedEntry != selection)
                SetSource(selection);
        }

        private void lbFiles_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int index = lbFiles.IndexFromPoint(e.Location);
                if (index != ListBox.NoMatches)
                {
                    lbFiles.SelectedIndex = index;
                    filecontextMenuStrip.Show(Cursor.Position);
                }
            }
        }

        private void lbFiles_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            int index = lbFiles.IndexFromPoint(e.Location);
            if (index != ListBox.NoMatches)
            {
                DBEntry entry = (DBEntry)((DataRowView)lbFiles.Items[index])["Key"];
                txtCurEntry.Text = entry.FileName;
                txtCurDefinition.Text = entry.BuildName;

                SetSource(GetEntry());
            }
        }
        #endregion

        #region Command Actions
        private void LoadFiles(IEnumerable<string> fileNames)
        {
            UpdateListBox();

            if (lbFiles.Items.Count == 0)
                return;

            if (LoadedEntry != null && fileNames.Any(x => x.Equals(LoadedEntry.FileName, IGNORECASE)))
            {
                var entry = (DBEntry)lbFiles.SelectedValue;
                txtCurEntry.Text = entry.FileName;
                txtCurDefinition.Text = entry.BuildName;
                txtStats.Text = string.Format("{0} fields, {1} rows{2}", entry.Data.Columns.Count, entry.Data.Rows.Count, _isLargeDataset ? "  |  Performance mode" : string.Empty);

                SetSource(GetEntry());
            }

            if (GetEntry() == null)
            {
                LoadedEntry = null;
                SetSource(null);
            }

            if (LoadedEntry == null && lbFiles.Items.Count > 0)
            {
                lbFiles.SetSelected(0, true);

                var entry = (DBEntry)lbFiles.SelectedValue;
                txtCurEntry.Text = entry.FileName;
                txtCurDefinition.Text = entry.BuildName;
                txtStats.Text = string.Format("{0} fields, {1} rows{2}", entry.Data.Columns.Count, entry.Data.Rows.Count, _isLargeDataset ? "  |  Performance mode" : string.Empty);

                SetSource(GetEntry());
            }

            if (LoadedEntry != null)
                txtCurDefinition.Text = LoadedEntry.BuildName;
        }

        private void SaveFile(bool saveas = true)
        {
            if (!IsLoaded) return;
            bool save = !saveas;

            if (saveas)
            {
                using (var sfd = new SaveFileDialog())
                {
                    sfd.InitialDirectory = Path.GetDirectoryName(LoadedEntry.SavePath);
                    sfd.FileName = LoadedEntry.SavePath;

                    string ext = Path.GetExtension(LoadedEntry.FilePath).TrimStart('.');
                    switch (ext.ToLower())
                    {
                        case "dbc":
                            sfd.Filter = "DBC Files|*.dbc";
                            break;
                        case "db2":
                            sfd.Filter = "DB2 Files|*.db2";
                            break;
                        case "adb":
                            sfd.Filter = "ADB Files|*.adb";
                            break;
                        case "wdb":
                        case "bin":
                            MessageBox.Show($"Saving is not implemented for {ext.ToUpper()} files.");
                            return;
                    }

                    if (sfd.ShowDialog(this) == DialogResult.OK)
                    {
                        save = true;
                        LoadedEntry.SavePath = sfd.FileName;
                    }
                }
            }

            if (save)
            {
                ProgressBarHandle(true, "Saving file...");
                Task.Factory.StartNew(() => new DBReader().Write(LoadedEntry, LoadedEntry.SavePath))
                .ContinueWith(x =>
                {
                    ProgressBarHandle(false);
                    LoadedEntry.Changed = false;
                    UpdateListBox();

                    if (x.IsFaulted)
                        MessageBox.Show($"Error exporting to file {x.Exception.InnerException.Message}");

                }, TaskScheduler.FromCurrentSynchronizationContext());
            }
        }

        private void SaveAll()
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog(this) == DialogResult.OK)
                {
                    ProgressBarHandle(true, "Saving files...");

                    Task.Run(() => Database.SaveFiles(fbd.SelectedPath))
                        .ContinueWith(x =>
                        {
                            if (x.Result.Count > 0)
                                new ErrorReport(x.Result).ShowDialog(this);

                            ProgressBarHandle(false);
                        }, TaskScheduler.FromCurrentSynchronizationContext());
                }
            }
        }

        private void GotoLine()
        {
            if (!IsLoaded) return;

            int id = 0;
            string res = "";
            if (ShowInputDialog("Id:", "Go to Id", 0.ToString(), ref res) == DialogResult.OK)
            {
                if (int.TryParse(res, out id))
                {
                    int index = _bindingsource.Find(LoadedEntry.Key, id);
                    if (index >= 0)
                        advancedDataGridView.SelectRow(index);
                    else
                        MessageBox.Show($"Id {id} doesn't exist.");
                }
                else
                    MessageBox.Show($"Invalid Id.");
            }
        }

        private void Find()
        {
            if (IsLoaded)
                FormHandler.Show<FindReplace>(false);
        }

        private void Replace()
        {
            if (IsLoaded)
                FormHandler.Show<FindReplace>(true);
        }

        private void Reload()
        {
            if (!IsLoaded) return;

            ProgressBarHandle(true, "Reloading file...");
            Task.Run(() => Database.LoadFiles(new string[] { LoadedEntry.FilePath }))
            .ContinueWith(x =>
            {
                if (x.Result.Count > 0)
                    new ErrorReport(x.Result).ShowDialog(this);

                LoadFiles(openFileDialog.FileNames);
                ProgressBarHandle(false);

            }, TaskScheduler.FromCurrentSynchronizationContext());
        }

        private void CloseFile()
        {
            if (!string.IsNullOrWhiteSpace(_bindingsource.Filter))
                _bindingsource.RemoveFilter();

            if (_bindingsource.IsSorted)
                _bindingsource.RemoveSort();

            if (LoadedEntry != null)
            {
                LoadedEntry.Dispose();
                Database.Entries.Remove(LoadedEntry);
                Database.Entries.TrimExcess();
            }

            SetSource(null);
            UpdateListBox();
        }

        private void CloseAllFiles()
        {
            if (!string.IsNullOrWhiteSpace(_bindingsource.Filter))
                _bindingsource.RemoveFilter();

            if (_bindingsource.IsSorted)
                _bindingsource.RemoveSort();

            for (int i = 0; i < Database.Entries.Count; i++)
                Database.Entries[i].Dispose();

            Database.Entries.Clear();
            Database.Entries.TrimExcess();

            SetSource(null);
            UpdateListBox();
        }

        private void Undo()
        {
            advancedDataGridView.Undo();
        }

        private void Redo()
        {
            advancedDataGridView.Redo();
        }

        private void InsertLine()
        {
            if (!IsLoaded) return;

            string res = "";
            if (ShowInputDialog("Id:", "Id to insert", "1", ref res) == DialogResult.OK)
            {
                int keyIndex = advancedDataGridView.Columns[LoadedEntry.Key].Index;

                if (!int.TryParse(res, out int id) || id < 0)
                {
                    MessageBox.Show($"Invalid Id. Out of range of the column min/max value.");
                }
                else
                {
                    int index = _bindingsource.Find(LoadedEntry.Key, id);
                    if (index < 0)
                    {
                        index = NewLine();
                        advancedDataGridView.Rows[index].Cells[LoadedEntry.Key].Value = id;
                        DefaultRowValues(index);

                        advancedDataGridView.OnUserAddedRow(advancedDataGridView.Rows[index]);

                        if (!LoadedEntry.Changed)
                        {
                            LoadedEntry.Changed = true;
                            UpdateListBox();
                        }
                    }

                    advancedDataGridView.SelectRow(index);
                }
            }
        }

        private void DefaultRowValues(int index = -1)
        {
            if (!IsLoaded)
                return;

            if (advancedDataGridView.SelectedRows.Count == 1)
                index = advancedDataGridView.CurrentRow.Index;
            else if (advancedDataGridView.SelectedCells.Count == 1)
                index = advancedDataGridView.CurrentCell.OwningRow.Index;

            if (index == -1)
                return;

            for (int i = 0; i < advancedDataGridView.Columns.Count; i++)
            {
                if (advancedDataGridView.Columns[i].Name == LoadedEntry.Key)
                    continue;

                advancedDataGridView.Rows[index].Cells[i].Value = advancedDataGridView.Columns[i].ValueType.DefaultValue();

                if (!LoadedEntry.Changed)
                {
                    LoadedEntry.Changed = true;
                    UpdateListBox();
                }
            }
        }

        private int NewLine()
        {
            if (!IsLoaded) return 0;

            var row = LoadedEntry.Data.NewRow();
            LoadedEntry.Data.Rows.Add(row);
            int index = _bindingsource.Find(LoadedEntry.Key, row[LoadedEntry.Key]);
            DefaultRowValues(index);
            advancedDataGridView.SelectRow(index);
            return index;
        }
        #endregion

        #region File Filter
        private void LoadBuilds()
        {
            var tables = lbFiles.Items.Cast<DataRowView>()
                            .Select(x => ((DBEntry)x["Key"]).TableStructure)
                            .OrderBy(x => x.Build)
                            .Select(x => x.BuildText).Distinct();

            cbBuild.Items.Clear();
            cbBuild.Items.Add("");
            cbBuild.Items.AddRange(tables.ToArray());
        }

        private void txtFilter_TextChanged(object sender, EventArgs e)
        {
            ApplyFileFilter();
        }

        private void cbBuild_SelectedIndexChanged(object sender, EventArgs e)
        {
            ApplyFileFilter();
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            txtFilter.Text = "";
            cbBuild.Text = "";
        }
        #endregion

        private void LoadColumnSizeDropdown()
        {
            cbColumnMode.Items.Clear();
            cbColumnMode.Items.Add(new KeyValuePair<string, DataGridViewAutoSizeColumnsMode>("None", DataGridViewAutoSizeColumnsMode.None));
            cbColumnMode.Items.Add(new KeyValuePair<string, DataGridViewAutoSizeColumnsMode>("Column Header", DataGridViewAutoSizeColumnsMode.ColumnHeader));
            cbColumnMode.Items.Add(new KeyValuePair<string, DataGridViewAutoSizeColumnsMode>("Displayed Cells", DataGridViewAutoSizeColumnsMode.DisplayedCells));
            cbColumnMode.Items.Add(new KeyValuePair<string, DataGridViewAutoSizeColumnsMode>("Displayed Cells Except Header", DataGridViewAutoSizeColumnsMode.DisplayedCellsExceptHeader));

            cbColumnMode.ValueMember = "Value";
            cbColumnMode.DisplayMember = "Key";
            cbColumnMode.SelectedIndex = 0;
        }

        private void UpdateListBox()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Key", typeof(DBEntry));
            dt.Columns.Add("Value", typeof(string));

            var entries = Database.Entries.OrderBy(x => x.Build).ThenBy(x => x.FileName);
            foreach (var entry in entries)
                dt.Rows.Add(entry, $"{entry.FileName} - {entry.BuildName}{(entry.Changed ? "*" : "")}");

            lbFiles.BeginUpdate();
            lbFiles.DataSource = new BindingSource(dt, null);

            if (Database.Entries.Count > 0)
            {
                lbFiles.ValueMember = "Key";
                lbFiles.DisplayMember = "Value";
            }
            else
            {
                ((BindingSource)lbFiles.DataSource).DataSource = null;
                ((BindingSource)lbFiles.DataSource).Clear();
            }

            lbFiles.EndUpdate();

            LoadBuilds();
        }

        private void KeyDownEvent(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode == Keys.S)
                SaveFile(false);
            else if (e.Control && e.KeyCode == Keys.G)
                GotoLine();
            else if (e.Control && e.Shift && e.KeyCode == Keys.S)
                SaveAll();
            else if (e.Control && e.KeyCode == Keys.F)
                Find();
            else if (e.Control && e.KeyCode == Keys.H)
                Replace();
            else if (e.Control && e.KeyCode == Keys.R)
                Reload();
            else if (e.Control && e.KeyCode == Keys.W)
                CloseFile();
            else if (e.Control && e.KeyCode == Keys.Z)
                Undo();
            else if ((e.Control && e.Shift && e.KeyCode == Keys.Z) || (e.Control && e.KeyCode == Keys.Y))
                Redo();
            else if (e.Control && e.KeyCode == Keys.N)
                NewLine();
            else if (e.Control && e.KeyCode == Keys.I)
                InsertLine();
            else if (e.KeyCode == Keys.F12)
                SaveFile();
            else if (e.Control && e.Shift && e.KeyCode == Keys.W)
                CloseAllFiles();
        }

        public void ProgressBarHandle(bool start, string currentTask = "", bool clear = true)
        {
            if (start)
                progressBar.Start();
            else
                progressBar.Stop(clear);

            lblCurrentProcess.Text = currentTask;
            lblCurrentProcess.Visible = !string.IsNullOrWhiteSpace(currentTask) && start;

            menuStrip.Enabled = !start;
            columnFilter.Enabled = !start;
            gbSettings.Enabled = !start;
            gbFilter.Enabled = !start;
            advancedDataGridView.ReadOnly = start;
            advancedDataGridView.Refresh();
        }

        private void AutoRun()
        {
            if (InstanceManager.AutoRun.Any(x => File.Exists(x)))
            {
                IEnumerable<string> filenames = InstanceManager.GetFilesToOpen();

                var loaddef = FormHandler.GetForm<LoadDefinition>();
                if (loaddef != null)
                {
                    loaddef.UpdateFiles(filenames);
                    return;
                }

                using (var loaddefs = new LoadDefinition())
                {
                    loaddefs.Files = filenames;
                    if (loaddefs.ShowDialog(this) != DialogResult.OK)
                        return;
                    else
                        filenames = loaddefs.Files;
                }

                ProgressBarHandle(true, "Loading files...");
                Task.Run(() => Database.LoadFiles(filenames))
                .ContinueWith(x =>
                {
                    if (x.Result.Count > 0)
                        new ErrorReport(x.Result).ShowDialog(this);

                    LoadFiles(filenames);
                    ProgressBarHandle(false);

                }, TaskScheduler.FromCurrentSynchronizationContext());
            }
        }

        private void Watcher()
        {
            watcher = new FileSystemWatcher
            {
                Path = Path.GetDirectoryName(DEFINITION_DIR),
                NotifyFilter = NotifyFilters.LastWrite,
                Filter = "*.xml",
                EnableRaisingEvents = true
            };
            watcher.Changed += delegate { Task.Run(() => Database.LoadDefinitions()); };
        }

        private void LoadRecentList()
        {
            recentToolStripMenuItem.DropDownItems.Clear();

            if (Properties.Settings.Default.RecentFiles == null)
            {
                Properties.Settings.Default.RecentFiles = new System.Collections.Specialized.StringCollection();
                Properties.Settings.Default.Save();
                recentToolStripMenuItem.Visible = false;
                return;
            }

            recentToolStripMenuItem.Visible = Properties.Settings.Default.RecentFiles.Count > 0;

            foreach (var recent in Properties.Settings.Default.RecentFiles)
            {
                if (!File.Exists(recent))
                    continue;

                ToolStripMenuItem menuItem = new ToolStripMenuItem(recent, null, loadRecentFilesToolStripMenuItem_Click)
                {
                    Tag = recent,
                    DisplayStyle = ToolStripItemDisplayStyle.Text,
                    ForeColor = Color.FromArgb(230, 232, 235)
                };

                recentToolStripMenuItem.DropDownItems.Add(menuItem);
            }
        }

        private void UpdateRecentList(string[] files)
        {
            string[] recentTmp = files;
            Array.Resize(ref recentTmp, Properties.Settings.Default.RecentFiles.Count + recentTmp.Length);
            Properties.Settings.Default.RecentFiles.CopyTo(recentTmp, files.Length);

            var recentFiles = recentTmp.Distinct().Where(x => File.Exists(x)).Take(10);

            Properties.Settings.Default.RecentFiles.Clear();
            Properties.Settings.Default.RecentFiles.AddRange(recentFiles.ToArray());

            LoadRecentList();
        }
    }

    internal class UiPreferences
    {
        public string LastColumnModeKey { get; set; } = "None";
        public int WindowLeft { get; set; } = -1;
        public int WindowTop { get; set; } = -1;
        public int WindowWidth { get; set; }
        public int WindowHeight { get; set; }
        public bool WindowMaximized { get; set; }
        public Dictionary<string, Dictionary<string, int>> ColumnWidths { get; set; } =
            new Dictionary<string, Dictionary<string, int>>(StringComparer.OrdinalIgnoreCase);
    }


    public class DarkToolStripRenderer : ToolStripProfessionalRenderer
    {
        private static readonly Color Bg = Color.FromArgb(30, 34, 40);
        private static readonly Color BgTop = Color.FromArgb(24, 28, 34);
        private static readonly Color Hover = Color.FromArgb(45, 52, 63);
        private static readonly Color Pressed = Color.FromArgb(52, 60, 72);
        private static readonly Color Border = Color.FromArgb(60, 68, 80);
        private static readonly Color Text = Color.FromArgb(230, 232, 235);
        private static readonly Color Accent = Color.FromArgb(245, 182, 66);

        public DarkToolStripRenderer() : base(new DarkColorTable())
        {
            RoundedEdges = false;
        }

        protected override void OnRenderToolStripBackground(ToolStripRenderEventArgs e)
        {
            Color fill = e.ToolStrip is MenuStrip ? BgTop : Bg;
            using (SolidBrush brush = new SolidBrush(fill))
                e.Graphics.FillRectangle(brush, e.AffectedBounds);
        }

        protected override void OnRenderMenuItemBackground(ToolStripItemRenderEventArgs e)
        {
            Rectangle rect = new Rectangle(Point.Empty, e.Item.Size);
            bool isTopLevel = e.Item.Owner is MenuStrip;
            bool isPressedMenu = e.Item is ToolStripMenuItem menuItem && menuItem.DropDown.Visible;

            Color fill;

            if (isPressedMenu)
                fill = Pressed;
            else if (e.Item.Selected)
                fill = Hover;
            else if (isTopLevel)
                fill = BgTop;
            else
                fill = Bg;

            using (SolidBrush brush = new SolidBrush(fill))
                e.Graphics.FillRectangle(brush, rect);

            if (e.Item.Selected || isPressedMenu)
            {
                using (Pen pen = new Pen(Border))
                    e.Graphics.DrawRectangle(pen, 0, 0, rect.Width - 1, rect.Height - 1);
            }
        }

        protected override void OnRenderItemText(ToolStripItemTextRenderEventArgs e)
        {
            bool hot = e.Item.Selected || (e.Item is ToolStripMenuItem mi && mi.DropDown.Visible);
            e.TextColor = hot ? Accent : Text;
            base.OnRenderItemText(e);
        }

        protected override void OnRenderArrow(ToolStripArrowRenderEventArgs e)
        {
            bool hot = e.Item.Selected || (e.Item is ToolStripMenuItem mi && mi.DropDown.Visible);
            e.ArrowColor = hot ? Accent : Text;
            base.OnRenderArrow(e);
        }

        protected override void OnRenderImageMargin(ToolStripRenderEventArgs e)
        {
            using (SolidBrush brush = new SolidBrush(Bg))
                e.Graphics.FillRectangle(brush, e.AffectedBounds);
        }

        protected override void OnRenderSeparator(ToolStripSeparatorRenderEventArgs e)
        {
            int y = e.Item.Height / 2;
            using (Pen pen = new Pen(Border))
                e.Graphics.DrawLine(pen, 6, y, e.Item.Width - 6, y);
        }

        protected override void OnRenderToolStripBorder(ToolStripRenderEventArgs e)
        {
            if (e.ToolStrip is ToolStripDropDownMenu)
            {
                Rectangle rect = new Rectangle(Point.Empty, e.ToolStrip.Size);
                rect.Width -= 1;
                rect.Height -= 1;
                using (Pen pen = new Pen(Border))
                    e.Graphics.DrawRectangle(pen, rect);
            }
        }
    }

    public class DarkColorTable : ProfessionalColorTable
    {
        private readonly Color bg = Color.FromArgb(30, 34, 40);
        private readonly Color bgTop = Color.FromArgb(24, 28, 34);
        private readonly Color bgHover = Color.FromArgb(45, 52, 63);
        private readonly Color bgPressed = Color.FromArgb(52, 60, 72);
        private readonly Color border = Color.FromArgb(60, 68, 80);

        public override Color MenuStripGradientBegin => bgTop;
        public override Color MenuStripGradientEnd => bgTop;

        public override Color MenuItemSelected => bgHover;
        public override Color MenuItemSelectedGradientBegin => bgHover;
        public override Color MenuItemSelectedGradientEnd => bgHover;

        public override Color MenuItemPressedGradientBegin => bgPressed;
        public override Color MenuItemPressedGradientMiddle => bgPressed;
        public override Color MenuItemPressedGradientEnd => bgPressed;

        public override Color MenuItemBorder => border;

        public override Color ToolStripDropDownBackground => bg;
        public override Color ImageMarginGradientBegin => bg;
        public override Color ImageMarginGradientMiddle => bg;
        public override Color ImageMarginGradientEnd => bg;

        public override Color SeparatorDark => border;
        public override Color SeparatorLight => border;

        public override Color CheckBackground => bgHover;
        public override Color CheckSelectedBackground => bgHover;
        public override Color CheckPressedBackground => bgPressed;

        public override Color ButtonSelectedHighlight => bgHover;
        public override Color ButtonSelectedHighlightBorder => border;
        public override Color ButtonPressedHighlight => bgPressed;
        public override Color ButtonPressedHighlightBorder => border;
    }
}
