﻿using System.Collections.Generic;
using System.Windows.Forms;
using System;
using System.Linq;
using System.Drawing;
using System.ComponentModel;
using System.Text.RegularExpressions;

public partial class CustomGUIs
{
    public static List<Dictionary<string, object>> DataGrid(
        List<Dictionary<string, object>> entries,
        List<string> propertyNames,
        bool spanAllScreens,
        List<int> initialSelectionIndices = null)
    {
        List<Dictionary<string, object>> filteredEntries = new List<Dictionary<string, object>>();
        bool escapePressed = false;

        Form form = new Form();
        form.StartPosition = FormStartPosition.CenterScreen;
        form.Text = $"Total Entries: {entries.Count}";
        form.BackColor = Color.White;

        DataGridView dataGridView = new DataGridView();
        dataGridView.Dock = DockStyle.Fill;
        dataGridView.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
        dataGridView.MultiSelect = true;
        dataGridView.ReadOnly = true;
        dataGridView.AutoGenerateColumns = false;
        dataGridView.BackgroundColor = Color.White;
        dataGridView.AllowUserToAddRows = false;

        foreach (string propertyName in propertyNames)
        {
            dataGridView.Columns.Add(new DataGridViewTextBoxColumn
            {
                HeaderText = propertyName,
                DataPropertyName = propertyName
            });
        }

        foreach (Dictionary<string, object> entry in entries)
        {
            DataGridViewRow row = new DataGridViewRow();
            row.CreateCells(dataGridView);
            for (int i = 0; i < propertyNames.Count; i++)
            {
                row.Cells[i].Value = entry.ContainsKey(propertyNames[i]) ? entry[propertyNames[i]] : null;
            }
            dataGridView.Rows.Add(row);
        }

        // Helper function to get first visible column index
        int GetFirstVisibleColumnIndex()
        {
            var visibleColumn = dataGridView.Columns
                .Cast<DataGridViewColumn>()
                .FirstOrDefault(c => c.Visible);
            
            return visibleColumn?.Index ?? -1;
        }

        // Initial selection handling
        if (initialSelectionIndices != null && initialSelectionIndices.Any())
        {
            int firstVisibleCol = GetFirstVisibleColumnIndex();
            foreach (int index in initialSelectionIndices)
            {
                if (index >= 0 && index < dataGridView.Rows.Count && firstVisibleCol != -1)
                {
                    dataGridView.Rows[index].Selected = true;
                    dataGridView.CurrentCell = dataGridView.Rows[index].Cells[firstVisibleCol];
                }
            }
        }

        TextBox searchBox = new TextBox { Dock = DockStyle.Top };

        searchBox.TextChanged += (sender, e) =>
        {
            string searchText = searchBox.Text;

            // Parse search tokens
            var tokens = Regex.Matches(searchText,
                @"(\$""[^""]+"":""[^""]+""|\$[^ ]+?:""[^""]+""|\$""[^""]+"":[^ ]+|\$[^ ]+?:[^ ]+|\$""[^""]+""|\$[^ ]+|""[^""]+""|\S+)")
                .Cast<Match>()
                .Select(m => m.Value.Trim())
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList();

            List<List<string>> columnVisibilityFilters = new List<List<string>>();
            List<(List<string> colParts, string valFilter)> columnValueFilters = new List<(List<string>, string)>();
            List<string> generalFilters = new List<string>();

            foreach (string token in tokens)
            {
                if (token.StartsWith("$"))
                {
                    int colonIndex = token.IndexOf(':');
                    if (colonIndex > 0)
                    {
                        // Column value filter
                        string colPart = token.Substring(1, colonIndex - 1);
                        string valPart = token.Substring(colonIndex + 1);

                        // Process column parts
                        bool wasQuoted = colPart.StartsWith("\"") && colPart.EndsWith("\"");
                        string cleanCol = StripQuotes(colPart).ToLower();
                        List<string> colParts = wasQuoted
                            ? cleanCol.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList()
                            : new List<string> { cleanCol };

                        // Process value filter
                        string valFilter = StripQuotes(valPart).ToLower();

                        if (colParts.Count > 0 && !string.IsNullOrWhiteSpace(valFilter))
                        {
                            columnValueFilters.Add((colParts, valFilter));
                        }
                    }
                    else
                    {
                        // Column visibility filter
                        string content = token.Substring(1);
                        bool wasQuoted = content.StartsWith("\"") && content.EndsWith("\"");
                        string cleanContent = StripQuotes(content).ToLower();
                        List<string> parts = wasQuoted
                            ? cleanContent.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).ToList()
                            : new List<string> { cleanContent };

                        if (parts.Count > 0)
                        {
                            columnVisibilityFilters.Add(parts);
                        }
                    }
                }
                else
                {
                    // General row filter
                    string filter = StripQuotes(token).ToLower();
                    if (!string.IsNullOrWhiteSpace(filter))
                    {
                        generalFilters.Add(filter);
                    }
                }
            }

            // Apply column visibility
            foreach (DataGridViewColumn column in dataGridView.Columns)
            {
                string columnName = column.HeaderText.ToLower();
                bool shouldShow = columnVisibilityFilters.Count == 0 ||
                    columnVisibilityFilters.Any(filterParts =>
                        filterParts.All(part => columnName.Contains(part)));
                column.Visible = shouldShow;
            }

            // Filter entries
            List<Dictionary<string, object>> filtered = entries.Where(entry =>
            {
                // 1. Apply column value filters
                foreach (var (colParts, valFilter) in columnValueFilters)
                {
                    List<string> matchingColumns = propertyNames
                        .Where(p => colParts.All(part => p.ToLower().Contains(part)))
                        .ToList();

                    if (matchingColumns.Count == 0) return false;

                    bool valueFound = matchingColumns.Any(col =>
                        entry.ContainsKey(col) &&
                        entry[col]?.ToString().ToLower().Contains(valFilter) == true);

                    if (!valueFound) return false;
                }

                // 2. Apply general filters
                if (generalFilters.Count > 0)
                {
                    string joinedValues = string.Join(" ", entry.Values
                        .Where(v => v != null)
                        .Select(v => v.ToString().ToLower()));

                    bool hasExclusions = generalFilters.Any(f => f.StartsWith("!"));
                    bool hasInclusions = generalFilters.Any(f => !f.StartsWith("!"));

                    // Process inclusions
                    if (hasInclusions && !generalFilters
                        .Where(f => !f.StartsWith("!"))
                        .All(f => joinedValues.Contains(f)))
                    {
                        return false;
                    }

                    // Process exclusions
                    if (hasExclusions && generalFilters
                        .Where(f => f.StartsWith("!"))
                        .Select(f => f.Substring(1))
                        .Any(excl => joinedValues.Contains(excl)))
                    {
                        return false;
                    }
                }

                return true;
            }).ToList();

            // Update grid
            dataGridView.Rows.Clear();
            foreach (Dictionary<string, object> item in filtered)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(dataGridView);
                for (int i = 0; i < propertyNames.Count; i++)
                {
                    row.Cells[i].Value = item.ContainsKey(propertyNames[i]) ? item[propertyNames[i]] : null;
                }
                dataGridView.Rows.Add(row);
            }

            // Maintain UI
            dataGridView.AutoResizeColumns();
            int requiredWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                + SystemInformation.VerticalScrollBarWidth + 50;
            form.Width = Math.Min(requiredWidth, Screen.PrimaryScreen.WorkingArea.Width - 20);
        };

        void HandleSelection()
        {
            filteredEntries.Clear();
            foreach (DataGridViewRow row in dataGridView.SelectedRows)
            {
                Dictionary<string, object> entry = new Dictionary<string, object>();
                for (int i = 0; i < propertyNames.Count; i++)
                {
                    entry[propertyNames[i]] = row.Cells[i].Value;
                }
                filteredEntries.Add(entry);
            }
            escapePressed = true;
            form.Close();
        }

        dataGridView.CellDoubleClick += (sender, e) => HandleSelection();

        dataGridView.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                escapePressed = true;
                form.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                HandleSelection();
            }
            else if (e.KeyCode == Keys.Tab && !e.Shift)
            {
                searchBox.Focus();
                e.Handled = true;
            }
            else if ((e.KeyCode == Keys.Right || e.KeyCode == Keys.Left) && e.Shift)
            {
                int scrollOffset = 1000;
                dataGridView.HorizontalScrollingOffset += e.KeyCode == Keys.Right ? scrollOffset : -scrollOffset;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Right)
            {
                dataGridView.HorizontalScrollingOffset += 50;
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Left)
            {
                dataGridView.HorizontalScrollingOffset = Math.Max(dataGridView.HorizontalScrollingOffset - 50, 0);
                e.Handled = true;
            }
        };

        searchBox.KeyDown += (sender, e) =>
        {
            if (e.KeyCode == Keys.Escape)
            {
                escapePressed = true;
                form.Close();
            }
            else if (e.KeyCode == Keys.Enter)
            {
                HandleSelection();
            }
            else if (e.KeyCode == Keys.Tab && !e.Shift)
            {
                dataGridView.Focus();
                e.Handled = true;
            }
            else if (e.KeyCode == Keys.Down || e.KeyCode == Keys.Up)
            {
                if (dataGridView.Rows.Count > 0)
                {
                    dataGridView.Focus();

                    int newIndex = 0;
                    int firstVisibleCol = GetFirstVisibleColumnIndex();
                    
                    if (dataGridView.SelectedRows.Count > 0)
                    {
                        int currentIndex = dataGridView.SelectedRows[0].Index;
                        newIndex = e.KeyCode == Keys.Down
                            ? Math.Min(currentIndex + 1, dataGridView.Rows.Count - 1)
                            : Math.Max(currentIndex - 1, 0);
                    }

                    if (firstVisibleCol != -1)
                    {
                        dataGridView.ClearSelection();
                        dataGridView.Rows[newIndex].Selected = true;
                        dataGridView.CurrentCell = dataGridView.Rows[newIndex].Cells[firstVisibleCol];
                    }
                }
                e.Handled = true;
            }
        };

        form.Load += (sender, e) =>
        {
            dataGridView.Sort(dataGridView.Columns[0], ListSortDirection.Ascending);
            dataGridView.AutoResizeColumns();

            int padding = 20;
            int requiredHeight = dataGridView.Rows.GetRowsHeight(DataGridViewElementStates.Visible)
                + dataGridView.ColumnHeadersHeight + 2 * dataGridView.Rows[0].Height
                + SystemInformation.HorizontalScrollBarHeight;
            int availableHeight = Screen.PrimaryScreen.WorkingArea.Height - padding * 2;

            form.Height = Math.Min(requiredHeight, availableHeight);

            if (spanAllScreens)
            {
                form.Width = Screen.AllScreens.Sum(s => s.WorkingArea.Width);
                form.Location = new Point(
                    Screen.AllScreens.Min(s => s.Bounds.X),
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2
                );
            }
            else
            {
                int requiredWidth = dataGridView.Columns.GetColumnsWidth(DataGridViewElementStates.Visible)
                    + SystemInformation.VerticalScrollBarWidth + 43;
                form.Width = Math.Min(requiredWidth, Screen.PrimaryScreen.WorkingArea.Width - padding * 2);
                form.Location = new Point(
                    (Screen.PrimaryScreen.WorkingArea.Width - form.Width) / 2,
                    (Screen.PrimaryScreen.WorkingArea.Height - form.Height) / 2
                );
            }
        };

        form.Controls.Add(dataGridView);
        form.Controls.Add(searchBox);
        searchBox.Select();
        form.ShowDialog();

        return filteredEntries;
    }

    private static string StripQuotes(string input)
    {
        return input.StartsWith("\"") && input.EndsWith("\"") && input.Length > 1
            ? input.Substring(1, input.Length - 2)
            : input;
    }
}
