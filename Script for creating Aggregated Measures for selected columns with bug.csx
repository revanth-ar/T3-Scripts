using System;
using System.Windows.Forms;
using System.Linq;
using System.Drawing;
using System.Collections.Generic;

// Custom form for selecting columns and aggregation types with search functionality
public class SelectionForm : Form
{
    public Label ColumnsLabel;
    public Label AggregationsLabel;
    public CheckedListBox AggregationsCheckedListBox;
    public CheckBox SelectAllCheckBox;
    public TextBox SearchTextBox;
    public TreeView ColumnsTreeView;
    public Button OkButton;
    public Button CancelButton;
    public CheckBox ShowHiddenColumnsCheckBox;
    public Dictionary<string, string[]> columnsByTableMain;

    public SelectionForm(string[] tableNames, Dictionary<string, string[]> columnsByTable)
    {
        columnsByTableMain = columnsByTable;

        Text = "Select Columns and Aggregation Types";
        Width = 800;
        Height = 800;
        Padding = new Padding(20);

        AggregationsCheckedListBox = new CheckedListBox
        {
            Dock = DockStyle.Top,
            Height = 150,
            Items = { "Sum", "Min", "Max", "Avg" }
        };
        Controls.Add(AggregationsCheckedListBox);

        SelectAllCheckBox = new CheckBox
        {
            Text = "Select All",
            Dock = DockStyle.Top
        };
        SelectAllCheckBox.CheckedChanged += (sender, args) => ToggleSelectAll();
        Controls.Add(SelectAllCheckBox);

        AggregationsLabel = new Label
        {
            Text = "Select aggregations",
            Dock = DockStyle.Top,
            Height = 40,
            TextAlign = ContentAlignment.MiddleLeft
        };
        AggregationsLabel.Font = new Font(AggregationsLabel.Font, FontStyle.Bold); 
        Controls.Add(AggregationsLabel);

        ColumnsTreeView = new TreeView
        {
            Dock = DockStyle.Top,
            Height = 200,
            CheckBoxes = true
        };
        foreach (var tableName in tableNames)
        {
            var tableNode = new TreeNode(tableName);
            foreach (var columnName in columnsByTable[tableName])
            {
                var columnNode = new TreeNode(columnName);
                tableNode.Nodes.Add(columnNode);
            }
            ColumnsTreeView.Nodes.Add(tableNode);
        }
        Controls.Add(ColumnsTreeView);

        SearchTextBox = new TextBox
        {
            Dock = DockStyle.Top,
            PlaceholderText = "Search columns..."
        };
        SearchTextBox.TextChanged += (sender, args) => FilterColumns();
        Controls.Add(SearchTextBox);

        ShowHiddenColumnsCheckBox = new CheckBox
        {
            Text = "Show Hidden Columns",
            Dock = DockStyle.Top
        };
        ShowHiddenColumnsCheckBox.CheckedChanged += (sender, args) => UpdateNumericColumns();
        Controls.Add(ShowHiddenColumnsCheckBox);

        ColumnsLabel = new Label
        {
            Text = "Select columns",
            Dock = DockStyle.Top,
            Height = 40,
            TextAlign = ContentAlignment.MiddleLeft
        };
        ColumnsLabel.Font = new Font(ColumnsLabel.Font, FontStyle.Bold); 
        Controls.Add(ColumnsLabel);

        OkButton = new Button
        {
            Text = "OK",
            Height = 40,
            DialogResult = DialogResult.OK,
            Dock = DockStyle.Bottom
        };
        Controls.Add(OkButton);

        CancelButton = new Button
        {
            Text = "Cancel",
            Height = 40,
            DialogResult = DialogResult.Cancel,
            Dock = DockStyle.Bottom
        };
        Controls.Add(CancelButton);
    }

    private void FilterColumns()
    {
        var expandedNodes = new HashSet<string>();
        foreach (TreeNode node in ColumnsTreeView.Nodes)
        {
            if (node.IsExpanded)
            {
                expandedNodes.Add(node.Text);
            }
        }

        ColumnsTreeView.Nodes.Clear();
        foreach (var tableName in columnsByTableMain.Keys)
        {
            var tableNode = new TreeNode(tableName);
            foreach (var columnName in columnsByTableMain[tableName])
            {
                if (columnName.ToLower().Contains(SearchTextBox.Text.ToLower()))
                {
                    var columnNode = new TreeNode(columnName);
                    tableNode.Nodes.Add(columnNode);
                }
            }
            if (tableNode.Nodes.Count > 0)
            {
                ColumnsTreeView.Nodes.Add(tableNode);
                if (SearchTextBox.Text.Length > 0 || expandedNodes.Contains(tableNode.Text))
                {
                    tableNode.Expand();
                }
            }
        }
    }

    private void ToggleSelectAll()
    {
        for (int i = 0; i < AggregationsCheckedListBox.Items.Count; i++)
        {
            AggregationsCheckedListBox.SetItemChecked(i, SelectAllCheckBox.Checked);
        }
    }

    public List<string> GetSelectedColumns()
    {
        var selectedColumns = new List<string>();
        foreach (TreeNode tableNode in ColumnsTreeView.Nodes)
        {
            foreach (TreeNode columnNode in tableNode.Nodes)
            {
                if (columnNode.Checked)
                {
                    selectedColumns.Add($"{tableNode.Text}.{columnNode.Text}");
                }
            }
        }
        return selectedColumns;
    }

    // Method to update numeric columns based on the Show Hidden Columns checkbox
    private void UpdateNumericColumns()
    {
        var showHidden = ShowHiddenColumnsCheckBox.Checked;

        //T3 picking error here when saved as Macro
        var numericColumnsnew = Model.AllColumns
            .Where(col => (col.DataType == DataType.Int64 || col.DataType == DataType.Double || col.DataType == DataType.Decimal) && (!col.IsHidden || showHidden))
            .ToList();
        //

        columnsByTableMain = numericColumnsnew
            .GroupBy(col => col.Table.Name)
            .ToDictionary(g => g.Key, g => g.Select(col => col.Name).ToArray());
        
        // Update the columns tree view
        FilterColumns();
    }
}

// Check if the model is loaded
if (Model != null)
{
    // Get the table on which the script is run
    var selectedTable = Selected.Table;

    if (selectedTable != null)
    {
        // Get all numeric columns in the selected table
        var numericColumns = Model.AllColumns
            .Where(col => (col.DataType == DataType.Int64 || col.DataType == DataType.Double || col.DataType == DataType.Decimal) && !col.IsHidden)
            .ToList();

        // Group columns by table names
        var columnsByTable = numericColumns
            .GroupBy(col => col.Table.Name)
            .ToDictionary(g => g.Key, g => g.Select(col => col.Name).ToArray());

        // Show the custom form to select columns and aggregation types
        var form = new SelectionForm(columnsByTable.Keys.ToArray(), columnsByTable);
        if (form.ShowDialog() == DialogResult.OK)
        {
            var selectedColumns = form.GetSelectedColumns();
            var selectedAggregations = form.AggregationsCheckedListBox.CheckedItems.Cast<string>().ToList();

            // Check if any columns and aggregation types were selected
            if (selectedColumns.Count > 0 && selectedAggregations.Count > 0)
            {
                foreach (var columnFullName in selectedColumns)
                {
                    var parts = columnFullName.Split('.');
                    var tableName = parts[0];
                    var columnName = parts[1];
                    var column = numericColumns.First(col => col.Table.Name == tableName && col.Name == columnName);
                    var columnFolder = $"Key Measures\\{tableName} {column.Name}";

                    foreach (var selectedAggregation in selectedAggregations)
                    {
                        var measureName = $"{selectedAggregation} of {column.Name}";

                        // Check if the measure already exists
                        if (selectedTable.Measures.Any(m => m.Name == measureName))
                        {
                            //Info($"Measure '{measureName}' already exists. Skipping creation.");
                            //continue;
                            measureName = $"{selectedAggregation} of {column.Name} [{tableName}]";
                        }

                        string daxExpression = string.Empty;
                        switch (selectedAggregation)
                        {
                            case "Sum":
                                daxExpression = $"SUM({column.DaxObjectFullName})";
                                break;
                            case "Min":
                                daxExpression = $"MIN({column.DaxObjectFullName})";
                                break;
                            case "Max":
                                daxExpression = $"MAX({column.DaxObjectFullName})";
                                break;
                            case "Avg":
                                daxExpression = $"AVERAGE({column.DaxObjectFullName})";
                                break;
                            default:
                                Warning($"Invalid aggregation type: {selectedAggregation}");
                                continue;
                        }

                        var newMeasure = selectedTable.AddMeasure(measureName, daxExpression);
                        newMeasure.FormatString = "0.00";
                        newMeasure.Description = $"This measure is the {selectedAggregation.ToLower()} of column {column.DaxObjectFullName}.";
                        newMeasure.DisplayFolder = columnFolder;
                    }
                }

                Info("All set! Your measures are ready to rock and roll!");
            }
            else
            {
                Warning("No columns or aggregation types were selected.");
            }
        }
        else
        {
            Warning("Selection was cancelled.");
        }
    }
    else
    {
        Warning("No table is currently selected.");
    }
}
else
{
    Warning("No model is currently loaded!");
}
