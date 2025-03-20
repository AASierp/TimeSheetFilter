using System;
using System.IO;
using ClosedXML.Excel;
using System.Windows.Forms;
using System.Data;


[System.Runtime.Versioning.SupportedOSPlatform("windows")]
public class FileHandling
{
	public string? FilePath { get; private set; }
	public bool SelectFile()
	{
		using (OpenFileDialog openFileDialog = new OpenFileDialog())
		{
			openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm";
			openFileDialog.Title = "Please Select A File";
			if (openFileDialog.ShowDialog() == DialogResult.OK)
			{
				FilePath = openFileDialog.FileName;
				return true;
			}
			else
			{
				return false;
			}
		}

	}

	public DataTable LoadFileGridView()
	{
		if (string.IsNullOrWhiteSpace(FilePath))
			throw new InvalidOperationException("No File Selected");

		DataTable dt = new DataTable();

		try
		{
			using (XLWorkbook workbook = new XLWorkbook(FilePath))
			{
				IXLWorksheet worksheet = workbook.Worksheets.Worksheet(1);

				var firstUsedRow = worksheet.FirstRowUsed();
				if (firstUsedRow == null)
				{
					MessageBox.Show("The Excel file has no usable data.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return dt;
				}


				int columnIndex = 1;

				foreach (IXLCell column in firstUsedRow.CellsUsed())
				{
					string? columnName = column.Value.ToString().Trim();

					if (string.IsNullOrEmpty(columnName) || dt.Columns.Contains(columnName))
					{
						columnName = $"Column{columnIndex++}";
					}

					dt.Columns.Add(columnName);
				}

				int expectedColumnCount = dt.Columns.Count;

				foreach (IXLRow row in worksheet.RowsUsed().Skip(1))
				{
					var dataRow = dt.NewRow();
					var cells = row.Cells(1, expectedColumnCount).ToList();

					for (int i = 0; i < cells.Count; i++)
					{
						dataRow[i] = i < cells.Count ? cells[i].Value.ToString(): "";
					}

					dt.Rows.Add(dataRow);
				}
			}

		}
		catch (Exception ex)
		{
			MessageBox.Show("Error Loading Excel File: " + ex.Message);
		}

		return dt;
	}

	public bool SaveFile(DataTable dt)
	{
		if (dt == null || dt.Rows.Count == 0)
		{
			MessageBox.Show("Data Table is Empty", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
			return false;
		}

		using (SaveFileDialog saveFileDialog = new SaveFileDialog())
		{
			saveFileDialog.Filter = "Excel Files|*.xlsx";
			saveFileDialog.Title = "Save The Filtered Data";
			saveFileDialog.FileName = "Filtered_TimeSheet.xlsx";
			if (saveFileDialog.ShowDialog() == DialogResult.OK)
			{
				try
				{
					using (XLWorkbook workbook = new XLWorkbook())
					{
						IXLWorksheet worksheet = workbook.Worksheets.Add("Filtered Data");

						for (int col = 0; col < dt.Columns.Count; col++)
						{
							worksheet.Cell(1, col + 1).Value = dt.Columns[col].ColumnName;
							worksheet.Cell(1, col + 1).Style.Font.Bold = true;
						}

						for (int row = 0; row < dt.Rows.Count; row++)
						{
							for (int col = 0; col < dt.Columns.Count; col++)
							{
								worksheet.Cell(row + 2, col + 1).Value = dt.Rows[row][col]?.ToString();
							}
						}

						worksheet.Columns().AdjustToContents();
						workbook.SaveAs(saveFileDialog.FileName);
					}

					MessageBox.Show("File save successful", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
					return true;
				}
				catch (Exception ex)
				{
					MessageBox.Show("Error Saving File" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
					return false;
				}
			}
		}
		return false;
	}

	public DataTable FilterFile(DataTable dt)
	{

		TimeSpan lunchDuration = new TimeSpan(0, 30, 59);

		TimeSpan breakDuration = new TimeSpan(0, 20, 59);

		TimeSpan incidentalCumulative = new TimeSpan(0, 10, 0);

		Dictionary<string, TimeSpan> incidentalTimes = new Dictionary<string, TimeSpan>();

		HashSet<string> firstReadyRecorded = new HashSet<string>();

		DataTable filteredTable = dt.Clone();

		foreach (DataRow row in dt.Rows)
		{
			string? employee = row["Agent Name"].ToString();

			string? type = row["Reason"].ToString();

			TimeSpan duration;

			if (!TimeSpan.TryParse(row["Duration"].ToString(), out duration))
				continue;

			if (type == "Lunch" && duration >= lunchDuration)
			{
				filteredTable.ImportRow(row);
			}

			else if (type == "Break" && duration >= breakDuration)
			{
				filteredTable.ImportRow(row);
			}

			else if (type == "Incidental")
			{
				if (!incidentalTimes.ContainsKey(employee))
				{
					incidentalTimes[employee] = TimeSpan.Zero;
				}
				incidentalTimes[employee] += duration;

			}

			else if (type == "Ready")
			{
				if (!firstReadyRecorded.Contains(employee))
				{
					filteredTable.ImportRow (row);
					firstReadyRecorded.Add(employee);
				}
			}
		}

		foreach (DataRow row in dt.Rows)
		{
			string? employee = row["Agent Name"].ToString();
			string? type = row["Reason"].ToString();
			TimeSpan duration;

			if (type == "Incidental" &&
				TimeSpan.TryParse(row["Duration"].ToString(), out duration) &&
				incidentalTimes.ContainsKey(employee) &&
				incidentalTimes[employee] > incidentalCumulative)
			{
				filteredTable.ImportRow(row);
			}
		}

		return filteredTable;
	}

}
