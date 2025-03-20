using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TimeSheetFilter
{
	public partial class Form1 : Form
	{
		private FileHandling fileHandler = new FileHandling();
		private DataTable dt = new DataTable();
		private DataTable filteredData = new DataTable(); 
		public Form1()
		{
			InitializeComponent();
		}

		private void btnLoadFile_click(object sender, EventArgs e)
		{
			if (fileHandler.SelectFile())
			{
				dt = fileHandler.LoadFileGridView();
				dataGridView1.DataSource = dt; 
				txtFilePath.Text = fileHandler.FilePath;
			}

		}

		private void Form1_Load(object sender, EventArgs e)
		{
			

		}

		private void btn_FilterFile(object sender, EventArgs e)
		{
			try
			{
				if (dt.Rows.Count > 0)
				{
					filteredData = fileHandler.FilterFile(dt);
					dataGridView1.DataSource = filteredData;
				}
				else
				{
					MessageBox.Show("No data available to filter.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void btn_SaveFile(object sender, EventArgs e)
		{
			try
			{
				if (filteredData != null && filteredData.Rows.Count > 0)
				{
				
					bool result = fileHandler.SaveFile(filteredData);

				}
				else
				{
					MessageBox.Show("No data available to save.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning);
				}
			}
			catch (Exception ex)
			{
				MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void textBox1_TextChanged(object sender, EventArgs e)
		{

		}

		private void label1_Click(object sender, EventArgs e)
		{

		}
	}
}
