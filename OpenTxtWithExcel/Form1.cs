using OfficeOpenXml;
using OfficeOpenXml.Style;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace OpenTxtWithExcel
{
	public partial class Form1 : Form
	{
		public FileDropHandler FileDroper;
		public Form1()
		{
			InitializeComponent();
			OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
			FileDroper = new FileDropHandler(this);
			comboBox1.SelectedIndex = 0;
		}

		private void Form1_DragEnter(object sender, DragEventArgs e)
		{
			e.Effect = DragDropEffects.Copy;
			var s = e.Data.GetData(typeof(string[])) as string[];
			label1.Text = "";
			List<string> filelist = new List<string>();
			List<string> resultlist = new List<string>();
			foreach (var item in s)
			{
				if (Directory.Exists(item))
				{
					string[] txtFile = Directory.GetFiles(item, "*.txt", SearchOption.AllDirectories);
					foreach (var file in txtFile)
					{
						filelist.Add(file);
					}
					string[] csvFile = Directory.GetFiles(item, "*.csv", SearchOption.AllDirectories);
					foreach (var file in csvFile)
					{
						filelist.Add(file);
					}
				}
				if (File.Exists(item) && (Path.GetExtension(item).Equals(".txt", StringComparison.OrdinalIgnoreCase) || Path.GetExtension(item).Equals(".csv", StringComparison.OrdinalIgnoreCase)))
				{
					filelist.Add(item);
				}
			}
			if (filelist.Count == 0)
			{
				label1.Text = "无有效文件";
				return;
			}
			label1.Text = DateTime.Now.ToLongTimeString() + " 处理中" + "\n";
			DateTime startTime = DateTime.Now;
			foreach (var item in filelist)
			{
				label1.Text += DateTime.Now.ToLongTimeString() +" "+ Path.GetFileName(item) + "\n";
				resultlist.Add(ProcessFile(item));
				this.Refresh();
			}

			if (filelist.Count == 1)
			{
				if (checkBox2.Checked)
				{
					Process.Start(resultlist[0]);
				}
			}
			else
			{
				if (checkBox3.Checked)
				{
					foreach (var item in resultlist)
					{
						Process.Start(item);
					}
				}
			}
			label1.Text += DateTime.Now.ToLongTimeString() + " 处理完毕" + "\n";
			label1.Text += filelist.Count + "个文件共用时" + (DateTime.Now - startTime).ToString("T") + "\n";

			return;
		}

		string ProcessFile(string file)
		{
			using (OfficeOpenXml.ExcelPackage excel = new OfficeOpenXml.ExcelPackage())
			{
				char[] splitchar = new char[2] { checkBox5.Checked ? ',' : '\n', checkBox6.Checked ? '\t' : '\n' };
				OfficeOpenXml.ExcelWorksheet ws = excel.Workbook.Worksheets.Add(Path.GetFileNameWithoutExtension(file));
				string[] s = File.ReadAllLines(file,Encoding.GetEncoding(comboBox1.SelectedItem.ToString()));
				if (s.Count() == 0)
				{
					label1.Text = "读取失败,文件可能为空";
				}
				DataTable dataTable = new DataTable();
				dataTable.BeginLoadData();
				foreach (var item in s[0].Split(splitchar))
				{
					try
					{
					dataTable.Columns.Add(item, typeof(string));

					}
					catch (Exception)
					{
						label1.Text += "读取失败,表头重复"+"\n";
						throw;
					}
				}
				foreach (var item in s.Skip(1))
				{
					dataTable.Rows.Add(item.Split(splitchar));
				}
				dataTable.EndLoadData();
				//识别数字的列 弃用 改为全部以字符串方式读取
				//DataTable dataTable2 = new DataTable();
				//foreach (DataColumn item in dataTable.Columns)
				//{
				//	if (dataTable.AsEnumerable().Select(x => x.Field<string>(item)).Where(x => int.TryParse(x, out _)).Count() == dataTable.Rows.Count)
				//	{ dataTable2.Columns.Add(item.ColumnName, typeof(int)); }
				//	else
				//	{
				//		dataTable2.Columns.Add(item.ColumnName, typeof(string));
				//	}
				//}
				//dataTable2.Load(dataTable.CreateDataReader());
				ws.Cells.LoadFromDataTable(dataTable, true);
				ws.Cells.Style.Numberformat.Format = "@";
				excel.SaveAs(new FileInfo(file.Remove(file.Length - 4) + ".xlsx"));
				return file.Remove(file.Length - 4) + ".xlsx";
			}
		}

		private void Form1_FormClosed(object sender, FormClosedEventArgs e)
		{
			Properties.Settings.Default.Save();
		}
	}
}
