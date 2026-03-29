using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Nimaime.Helper.Data;

namespace Nimaime.Helper.File
{
	public static class CSVHelper
	{
		/// <summary>
		/// CSV转换为DataTable
		/// </summary>
		/// <param name="fileName">CSV文件</param>
		/// <returns></returns>
		public static DataTable? CSV2DataTable(string fileName)
		{
			return CSV2DataTable(fileName, CommonFileHelper.GetFileEncoding(fileName));
		}

		/// <summary>
		/// CSV转换为DataTable
		/// </summary>
		/// <param name="fileName">CSV文件</param>
		/// <param name="encoding">指定编码</param>
		/// <returns></returns>
		public static DataTable? CSV2DataTable(string fileName, Encoding encoding)
		{
			if (!System.IO.File.Exists(fileName))
			{
				return null;
			}

			DataTable dataTable = new DataTable();
			FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read);
			StreamReader streamReader = new StreamReader(fileStream, encoding);
			int num = 0;
			bool flag = true;
			string pattern = ",(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)";
			string text;
			while ((text = streamReader.ReadLine()) != null)
			{
				string[] array = text.Split(new char[1] { ',' });
				if (flag)
				{
					flag = false;
					num = array.Length;
					for (int i = 0; i < num; i++)
					{
						DataColumn column = new DataColumn(array[i].Replace("\"", ""));
						dataTable.Columns.Add(column);
					}
				}
				else
				{
					string[] array2 = Regex.Split(text, pattern);
					DataRow dataRow = dataTable.NewRow();
					for (int j = 0; j < num; j++)
					{
						dataRow[j] = array2[j].Replace("\"", "");
					}

					dataTable.Rows.Add(dataRow);
				}
			}

			streamReader.Close();
			fileStream.Close();
			return dataTable;
		}

		/// <summary>
		/// 将DataTable写入CSV
		/// </summary>
		/// <param name="dt">DataTable</param>
		/// <param name="AbosultedFilePath">输出文件绝对路径</param>
		public static void DataTable2CSV(DataTable dt, string AbosultedFilePath)
		{
			if (!Directory.Exists(Path.GetDirectoryName(AbosultedFilePath)))
			{
				Directory.CreateDirectory(Path.GetDirectoryName(AbosultedFilePath));
			}
			FileStream fs = new FileStream(AbosultedFilePath, FileMode.Create, FileAccess.Write);
			StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
			//Tabel header
			string line = "";
			for (int i = 0; i < dt.Columns.Count; i++)
			{
				line += dt.Columns[i].ColumnName + ',';
			}
			line = line.TrimEnd(',');
			sw.WriteLine(line);
			//Table body
			for (int i = 0; i < dt.Rows.Count; i++)
			{
				line = "";
				for (int j = 0; j < dt.Columns.Count; j++)
				{
					line += "\"" + dt.Rows[i][j].ToString().DelQuota() + "\",";
				}
				line = line.Substring(0, line.Length - 1);
				sw.WriteLine(line);
			}
			sw.Flush();
			sw.Close();
		}
	}
}
