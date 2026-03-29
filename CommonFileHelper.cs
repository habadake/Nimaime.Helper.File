using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Nimaime.Helper.File
{
	/// <summary>
	/// 通用文件帮助类
	/// </summary>
	public static class CommonFileHelper
	{
		/// <summary>
		/// 检测字节流是否是指定格式
		/// </summary>
		/// <param name="data">字节数据</param>
		/// <param name="fileType">期望的文件格式</param>
		/// <returns></returns>
		public static bool IsFileType(byte[] data, FileType fileType)
		{
			switch (fileType)
			{
				case FileType.OTHER:
					return true;
				case FileType.Excel:
					if (data == null || data.Length < 4)
						return false;

					// XLS
					if (data[0] == 0xD0 &&
						data[1] == 0xCF &&
						data[2] == 0x11 &&
						data[3] == 0xE0)
						return true;

					// XLSX (ZIP)
					if (data[0] == 0x50 &&
						data[1] == 0x4B &&
						data[2] == 0x03 &&
						data[3] == 0x04)
						return true;

					return false;
				default:
					return true;
			}
		}

		/// <summary>
		/// 文件类型枚举
		/// </summary>
		public enum FileType
		{
			/// <summary>
			/// 其他未指定格式
			/// </summary>
			OTHER = 0,
			/// <summary>
			/// Excel 工作簿（包括 XLS 和 XLSX）
			/// </summary>
			Excel = 1,
			/// <summary>
			/// PDF 文档
			/// </summary>
			PDF = 2,
			/// <summary>
			/// TXT 文本文档
			/// </summary>
			TXT = 3,
		}
		
		/// <summary>
		/// 判断文件编码
		/// </summary>
		/// <param name="filePath">文件路径</param>
		/// <returns></returns>
		public static Encoding GetFileEncoding(string filePath)
		{
			using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
			{
				if (fileStream.Length >= 4)
				{
					byte[] buffer = new byte[4];
					fileStream.Read(buffer, 0, 4);

					if (buffer[0] == 0xEF && buffer[1] == 0xBB && buffer[2] == 0xBF)
					{
						return Encoding.UTF8;
					}
					else if (buffer[0] == 0xFF && buffer[1] == 0xFE)
					{
						return Encoding.Unicode; // UTF-16LE
					}
					else if (buffer[0] == 0xFE && buffer[1] == 0xFF)
					{
						return Encoding.BigEndianUnicode; // UTF-16BE
					}
					else if (buffer[0] == 0x1B && buffer[1] == 0x24 && buffer[2] == 0x40)
					{
						return Encoding.GetEncoding("Shift_JIS"); // Shift JIS
					}
					else if (buffer[0] >= 0xB0 && buffer[0] <= 0xF7 && buffer[1] >= 0xA1 && buffer[1] <= 0xFE)
					{
						return Encoding.GetEncoding("GB2312"); // GB2312
					}

				}
			}

			return Encoding.Default; // Default system encoding
		}
	}
}
