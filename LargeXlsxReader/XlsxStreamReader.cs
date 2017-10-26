using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Data;
using System.IO;

namespace LargeXlsxReader
{
    public class XlsxStreamReader
    {
		List<string> sStrings;
		List<WorksheetsRelationships> wRelationships;
		string id;
		string _fileName;
		/// <summary>
		/// Creates a new instance of XlsxStreamReader
		/// </summary>
		/// <param name="fileName">xlsx file name</param>
		/// <param name="sheetName">Name of the sheet</param>
		public XlsxStreamReader(string fileName, string sheetName)
		{
			_fileName = fileName;
			sStrings = GetSharedStrings(fileName);
			wRelationships = GetWorksheetsIds(fileName);
			id = wRelationships.FirstOrDefault(x => x.SheetName == sheetName)?.RelationshipId;
			if (id == null)
			{
				throw new ArgumentException($"The sheet {sheetName} doesn't exists");
			}
		}
		/// <summary>
		/// Creates a new instance of XlsxStreamReader
		/// </summary>
		/// <param name="fileName">xlsx file name</param>
		/// <param name="sheetId">Index of the sheet, base 1</param>
		public XlsxStreamReader(string fileName, int sheetId)
		{
			_fileName = fileName;
			sStrings = GetSharedStrings(fileName);
			wRelationships = GetWorksheetsIds(fileName);
			id = wRelationships.FirstOrDefault(x => x.SheetId == sheetId)?.RelationshipId;
			if (id == null)
			{
				throw new ArgumentException($"The sheet {sheetId} doesn't exists");
			}
		}

		/// <summary>
		/// IEnumerable with Excel's rows as an object array
		/// </summary>
		public IEnumerable<object[]> Rows
		{
			get
			{
				
				DataTable dt = new DataTable();
				int LastRow = -1;
				int LastColumn = -1;
				int FirstRow = -1;
				int FirstColumn = -1;
				
				using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(_fileName, false))
				{
					WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
					WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(id);
					OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
					//Get sheet dimensions
					while (reader.Read())
					{
						if (reader.IsStartElement && reader.ElementType == typeof(SheetDimension))
						{
							int[] dim = TranslateAddress(reader.Attributes[0].Value);
							FirstRow = dim[0];
							FirstColumn = dim[1];
							//Range as the sheet dimension
							if (dim.Length == 4)
							{
								LastRow = dim[2];
								LastColumn = dim[3];
							}
							//Cell as the sheet dimension (only 1 cell has been initialized)
							else
							{
								LastRow = dim[0];
								LastColumn = dim[1];
							}
							break;
						}
					}
					//Define the columns. The column name is in Excel format (A,B,C...)
					for (int c = 1; c <= LastColumn; c++)
					{
						dt.Columns.Add(ColumnName(c));
					}

					int ll = 0;
					object[] currentRow = new object[LastColumn];
					while (reader.Read())
					{
						if (reader.IsStartElement)
						{
							if (reader.ElementType == typeof(Row))
							{
								//Add rows that aren't in the file (unitialized rows aren't saved in xlsx)
								var line = int.Parse(reader.Attributes.First(x => x.LocalName == "r").Value);
								while (++ll < line)
								{
									yield return new object[LastColumn];
								}
								//Initialize a new row
								currentRow = new object[LastColumn];
							}
							else if (reader.ElementType == typeof(Cell))
							{
								//Pick the data type of the Cell. s for sharedString, str for string, 
								//b for bool and absent for numbers
								var tipo = reader.Attributes.FirstOrDefault(x => x.LocalName == "t").Value;
								var endereco = TranslateAddress(
									reader.Attributes.First(x => x.LocalName == "r").Value
									);
								//Search for the CellValue or quit if it find the closing Cell tag
								while (reader.ElementType != typeof(CellValue)
										|| (reader.ElementType != typeof(Cell) && reader.IsEndElement))
								{
									reader.Read();
								}
								if (reader.ElementType == typeof(CellValue))
								{
									switch (tipo)
									{
										//Shared string
										case "s":
											currentRow[endereco[1] - 1] =
												sStrings[int.Parse(reader.GetText())];
											break;
										//Number
										case null:
											if (int.TryParse(reader.GetText(), out int saida))
												currentRow[endereco[1] - 1] = saida;
											else
												currentRow[endereco[1] - 1] =
													Convert.ToDouble(reader.GetText(),
																		CultureInfo.InvariantCulture);
											break;
										//String
										case "str":
											currentRow[endereco[1] - 1] = reader.GetText();
											break;
										//Boolean
										case "b":
											currentRow[endereco[1] - 1] =
												reader.GetText() == "1" ? true : false;
											break;
									}
								}
							}
						}
						else
						{
							//Return created row when found the closing Row tag
							if (reader.ElementType == typeof(Row))
							{
								yield return currentRow;
							}
						}
					}

				}
			}
		}

		/// <summary>
		/// Creates a CSV from the sheetName of the xlsxFile, using the separator
		/// </summary>
		/// <param name="xlsxFile">xlsx file to read from</param>
		/// <param name="sheetName">Name of the sheet</param>
		/// <param name="csvFile">Destination csv file</param>
		/// <param name="separator">csv separator. Default ','</param>
		public static void CreateCsv(string xlsxFile, string sheetName, string csvFile, char separator = ',')
		{
			var relationships = GetWorksheetsIds(xlsxFile);
			var id = relationships.FirstOrDefault(x => x.SheetName == sheetName);
			if (id == null)
			{
				throw new ArgumentException($"The sheet {sheetName} doesn't exists");
			}
			CreateCsv(
				xlsxFile, 
				relationships.First(x => x.SheetName == sheetName).SheetId, 
				csvFile, 
				separator
			);
		}

		/// <summary>
		/// Creates a CSV from the sheet number sheetId of the xlsxFile, using the separator
		/// </summary>
		/// <param name="xlsxFile">xlsx file to read from</param>
		/// <param name="sheetId">Index of the sheet, base 1</param>
		/// <param name="csvFile">Destination csv file</param>
		/// <param name="separator">csv separator. Default ','</param>
		public static void CreateCsv(string xlsxFile, int sheetId, string csvFile, char separator = ',')
		{
			int LastRow = -1;
			int LastColumn = -1;
			int FirstRow = -1;
			int FirstColumn = -1;

			List<string> dicionario = GetSharedStrings(xlsxFile);
			var relationships = GetWorksheetsIds(xlsxFile);

			var id = relationships.FirstOrDefault(x => x.SheetId == sheetId);
			if (id == null)
			{
				throw new ArgumentException($"The sheet {sheetId} doesn't exists");
			}
			using(StreamWriter sw = new StreamWriter(csvFile,false))
			{
				using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(xlsxFile, false))
				{
					WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
					WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(
						relationships.First(x => x.SheetId == sheetId).RelationshipId
						);
					OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
					//Get sheet dimensions
					while (reader.Read())
					{
						if (reader.IsStartElement && reader.ElementType == typeof(SheetDimension))
						{
							int[] dim = TranslateAddress(reader.Attributes[0].Value);
							FirstRow = dim[0];
							FirstColumn = dim[1];
							//Range as the sheet dimension
							if (dim.Length == 4)
							{
								LastRow = dim[2];
								LastColumn = dim[3];
							}
							//Cell as the sheet dimension (only 1 cell has been initialized)
							else
							{
								LastRow = dim[0];
								LastColumn = dim[1];
							}
							break;
						}
					}
					
					int ll = 0;
					string[] currentRow = new string[LastColumn];
					while (reader.Read())
					{
						if (reader.IsStartElement)
						{
							if (reader.ElementType == typeof(Row))
							{
								//Initialize a new row
								currentRow = new string[LastColumn];

								//Add rows that aren't in the file (unitialized rows aren't saved in xlsx)
								var line = int.Parse(reader.Attributes.First(x => x.LocalName == "r").Value);
								while (++ll < line)
								{
									var lin = string.Join(separator.ToString(), currentRow);
									sw.WriteLine(lin);
								}							
							}
							else if (reader.ElementType == typeof(Cell))
							{
								//Pick the data type of the Cell. s for sharedString, str for string, 
								//b for bool and absent for numbers
								var tipo = reader.Attributes.FirstOrDefault(x => x.LocalName == "t").Value;
								var endereco = TranslateAddress(
									reader.Attributes.First(x => x.LocalName == "r").Value
									);
								//Search for the CellValue or quit if it find the closing Cell tag
								while (reader.ElementType != typeof(CellValue)
									   || (reader.ElementType != typeof(Cell) && reader.IsEndElement))
								{
									reader.Read();
								}
								if (reader.ElementType == typeof(CellValue))
								{
									switch (tipo)
									{
										//Shared string
										case "s":
											currentRow[endereco[1] - 1] =
												dicionario[int.Parse(reader.GetText())];
											break;
										//Number
										case null:
										//String
										case "str":
											currentRow[endereco[1] - 1] = reader.GetText();
											break;
										//Boolean
										case "b":
											currentRow[endereco[1] - 1] =
												reader.GetText() == "1" ? "true" : "false";
											break;
									}
								}

							}
						}
						else
						{
							//Add the created row when found the closing Row tag
							if (reader.ElementType == typeof(Row))
							{
								sw.WriteLine(string.Join(separator.ToString(),currentRow));
							}
						}
					}

				}
			}
		}

		/// <summary>
		/// Creates a DataTable from the sheet sheetName of the file xlsxFile
		/// </summary>
		/// <param name="xlsxFile">xlsx file to read</param>
		/// <param name="sheetName">Name of the sheet</param>
		/// <returns>DataTable with column names as Excel column address</returns>
		public static DataTable CreateDataTable(string xlsxFile, string sheetName)
		{
			var relationships = GetWorksheetsIds(xlsxFile);
			var id = relationships.FirstOrDefault(x => x.SheetName == sheetName);
			if (id == null)
			{
				throw new ArgumentException($"The sheet {sheetName} doesn't exists");
			}
			return CreateDataTable(xlsxFile, relationships.First(x => x.SheetName == sheetName).SheetId);
		}

		/// <summary>
		/// Creates a DataTable from the sheet sheetName of the file xlsxFile
		/// </summary>
		/// <param name="xlsxFile">xlsx file to read</param>
		/// <param name="sheetId">Index of the sheet, base 1</param>
		/// <returns>DataTable with column names as Excel string column address</returns>
		public static DataTable CreateDataTable(string xlsxFile, int sheetId)
		{
			DataTable dt = new DataTable();
			int LastRow = -1;
			int LastColumn = -1;
			int FirstRow = -1;
			int FirstColumn = -1;

			List<string> dicionario = GetSharedStrings(xlsxFile);
			var relationships = GetWorksheetsIds(xlsxFile);

			var id = relationships.FirstOrDefault(x => x.SheetId == sheetId);
			if (id == null)
			{
				throw new ArgumentException($"The sheet {sheetId} doesn't exists");
			}

			using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(xlsxFile, false))
			{
				WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
				WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(
					relationships.First(x => x.SheetId == sheetId).RelationshipId
					);
				OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
				//Get sheet dimensions
				while (reader.Read())
				{
					if (reader.IsStartElement && reader.ElementType == typeof(SheetDimension))
					{
						int[] dim = TranslateAddress(reader.Attributes[0].Value);
						FirstRow = dim[0];
						FirstColumn = dim[1];
						//Range as the sheet dimension
						if (dim.Length == 4)
						{
							LastRow = dim[2];
							LastColumn = dim[3];
						}
						//Cell as the sheet dimension (only 1 cell has been initialized)
						else
						{
							LastRow = dim[0];
							LastColumn = dim[1];
						}
						break;
					}
				}
				//Define the columns. The column name is in Excel format (A,B,C...)
				for(int c = 1; c <= LastColumn; c++)
				{
					dt.Columns.Add(ColumnName(c));
				}
				
				int ll = 0;
				object[] currentRow = new object[LastColumn];
				while (reader.Read())
				{
					if (reader.IsStartElement)
					{
						if (reader.ElementType == typeof(Row))
						{
							//Add rows that aren't in the file (unitialized rows aren't saved in xlsx)
							var line = int.Parse(reader.Attributes.First(x => x.LocalName == "r").Value);
							while (++ll < line)
							{
								dt.Rows.Add(dt.NewRow());
							}
							//Initialize a new row
							currentRow = new object[LastColumn];
						}
						else if (reader.ElementType == typeof(Cell))
						{
							//Pick the data type of the Cell. s for sharedString, str for string, 
							//b for bool and absent for numbers
							var tipo = reader.Attributes.FirstOrDefault(x => x.LocalName == "t").Value;
							var endereco = TranslateAddress(
								reader.Attributes.First(x => x.LocalName == "r").Value
								);
							//Search for the CellValue or quit if it find the closing Cell tag
							while (reader.ElementType != typeof(CellValue) 
								   || (reader.ElementType != typeof(Cell) && reader.IsEndElement))
							{
								reader.Read();
							}
							if (reader.ElementType == typeof(CellValue))
							{
								switch (tipo)
								{
									//Shared string
									case "s":
										currentRow[endereco[1] - 1] = 
											dicionario[int.Parse(reader.GetText())];
										break;
									//Number
									case null:
										if (int.TryParse(reader.GetText(), out int saida))
											currentRow[endereco[1] - 1] = saida;
										else
											currentRow[endereco[1] - 1] = 
												Convert.ToDouble(reader.GetText(), 
																 CultureInfo.InvariantCulture);
										break;
									//String
									case "str":
										currentRow[endereco[1] - 1] = reader.GetText();
										break;
									//Boolean
									case "b":
										currentRow[endereco[1] - 1] = 
											reader.GetText() == "1" ? true : false;
										break;
								}
							}
							
						}
					}
					else
					{
						//Add the created row when found the closing Row tag
						if (reader.ElementType == typeof(Row))
						{
							dt.Rows.Add(currentRow);
						}
					}
				}
				
			}
			return dt;
		}

		private static List<WorksheetsRelationships> GetWorksheetsIds(string xlsxFile)
		{
			List<WorksheetsRelationships> relationships = new List<WorksheetsRelationships>();
			using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(xlsxFile, false))
			{
				WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
				var sheets = workbookPart.Workbook.Sheets.Cast<Sheet>().ToList();
				sheets.ForEach(x => 
					relationships.Add(
						new WorksheetsRelationships()
						{
							RelationshipId = x.Id.Value,
							SheetName = x.Name.Value,
							SheetId = Convert.ToInt32(x.SheetId.Value)
						}
					)
				);
			}
			return relationships;
		}

		private static List<string> GetSharedStrings(string xlsxFile)
		{
			List<string> dicionario = new List<string>();
			using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(xlsxFile, false))
			{
				

				WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

				//Get the string List, which is stored in a separate file
				var sharedStringPart = workbookPart.SharedStringTablePart;
				OpenXmlReader reader = OpenXmlReader.Create(sharedStringPart);
				while (reader.Read())
				{
					if (reader.IsStartElement && reader.ElementType == typeof(Text))
					{
						dicionario.Add(reader.GetText());
					}
				}
			}
			return dicionario;
		}
		
		/// <summary>
		/// Translates an address of type excel address to an numeric array, base 1. 
		/// If it's a cell, the returning array has 2 elements, line and column;
		/// If it's a range, the returnig array has 4 elements:
		/// initial line, initial column, final line, final column
		/// </summary>
		/// <param name="address">Address in one of the formats A1, A1:B1, $A$1 etc</param>
		/// <returns>initial line, initial column, [final line, final column]. Base 1</returns>
		public static int[] TranslateAddress(string address)
		{
			address = address.ToUpper();
			int[] retorno = new int[2];
			string[] partes = address.Split(':');

			if (partes.Length == 2)
			{
				retorno = new int[4];
				retorno[2] = int.Parse(Regex.Match(partes[1], "[0-9]+").Value);
				retorno[3] = ColumnIndex(Regex.Match(partes[1], "[A-Z]+").Value);
			}
			retorno[0] = int.Parse(Regex.Match(partes[0], "[0-9]+").Value);
			retorno[1] = ColumnIndex(Regex.Match(partes[0], "[A-Z]+").Value);
			return retorno;
		}

		/// <summary>
		/// Translates an address of type excel address to an numeric array, base 1.
		/// Returns true if it could parse, false otherwise
		/// If it's a cell, the returning array has 2 elements, line and column;
		/// If it's a range, the returnig array has 4 elements:
		/// initial line, initial column, final line, final column
		/// </summary>
		/// <param name="address">Address in one of the formats A1, A1:B1, $A$1 etc</param>
		/// <param name="matrix">initial line, initial column, [final line, final column]. Base 1</param>
		/// <returns></returns>
		public static bool TryTranslateAddress(string address, out int[] matrix)
		{
			address = address.ToUpper();
			if (!Regex.IsMatch(address, @"(\$?[A-Z]+\$?[1-9][0-9]*(:\$?[A-Z]+\$?[1-9][0-9]*)?|\$?[1-9][0-9]*:\$?[1-9][0-9]*|\$?[A-Z]+:\$?[A-Z]+)"))
			{
				matrix = null;
				return false;
			}
			int[] retorno = new int[2];
			string[] partes = address.Split(':');

			if (partes.Length == 2)
			{
				retorno = new int[4];
				retorno[2] = int.Parse(Regex.Match(partes[1], "[0-9]+").Value);
				retorno[3] = ColumnIndex(Regex.Match(partes[1], "[A-Z]+").Value);
			}
			retorno[0] = int.Parse(Regex.Match(partes[0], "[0-9]+").Value);
			retorno[1] = ColumnIndex(Regex.Match(partes[0], "[A-Z]+").Value);
			matrix = retorno;
			return true;
		}

		/// <summary>
		/// Parses an column address like A, AB, DV to the corresponding column index, base 1
		/// </summary>
		/// <param name="columnAddress">Column address</param>
		/// <returns>Corresponding column index, base 1</returns>
		public static int ColumnIndex(string columnAddress)
		{
			int columnIndex = 0;
			int multiplier = 0;
			var chars = columnAddress.ToList();
			chars.Reverse();
			foreach (char c in chars)
			{
				if (c < 65 || c > 90)
				{
					throw new ArgumentException($"The column address {columnAddress} is not valid");
				}
				columnIndex += (c - 64) * (int)Math.Round(Math.Pow(26, multiplier));
				multiplier++;
			}
			return columnIndex;
		}

		/// <summary>
		/// Returns the column string address from the given base 1 column index
		/// </summary>
		/// <param name="index">Index of the column, base 1</param>
		/// <returns>The column string address</returns>
		public static string ColumnName(int index)
		{
			string address = "";
			if (index <= 0)
			{
				throw new ArgumentOutOfRangeException("The index must be greater than 0");
			}
			if (index <= 26)
				address += (char)(index + 64);
			else
			{
				address = ColumnName((index-1)/26);
				address += (char)((index - 1)%26 + 65);
			}
			return address;
		}
	}
}
