# LargeXlsxReader

## What is LargeXlsxReader

LargeXlsxReader is a library to help read huge xlsx files without a great memory footprint. 
All xlsx file accesses are made through Stream. LargeXlsxReader uses Open XML SDK to work.

## Features and usage

* Convert xlsx to csv:

	`XlsxStreamReader.CreateCsv(xlsxFileName, sheetName, csvFileName, charSeparator);`
	`XlsxStreamReader.CreateCsv(xlsxFileName, sheetIndex, csvFileName, charSeparator);`

* Create a DataTable from a xlsx:

	`DataTable dt = XlsxStreamReader.CreateDataTable(fileName, sheetName);`
	`DataTable dt = XlsxStreamReader.CreateDataTable(fileName, sheetIndex);`

* Iterate through xlsx rows using an IEnumerable<object[]>:

```
	//using sheet index...
	XlsxStreamReader xlsx = new XlsxStreamReader(fileName, sheetIndex);
	foreach (var r in xlsx.Rows)
	{
		var data = r[0];
	}
	
	//...or sheet name
	XlsxStreamReader xlsx = new XlsxStreamReader(fileName, sheetName);
	foreach (var r in xlsx.Rows)
	{
		var data = r[0];
	}
```

* Convert an Excel Addresses ("D5", "C4:Z5", "B:B", "20:20", ...) to a matrix addresses 
([5,4], [4,3,5,26], [1,2,1048576,2], [20,1,20,16384], ...):

```
	//For well formed addresses...
	int[] addressMatrix = XlsxStreamReader.TranslateAddress("D5");
	
	//...or when you don't know if it's valid
	int[] addressMatriz;
	bool valid = XlsxStreamReader.TryTranslateAddress("D5",out addressMatrix);
```

* Convert column names (A,C,AA) to column indexes (1,3,27):
	
	`int cIndex = XlsxStreamReader.ColumnIndex("AA");`

* Convert from column indexes (1,3,27) to column names (A,C,AA):

	`string cName = XlsxStreamReader.ColumnName(27);`
