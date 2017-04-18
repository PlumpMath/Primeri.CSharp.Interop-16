using System;
using InteropExcel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Excel
{
	public class IOWrite
	{
		private DataStruct _data;

		private InteropExcel.Application excel;

		public IOWrite(DataStruct data)
		{
			_data = data;

		}

		public bool exportTable()
		{
			try
			{
				//Podgotovka
				excel = new InteropExcel.ApplicationClass ();

				if ( excel == null ) return false;

				InteropExcel.Workbook workbook = excel.Workbooks.Add();
				if ( excel == null ) return false;

				excel.Visible = false;


				InteropExcel.Worksheet sheet = (InteropExcel.Worksheet) workbook.Worksheets [1];
				sheet.Name = "Tabliza1";

				//Popalvane na tablizata

				int i = 1;

				addRow(new DataRow("Ime ", "Familia ", "Godini " ), i++, true, 50 );i++;

				foreach (DataRow row in _data.table)
				{

					addRow( row, i++, false, -1 );

				}
				i++; addRow(new DataRow("Broi redove ", "", _data.table.Count.ToString () ), i++, true, -1 );

				//Zapametiavane i zatvariane
				workbook.SaveCopyAs( getPath ());
				        
				excel.DisplayAlerts = false; //Izkliuchvame vsichki saobshtenia na Excel
				workbook.Close();

				excel.Quit ();

				//Osvobojdavane na pametta ot excel

				if (workbook != null) Marshal.ReleaseComObject(workbook);
				if (sheet != null) Marshal.ReleaseComObject(sheet);
				if (excel != null) Marshal.ReleaseComObject(excel);

				workbook = null;
				sheet = null;
				excel = null;

				GC.Collect();

				return true;
			}
			catch
			{
			}
			return false;

		}

		public void addRow ( DataRow _dataRow, int _indexRow, bool isBold, int color )
		{
			try
			{

				InteropExcel.Range range;

				//Formatirane
				range = excel.Range["A" + _indexRow.ToString(), "C" + _indexRow.ToString()];
				if (color > 0)    range.Interior.ColorIndex = color; //-1
				if (isBold)       range.Font.Bold = isBold;

				//Vavejdane na danni kletka po kletka
				range = excel.Range["A" + _indexRow.ToString(),"A" + _indexRow.ToString() ];
				range.Value2 = _dataRow.firstName;

				range = excel.Range["B" + _indexRow.ToString(), "B" + _indexRow.ToString()];
				range.Value2 = _dataRow.lastName;

				range = excel.Range["C" + _indexRow.ToString(), "C" + _indexRow.ToString()];
				range.Value2 = _dataRow.age;

			}
			catch
			{

			}
		}
		public void runFile()
		{
			try{
				System.Diagnostics.Process.Start (getPath ());	



			}catch {

			}
		}

		private string getPath()
		{

			return System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Table.xlsx");
		}
	}
}
