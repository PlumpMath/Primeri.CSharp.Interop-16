using System;

namespace Excel
{
	class MainClass
	{
		public static void Main(string[] args)
		{
			DataStruct data = new DataStruct();

			IOWrite write = new IOWrite(data);




			//Nabirane na danni v osnovnata tabliza
			data.addRow ("Martin", "Simeonov", "33");
			data.addRow("George", "Marinov", "37");

			//Proverka na tablizata
			data.printTable();



		}
	}
}
