using System;
using System.Windows.Forms;
using ClosedXML.Excel;

namespace ExcelFileModification
{
  class FileModifier
  {
	[STAThread]
	static void Main()
	{
	  Console.WriteLine("Select Excel file from your computer");
	  string filePath = SelectFile();

	  if (filePath != null)
	  {
		Transpose(filePath);
	  }
	  else 
	  {
		Console.WriteLine("file has not been selected, Program exits");
	  }
	}



	//----------------------------------------------------Select Any Excel File-----------------------------------------------------//
	private static string SelectFile()
	{
	  OpenFileDialog fileDialog = new OpenFileDialog();
	  fileDialog.FileName = "";
	  fileDialog.DefaultExt = ".xlsx";
	  fileDialog.Filter = "Excel Files (.xlsx)|*.xlsx";
	  fileDialog.InitialDirectory = @"C:\";
	  fileDialog.Title = "Select Excel File";
	  fileDialog.RestoreDirectory = true;
	  fileDialog.Multiselect = false;
	  string filePath = null;
	  if (fileDialog.ShowDialog() == DialogResult.OK)
	  {
		filePath = fileDialog.FileName;
		Console.WriteLine("file: {0} has been selected", filePath);
	  }
	  return filePath;
	}


	//----------------------------------------------------Transpose Given File-----------------------------------------------------//
	private static void Transpose(string filePath)
	{
	  try
	  {
		var workBook = new XLWorkbook(filePath);
		var workSheet = workBook.Worksheet(1);
		//var range = workSheet.Range("A1:D7");               // range of given test excel file
		var range = workSheet.RangeUsed();                    // used range of any selected excel file
		range.Transpose(XLTransposeOptions.MoveCells);
		workSheet.Columns().AdjustToContents();
		string file = filePath.Substring(0, (filePath.Length - 5));
		file = file + "_modified.xlsx";
		workBook.SaveAs(file);
		Console.WriteLine("file: {0} has been successfully modified and saved as {1}", filePath, file);
	  }
	  catch (Exception excp)
	  {
		Console.WriteLine("Error! {0}", excp.Message);
	  }
	}
  }
}
