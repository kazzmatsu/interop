using System;
using System.IO;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace WildSeven
{
	public enum ExDev { ExcelApp, ExcelBooks, ExcelBook, ExcelSheet, ExcelRange, ExcelCells, ExcelCell }
	public enum xl { No = 0x0, Yes = 0x1 }
	public class ExcelBase
	{
		public static void Main ( string [ ] args )
		{
			var mmm = new ExcelBase( );
			mmm.ClsMain( );
		}
		internal void ClsMain ( )
		{
			var ExApp = @"Excel.Application";
			var ExFile = @"tBase.xlsx";
			ExcelCreate( ExApp, ExFile );
			ExcelOpen( ExApp, ExFile );
		}
		internal void ExcelOpen ( string exApp, string ExFil )
		{
			var ExBase = new SortedList<string, dynamic>();
			var BookPath = Path.Combine( Environment.CurrentDirectory, ExFil );	//Excel Path
			try {
				ExBase["004ExcelApp"] = Activator.CreateInstance ( Type.GetTypeFromProgID ( exApp ));	// Excel Application
				try {
					ExBase["004ExcelApp"].Visible = true;	// Excel Display
					ExBase["004ExcelApp"].DisplayAlerts = false;	// Alert
					// var xlCellTypeLastCell = 11;
					ExBase["003ExcelBooks"] = ExBase["004ExcelApp"].Workbooks;
					ExBase["002ExcelBook"]  = ExBase["003ExcelBooks"].Open( BookPath );
					ExBase["001ExcelSheet"] = ExBase["002ExcelBook"].Worksheets;	// create 1 sheet
					Console.WriteLine("{0}",BookPath);
					Console.WriteLine("Book.Name := {0}", ExBase["002ExcelBook"].Name );
					Console.WriteLine("Book.Sheet.Count := {0}", ExBase["001ExcelSheet"].Count );
				}
				catch ( Exception ex )
				{
					ExBase["004ExcelApp"].Quit();
					Marshal.ReleaseComObject( ExBase["004ExcelApp"] );
					Console.WriteLine ( ex.Message );
				}
				finally
				{
					foreach ( string key in ExBase.Keys )
					{
						Console.WriteLine( "Debug Release key := {0} ", key );
						switch( key )
						{
							case "004ExcelApp":
									ExBase[key].Quit();
									Marshal.FinalReleaseComObject( (object)ExBase[key]);
									break;
							case "002ExcelBook":
									ExBase[key].Close();
									break;
							default:
									Marshal.FinalReleaseComObject( (object)ExBase[key]);
									break;
						}
					}	//	Marshal.FinalReleaseComObject( (object)ExBase["004ExcelApp"]);
				}
			}
			catch ( Exception e )
			{
				ExBase["004ExcelApp"].Quit();
				Console.WriteLine("Excel not Install");
				Console.WriteLine(e.Message);
				return;
			}
		}
		internal void ExcelCreate ( string exApp , string ExFil )
		{
			var ExBase = new SortedList<string, dynamic>();
			try {
				ExBase["009ExcelApp"] = Activator.CreateInstance ( Type.GetTypeFromProgID ( exApp ));	// Excel Application
				try {
					var BookPath = Path.Combine( Environment.CurrentDirectory, ExFil );		// Excel Path
					var xlSrcRange = 0x01;
					ExBase["009ExcelApp"].Visible = false;
					ExBase["009ExcelApp"].DisplayAlerts = false;
					ExBase["009ExcelApp"].Workbooks.Add();
					ExBase["008ExcelBook"] = ExBase["009ExcelApp"].Workbooks.Add();
					ExBase["008ExcelSheet"] = ExBase["008ExcelBook"].Worksheets( 1 );
					ExBase["008ExcelShee2"] = ExBase["008ExcelBook"].Worksheets( 1 );
					var RowCount = 88000;
					var ColumCount = 250;
					ExBase["008ExcelSheet"].Name = "C#の処理";
					
					var xValues = new object[ RowCount, ColumCount + 1 ];
					for ( var i = 0; i < RowCount; i++ )
					{
						for ( var j = 0; j < ColumCount; j++ )
						{
							xValues[i,j] = (i == 0 ) ? "TiTle Row" + j : "処理 : " + i + j ;
						}
					}
					xValues[0,ColumCount] = "子";
					var x = ExBase["008ExcelSheet"].Range ( ExBase["008ExcelSheet"].Cells( 1, 1), ExBase["008ExcelSheet"].Cells( RowCount, ColumCount + 1 ) ) ;
					x.Value = xValues;
					//Sheet.getCells( Sheet.Cells( 1, 1 ), Sheet.Cells( 1, ColumCount + 1 ) ).setRange(xValues);
					//ExBase["007SourceRange"] = ExBase["008ExcelSheet"].Cells;
					//ExBase["007SourceRange"].Cells( 1, ColumCount + 1 ).Value = "子";	//１つのセルから AutoFit で値をセット
					Console.WriteLine("----- Debug -3----");
					ExBase["008ExcelShee3"] = ExBase["008ExcelSheet"].Range( ExBase["008ExcelSheet"].Cells( 1, 1), ExBase["008ExcelSheet"].Cells( 1, ColumCount + 1) );	//基となるセル範囲
					Console.WriteLine("----- Debug -4----");
					ExBase["006FillRange"]  = ExBase["008ExcelShee3"].Range( ExBase["008ExcelShee3"].Cells( 1, 1), ExBase["008ExcelShee3"].Cells( RowCount, ColumCount + 1) );	//基となるセル範囲
					Console.WriteLine("----- Debug -5----");
					//Sheet.Value(xValue);
					//ExBase["007SourceRange"].AutoFill ( ExBase["006FillTrange"] );
					Console.WriteLine("----- Debug -6----");
					var listObj = ExBase["008ExcelSheet"].ListObjects;
					Console.WriteLine("----- Debug -7----");
					listObj.Add ( xlSrcRange, ExBase["006FillRange"], xl.No, xl.Yes ).Name = "TableList";
					Console.WriteLine("----- Debug -8----");
					ExBase["008ExcelBook"].SaveAs( BookPath );	// save
					Console.WriteLine("----- Debug -9----");
					Console.WriteLine( "{0}", BookPath );
				}
				catch ( Exception ex)
				{
					ExBase["009ExcelApp"].Quit();
					Marshal.ReleaseComObject ( ExBase["009ExcelApp"] );
					Console.WriteLine ( ex.Message );
				}
				finally
				{
					foreach ( string key in ExBase.Keys )
					{
						Console.WriteLine( "Debug Release key := {0} ", key );
						switch( key )
						{
							case "009ExcelApp":
									ExBase[key].Quit();
									Marshal.FinalReleaseComObject( (object)ExBase[key]);
									break;
							case "008ExcelBook":
									ExBase[key].Close();
									break;
							default:
									Marshal.FinalReleaseComObject( (object)ExBase[key]);
									break;
						}
					}
				}
			}
			catch ( Exception e )
			{
				ExBase["009ExcelApp"].Quit();
				Console.WriteLine("Excel not Install");
				Console.WriteLine(e.Message);
			}
		}
	}
}
