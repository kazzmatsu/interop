using namespace System 
using namespace System.IO 
using namespace System.Text 
using namespace System.Console 
using namespace System.Collections 
using namespace System.Collections.Generic 

class SheetTable 
{
  [object]$SheetTableObj 
  [string]$SheetTableName 
  [object]$HeaderObj 
  [string[]]$Header = @() 
  [object]$DtBodyObj 
  [string[]]$DtBodySub = @() 
  [string[][]]$DtBody = @() 
  [int]$ColNum 
  [int]$RowNum SheetTable( [Object]$ShtTblObj ) 
  {
    $This.SheetTableObj = $ShtTblObj 
    $This.SheetTableName = $ShtTblObj.Name 
    $This.ColNum = $ShtTblObj.Listcolumns.Count 
    $This.RowNum = $ShtTblObj.Listrows.Count 
    $This.HeaderObj = $ShtTblObj.HeaderRowRange 
    $This.DtBodyObj = ShtTblObj.DataBodyRange
    1..($This.ColNum) | %{
      This.Header+=(ShtTblObj.HeaderRowRange[].Value2).ToString() 
    } 
    $This.Header -Join $This.Header 1..$($This.RowNum) | %{
      $Rw = $PSItem 1..$($This.ColNum) | %{
        #Set-PSDebug -Trace 2 
        $This.DtBodySub += $($ShtTblObj.DataBodyRange[$Rw,$].Value2).ToString() 
        #Set-PSDebug -Trace 0 
      }
      $This.DtBody += ,@($This.DtBodySub) $This.DtBodySub = @() 
    }
    Set-PSDebug -Step
  }
}
class ExcelSheets 
{
 [Object]$Sheets 
 [string]$SheetsName 
 $SheetTable = [List[SheetTable]]::New() 
  [int]$SheetTableNum ExcelSheets( [object] $WorkSheet ) {
    $this.Sheets = $WorkSheet 
    $this.SheetsName = $WorkSheet.Name 
    $this.SheetTableNum = $WorkSheet.ListObjects.Count 1..$this.SheetTableNum | %{
      $This.SheetTable.Add( [SheetTable]::New($WorkSheet.ListObjects($PSItem)) ) 
    }
  }
} 
class ExcelBooks {
 [object]$exBs 
 [object]$RefBook 
 [string]$ExcelName 
 $ExSheets = [List[ExcelSheets]]::New() 
 [int]$ExSheetNum 
 ExcelBooks ( [object] $ExBooks ) {
    Try {
      $ThisexBs = New-Object -ComObject Excel.Application 
      $This.ExcelName = $ExBooks 
      $This.RefBook = $ThisexBs.Workbooks.Open( $ExBooks ) 
      $This.ExSheetNum = $This.RefBook.Sheets.Count 
      1..$This.ExSheetNum | %{
        $This.ExSheets.Add( [ExcelSheets]::new($This.RefBook.Sheets($PSItem)) )
      }
    }
    Catch { Write-Output( "Oops!" ) } 
    Finally { $This.DestExcelBooks() }
  }
  [string]Debug() {
  $ret = 1..$this.SheetTableNum | %{
    "$($This.ExSheets)"} 
    return $ret 
  }
  [void]DestExcelBooks() {
    $This.RefBook.Close() 
    if ( $This.exBs ) { $This.exBs.Quit() } 
    gv | ? Value -is [__ComObject] | clv 
    [gc]::Collect() 
    [gc]::WaitForPendingFinalizers() 
  } 
}
using namespace System.Collections.Generic 
using module ".\ExcelBooks.PSM1" 
Function ExcelDataView ( [ExcelBooks]$ExData ) 
{ 
  0..( $ExData.ExSheetNum - 1) | %{
    $sc = PSItemWrite−Output("SheetName:=($ExData.ExSheets[$sc].SheetsName)" ) 
Set-PSDebug -Trace 0 
    0..( $ExData.ExSheets[$sc].SheetTableNum - 1) | %{
      $st = PSItemWrite−Output("SheetTable:=($ExData.ExSheets[$sc].SheetTable[$st].SheetTableName)" ) 
      [string]" $($ExData.ExSheets[$sc].SheetTable[$st].Header)" 
      0..$ExData.ExSheets[$sc].SheetTable[$st].RowNum | %{
        $rw = PSItem[string]"($ExData.ExSheets[$sc].SheetTable[$st].DtBody[$rw])" 
      }
    }
  }
  Set-PSDebug -Trace 0 
} 
Function ExcelDataRead ( [string]$ExcelFileBook ) 
{
  return [ExcelBooks]::new( $ExcelFileBook ) 
}
Function Main ( [string[]]$refargs )
{
  $inFile = $refargs[0] 
  exBooks="(PSScriptRoot)($inFile)"
  $ExData = ExcelDataRead 
  $exBooks ExcelDataView $ExData 
} 

Main $args

@ECHO OFF 
@SETLOCAL ENABLEEXTENSIONS ENABLEDELAYEDEXPANSION 
@REM 
@REM Powershell Wraper Script 
@REM 
@SET SHELL=powershell.exe 
@SET BASEPARM=-ExecutionPolicy Bypass -Command 
@SET SCRIPTPATH=\vdinas03\FolderRedirect\NSS\fbs_matsumoto_kazuhi\home\fbsmatsu\develop\DevBase01 
@SET MainScript=!SCRIPTPATH!\DevBase01.PS1 
@SET ExBASE=DAT 
@SET ExBook=ServerLoginList.xlsx 
@SET ExSht1=DefineLoginSheet 
@SET ExTbl1=DefineLogin
@REM +--------------------------------------------------------------------------------+ 
@REM +------+ Program Gimmick PowerShell ExcelBook ExcelSheet ExcelTable 
@REM +--------------------------------------------------------------------------------+ 
@SET LAUNCHER=!SHELL! !BASEPARM! !MainScript! !ExBASE!!ExBook! !ExSht1! !ExTbl1! 
@REM +--------------------------------------------------------------------------------+
@REM @ECHO !LAUNCHER!
!LAUNCHER!
@ENDLOCAL 
@ECHO ON 
@REM 
@PAUSE
