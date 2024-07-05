using namespace System
using namespace System.IO
using namespace System.Text
using namespace System.Console
using namespace System.Collection
using namespace System.Collection.Generic
using namespace Microsoft.Office.Interop.Excel
#using module    ".\DefineLoginSheet.PSM1"
using module    ".\ListLoginSheet.PSM1"
using module    ".\MakeTeraTermMacro.PSM1"
####################################################################################
function ExcelTableClassRead([object[,]]$TargetDts )
{
  $ETCR = [LoginSheet]::new()
# Get-Variable ETCR
  $RStart = 1
# Get-Variable TargetDts | Format-List
  $RStart..$TargetDts.GetLength(0) | %{ $Row = $_ 
    if ( $Row -gt 1 ){ # Ignore the first Record 
<#
        $ETCR.Add([DefineLoginSheet]::new(
          DefineLoginSheet(  $TargetDts[$Row, 1] $TargetDts[$Row, 2] $TargetDts[$Row, 3] $TargetDts[$Row, 4] $TargetDts[$Row, 5] )
          )
        )
#        $ETCR.No = $TargetDts[$Row, 1]
#>
    }
  }
  Get-Variable ETCR
}
####################################################################################
function ExecuterTeraTermMacro([object]$MacroList)
{
  $intervalTime = 1500
  $TermModule = 'C:\Program Files (x86)\teraterm\ttpmacro.exe'
  $MacroList | %{$cnt = 1
    Write-Output ( " MacroScript := {0,2} $_" -f ($cnt++) )
    Start-Process -FilePath $TermModule -ArgumentList $_
    Start-Sleep -Milliseconds $intervalTime
  }
}
####################################################################################
function ExcelTableRead([object[,]]$TargetDts )
{
  $RStart = 1
  $TTFile = @()
  $RStart..$TargetDts.GetLength(0) | %{ $Row = $_ 
    if ( $Row -gt 1 ){ # Ignore the first Record 
      $TTFile += MakeTeraTermMacro $TargetDts[$Row, 2], $TargetDts[$Row, 3], $TargetDts[$Row, 4], $TargetDts[$Row, 5]
    }
#    Write-Output ( " # := {0,2}: $($TargetDts[$Row, 4]):$($TargetDts[$Row, 5])" -f ([int32]$Row - 1) )
  }
  return $TTFile
}
####################################################################################
function ExcelOpen([string]$exBook, [string]$exSheet, [string]$exTable )
{
  $exBs = New-Object -ComObject Excel.Application
  $exBs.Visible = $false # Debug $true
  $RefBook = $exBs.Workbooks.Open( "$($PSScriptRoot)\$($exBook)" )
  $TargetTable = $RefBook.Sheets( $exSheet ).ListObjects( $exTable )

  $TargetTableBase = $RefBook.Sheets( "TeratermMacro" )
  $VariablesLB =  $TargetTableBase.ListObjects( "LBCMD" ).Range.Value2
  $VariablesL2 =  $TargetTableBase.ListObjects( "L2CMD" ).Range.Value2
  $VariablesL3 =  $TargetTableBase.ListObjects( "L3CMD" ).Range.Value2
  $VariablesPS =  $TargetTableBase.ListObjects( "ENCPASS" ).Range.Value2

  TeratermEncData $VariablesPS

  $TeratermMacro = ExcelTableRead $TargetTable.Range.Value2

  $RefBook.Close()
  $exBs.Quit()
  Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable

  ExecuterTeraTermMacro $TeratermMacro

}
####################################################################################
function PreConfig ()
{
  Set-Location $($PSScriptRoot)
  @([DefFiles].GetEnumNames()) | % { 
    New-Item -Name $($_) -ItemType Directory -ErrorAction SilentlyContinue
  }
}
####################################################################################
function Main ( [string[]] $RefArgs )
{
  $ExcelBook  = $RefArgs[0]
  $ExcelSheet = $RefArgs[1]
  $ExcelTable = $RefArgs[2]
  PreConfig
  ExcelOpen $ExcelBook $ExcelSheet $ExcelTable
}
####################################################################################
Main ( $args )
####################################################################################
