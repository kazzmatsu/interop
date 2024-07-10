using namespace System
using namespace System.IO
using namespace System.Text
using namespace System.Console
using namespace System.Collections
using namespace System.Collections.Generic
using namespace Microsoft.Office.Interop.Excel
#using module    ".\DefineLoginSheet.PSM1"
using module    ".\ListLoginSheet.PSM1"
using module    ".\MakeTeraTermMacro.PSM1"
####################################################################################
function ExcelTableClassRead([string]$passFile, [string[]]$IncFlNm, [object[,]]$TargetDts )
{
  $ETCR = [List[LoginSheet]]::new()
  $RStart = 1
  $RStart..$TargetDts.GetLength(0) | %{ $Row = $PSItem
    if ( $Row -gt 1 ){ # Ignore the first Record 
        $ETCR.Add("","","","","")
    }
  }
}
####################################################################################
function ExecuterTeraTermMacro([object]$MacroList)
{
  $intervalTime = 1500
  $TermModule = 'C:\Program Files (x86)\teraterm\ttpmacro.exe'
  $MacroList | %{$cnt = 1
    Write-Output ( " TTMacro := $PSItem" )
    Start-Process -FilePath $TermModule -ArgumentList $PSItem
    Start-Sleep -Milliseconds $intervalTime
  }
}
####################################################################################
function ExcelTableRead([string]$passFile, [string[]]$IncFlNm, [object[,]]$TargetDts )
{
  $RStart = 1
  $TTFile = @()

  $RStart..$TargetDts.GetLength(0) | %{ $Row = $PSItem 
    if ( $Row -gt 1 ){ # Ignore the first Record 
      $TTFile += MakeTeraTermMacro $($passFile).TrimStart() $IncFlNm $TargetDts[$Row, 2], $TargetDts[$Row, 3], $TargetDts[$Row, 4], $TargetDts[$Row, 5]
    }
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

  $IncCmdFile = @()

  $TargetTableBase = $RefBook.Sheets( "TeratermMacro" )
  @([TTMCMD].GetEnumNames()) | % {
    $TargetCmd = $TargetTableBase.ListObjects( $PSItem ).Range.Value2
    if ( $PSItem -eq "ENCPASS") {
      $PassFile = TeratermEncData $PSItem  $TargetCmd
    } else {
      $IncCmdFile += CommonCmd $PSItem $TargetCmd
    }
  }

  $TeratermMacro = ExcelTableRead $PassFile ($IncCmdFile -ne $null) $TargetTable.Range.Value2

  $RefBook.Close()
  $exBs.Quit()

  gv | ? Value -is [__ComObject] | clv

  ExecuterTeraTermMacro $TeratermMacro

}
####################################################################################
function PreConfig ()
{
  Set-Location $($PSScriptRoot)
  @([DefFiles].GetEnumNames()) | % { 
    New-Item -Name $($PSItem) -ItemType Directory -ErrorAction SilentlyContinue
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
