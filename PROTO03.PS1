using namespace System
using namespace System.IO
using namespace System.Text
using namespace System.Console
using namespace System.Collections
using namespace System.Collections.Generic
using namespace Microsoft.Office.Interop.Excel
using module    ".\DefineLoginSheet.PSM1"
using module    ".\ListLoginSheet.PSM1"
using module    ".\MakeTeraTermMacro.PSM1"
####################################################################################
function ExcelTableClassRead ([string]$passFile, [string[]]$IncFlNm, [object[,]]$TrgDt )
{ ### Dead Code ### 
  $ETCR = [List[DefineLoginSheet]]::new()
  $RStart = 1
  $RStart..$TrgDt.GetLength(0) | %{ $Row = $PSItem
    if ( $Row -gt 1 ){ # Ignore the first Record 
      $ETCR.Add( 
        [DefineLoginSheet]::new( "$($TrgDt[$Row,1])",
                                 "$($TrgDt[$Row,2])",
                                 "$($TrgDt[$Row,3])",
                                 "$($TrgDt[$Row,4])",
                                 "$($TrgDt[$Row,5])"
        )
      )
    }
  }
<#
  $ETCR | %{
    Write-Output ("{0,3}:{1,-5}:{2,-12}:{3,-16}:{4}" -f $PSItem.No, $PSItem.Category, $PSItem.hostname, $PSItem.IPAddress, $PSItem.LoginUser)
#    Write-Output ("$($PSItem.No), $($PSItem.Category), $($PSItem.hostname), $($PSItem.IPAddress), $($PSItem.LoginUser)")
  }
#>
}
####################################################################################
function ExecuterTeraTermMacro([object]$MacroList)
{
  $intervalTime = 3000
  $TermModule = 'C:\Program Files (x86)\teraterm\ttpmacro.exe'
  $MacroList | %{
    Write-Output ( " start TeratermMacro := $PSItem" )
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

  #$IncCmdFile | %{Write-Output($PSItem)}

#  ExcelTableClassRead $PassFile ($IncCmdFile -ne $null) $TargetTable.Range.Value2

  $TeratermMacro = ExcelTableRead $PassFile ($IncCmdFile -ne $null) $TargetTable.Range.Value2

  $RefBook.Close()
  $exBs.Quit()

  gv | ? Value -is [__ComObject] | clv
  [gc]::Collect()
  [gc]::WaitForPendingFinalizers()

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
