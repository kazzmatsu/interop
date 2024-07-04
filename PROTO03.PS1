using namespace System
using namespace System.IO
using namespace System.Text
using namespace System.Console
using namespace System.Collection
using namespace System.Collection.Generic
using namespace Microsoft.Office.Interop.Excel
#using module    ".\DefineLoginSheet.PSM1"
using module    ".\ListLoginSheet.PSM1"
####################################################################################
function MakeTeraTermMacro ( [string[]]$params )
{
  $categy    = $params[0]
  $hostnm    = $params[1]
  $ipa       = $params[2]
  $user      = $params[3]
  $MacroDir  = $([DefFiles]::TetaTermMacro)
  $CommonDir = "..\$([DefFiles]::Common)"
  $fileName  = "$($PSScriptRoot)\$($MacroDir)\$($categy)_$($hostnm).TTL"
  if ( !( Test-Path -Path $fileName ) )
  {
    $fl = New-Object StreamWriter($fileName, $false, [Encoding]::GetEncoding("sjis"))
    $fl.WriteLine("USERNAME = '$user'")
    $fl.WriteLine("PASSFILE = '$($PSScriptRoot)\$([DefFiles]::PASDAT)\PASSWORD.DAT'")
    $fl.WriteLine("LOGSPATH = '$($PSScriptRoot)\$([DefFiles]::LOG)\'")
    $fl.WriteLine("HOSTADDR = '$ipa'")
    $fl.WriteLine("TITLE = '$hostnm'")
    $fl.WriteLine("getpassword PASSFILE USERNAME PASSWORD")
    $fl.WriteLine("COMMAND = HOSTADDR")
    $auterparam = switch($categy){"LB"{"challenge"}Default{"password"}}
    $fl.WriteLine("strconcat COMMAND ':22 /ssh /2 /auth=$($auterparam) /user=$user /passwd=' ")
    $fl.WriteLine("strconcat COMMAND PASSWORD")
    $fl.WriteLine("getdate LOG_NAME '$($categy)_$($hostnm)_%Y%m%d_%H%M%S_&h.log'")
    $fl.WriteLine("FULLPATH = LOGSPATH")
    $fl.WriteLine("strconcat FULLPATH LOG_NAME")
    $fl.WriteLine("connect COMMAND")
    $fl.WriteLine("if result <> 2 then ")
    $fl.WriteLine("  messagebox 'It could not be connected.' 'Connection Error' ")
    $fl.WriteLine("  end")
    $fl.WriteLine("endif")
    $fl.WriteLine("settitle TITLE")
    $fl.WriteLine("logopen FULLPATH")
    $fl.WriteLine("timeout = 300")
    $fl.WriteLine("wait '$' '#'")
    $fl.WriteLine("if result = 0 then")
    $fl.WriteLine("  end")
    $fl.WriteLine("endif")
    $fl.WriteLine("include '$($CommonDir)\$($categy)UpDate01.TTL'")
    $fl.WriteLine("end")
    $fl.Close()
  }
  return $fileName
}
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

  $TeratermMacro = ExcelTableRead $TargetTable.Range.Value2
  $RefBook.Close()
  $exBs.Quit()

  ExecuterTeraTermMacro $TeratermMacro

  Get-Variable | Where-Object Value -is [__ComObject] | Clear-Variable
}
####################################################################################
function PreConfig ()
{
  # Write-Output ( " koko := $($(Get-Location).ProviderPath)" )
  Set-Location $($PSScriptRoot)
  $MacroDirs = @([DefFiles].GetEnumNames())
  $MacroDirs | % { 
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
