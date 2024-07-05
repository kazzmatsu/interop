using namespace System
using namespace System.IO
using namespace System.Text
using module    ".\ListLoginSheet.PSM1"
####################################################################################
function TeratermEncData ( [object[,]] $EncDat )
{
  $Rstart = 1
  $PASDTDir  = $([DefFiles]::PASDAT)
  $PsfileName  = "$($PSScriptRoot)\$($PASDTDir)\PASSWORDs.DAT"
  if ( !( Test-Path -Path $PsfileName ) )
  {
    $fl = New-Object StreamWriter($PsfileName, $false, [Encoding]::GetEncoding("sjis"))
    $fl.WriteLine("[Password]")
<#
    $Rstart..$EncDat.GetLength(0) | % { $Row = $_ 
      if ( $Row -gt 1 ){ # Ignore the first Record
        $fl.WriteLine("{0}={1}" -f [string]$EncDat[$Row,2] , [string]$EncDat[$Row,3] )
      }
    }
#>
    $fl.Close()
  }
}
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
    $fl.WriteLine("getdate LOG_NAME '$($categy)_$($hostnm)_&h_%Y%m%d_%H%M%S.log'")
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
