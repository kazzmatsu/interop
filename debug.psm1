using namespace System.Collections.Generic
$exBs = New-Object -ComObject Excel.Application
$RefBook = $exBs.Workbooks.Open( "\\vdinas03\FolderRedirect\NSS\fbs_matsumoto_kazuhi\home\fbsmatsu\develop\AutoLogin\DAT\ServerLoginList.xlsx" )
$TargetTable = $RefBook.Sheets( "DefineLoginSheet" ).ListObjects( "DefineLogin")
#$TargetTable.Range.Value2

class DefineLoginSheet
{
  [string]$No
  [string]$Category
  [string]$hostname
  [string]$IPAddress
  [string]$LoginUser
  DefineLoginSheet( )  {}
  DefineLoginSheet( [string]$No, [string]$Category, [string]$hostname, [string]$IPAddress, [string]$LoginUser )
  {
    $this.No        = $No
    $this.Category  = $Category
    $this.hostname  = $hostname
    $this.IPAddress = $IPAddress
    $this.LoginUser = $LoginUser
  }
}

$ETCR = [List[DefineLoginSheet]]::new()
1..$TargetTable.Range.Value2.GetLength(0) | %{ $Row = $PSItem
  if ( $Row -gt 1 ){ # Ignore the first Record 
    $ETCR.Add( 
      [DefineLoginSheet]::new( "$($TargetTable.Range.Value2[$Row,1])",
                               "$($TargetTable.Range.Value2[$Row,2])",
                               "$($TargetTable.Range.Value2[$Row,3])",
                               "$($TargetTable.Range.Value2[$Row,4])",
                               "$($TargetTable.Range.Value2[$Row,5])"
      )
    )
  }
}

0..$ETCR.Count | %{ $ETCR[$_] | Format-Table }
$ETCR | Format-Table
$ETCR[0] | Format-Table

$RefBook.Close()
$exBs.Quit()
gv | ? Value -is [__ComObject] | clv
[gc]::Collect()
[gc]::WaitForPendingFinalizers()
