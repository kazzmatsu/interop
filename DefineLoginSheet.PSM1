using namespace System
using namespace System.IO
using namespace System.Text
using namespace System.Console
using namespace System.Collections
using namespace System.Collections.Generic
####################################################################################
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
####################################################################################
class LoginSheet : DefineLoginSheet
{
}
####################################################################################
class EncTable
{
  [int]$tNo
  [string]$User
  [string]$EncPass
  EncTable ( ) { }
  EncTable ( [int]$tNo, [string]$User, [string]$EncPass )
  {
    this.tNo      = $tNo
    this.User     = $User
    this.EncPass  = $EncPass
  }
}
####################################################################################
class CmdTable
{
  [int]$No
  [string]$Send
  [string]$Wait
  CmdTable ( ) { }
  CmdTable ( [int]$No, [string]$Send, [string]$Wait )
  {
    this.No    = $No
    this.Send  = $Send
    this.Wait  = $Wait
  }
}
####################################################################################
