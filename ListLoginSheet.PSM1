using namespace System.Collections.Generic
using module    ".\DefineLoginSheet.PSM1"
####################################################################################
#class LoginSheet : System.Collections.Generic.List[DefineLoginSheet] {}
####################################################################################
enum DefFiles 
{
  TetaTermMacro
  Common
  PASDAT
  DAT
  LOG
}
####################################################################################
enum TTMS
{
  No        = 1
  Category  = 2
  hostname  = 3
  IPAddress = 4
  LoginUser = 5
}
####################################################################################
enum LayerCategory
{
  LB
  L2SW
  L3SW
}
####################################################################################
enum TTMCMD
{
  LBCMD
  L2CMD
  L3CMD
  ENCPASS
}
####################################################################################
