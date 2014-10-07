Attribute VB_Name = "MSTSShare"
Option Explicit

Public TSLServer As String

Global Const ConstMaxShare = 4 'Max client connection and Max WorkStation STS share

'STS network control state----------------------
Global Const STS_BUSY = 0
Global Const STS_FREE = 1
Global Const STS_INIT_OK = 2
Global Const STS_INIT_NG = 3
Global Const STS_CALI_OK = 4
Global Const STS_CALI_NG = 5
Global Const STS_LINK_OK = 6
Global Const STS_LINK_NG = 7
'STS system control state-------------------------
Global Const STS_PAR_OK = 8
Global Const STS_PAR_NG = 9
Global Const STS_SCA_OK = 10
Global Const STS_SCA_NG = 11
Global Const STS_TRC_OK = 12
Global Const STS_TRC_NG = 13



Global Const STS_OFFLINE_OK = 3


'------------------------------
'NC:INIT
'NC:INIT_OK
'NC:INIT_NG
'------------------------------
'NC:CALI
'NC:CALI_OK
'NC:CALI_NG
'------------------------------
'NC:LINK
'NC:LINK_OK
'NC:LINK_NG
'-----------------------------
'NC:USE
'NC:BUSY
'NC:FREE
'NC:OFFLINE
'NC:USE_TIMEOUT
'NC:OFFLINE_OK
'NC:OFFLINE_NG

'NC:STOP
'NC:STOP_OK
'NC:STOP_NG

'------------------------------
'SC:PAR_XXX------>Set Testing Parameter
'SC:PAR_OK------->Set STS Parameter OK
'SC:PAR_NG------->Set STS Parameter NG
'-------------------------------
'SC:SCA_TEST
'SC:SCA_OK
'SC:SCA_NG
'------------------------------
'SC:TRC_GET
'SC:TRC_LEN_XXXXX
'SC:TRC_RECV_READY
'SC:TRC_SEND_OK
'SC:TRC_RECV_OK





