#Region ;**** Directives created by AutoIt3Wrapper_GUI ****
#AutoIt3Wrapper_Outfile_type=a3x
#AutoIt3Wrapper_Icon=Res\ohiohealth.ico
#AutoIt3Wrapper_Outfile=:\<redacted>\Utils\_.Sources\AutoCollections.a3x
#AutoIt3Wrapper_Outfile_x64=..\_.rc\AutoCollections.a3x
#AutoIt3Wrapper_UseUpx=y
#AutoIt3Wrapper_Res_Description=SCCM Deployment Helper
#AutoIt3Wrapper_Res_Fileversion=23.322.956.53
#AutoIt3Wrapper_Res_Fileversion_AutoIncrement=y
#AutoIt3Wrapper_Res_Fileversion_First_Increment=y
#AutoIt3Wrapper_Res_ProductName=AutoCollections
#AutoIt3Wrapper_Res_Language=1033
#AutoIt3Wrapper_Run_Au3Stripper=y
#EndRegion ;**** Directives created by AutoIt3Wrapper_GUI ****

#cs ----------------------------------------------------------------------------

 AutoIt Version: 3.3.14.5
 Author:         BiatuAutMiahn[@outlook.com]

 Script Function:

[TODO]
-New Auth, when loading auth tokens and more than one token available, show dropdown /w last element being Add button.
-If single token available only show that token as text field.
-If unchecking remember, delete only that token from the AuthDB.

-Implement confirmation dialog /w listview for Add/Remove/Install/Uninstall when MultiSel. _ArrayToString
    -ConfDlg always when Adding. MsgBox ConfDlg when not MultiSel.
-Bug where toast appears offset on separate monitor from gui, and monitors differ in resolution.
    -Solution is to get resolution of primary monitor, not the default that autoit checks. (Deprioritized)
-Implement ObjEvent for InstallationStatus and get rid of aWatch/Watch2 (Deprioritized)
    -Listview needs to update when watching for app installation.
-Using Space to select in AddGui does not append item to queue.
-Disable Gui when performing actions (Run all Cycles)
-Add "Install Client" > "Install|Reinstall"
-Clear watch queue when closing host.
-The Adding... status in AddGui needs space between name and parentheses.
-Move collection removal functionality to AddGui.
-If device is offline, disable functionality that requires online device, ie adding/removing collections.
-Change alt status for Pending reboot.
-Add reboot host to actions.
#ce ----------------------------------------------------------------------------
Opt("TrayAutoPause", 0)
Opt("TrayIconHide", 1)
#include <Debug.au3>
#include <Array.au3>
#include <ButtonConstants.au3>
#include <EditConstants.au3>
#include <GUIConstantsEx.au3>
#include <GuiStatusBar.au3>
#include <ListViewConstants.au3>
#include <StaticConstants.au3>
#include <WindowsConstants.au3>
#include <GuiEdit.au3>
#include <GuiListView.au3>

#include "Includes\Common\CryptProtect.au3"
#include "Includes\Common\InfinityCommon.au3"
#include "Includes\Common\Helpers.au3"
#include "Includes\Common\Toast.au3"

_InitBasicTray()

Global Const $VERSION = "23.322.956.53"
Global $sAlias="AutoCollections v"&$VERSION

_Log($sAlias)

; Globals
Global $bHasToken=False
Global $bHostActive=False
Global $bWarnRemember=True
Global $bRemember=False
Global $bHostPing=False, $bHostResolve=False, $bHostWmi=False, $sHost="", $aHostCollections, $sHostResourceId
Global $bCanInstall,$bCanUninstall,$bCanRemove,$iSelIdx=0,$bEnMulti=False, $aSel[0], $bEnColMgmt=True
Global $idCtx, $idCtxIn, $idCtxUn, $idCtxRm, $idCtxDesc=-9999, $bExit=False, $aWatch[0][4], $bWatch=False, $tWatch=0
Global $bForceSync=False, $bWaitApp=False, $bForceSyncLast=True, $sSearchLast="", $tSearch, $sSearch="", $bSeach=False
Global $tSearch, $hAddGui, $idAddSearch, $hAddSearch, $idAddSearch, $idAddAppsLV, $hAddAppsLV, $idAddSync, $idAddWait, $idAddRefresh, $idAddDone, $idAddStatus, $idAddAdd
Global $bAddSync, $bWaitApp, $bSearch, $aSearch, $bRemoveWarn=True, $bWarnPurpose=True, $bSearchAbort=False, $bAddMod=False
Global $aAddQueue[0][3], $aModQueue[0][3], $tWatch2=0, $bWatch2=False, $bWatch2Proc=False, $sSelMagic=""
Global $iStatDelay=125/2
Global $sAuthIni=@UserProfileDir&"\ohAuthToken.ini"
;Global $g_aAuth[2]
;Global $bGuiAddDefSync=True
Global $oWMI, $aAppsSMS
;026|027|101|102|103|104|105|107|122|123|
Local $sSched="001|002|003|010|021|022|031|032|108|111|113|114|121|221|222"
Local $aSched=StringSplit($sSched,'|')

Local $aSchedDesc[]=[ _
    "Hardware Inventory Cycle", _
    "Software Inventory Cycle", _
    "Discovery Data Collection Cycle", _
    "File Collection Cycle", _
    "Machine Policy Retrieval Cycle", _
    "Machine Policy Evaluation Cycle", _
    "Software Metering Usage Report Cycle", _
    "Windows Installers Source List Update Cycle", _
    "Software Updates Assignments Evaluation Cycle", _
    "State Message Refresh", _
    "Software Update Scan Cycle", _
    "Update Store Policy", _
    "Application Deployment Evaluation Cycle", _
    "Endpoint deployment reevaluate", _
    "Endpoint AM policy reevaluate" _
]


$iWidth=768+64+8
$iHeight=256+128
$iMargin=8
$iBtnCnt=7
$iBtnWidth=Int(($iWidth-32-(8*$iBtnCnt))/$iBtnCnt)
$iBtnTop=9+20+8
$iBtnLeft=4+($iWidth/2)-($iBtnWidth+$iMargin)*$iBtnCnt/2

; Initial Warning.
;~ Local $sMsg=""
;~ $sMsg&="This software does not directly modify any device's"&@LF
;~ $sMsg&="settings or configuration. It only reads information"&@LF
;~ $sMsg&="from the target device and issues commands via SCCM"&@LF
;~ $sMsg&="and WMI. In effect it is equivelent to 'Add Resource'"&@LF
;~ $sMsg&="in SCCM, and 'Install' or 'Uninstall' from software"&@LF
;~ $sMsg&="center."
;~ If @Compiled Then MsgBox(48,$sAlias,$sMsg)
Local $vOpt=getOpt('sHost')
_Toast_Set(0,-1,-1,-1,-1,-1,"Consolas",125,125)
; Initialize GUI
$hWnd = GUICreate($sAlias, $iWidth, $iHeight)
GUISetFont(10, 400, 0, "Consolas")
$idHost = GUICtrlCreateInput("", 8, 9, 124, 20,$ES_CENTER)
If $vOpt Then GUICtrlSetData(-1,$vOpt)
$idHotKey = GUICtrlCreateDummy()
Dim $AccelKeys[1][2] = [["{ENTER}", $idHotKey]]; Set accelerators
GUISetAccelerators($AccelKeys)
GUICtrlSetTip(-1,"Enter hostname to manage here.")
_GUICtrlEdit_SetCueBanner($idHost, "Host", True)
$idOpid = GUICtrlCreateInput("", 8+124+4, 9, 124, 20,$ES_CENTER)
_GUICtrlEdit_SetCueBanner($idOpid, "OPID", True)
GUICtrlSetTip(-1,"Enter your OPID here.")
$idPass = GUICtrlCreateInput("", 8+128+124+4, 9, 124, 20, $ES_PASSWORD)
_GUICtrlEdit_SetCueBanner($idPass, "Password", True)
GUICtrlSetTip(-1,"Enter your password here.")
$idRemember = GUICtrlCreateCheckbox("Remember (Encrypted)",396, 10, 128+32, 17)
GUICtrlSetTip(-1,"Remember my credentials.")
$idEnMulti = GUICtrlCreateCheckbox("Multiselect",396+128+32+4, 10, 96, 17)
GUICtrlSetTip(-1,"Enable multiple selections."&@LF&"(When installing/uninstalling/removing multiple items.)")
$idInit = GUICtrlCreateButton("Start", $iBtnLeft, $iBtnTop, $iBtnWidth, 25)
GUICtrlSetTip(-1,"Starts collection managment.")
$idUninit = GUICtrlCreateButton("Done", $iBtnLeft, $iBtnTop, $iBtnWidth, 25)
GUICtrlSetTip(-1,"Stops collection managment.")
$idAdd = GUICtrlCreateButton("Add", $iBtnLeft+$iBtnWidth+$iMargin, $iBtnTop, $iBtnWidth, 25)
GUICtrlSetTip(-1,"Adds host to an APP collection.")
$idRemove = GUICtrlCreateButton("Remove", $iBtnLeft+($iBtnWidth+$iMargin)*2, $iBtnTop, $iBtnWidth, 25)
GUICtrlSetTip(-1,"Remove host from selected collection.")
$idInstall = GUICtrlCreateButton("Install", $iBtnLeft+($iBtnWidth+$iMargin)*3, $iBtnTop, $iBtnWidth, 25)
GUICtrlSetTip(-1,"Trigger software installation.")
$idUninstall = GUICtrlCreateButton("Uninstall", $iBtnLeft+($iBtnWidth+$iMargin)*4, $iBtnTop, $iBtnWidth, 25)
GUICtrlSetTip(-1,"Trigger software Uninstallation.")
$idHostAct = GUICtrlCreateButton("Actions", $iBtnLeft+($iBtnWidth+$iMargin)*5, $iBtnTop, $iBtnWidth, 25)
GUICtrlSetTip(-1,"Various actions to perform on host.")
$idHostActCtx = GUICtrlCreateContextMenu($idHostAct)
$hHostActCtx=GUICtrlGetHandle($idHostActCtx)
;$idHostActCtx_Pol = GUICtrlCreateMenu("Policy", $idHostActCtx)
$idHostActCtx_PolPurge = GUICtrlCreateMenuItem("Policy Purge/Reset", $idHostActCtx)
$idHostActCtx_PolSync = GUICtrlCreateMenuItem("Policy Full Sync", $idHostActCtx)
$idHostActCtx_PolCycle = GUICtrlCreateMenuItem("Run CCM Cycles", $idHostActCtx)
$idHostActCtx_PolAllCycle = GUICtrlCreateMenuItem("Run All Cycles", $idHostActCtx)
GUICtrlCreateMenuItem("", $idHostActCtx)
$idHostActCtx_ExportColl = GUICtrlCreateMenuItem("Colls2Clip", $idHostActCtx)
;$idHostActCtx_ActCCM = GUICtrlCreateMenu("CCM Actions", $idHostActCtx)
;$idHostActCtx_ActCCM = GUICtrlCreateMenuItem("Purge/Reset", $idHostActCtx_Pol)

;GUICtrlSetTip(-1,"Tells host to discard current policy and reload.")
;$idSyncPol = GUICtrlCreateButton("Sync Policy", $iBtnLeft+($iBtnWidth+$iMargin)*6, $iBtnTop, $iBtnWidth, 25)
;GUICtrlSetTip(-1,"Tells host to grab full policy instead of policy changes.")
$idRefresh = GUICtrlCreateButton("Refresh", $iBtnLeft+($iBtnWidth+$iMargin)*6, $iBtnTop, $iBtnWidth, 25)
GUICtrlSetTip(-1,"Refresh host's collections.")
$idStatus = _GUICtrlStatusBar_Create($hWnd)
_GUICtrlStatusBar_SetText($idStatus,"Initializing")
GUICtrlCreateLabel("Host Collections:", 8, $iHeight-256-32-16-4, 128, 19)
$idHostLV = GUICtrlCreateListView("CID|State|Last Installed|Version|Name", 8, $iHeight-256-32, $iWidth-16, 256)
$hHostLV = GUICtrlGetHandle($idHostLV)
_GUICtrlListView_SetColumnWidth($hHostLV,0,64+32)
_GUICtrlListView_SetColumnWidth($hHostLV,1,64+32+8)
_GUICtrlListView_SetColumnWidth($hHostLV,2,96+16)
_GUICtrlListView_SetColumnWidth($hHostLV,3,64+32)
;$hHostLV=_GUICtrlListView_Create($hWnd,"CID|State|Last Installed|Name", 8, $iHeight-256-32, $iWidth-16, 256 ,BitOr($LVS_LIST,$LVS_NOSORTHEADER,$LVS_REPORT))

GuiSetStates(0)
GUISetState(@SW_SHOW)

; Initialize Credentials...
_Log("ohAuth")
$bHasToken=ohAuth_loadToken()
if $bHasToken Then
	$bRemember=True
	GuiCtrlSetData($idOpid,$g_aAuth[0][0])
	GuiCtrlSetData($idPass,$g_aAuthSalts[@HOUR])
EndIf
_GUICtrlStatusBar_SetText($idStatus,"Initializing SMS...")
_Log("EnsureSMS")
_EnsureSMS()
_GUICtrlStatusBar_SetText($idStatus,"Initializing SMS...Done")
Sleep($iStatDelay)
;~ _GUICtrlStatusBar_SetText($idStatus,"Getting App Deployments...")
;~ _Log("SMSGetCollections")
;~ $aAppsSMS=_SMSGetCollections()
;~ _ArraySort($aAppsSMS,0,0,0,0)
;~ _GUICtrlStatusBar_SetText($idStatus,"Getting App Deployments...Done")
;~ Sleep($iStatDelay)
_GUICtrlStatusBar_SetText($idStatus,"Ready")
_Log("Initialized")
GuiSetStates(1)
#EndRegion ### END Koda GUI section ###
;MsgBox(64,"",_CheckHost(GUICtrlRead($idHost)))
If GUICtrlRead($idHost)<>"" Then _InitHost()
;~ GuiSetStates(3)
;~ _AddCollectionGui()
;OHP01D59
;~ Local $aHosts[]=[1,"LT500016"]
;~ Local $sHostResourceId
;~ For $i=1 To $aHosts[0]
;~     $sHost=$aHosts[$i]
;~ 	$oResults=$g_oSMS.ExecQuery("SELECT * FROM SMS_R_System WHERE Name = '"&$sHost&"'")
;~ 	_Log("InitHost,SMS")
;~ 	If $oResults.Count==0 Then
;~         ConsoleWrite("NoRID:"&$sHost&@CRLF)
;~         ContinueLoop
;~     EndIf
;~ 	If $oResults.Count>1 Then
;~         ConsoleWrite("TooManyRid:"&$sHost&@CRLF)
;~         ContinueLoop
;~     EndIf
;~     For $oObj In $oResults
;~         $sHostResourceId=$oObj.ResourceID
;~         _Log("InitHost,ResourceID:"&$sHostResourceId)
;~     Next
;~     $iRet=_CollectionAddResource(StringUpper($sHost),$sHostResourceId,"OHP01D59",True)
;~     ConsoleWrite($iRet&@CRLF)
;~ Next
;~ For $i=1 To $aHosts[0]
;~     Local $sHost=$aHosts[$i]
;~     ConsoleWrite($sHost&@CRLF)
;~ 	_EnsureWMI($sHost,"/root/ccm/ClientSDK",True)
;~     Local $oNS=_EnsureWMI($sHost,"/root/ccm",True)
;~     _HostResetPolicy($sHost,False)
;~     Local $oClass = $oNS.Get("SMS_Client")
;~     Local $aSched[]=["121","021","022","002"]
;~     For $j=0 To UBound($aSched)-1
;~         _HostTriggerSchedule($sHost,$aSched[$j],$oClass,$oNS)
;~         Sleep($iStatDelay)
;~     Next
;~     ;_HostRefreshPolicy($sHost,$idStatus)
;~ Next
;~ ;Get ResourceId

;~ Exit
While 1
	$nMsg = GUIGetMsg()
	Switch $nMsg
        Case $GUI_EVENT_CLOSE
            _GUICtrlStatusBar_SetText($idStatus,"Exiting...")
			$bExit=True
			_Exit()
        Case $idHostAct
            Local $aPos, $x, $y
            $aPos=ControlGetPos($hWnd,"",$nMsg)
            $x=$aPos[0]
            $y=$aPos[1]+$aPos[3]
            ClientToScreen($hWnd,$x,$y)
            TrackPopupMenu($hWnd,$hHostActCtx,$x,$y)
        Case $idHostActCtx_PolAllCycle
            _Log("CCMCycleAll")
            ;Local $aSched[]=["021","022"]
            ;_HostTriggerSchedule($sHost,$aSched)
            _HostRefreshPolicy($sHost,$idStatus,$sSched)
            Sleep($iStatDelay)
			_GUICtrlStatusBar_SetText($idStatus,"Ready")
        Case $idHostActCtx_PolCycle
            _Log("CCMCycles")
            _HostRefreshPolicy($sHost,$idStatus)
            Sleep($iStatDelay)
			_GUICtrlStatusBar_SetText($idStatus,"Ready")
        Case $idHostActCtx_PolPurge
            _Log("PurgePolicy")
            GuiSetStates(3)
			Local $sMsg=""
			$sMsg&="Warning! Forcing a complete policy reset will"&@LF
			$sMsg&="purge existing policies. If there are legacy"&@LF
			$sMsg&="deployments that are no longer available they may be removed."&@LF&@LF
			$sMsg&="Are you sure you want to continue?"
            $iRet=MsgBox(48+1,"Policy Reset",$sMsg)
            If $iRet<>1 Then
                _GUICtrlStatusBar_SetText($idStatus,"Ready")
                GuiSetStates(2)
                ContinueLoop
            EndIf
            _GUICtrlStatusBar_SetText($idStatus,"Purge Host's CM Policy...")
            _HostResetPolicy($sHost,True)
            Sleep($iStatDelay)
            _GUICtrlStatusBar_SetText($idStatus,"Purge Host's CM Policy...Done")
            Sleep($iStatDelay)
            _HostRefreshPolicy($sHost,$idStatus)
            MsgBox(64,"Policy Reset","Note: this action may take a long time.")
			_GUICtrlStatusBar_SetText($idStatus,"Ready")
            GuiSetStates(2)
        Case $idHostActCtx_PolSync
            _Log("FullPolicySync")
            GuiSetStates(3)
            _GUICtrlStatusBar_SetText($idStatus,"Host Request Full CM Policy...")
            _HostResetPolicy($sHost,False)
            Sleep($iStatDelay)
            _GUICtrlStatusBar_SetText($idStatus,"Host Request Full CM Policy...Done")
            Sleep($iStatDelay)
            _HostRefreshPolicy($sHost,$idStatus)
			_GUICtrlStatusBar_SetText($idStatus,"Ready")
            GuiSetStates(2)
		Case $idInit
			_InitHost()
        Case $idHotKey
            ;ConsoleWrite(_GuiCtrlGetFocus($hWnd)&"=="&$idHost&@CRLF)
            If _GuiCtrlGetFocus($hWnd)<>$idHost And _GuiCtrlGetFocus($hWnd)<>$idOpid And _GuiCtrlGetFocus($hWnd)<>$idPass Then ContinueLoop
            $sHost=StringStripWS(GUICtrlRead($idHost),3)
            ;If $sHost=="" Then ContinueLoop
            _InitHost()
		Case $idRefresh
            _Log("RefreshHost")
			GuiSetStates(3)
			_GUICtrlListView_SetItemSelected($hHostLV,-1,False,False)
            ;$sSelMagic=$aHostCollections[$iSelIdx][15]
            RefreshHost()
            For $i=0 To UBound($aHostCollections,1)-1
                If $sSelMagic==$aHostCollections[$i][15] Then
                    $iSelIdx=$i
                    ExitLoop
                EndIf
            Next
            If Not @Compiled Then _ArrayDisplay($aHostCollections)
			_GUICtrlStatusBar_SetText($idStatus,"Ready")
			GuiSetStates(2)
		Case $idUninit
            _Log("UninitializeHost")
			GuiSetStates(1)
            ObjEvent($oWMI.CCM_InstanceEvent)
            $ogCcmEvent=-1
			$bHostActive=False
			$bHostPing=False
			$bHostResolve=False
			$bHostWmi=False
			$sHost=""
            $bWatch2=False
            AdlibUnregister("_AppWatch2")
            While $bWatch2Proc
                If TimerDiff($tWatch2)>=2000 Then
                    $bWatch2Proc=False
                    ExitLoop
                EndIf
                Sleep(1)
            WEnd
            ;;$bWatch2=True
            Local $aEmpty[0][0]
			$aHostCollections=$aEmpty
			GUIRegisterMsg($WM_NOTIFY, "")
			_GUICtrlListView_BeginUpdate($hHostLV)
			_GUICtrlListView_DeleteAllItems($hHostLV)
			_GUICtrlListView_EndUpdate($hHostLV)
		Case $idRemember
            _Log("RememberAuth")
;~             _Log("TriggeredDescription,"&_ArrayToString($aHostCollections,'|',$iSelIdx,$iSelIdx))
			If GuiCtrlRead($idRemember)==$GUI_CHECKED Then
				$bRemember=True
				If $bWarnRemember Then
					$sMsg="Credentials will be stored in your user profile directory and will only be accessible from this account on this machine."
					MsgBox(48,"ohAuth - Remember Credentials",$sMsg,0,$hWnd)
					$bWarnRemember=False
				EndIf
			Else
				$bRemember=False
				if $bHasToken Then
					If MsgBox(BitOr($MB_ICONWARNING,$MB_OKCANCEL),"Warning","This will delete your cached token.",0,$hWnd)==1 Then
                        _Log("RememberAuth,Disabled,DeleteToken")
						_GUICtrlStatusBar_SetText($idStatus,"Deleting token...")
                        FileDelete($sAuthIni)
						Sleep(500)
						_GUICtrlStatusBar_SetText($idStatus,"Deleting token...Done")
						Sleep($iStatDelay)
						$bHasToken=False
						GuiCtrlSetData($idOpid,"")
						GuiCtrlSetData($idPass,"")
						GuiSetStates(1)
						Sleep(500)
						_GUICtrlStatusBar_SetText($idStatus,"Ready")
					Else
                        _Log("RememberAuth,Enabled")
						GuiCtrlSetState($idRemember,$GUI_CHECKED)
					EndIf
				EndIf
			EndIf
		Case $idCtxDesc
            _Log("TriggeredDescription,"&_ArrayToString($aHostCollections,'|',$iSelIdx,$iSelIdx))
			If $idEnMulti==True Then
				DisposeCtx()
				ContinueLoop
			EndIf
			MsgBox(64,"Description - "&$aHostCollections[$iSelIdx][0],$aHostCollections[$iSelIdx][1])
		Case $idEnMulti
			If GuiCtrlRead($idEnMulti)==$GUI_CHECKED Then
				$bEnMulti=True
				DisposeCtx()
                _Log("EnabledMultiSelect")
			Else
				$bEnMulti=False
                _Log("DisabledMultiSelect")
			EndIf
			_GUICtrlListView_SetExtendedListViewStyle($hHostLV,BitOr($LVS_EX_TWOCLICKACTIVATE,$WS_EX_CLIENTEDGE,$LVS_EX_DOUBLEBUFFER,$LVS_EX_FULLROWSELECT,$LVS_EX_GRIDLINES,$LVS_EX_GRIDLINES,$bEnMulti ? $LVS_EX_CHECKBOXES : 0))
			;Tidy_Parameters$iHostStateLV=GUICtrlGet($idHostLV)
			;GUICtrlSetStyle($idHostLV,$bEnMulti ? BitOr($iHostStateLV,$LVS_SINGLESEL) : BitAND($iHostStateLV,$LVS_SINGLESEL))
		Case $idInstall
            _Log("InstallApp,"&_ArrayToString($aHostCollections,'|',$iSelIdx,$iSelIdx))
			GuiSetStates(3)
			;1|sName|sDesc|$iRev|$sStat|$sLastInst|$sInProgressActions|$sAvailableActions|$bIsMachineTarget|$sId|$sCollectionId|$sVersion
			$colItems=$oWMI.ExecQuery("SELECT * FROM CCM_Application", "WQL", $wbemFlagReturnImmediately)
			Local $vRet=Null
			For $objItem In $colItems
				If $aHostCollections[$iSelIdx][8]<>$objItem.Id Then ContinueLoop
				$vRet=$objItem.Install($objItem.Id,$objItem.Revision,$objItem.IsMachineTarget,0,"High",False)
				ExitLoop
			Next
			If $vRet==0 Then
                _Toast_Hide()
				_Toast_Show(0, $sAlias, "Triggered Install for "&$aHostCollections[$iSelIdx][0],-10,False)
				;_GUICtrlStatusBar_SetText("Installing "&$aHostCollections[$iSelIdx][0]&"...")
				$iMax=UBound($aWatch)
				ReDim $aWatch[$iMax+1][6]
				$aWatch[$iMax][0]=$aHostCollections[$iSelIdx][8]
				$aWatch[$iMax][1]="Install"
				$aWatch[$iMax][2]=$aHostCollections[$iSelIdx][0]
				$aWatch[$iMax][3]=True
				$aWatch[$iMax][4]=$aHostCollections[$iSelIdx][11]
				$aWatch[$iMax][5]=$aHostCollections[$iSelIdx][12]
				$bWatch=True
				$tWatch=TimerInit()
                If Not $bEnMulti Then $sSelMagic=$aHostCollections[$iSelIdx][15]
                RefreshHost()
                If Not $bEnMulti Then
                    For $i=0 To UBound($aHostCollections,1)-1
                        If $sSelMagic==$aHostCollections[$i][15] Then
                            $iSelIdx=$i
                            ExitLoop
                        EndIf
                    Next
                EndIf
			Else
				MsgBox(64,"Install - "&$aHostCollections[$iSelIdx][0],"Installation not started. Unknown Error.")
			EndIf
;~ 			If $vRet==0 Then AdlibRegister("_ActWatch",1000)
			_GUICtrlStatusBar_SetText($idStatus,"Ready")
			GuiSetStates(2)
			GUICtrlSetState($idHostLV, $GUI_FOCUS)
            _GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
            _GUICtrlListView_EnsureVisible($hHostLV,$iSelIdx)
            ;_GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
		Case $idUninstall
            _Log("UninstallApp,"&_ArrayToString($aHostCollections,'|',$iSelIdx,$iSelIdx))
			GuiSetStates(3)
			;1|sName|sDesc|$iRev|$sStat|$sLastInst|$sInProgressActions|$sAvailableActions|$bIsMachineTarget|$sId|$sCollectionId|$sVersion
			$colItems=$oWMI.ExecQuery("SELECT * FROM CCM_Application", "WQL", $wbemFlagReturnImmediately)
			Local $vRet=Null
			For $objItem In $colItems
				If $aHostCollections[$iSelIdx][8]<>$objItem.Id Then ContinueLoop
				$vRet=$objItem.Uninstall($objItem.Id,$objItem.Revision,$objItem.IsMachineTarget,0,"High",False)
				ExitLoop
			Next
			If $vRet==0 Then
                _Toast_Hide()
				_Toast_Show(0, $sAlias, "Triggered Uninstall for "&$aHostCollections[$iSelIdx][0],-10,False)
				;_Toast_Hide()
				;_GUICtrlStatusBar_SetText("Uninstalling "&$aHostCollections[$iSelIdx][0]&"...")
				$iMax=UBound($aWatch)
				ReDim $aWatch[$iMax+1][6]
				$aWatch[$iMax][0]=$aHostCollections[$iSelIdx][8]
				$aWatch[$iMax][1]="Uninstall"
				$aWatch[$iMax][2]=$aHostCollections[$iSelIdx][0]
				$aWatch[$iMax][3]=True
				$aWatch[$iMax][4]=$aHostCollections[$iSelIdx][11]
				$aWatch[$iMax][5]=$aHostCollections[$iSelIdx][12]
				$bWatch=1
				$tWatch=TimerInit()
                If Not $bEnMulti Then $sSelMagic=$aHostCollections[$iSelIdx][15]
                RefreshHost()
                If Not $bEnMulti Then
                    For $i=0 To UBound($aHostCollections,1)-1
                        If $sSelMagic==$aHostCollections[$i][15] Then
                            $iSelIdx=$i
                            ExitLoop
                        EndIf
                    Next
                EndIf
			Else
				MsgBox(64,"Uninstall - "&$aHostCollections[$iSelIdx][0],"Uninstallation not started. Unknown Error.")
			EndIf
;~ 			If $vRet==0 Then AdlibRegister("_ActWatch",1000)
			_GUICtrlStatusBar_SetText($idStatus,"Ready")
			GuiSetStates(2)
			GUICtrlSetState($idHostLV, $GUI_FOCUS)
            _GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
            _GUICtrlListView_EnsureVisible($hHostLV,$iSelIdx)
		Case $idAdd
			GuiSetStates(3)
            $bWatch2=False
            ;$tWatch2=TimerInit()
            AdlibUnregister("_AppWatch2")
            While $bWatch2Proc
                If TimerDiff($tWatch2)>=2000 Then
                    $bWatch2Proc=False
                    ExitLoop
                EndIf
                Sleep(1)
            WEnd
            GUIRegisterMsg($WM_NOTIFY,"")
			;MsgBox(48,"Add - "&$aHostCollections[$iSelIdx][0],"Experimental")
			_AddCollectionGui()
            ;GUIRegisterMsg($WM_COMMAND, "WM_COMMAND");only used for EN_CHANGE so far
            GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
            If $bAddMod Then
                If Not $bEnMulti Then $sSelMagic=$aHostCollections[$iSelIdx][15]
                RefreshHost()
                If Not $bEnMulti Then
                    For $i=0 To UBound($aHostCollections,1)-1
                        If $sSelMagic==$aHostCollections[$i][15] Then
                            $iSelIdx=$i
                            ExitLoop
                        EndIf
                    Next
                EndIf

            EndIf
            ;$tWatch2=TimerInit()
            ;AdlibRegister("_AppWatch2",250)
            ;;$bWatch2=True
			; Get ResourceId
			Sleep($iStatDelay)
			GuiSetStates(2)
			GUICtrlSetState($idHostLV, $GUI_FOCUS)
            _GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
            _GUICtrlListView_EnsureVisible($hHostLV,$iSelIdx)
            _GUICtrlStatusBar_SetText($idStatus,"Ready")
		Case $idRemove
            GuiSetStates(3)
            _Log("RemoveCollection,"&_ArrayToString($aHostCollections,'|',$iSelIdx,$iSelIdx))
			GuiSetStates(3)
			If $bEnMulti Then
                If Not @Compiled Then
                    Local $iModQueue=UBound($aModQueue,1)
                    If Not @Compiled Then _ArrayDisplay($aModQueue,$iModQueue)
                    If Not _ConfirmActMulti("Remove from Collections","Are you sure you want to remove this host from these collections?") Then
                        GuiSetStates(2)
                        GUICtrlSetState($idHostLV, $GUI_FOCUS)
                        _EvalItemActs($iSelIdx)
                        _GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
                        _GUICtrlListView_EnsureVisible($hHostLV,$iSelIdx)
                        _GUICtrlStatusBar_SetText($idStatus,"Ready")
                        ContinueLoop
                    EndIf
                    _Log("RemoveCollection,"&_ArrayToString($aModQueue))
                    For $i=0 To UBound($aModQueue,1)-1
                        If Not $aModQueue[$i][2] Then ContinueLoop
                        _GUICtrlStatusBar_SetText($idStatus,"Removing "&$sHost&" from "&$aModQueue[$i][1]&" ("&$aModQueue[$i][0]&")...")
                        Sleep(1000)
                        _GUICtrlStatusBar_SetText($idStatus,"Removing "&$sHost&" from "&$aModQueue[$i][1]&" ("&$aModQueue[$i][0]&")...Done")
                        Sleep(250)
                    Next
                    If Not $bEnMulti Then $sSelMagic=$aHostCollections[$iSelIdx][15]
                    RefreshHost()
                    If Not $bEnMulti Then
                        For $i=0 To UBound($aHostCollections,1)-1
                            If $sSelMagic==$aHostCollections[$i][15] Then
                                $iSelIdx=$i
                                ExitLoop
                            EndIf
                        Next
                    EndIf
                Else
                    MsgBox(48,"RemoveMulti - "&$aHostCollections[$iSelIdx][0],"Not yet Implemented")
                EndIf
                GuiSetStates(2)
                GUICtrlSetState($idHostLV, $GUI_FOCUS)
                _EvalItemActs($iSelIdx)
                _GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
                _GUICtrlListView_EnsureVisible($hHostLV,$iSelIdx)
                _GUICtrlStatusBar_SetText($idStatus,"Ready")
                ContinueLoop
			Else
                If Not StringInStr($aHostCollections[$iSelIdx][3],"NotInstalled",0) Then
                    MsgBox(48,"Remove - "&$aHostCollections[$iSelIdx][0],"You must uninstall this deployment before removing it!")
					GuiSetStates(2)
                    GUICtrlSetState($idHostLV, $GUI_FOCUS)
                    _EvalItemActs($iSelIdx)
                    _GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
                    _GUICtrlListView_EnsureVisible($hHostLV,$iSelIdx)
					_GUICtrlStatusBar_SetText($idStatus,"Ready")
                    ContinueLoop
                EndIf
				;If $bRemoveWarn Then
                ;    MsgBox(48,"Remove - "&$aHostCollections[$iSelIdx][0],"Warning! Even if you remove this software deployment, the host may retain the last version installed/added. Only forcing a complete policy reset will remove this software. Forcing a complete policy reset will remove all locally installed deployments, purge existing policies, then download/install assigned deployments.")
                ;    $bRemoveWarn=False
                ;EndIf
				_GUICtrlStatusBar_SetText($idStatus,"Removing "&$aHostCollections[$iSelIdx][0]&" ("&$aHostCollections[$iSelIdx][9]&")...")
				_CollectionRemoveResource(StringUpper($sHost),$sHostResourceId,$aHostCollections[$iSelIdx][9])
                If @error Then
                    If @extended==-2147352567 Then MsgBox(48,"Remove - "&$aHostCollections[$iSelIdx][0],"Host is not a member of this deployment.")
                EndIf
				_GUICtrlStatusBar_SetText($idStatus,"Removing "&$aHostCollections[$iSelIdx][0]&" ("&$aHostCollections[$iSelIdx][9]&")...Done")
                _GUICtrlStatusBar_SetText($idStatus,"Waiting 5 sec...")
                Sleep(5000)
                _HostRefreshPolicy($sHost,$idStatus)
			EndIf
			Sleep($iStatDelay)
            If Not $bEnMulti Then $sSelMagic=$aHostCollections[$iSelIdx][15]
            RefreshHost()
            If Not $bEnMulti Then
                For $i=0 To UBound($aHostCollections,1)-1
                    If $sSelMagic==$aHostCollections[$i][15] Then
                        $iSelIdx=$i
                        ExitLoop
                    EndIf
                Next
            EndIf
			GuiSetStates(2)
			GUICtrlSetState($idHostLV, $GUI_FOCUS)
            _GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
            _GUICtrlListView_EnsureVisible($hHostLV,$iSelIdx)
			_GUICtrlStatusBar_SetText($idStatus,"Ready")
;~ 		Case $LVN_ITEMACTIVATE
;~ 			_Log
        Case $idHostActCtx_ExportColl
            Local $iCount=_GUICtrlListView_GetItemCount ($hHostLV)
            Local $sStr=""
            For $i=0 To $iCount-1
                $aItem = _GUICtrlListView_GetItemTextArray($hHostLV, $i)
                $sStr &= StringFormat("[%-10s][%-13s]: %s v%s", $aItem[1], $aItem[2], $aItem[5], $aItem[4]) & @CRLF
            Next
            ClipPut($sStr)
	EndSwitch
	If $bWatch Then
		If TimerDiff($tWatch)<=1000 Or UBound($aWatch,1)==0 Then ContinueLoop
        $tTimerRet=TimerInit()
		$tTimer=TimerInit()
		$cItems=$oWMI.ExecQuery("SELECT * FROM CCM_Application", "WQL", $wbemFlagReturnImmediately)
		_Log("oWmi:"&TimerDiff($tTimer),"aWatch")
		For $oItem In $cItems
			$bHasId=True
			$iIdx=null
			For $i = 0 To UBound($aWatch,1)-1
				If $oItem.Id==$aWatch[$i][0] Then
					$iIdx=$i
					ExitLoop
				EndIf
				$bHasId=False
			Next
			If Not $bHasId Then ContinueLoop
            _Log($oItem.EvaluationState)
            _Log($oItem.PercentComplete)
			$sActs=objToArray($oItem.InProgressActions,True)
			If $iIdx>=UBound($aWatch,1) Then
				$bWatch=False
				GuiSetStates(3)
                RefreshHost()
                If Not $bEnMulti Then
                    For $i=0 To UBound($aHostCollections,1)-1
                        If $sSelMagic==$aHostCollections[$i][15] Then
                            $iSelIdx=$i
                            ExitLoop
                        EndIf
                    Next
                EndIf
				_GUICtrlStatusBar_SetText($idStatus,"Ready")
				GuiSetStates(2)
                GUICtrlSetState($idHostLV, $GUI_FOCUS)
                _EvalItemActs($iSelIdx)
                _GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
                _GUICtrlListView_EnsureVisible($hHostLV,$iSelIdx)
				ExitLoop
			EndIf
			If Not StringInStr($sActs,$aWatch[$i][1]) and $aWatch[$i][3]==True Then
				$aWatch[$i][3]=False
                _Log("ProgressWatch,"&$aWatch[$i][2]&","&$aWatch[$i][1]&" finished")
                _Toast_Hide()
				_Toast_Show(0, $sAlias,$aWatch[$i][1]&" for "&$aWatch[$i][2]&" has completed.",-20,False)
				Sleep($iStatDelay)
			EndIf
			Local $bFinished=True
			For $j=0 to $aWatch[0][0]
				If $aWatch[$i][3]==True Then
					$bFinished=False
					ExitLoop
				EndIf
			Next
			If $bFinished Then
				Sleep($iStatDelay)
				$bWatch=False
				GuiSetStates(3)
                RefreshHost()
                If Not $bEnMulti Then
                    For $i=0 To UBound($aHostCollections,1)-1
                        If $sSelMagic==$aHostCollections[$i][15] Then
                            $iSelIdx=$i
                            ExitLoop
                        EndIf
                    Next
                EndIf
				_GUICtrlStatusBar_SetText($idStatus,"Ready")
				GuiSetStates(2)
                GUICtrlSetState($idHostLV, $GUI_FOCUS)
                _EvalItemActs($iSelIdx)
                _GUICtrlListView_SetItemSelected($hHostLV,$iSelIdx,True,True)
                _GUICtrlListView_EnsureVisible($hHostLV,$iSelIdx)
				ExitLoop
			EndIf
		Next
		_Log("vRet:"&TimerDiff($tTimerRet),"aWatch")
		$tWatch=TimerInit()
	EndIf
WEnd

Func _ConfirmActMulti($sTitle,$sMsg)
    Return True
EndFunc

Func _Exit()
	Exit
EndFunc

Func _HostRefreshPolicy($sHost,$idStatus,$sScheds="121|021|022|002")
    $sgErrorFunc="_HostRefreshPolicy"
    Local $aSched=StringSplit($sScheds,'|')
    Local $oNS=_EnsureWMI($sHost,"/root/ccm")
    Local $oClass = $oNS.Get("SMS_Client")
    For $i=1 To $aSched[0]
        _GUICtrlStatusBar_SetText($idStatus,"Trigger: "&$aSchedDesc[$i-1]&"...")
        $vRet=_HostTriggerSchedule($sHost,$aSched[$i-1],$oClass,$oNS)
        If @Error Then
            _GUICtrlStatusBar_SetText($idStatus,"Trigger: "&$aSchedDesc[$i-1]&"...Failed (0x"&Hex(@extended)&")")
            Sleep(1000)
            ContinueLoop
        EndIf
        _GUICtrlStatusBar_SetText($idStatus,"Trigger: "&$aSchedDesc[$i-1]&"...Done")
        Sleep($iStatDelay)
    Next
    Sleep(1000)
    _GUICtrlStatusBar_SetText($idStatus,"Waiting 10 sec...")
    Sleep(10000)
    $sgErrorFunc=""
EndFunc

Func _InitHost()
	; Sanity checks before continuing.
	$sHost=GUICtrlRead($idHost)
	$sOpid=GUICtrlRead($idOpid)
	$bHasHost=$sHost<>""
	$bHasOpid=$sOpid<>""
	$bHasPass=GUICtrlRead($idPass)<>""
	If Not $bHasHost Then
		If not $bHasOpid and not $bHasPass Then
			MsgBox(16,"Error","You must enter a Host, OPID, and Password to continue.",0,$hWnd)
		Else
			MsgBox(16,"Error","You must enter a Host to continue.",0,$hWnd)
		EndIf
		Return
	Else
		If not $bHasOpid and not $bHasPass Then
			MsgBox(16,"Error","You must enter an OPID, and Password to continue.",0,$hWnd)
			Return
		ElseIf $bHasOpid and not $bHasPass Then
			MsgBox(16,"Error","You must enter a Password to continue.",0,$hWnd)
			Return
		ElseIf $bHasPass and not $bHasOpid Then
			MsgBox(16,"Error","You must enter an OPID to continue.",0,$hWnd)
			Return
		EndIf
	EndIf
	GuiSetStates(0)
	; Attempt Authentication if we do not have a token.
	_Log("InitHost: '"&$sHost&"','"&$sOpid&"','"&$bHasPass)
	if not $bHasToken Then
		_Log("InitHost,noToken")
		$sUser=GUICtrlRead($idOpid)
		$sPass=GUICtrlRead($idPass)
		_GUICtrlStatusBar_SetText($idStatus,"Validating Credentials...")
		$iRet=_AD_Open("<redacted>\"&$sUser, $sPass,"","",1)
		If $iRet Then
			GuiCtrlSetData($idPass,$g_aAuthSalts[@HOUR])
			_Log("InitHost,noToken,Valid")
			_GUICtrlStatusBar_SetText($idStatus,"Validating Credentials...Success")
			Sleep($iStatDelay)
			_AD_Close()
			_GUICtrlStatusBar_SetText($idStatus,"Encrypting...")
			$g_aAuth[0][1]=_Base64Encode(_CryptProtectData($sPass))
			$sPass=""
			$g_aAuth[0][0]=$sUser
			$sUser=""
			_GUICtrlStatusBar_SetText($idStatus,"Encrypting...Done")
			_Log("InitHost,noToken,Valid,Encrypted")
			Sleep($iStatDelay)
			if $bRemember Then
				_GUICtrlStatusBar_SetText($idStatus,"Saving Token...")
				IniWrite($sAuthIni,"ohAuth","opid",$g_aAuth[0][0])
				IniWrite($sAuthIni,"ohAuth","token",$g_aAuth[0][1])
				_GUICtrlStatusBar_SetText($idStatus,"Saving Token...Done")
				_Log("InitHost,noToken,Valid,Encrypted,Saved")
				Sleep($iStatDelay)
			EndIf
		Else
			If @Error==8 Then
				_GUICtrlStatusBar_SetText($idStatus,"Validating Credentials...Authentication Failed")
				_Log("InitHost,noToken,AuthFail")
				Sleep($iStatDelay)
			Else
				_GUICtrlStatusBar_SetText($idStatus,"Validating Credentials...Internal Failure")
				_Log("InitHost,noToken,InternalFail")
				Sleep($iStatDelay)
			EndIf
			GuiSetStates(1)
			Return
		EndIf
	EndIf
	_GUICtrlStatusBar_SetText($idStatus,"Checking host: Resolve...")
    Sleep(125)
    $sHost=StringLower($sHost)
    $bIsIP=StringRegExp($sHost,"^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$")
    $bIsHostname=StringRegExp($sHost,"^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$")
    If Not $bIsIP And Not $bIsHostname Then
        MsgBox(16,"Error", "Please enter a valid Hostname or IP.")
        GuiSetStates(1)
        _GUICtrlStatusBar_SetText($idStatus,"Ready")
        Return False
    EndIf
    $aRes=_Resolve($sHost,10000)
	_Log("InitHost,Resolve: "&_ArrayToString($aRes))
	if $aRes[0]<>"" and  $aRes[1]<>"" Then $bHostResolve=True
    $iRet=Null
    If Not IsArray($aRes) Then
        $iRet=MsgBox(16,"Error", "There was an error resolving the host."&@LF&"Failed to get Hostname or IP from host specified."&@LF&@LF&"Would you like to continue anyway? (This may make program unresponsive)")
    Else
        $iErr=0
        If $aRes[0]==False Then $iErr+=10
        If $aRes[1]==False Then $iErr+=1
        If $iErr==1 Then
            $iRet=MsgBox(49,"Error", "There was an error resolving the host."&@LF&"Hostname lookup Failed."&@LF&@LF&"Would you like to continue anyway? (This may make program unresponsive)")
        ElseIf $iErr==10 Then
            $iRet=MsgBox(49,"Error", "There was an error resolving the host."&@LF&"IP lookup Failed."&@LF&@LF&"Would you like to continue anyway? (This may make program unresponsive)")
        ElseIf $iErr==11 Then
            $iRet=MsgBox(49,"Error", "There was an error resolving the host."&@LF&"Failed to get Hostname and IP."&@LF&@LF&"Would you like to continue anyway? (This may make program unresponsive)")
        EndIf
    EndIf
    If $iRet==2 Then
        GuiSetStates(1)
        _GUICtrlStatusBar_SetText($idStatus,"Ready")
        Return False
    EndIf
	_GUICtrlStatusBar_SetText($idStatus,"Checking host: Resolve, Ping...")
    Sleep($iStatDelay)
	$iRet=_Ping($sHost)
	if @Error Then
        $iRet=MsgBox(49,"Error", "The host does not respond to ping."&@LF&"Would you like to continue anyway? (This may make program unresponsive)")
        If $iRet==2 Then
            GuiSetStates(1)
            _GUICtrlStatusBar_SetText($idStatus,"Ready")
            Return False
        EndIf
    EndIf
	_GUICtrlStatusBar_SetText($idStatus,"Checking host: Resolve, Ping, SMS...")
    Sleep($iStatDelay)
	; Get Host's ResourceId
    _EnsureSMS(True)
	$oResults=$g_oSMS.ExecQuery("SELECT * FROM SMS_R_System WHERE Name = '"&$sHost&"'")
	_Log("InitHost,SMS")
    For $oI In $oResults
        ConsoleWrite($oI.Name&@CRLF)
        ConsoleWrite($oI.ResourceID&@CRLF)
    Next
	If $oResults.Count==0 Or Not IsObj($oResults) Then
		MsgBox(48,"Error - Initialize Host","Unable to retrieve Hosts's ResourceId from SCCM."&@LF&"Adding/Removing device to/from collections will be disabled.")
		$bEnColMgmt=False
	ElseIf $oResults.Count==1 Then
		For $oObj In $oResults
			$sHostResourceId=$oObj.ResourceID
			_Log("InitHost,ResourceID:"&$sHostResourceId)
		Next
		If $sHostResourceId="" Then
			MsgBox(48,"Error - Initialize Host","Unable to retrieve Hosts's ResourceId from SCCM."&@LF&"Adding/Removing device to/from collections will be disabled.")
			$bEnColMgmt=False
		EndIf
	Else
		MsgBox(48,"Error - Initialize Host","More than one host return for this hostname."&@LF&"Adding/Removing device to/from collections will be disabled.")
		$bEnColMgmt=False
	EndIf
	_GUICtrlStatusBar_SetText($idStatus,"Checking host: Resolve, Ping, SMS, WMI...")
    Sleep($iStatDelay)
	; Initialize WMI
	_Log("InitHost,EnsureWMI")
	$oWMI=_EnsureWMI($sHost,"/root/ccm/ClientSDK",True)
    ;CCM_InstanceEvent
    _EnsureWMI($sHost,"/root/ccm",True)
	_GUICtrlStatusBar_SetText($idStatus,"Checking host: Done")
    Sleep($iStatDelay)
	_Log("InitHost,Done")
	Sleep($iStatDelay)
	RefreshHost(True)
    ;$oCCM=_EnsureWMI($sHost,"/root/ccm/ClientSDK",True)
    ;$ogCcmEvent = ObjEvent($oCCM, "_CCM_Event","CCM_Event")
    ;AdlibRegister("_AppWatch2",250)
    ;$bWatch2=True
;~     $iTimerA=TimerInit()
;~     _Log("DGC-Test-A:"&TimerDiff($iTimerA),"_InitHost")
;~     _DevGetCollections($sHost,True)
;~     _Log("DGC-Test-A:"&TimerDiff($iTimerA),"_InitHost")
;~     $iTimerB=TimerInit()
;~     _Log("DGC-Test-B:"&TimerDiff($iTimerB),"_InitHost")
;~     _DevGetCollections($sHost,False)
;~     _Log("DGC-Test-B:"&TimerDiff($iTimerB),"_InitHost")
	GuiSetStates(2)
	GUIRegisterMsg($WM_NOTIFY, "WM_NOTIFY")
	$bHostActive=True
	_GUICtrlStatusBar_SetText($idStatus,"Ready")
	_Log($sHostResourceId&@CRLF)
EndFunc

;
; Watch for Status Changes and Update list.
;
Func _AppWatch2()
    If $bWatch2 And TimerDiff($tWatch2)>=1000 Then
        $bWatch2Proc=True
        ;$tWatch2=TimerInit()
        $tTimerRet=TimerInit()
        $bgNoLog=True
        $aNew=_DevGetCollections($sHost,False)
        If @error Then
            $bWatch2=False
            ;$tWatch2=TimerInit()
            AdlibUnRegister("_AppWatch2")
            $bgNoLog=False
            Return
        EndIf
        $bgNoLog=False
        Local $sStatusLast, $sStatusNew
        For $i=0 To UBound($aNew,1)-1
            $iIndex=-1
            For $j=0 To UBound($aHostCollections,1)-1
                If $aHostCollections[$j][8]==$aNew[$i][8] Then
                    $iIndex=$j
                    ExitLoop
                EndIf
            Next
            If $iIndex==-1 Then
                _Log("isNew:"&$iIndex,"aWatch2")
                ContinueLoop; New Entry Added
            EndIf
            ;For $j=0 To UBound($aHostCollections,2)-1 ; Existing Entry, check for differences.
            If $aHostCollections[$iIndex][3]<>$aNew[$i][3] Then
               _Log($aHostCollections[$iIndex][1]&": "&$aHostCollections[$iIndex][3]&" > "&$aNew[$i][3],"aWatch2")
            EndIf
            If $aHostCollections[$iIndex][6]<>$aNew[$i][6] Or $aHostCollections[$iIndex][11]<>$aNew[$i][11] Then
                $sStatusLast=_ResolveStatus($aHostCollections[$iIndex][3],$aHostCollections[$iIndex][6],$aHostCollections[$iIndex][11])
                $sStatusNew=_ResolveStatus($aHostCollections[$iIndex][3],$aHostCollections[$iIndex][6],$aHostCollections[$iIndex][11])
               If $sStatusLast<>$sStatusNew Then _Log($aHostCollections[$iIndex][1]&": "&$sStatusLast&" > "&$sStatusNew,"aWatch2")
            EndIf
            ;11, 6, 3
            ;Next
        Next
        ;$tWatch2=TimerInit()
        _Log("vRet:"&TimerDiff($tTimerRet),"Watch2")
    EndIf
    $bWatch2Proc=False
    ;$aHostCollections[$i][13]
EndFunc

Func DisposeCtx()
	GUICtrlDelete($idCtxDesc)
	GUICtrlDelete($idCtxIn)
	GUICtrlDelete($idCtxRm)
	GUICtrlDelete($idCtxUn)
	GUICtrlDelete($idCtx)
EndFunc

Func SpawnCtx()
	; Context Menu
	DisposeCtx()
	$idCtx = GUICtrlCreateContextMenu($idHostLV)
	$idCtxIn = GUICtrlCreateMenuItem("Install", $idCtx)
	GUICtrlSetState($idCtxIn,$bCanInstall ? $GUI_ENABLE : $GUI_DISABLE)
	$idCtxUn = GUICtrlCreateMenuItem("Uninstall", $idCtx)
	GUICtrlSetState($idCtxUn,$bCanUninstall ? $GUI_ENABLE : $GUI_DISABLE)
	GUICtrlCreateMenuItem("", $idCtx)
	$idCtxRm = GUICtrlCreateMenuItem("Remove", $idCtx)
	GUICtrlSetState($idCtxRm,$bCanRemove ? $GUI_ENABLE : $GUI_DISABLE)
	GUICtrlCreateMenuItem("", $idCtx)
	$idCtxDesc = GUICtrlCreateMenuItem("Description", $idCtx)
	GUICtrlSetState($idCtxDesc, $aHostCollections[$iSelIdx][1]<>"" ? $GUI_ENABLE : $GUI_DISABLE)
EndFunc

Func GetHostColl($sColl)
    Local $ivDim=UBound($aHostCollections,2)
    Local $aRet[$ivDim]
    For $i=0 To UBound($aHostCollections,1)-1
        If $aHostCollections[$i][9]==$sColl Then
            For $j=0 To $ivDim-1
                $aRet[$j]=$aHostCollections[$i][$j]
            Next
            Return SetError(0,0,$aRet)
        EndIf
    Next
    Return SetError(1,0,False)
EndFunc

Func WM_NOTIFY($hWnd, $iMsg, $iwParam, $ilParam)
    #forceref $hWnd, $iMsg, $iwParam
    Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR
	$tNMHDR = DllStructCreate($tagNMHDR, $ilParam)
    $hWndFrom = DllStructGetData($tNMHDR, "hWndFrom")
    $iIDFrom = DllStructGetData($tNMHDR, "IDFrom")
    $iCode = DllStructGetData($tNMHDR, "Code")
    Local $sCid, $sDesc, $bCheck, $bModHas, $iIdx, $iMax, $idItem, $tInfo, $iIndex
    Switch $hWndFrom
        Case $hHostLV
            $iIndex=-1
            ;
            ; On any click event, get item index.
            ;
            If $iCode==$NM_CLICK Or $iCode==$NM_DBLCLK Or $iCode==$NM_RCLICK Or $iCode==$NM_RDBLCLK Then
                $tInfo = DllStructCreate($tagNMITEMACTIVATE, $ilParam)
                $iIndex = DllStructGetData($tInfo, "Index")
            EndIf
            If $iIndex==-1 Then Return $GUI_RUNDEFMSG
            ;
            ; On left click, dynamically consider available actions on an item.
            ;
            if $iCode == $NM_CLICK Then
				_Log('>SngClk'&$iIndex&'|'&_EscapeString(_ArrayToString($aHostCollections,'|',$iIndex,$iIndex))&@CRLF)
				;[sName|sDesc|iRevision|sState|sDate|saAllowedActions|saInProgressActions|bIsMachineTarget|sScopeId|sCollectionId
                _EvalItemActs($iIndex)
;~ 				$bCanInstall
;~ 				$bCanUninstall
;~ 				$bCanRemove
				$iSelIdx=$iIndex
                $sSelMagic=$aHostCollections[$iSelIdx][15]
            ;
            ; Toggle item when double clicking it and MutiSelect is enabled.
            ;
			Elseif $iCode == $NM_DBLCLK Then
				If $bEnMulti==True Then
                    _Log('>DblClk'&$iIndex&'|'&_EscapeString(_ArrayToString($aHostCollections,'|',$iIndex,$iIndex))&@CRLF)
					$iState=_GUICtrlListView_GetItemChecked($hHostLV,$iIndex)
					_GUICtrlListView_SetItemChecked($hHostLV,$iIndex,$iState ? False : True)
				Else
;~ 					_Log($bEnMulti&@CRLF)
;~ 					SpawnCtx()
;~ 					DllCall("user32.dll", "ptr", "SendMessage", "hwnd", $hWnd, "int", $WM_CONTEXTMENU, "int", $hHostLV, "int", 0)
				EndIf
				$iSelIdx=$iIndex
            ;
            ; Show a context menu when MultiSelect is disabled and when right clicking an item.
            ;
			ElseIf $iCode == $NM_RCLICK Then
				$iIdx=_GUICtrlListView_GetSelectedIndices($hHostLV)
				If $iIdx=="" Then Return $GUI_RUNDEFMSG
				$iSelIdx=$iIdx
				If $bEnMulti==False Then SpawnCtx()
            EndIf
            ;
            ; On any type of click on a checkbox, or when double clicking an item (toggling the checkbox)
            ;  Add or remove it from a queue of items to act upon.
            ;
            If $iCode==$NM_CLICK Or $iCode==$NM_DBLCLK Or $iCode==$NM_RCLICK Or $iCode==$NM_RDBLCLK Then
                Local $iX = DllStructGetData($tInfo, "X")
                Local $aIconRect = _GUICtrlListView_GetItemRect($hHostLV, $iIndex, 1)
                If $iX < $aIconRect[0] And $iX >= 5 Or ($iCode==$NM_DBLCLK And $bEnMulti) Then
                    Local $sCid=_GUICtrlListView_GetItemText($hHostLV, $iIndex)
                    If $sCid=="" Then
                        _GUICtrlListView_SetItemChecked($hHostLV, $iIndex,True)
                        Return 0
                    EndIf
                    If $iCode==$NM_DBLCLK And $bEnMulti Then
                        $bCheck=_GUICtrlListView_GetItemChecked($hHostLV, $iIndex)==True
                    Else
                        $bCheck=_GUICtrlListView_GetItemChecked($hHostLV, $iIndex)==False
                    EndIf
                    $aCol=GetHostColl($sCid)
                    If Not @error Then
                        $sDesc=$aCol[13]
                    EndIf
                    ;$idItem=_GUICtrlListView_MapIndexToID($hHostLV, $iIndex); For color
                    $bModHas=False
                    $iIdx=-1
                    $iMax=UBound($aModQueue,1)
                    For $i=0 To $iMax-1
                        If $aModQueue[$i][0]<>$sCid Then ContinueLoop
                        $bModHas=True
                        $iIdx=$i
                        ExitLoop
                    Next
                    If $bModHas Then
                        If $bCheck Then
                            $aModQueue[$iIdx][2]=True
                        Else
                            ; Insanity Check
                            Local $bDup=False, $bDupChk,$iCount=_GUICtrlListView_GetItemCount($hHostLV)
                            For $k=0 To $iCount
                                If StringCompare(_GUICtrlListView_GetItemText($hHostLV,$k),$sCid)==0 Then
                                    If $k==$iIndex Then
                                        If $iCode==$NM_DBLCLK And $bEnMulti Then
                                            $bDupChk=_GUICtrlListView_GetItemChecked($hHostLV, $k)==True
                                        Else
                                            $bDupChk=_GUICtrlListView_GetItemChecked($hHostLV, $k)==False
                                        EndIf
                                    Else
                                        $bDupChk=_GUICtrlListView_GetItemChecked($hHostLV, $k)==True
                                    EndIf
                                    If $bDupChk Then
                                        $bDup=True
                                        ExitLoop
                                    EndIf
                                EndIf
                            Next
                            If Not $bDup Then $aModQueue[$iIdx][2]=False
                        EndIf
                        _Log($sDesc&' ('&$sCid&'): '&$aModQueue[$iIdx][2],"WM_NOTIFY")
                        ;If $bCheck Then
                        ;    GUICtrlSetBkColor($idItem,0xAAAAEE)
                        ;Else
                        ;    GUICtrlSetBkColor($idItem,0xEEEEAA)
                        ;EndIf
                    Else
                        ReDim $aModQueue[$iMax+1][3]
                        $aModQueue[$iMax][0]=$sCid
                        $aModQueue[$iMax][1]=$sDesc
                        $aModQueue[$iMax][2]=True
                    EndIf
                    Return 0
                EndIf
            EndIf
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc   ;==>WM_NOTIFY

;
; Evaulates Available Actions for an Item
;
Func _EvalItemActs($iIndex)
    $sState=$aHostCollections[$iIndex][3]
    $sAvailActs=$aHostCollections[$iIndex][5]
    $sActsInProg=$aHostCollections[$iIndex][6]
    $bCanRemove=True
    $bCanInstall=True
    $bCanUninstall=True
    if $sState=="Installed" Then
        $bCanInstall=False
        if StringInStr($sActsInProg,"Uninstall") Or StringInStr($sActsInProg,"Install") Then $bCanUninstall=False
    ElseIf $sState=="Installing" Then
        $bCanInstall=False
        $bCanUninstall=False
    ElseIf $sState=="Unstalling" Then
        $bCanInstall=False
        $bCanUninstall=False
    ElseIf $sState=="NotInstalled" Then
        if StringInStr($sActsInProg,"Install") Then $bCanInstall=False
        $bCanUninstall=False
        EndIf
    ; No Collection ID, Host is not member.
    If $aHostCollections[$iIndex][9]=="" Then
        $bCanRemove=False
        ;$bCanInstall=False
    EndIf
    GUICtrlSetState($idInstall,$bCanInstall ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($idUninstall,$bCanUninstall ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($idRemove,$bCanRemove ? $GUI_ENABLE : $GUI_DISABLE)
EndFunc

Func _EscapeString($sStr)
	$sStr=StringReplace($sStr,"\","\\")
	$sStr=StringReplace($sStr,@CR,"\r")
	$sStr=StringReplace($sStr,@LF,"\n")
	$sStr=StringReplace($sStr,@TAB,"\t")
	Return $sStr
EndFunc

Func RefreshHost($bFirstCall=False)
    _Log("GetHostCollections")
	; Get Host Collections
	If $bFirstCall Then
		_GUICtrlStatusBar_SetText($idStatus,"Getting host collections...")
	Else
		_GUICtrlStatusBar_SetText($idStatus,"Reloading host collections...")
	EndIf
    $bWatch2=False
    AdlibUnregister("_AppWatch2")
    While $bWatch2Proc
        If TimerDiff($tWatch2)>=2000 Then
            $bWatch2Proc=False
            ExitLoop
        EndIf
        Sleep(1)
    WEnd
	$aHostCollections=_DevGetCollections($sHost,True)
    If @error Then
        If $bFirstCall Then
            _GUICtrlStatusBar_SetText($idStatus,"Getting host collections...Failed")
        Else
            _GUICtrlStatusBar_SetText($idStatus,"Reloading host collections...Failed")
        EndIf
        Return
    EndIf
    If Not @Compiled Then _ArrayDisplay($aHostCollections)
	_GUICtrlListView_BeginUpdate($hHostLV)
	_GUICtrlListView_DeleteAllItems($hHostLV)
	For $i=0 to UBound($aHostCollections,1)-1
		$sItem=$aHostCollections[$i][9]
		$sState=$aHostCollections[$i][3]
		$sActProg=$aHostCollections[$i][6]
		$sName=$aHostCollections[$i][0]
		$sVer=$aHostCollections[$i][10]
		$sEvalState=$aHostCollections[$i][11]
		$sProg=$aHostCollections[$i][12]
		$sDate=StringStripWS(StringTrimRight($aHostCollections[$i][4],9),7)
		if StringInStr($sName,$sVer) Then $sName=StringReplace($sName,$sVer,"")
		$sName=StringStripWS($sName,7)
        $sItem&='|'&_ResolveStatus($sState,$sActProg,$sEvalState)
		$sItem&='|'&$sDate
		$sItem&='|'&$sVer
		$sItem&='|'&$sName
		$aHostCollections[$i][14]=GUICtrlCreateListViewItem($sItem,$idHostLV)
		If Mod($i, 2)==0 Then GUICtrlSetBkColor($aHostCollections[$i][14],0xEEEEEE)
        If $aHostCollections[$i][9]=="" Then GUICtrlSetBkColor($aHostCollections[$i][14],0xEEAAAA)
		GUICtrlSetTip($aHostCollections[$i][14],$aHostCollections[$i][1],$aHostCollections[$i][0])
	Next
	_GUICtrlListView_EndUpdate($hHostLV)
    ;$tWatch2=TimerInit()
    ;AdlibRegister("_AppWatch2",250)
    ;$bWatch2=True
	If $bFirstCall Then
		_GUICtrlStatusBar_SetText($idStatus,"Getting host collections...Done")
	Else
		_GUICtrlStatusBar_SetText($idStatus,"Reloading host collections...Done")
	EndIf
	Sleep($iStatDelay)
	ReDim $aSel[UBound($aHostCollections,1)]
	;_ArrayDisplay($aHostCollections)
EndFunc

Func _ResolveStatus($sState,$sActProg,$sEvalState)
        Local $sRet
        Local $bStates[]=[StringInStr($sActProg,"Install"),StringInStr($sActProg,"Update"),StringInStr($sActProg,"Uninstall")]
		Switch $sState
			Case "NotInstalled"
				If $bStates[0] Then
                    If $sEvalState==6 Or $sEvalState==7 Or $sEvalState==8 Or $sEvalState==23 Or $sEvalState==24 Then
                        $sRet="Downloading"
                    Else
                        $sRet="Installing"
                    EndIf
                Else
                    If $bStates[1] Then
                        $sRet="Updating"
                    Else
                        $sRet="Not Installed"
                    EndIf
                EndIf
			Case "Installed"
				If $bStates[1] Then
					$sRet="Updating"
                Else
                    If $bStates[2] Then
                        $sRet="Uninstalling"
                    Else
                        $sRet="Installed"
                    EndIf
				EndIf
			Case Else
				$sRet=$sState
        EndSwitch
        If $sEvalState==2 Then
                $sRet&=", Enforced/Resolved"
;~         ElseIf $sEvalState==3 Then
;~                 $sItem&=", Not Required"
;~         ElseIf $sEvalState==4 Then
;~                 $sItem&=", EnforceAvailable"
        ElseIf $sEvalState==5 Then
                $sRet&=", FailedEnforce"
        ElseIf $sEvalState==10 Then
                $sRet&=", AwaitSerialEnforcement"
        ElseIf $sEvalState==11 Then
                $sRet&=", AwaitSerialEnforcement"
        ElseIf $sEvalState==12 Then
                $sRet&=", EnforcingDeps"
        ElseIf $sEvalState==13 Then
                $sRet&=", Enforcing"
        ElseIf $sEvalState==14 Or $sEvalState==15  Or $sEvalState==10 Then
                $sRet&=", Pending Reboot"
        ElseIf $sEvalState==16 Then
                $sRet&=", Update Available"
        ElseIf $sEvalState==18 Then
                $sRet&=", AwaitUserSession"
        ElseIf $sEvalState==19 Then
                $sRet&=", AwaitAllUsersLogoff"
        ElseIf $sEvalState==20 Then
                $sRet&=", AwaitUserLogon"
        ElseIf $sEvalState==21 Then
                $sRet&=", AwaitRetry"
        EndIf
        Return $sRet
EndFunc

Func GuiSetStates($i)
	; [0] All Controls Disabled.
	; [1] PreInit, Host and Auth Enabled, All else disabled
	; [2] PostInit, Host and Auth Disabled, All Else disabled.
	Local $aStates[][]=[ _
		[0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0], _
		[1, 1, 1, 1, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0], _
		[0, 0, 0, 0, 2, 1, 0, 0, 0, 1, 1, 1, 1, 1], _
		[0, 0, 0, 0, 3, 0, 0, 0, 0, 0, 0, 0, 0, 0] _
	]
	GuiCtrlSetState($idHost,$aStates[$i][0] ? $GUI_ENABLE : $GUI_DISABLE)
	if $bHasToken Then
		GuiCtrlSetState($idOpid,$GUI_DISABLE)
		GuiCtrlSetState($idPass,$GUI_DISABLE)
		if $i=1 Then
			GuiCtrlSetState($idRemember,BitOr($GUI_ENABLE,$GUI_CHECKED))
		Else
			GuiCtrlSetState($idRemember,BitOr($GUI_DISABLE,$GUI_CHECKED))
		EndIf
	Else
		GuiCtrlSetState($idOpid,$aStates[$i][1] ? $GUI_ENABLE : $GUI_DISABLE)
		GuiCtrlSetState($idPass,$aStates[$i][2] ? $GUI_ENABLE : $GUI_DISABLE)
		GuiCtrlSetState($idRemember,$aStates[$i][3] ? $GUI_ENABLE : $GUI_DISABLE)
	EndIf

	$idiInit=$GUI_DISABLE
	$idiUninit=BitOr($GUI_HIDE,$GUI_DISABLE)
	If $aStates[$i][4]==1 Then
		$idiInit=BitOr($GUI_SHOW,$GUI_ENABLE)
	ElseIf $aStates[$i][4]>1 Then
		$idiInit=BitOr($GUI_HIDE,$GUI_DISABLE)
		$idiUninit=BitOr($GUI_SHOW,$GUI_ENABLE)
	EndIf
	If $aStates[$i][4]>2 Then $idiUninit=BitOr($GUI_SHOW,$GUI_DISABLE)
	GuiCtrlSetState($idInit,$idiInit)
	GuiCtrlSetState($idUninit,$idiUninit)
	If Not $bEnColMgmt Then
		$aStates[$i][5]=0
		$aStates[$i][6]=0
	EndIf
	GuiCtrlSetState($idAdd,$aStates[$i][5] ? $GUI_ENABLE : $GUI_DISABLE)
	GuiCtrlSetState($idRemove,$aStates[$i][6] ? $GUI_ENABLE : $GUI_DISABLE)
	GuiCtrlSetState($idInstall,$aStates[$i][7] ? $GUI_ENABLE : $GUI_DISABLE)
	GuiCtrlSetState($idUninstall,$aStates[$i][8] ? $GUI_ENABLE : $GUI_DISABLE)
	GuiCtrlSetState($idRefresh,$aStates[$i][9] ? $GUI_ENABLE : $GUI_DISABLE)
	GuiCtrlSetState($idHostLV,$aStates[$i][10] ? $GUI_ENABLE : $GUI_DISABLE)
    $aStates[$i][11] = $bEnMulti ? 1 : 0
	GuiCtrlSetState($idenMulti,$aStates[$i][11] ? $GUI_ENABLE : $GUI_DISABLE)
    GuiCtrlSetState($idHostAct,$aStates[$i][12] ? $GUI_ENABLE : $GUI_DISABLE)
    ;GuiCtrlSetState($idSyncPol,$aStates[$i][13] ? $GUI_ENABLE : $GUI_DISABLE)
;~ 	$iHostStateLV=GUICtrlGetState($idHostLV)
;~ 	GUICtrlSetStyle($idHostLV,$bEnMulti ? BitOr($iHostStateLV,$LVS_SINGLESEL) : BitAND($iHostStateLV,$LVS_SINGLESEL))
	_GUICtrlListView_SetExtendedListViewStyle($hHostLV,BitOr($LVS_EX_TWOCLICKACTIVATE,$WS_EX_CLIENTEDGE,$LVS_EX_DOUBLEBUFFER,$LVS_EX_FULLROWSELECT,$LVS_EX_GRIDLINES,$LVS_EX_GRIDLINES,$bEnMulti ? $LVS_EX_CHECKBOXES : 0))
	GUICtrlSetBkColor($idHostLV,0xFFFFFF)
	;GUICtrlSetBkColor($idHostLV, $GUI_BKCOLOR_LV_ALTERNATE)
	_GUICtrlListView_SetColumnWidth($hHostLV,0,64+32)
	_GUICtrlListView_SetColumnWidth($hHostLV,1,64+32+8)
	_GUICtrlListView_SetColumnWidth($hHostLV,2,96+16)
	_GUICtrlListView_SetColumnWidth($hHostLV,3,64+32)
	;_GUICtrlListView_SetColumnWidth($hHostLV,0,$LVSCW_AUTOSIZE)
	;_GUICtrlListView_SetColumnWidth($hHostLV,1,$LVSCW_AUTOSIZE)
	;_GUICtrlListView_SetColumnWidth($hHostLV,2,$LVSCW_AUTOSIZE)
	;_GUICtrlListView_SetColumnWidth($hHostLV,3,$LVSCW_AUTOSIZE)
	_GUICtrlListView_SetColumnWidth($hHostLV,4,$LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_JustifyColumn($hHostLV,0,2)
	_GUICtrlListView_JustifyColumn($hHostLV,1,2)
	_GUICtrlListView_JustifyColumn($hHostLV,2,2)
	Sleep($iStatDelay)
EndFunc

Func _AddCollectionGui()
    _Log("AddCollectionGui")
    Dim $aAddQueue[0][3]
    $bAddMod=False
	#Region ### START Koda GUI section ### Form=
	$iGuiAddW= 512+128
	$iGuiAddH = 256+128+32
	$hAddGui = GUICreate("Add to Collections",$iGuiAddW,$iGuiAddH,-1,-1,-1,-1,$hWnd)
	GUISwitch($hAddGui)
	GUISetFont(10, 400, 0, "Consolas")
	$idAddSearch = GUICtrlCreateInput("", 8, 8,  $iGuiAddW-16, 20)
	$hAddSearch = GUICtrlGetHandle($idAddSearch)
	GUICtrlSetTip(-1,"Search for App's Name or Description here.")
	_GUICtrlEdit_SetCueBanner($hAddSearch, "Search", True)
	$idAddAppsLV = GUICtrlCreateListView("CID|Name", 8, 8+20+2, $iGuiAddW-16, $iGuiAddH-64-16-6)
	$hAddAppsLV=GUICtrlGetHandle($idAddAppsLV)
	_GUICtrlListView_SetExtendedListViewStyle($hAddAppsLV,BitOr($LVS_EX_TWOCLICKACTIVATE,$WS_EX_CLIENTEDGE,$LVS_EX_DOUBLEBUFFER,$LVS_EX_FULLROWSELECT,$LVS_EX_GRIDLINES,$LVS_EX_GRIDLINES,$LVS_EX_CHECKBOXES))
	_GUICtrlListView_SetColumnWidth($hAddAppsLV,0,80)
	_GUICtrlListView_SetColumnWidth($hAddAppsLV,1,$LVSCW_AUTOSIZE_USEHEADER)
	_GUICtrlListView_JustifyColumn($hAddAppsLV,0,2)
	GUICtrlSetBkColor($idAddAppsLV,0xFFFFFF)
	;GUICtrlSetBkColor($idAddAppsLV, $GUI_BKCOLOR_LV_ALTERNATE)
	;GUICtrlSetBkColor(-1,0xEEEEEE)
	$idAddSync = GUICtrlCreateCheckbox("Run Actions", 8, $iGuiAddH-8-20-16-6, 128+8, 20)
	$idAddWait = GUICtrlCreateCheckbox("Wait for App(s) to Sync",8+8+128+8, $iGuiAddH-8-20-16-6, 256, 20)
	$idAddRefresh = GUICtrlCreateButton("Refresh", $iGuiAddW-128-64-16-8, $iGuiAddH-8-22-16-6, 64, 25)
	$idAddAdd = GUICtrlCreateButton("Add", $iGuiAddW-128-16, $iGuiAddH-8-22-16-6, 64, 25)
	$idAddDone = GUICtrlCreateButton("Return", $iGuiAddW-64-8, $iGuiAddH-8-22-16-6, 64, 25)
	$idAddStatus = _GUICtrlStatusBar_Create($hAddGui)
	_GUICtrlStatusBar_SetText($idAddStatus,"Initializing")
	#EndRegion ### END Koda GUI section ###
	_GuiAddStates(0)
	GUISetState(@SW_SHOW,$hAddGui)
	;GUICtrlSetData($idAddSearch,"Refreshing Apps...")
	If UBound($aAppsSMS,1)==0 Then
		_GuiAddRefresh()
	Else
		_RefreshCollections()
	EndIf
    Sleep($iStatDelay)
	_GUICtrlStatusBar_SetText($idAddStatus,"Ready")
	$tSearch=TimerInit()
	GUIRegisterMsg($WM_COMMAND, "GuiAdd_WM_COMMAND");only used for EN_CHANGE so far
	GUIRegisterMsg($WM_NOTIFY, "GuiAdd_WM_NOTIFY")
	_GuiAddStates(1)
	;_DebugArrayDisplay($aHostCollections)
	While 1
		$nMsg = GUIGetMsg()
		Switch $nMsg
            Case $GUI_EVENT_CLOSE, $idAddDone
                _Log("AddCollectionGui,Close")
                GUIRegisterMsg($WM_COMMAND,"");only used for EN_CHANGE so far
                GUIRegisterMsg($WM_NOTIFY,"")
                GUISetState(@SW_HIDE,$hAddGui)
				GUIDelete($hAddGui)
				GUISwitch($hWnd)
				Return
			Case $idAddWait
                _Log("AddCollectionGui,AddWait")
				$bWaitApp=GuiCtrlRead($idAddWait)==$GUI_CHECKED
				_Log($bWaitApp&@CRLF)
				If $bWaitApp Then
					$bForceSyncLast=GuiCtrlRead($idAddSync)
					$bAddSync=True
					GUICtrlSetState($idAddSync, BitOR($GUI_CHECKED,$GUI_DISABLE))
				Else
					GUICtrlSetState($idAddSync, BitOR($bForceSyncLast,$GUI_ENABLE))
					$bAddSync=GuiCtrlRead($idAddSync)==$GUI_CHECKED
				EndIf
			Case $idAddSync
                _Log("AddCollectionGui,AddSync")
				$bAddSync=GuiCtrlRead($idAddSync)==$GUI_CHECKED
			Case $idAddRefresh
                _Log("AddCollectionGui,Refresh")
				_GuiAddStates(0)
				_GuiAddRefresh()
				_GUICtrlStatusBar_SetText($idAddStatus,"Ready")
				_GuiAddStates(1)
			Case $idAddAdd
                _Log("AddCollectionGui,Add")
                _GuiAddStates(0)
                ;_ArrayDisplay($aAddQueue)
				;Sleep($iStatDelay)
				;_GUICtrlStatusBar_SetText($idAddStatus,"Ready")
                ;_GuiAddStates(1)
                ;ContinueLoop
                Local $aColls[0][3],$iMax
				For $i=0 To UBound($aAddQueue,1)-1
                    If $aAddQueue[$i][2]==False Then ContinueLoop
					_GUICtrlStatusBar_SetText($idAddStatus,"Adding "&$sHost&" to "&$aAddQueue[$i][1]&" ("&$aAddQueue[$i][0]&")...")
                    Sleep($iStatDelay)
					If _hostHasCollection($aAddQueue[$i][0]) Then
						_Log("AddMembershipRule,"&$aAddQueue[$i][0]&",AlreadyExists")
                        _GUICtrlStatusBar_SetText($idAddStatus,"Adding "&$sHost&" to "&$aAddQueue[$i][1]&" ("&$aAddQueue[$i][0]&")...Failed (Already Exists)")
                        Sleep(1000)
						ContinueLoop
					EndIf
                    $iMax=UBound($aColls,1)
                    ReDim $aColls[$iMax+1][3]
                    $aColls[$iMax][0]=$aAddQueue[$i][0]
                    $aColls[$iMax][1]=_GetCollection($aAddQueue[$i][0])
                    _Log($aColls[$iMax][0])
                    _Log(IsObj($aColls[$iMax][1]))
					$iRet=_CollectionAddResource(StringUpper($sHost),$sHostResourceId,$aColls[$iMax][1],True,True)
					$iRet=True
                    $aColls[$iMax][2]=$iRet
                    If $iRet Then
                        $bAddMod=True
						_GUICtrlStatusBar_SetText($idAddStatus,"Adding "&$sHost&" to "&$aAddQueue[$i][1]&" ("&$aAddQueue[$i][0]&")...Done")
                    Else
						_GUICtrlStatusBar_SetText($idAddStatus,"Adding "&$sHost&" to "&$aAddQueue[$i][1]&" ("&$aAddQueue[$i][0]&")...Failed")
                    EndIf
                    Log("AddCollectionGui,Add:"&$sHost&","&$sHostResourceId&","&$aAddQueue[$i][0]&","&$i&","&$iRet)
					Sleep($iStatDelay)
				Next
                _GUICtrlStatusBar_SetText($idAddStatus,"Waiting 5 sec...")
                Sleep(5000)
                For $i=0 To UBound($aColls,1)-1
                    If $aColls[$iMax][2]==False Then ContinueLoop
                    _GUICtrlStatusBar_SetText($idAddStatus,"Refreshing "&$aAddQueue[$i][1]&" ("&$aAddQueue[$i][0]&")...")
                    Sleep($iStatDelay)
                    $iRet=_CollectionRefresh($aColls[$i][1])
                    If @Error Then
                        _GUICtrlStatusBar_SetText($idAddStatus,"Refreshing "&$aAddQueue[$i][1]&" ("&$aAddQueue[$i][0]&")...Failed")
                    Else
                        _GUICtrlStatusBar_SetText($idAddStatus,"Refreshing "&$aAddQueue[$i][1]&" ("&$aAddQueue[$i][0]&")...Done")
                    EndIf
                    Sleep($iStatDelay)
                Next
                If $bAddMod Then
                    _GUICtrlStatusBar_SetText($idAddStatus,"Waiting 10 sec...")
                    Sleep(10000)
                EndIf
				If $bAddSync Then
					If Not $bAddMod Then
                        _GUICtrlStatusBar_SetText($idAddStatus,"Waiting 5 sec...")
                        Sleep(5000)
                    EndIf
					_HostRefreshPolicy($sHost,$idAddStatus)
				EndIf
				Sleep($iStatDelay)
				_GUICtrlStatusBar_SetText($idAddStatus,"Ready")
                _GuiAddStates(1)
		EndSwitch
	WEnd
EndFunc

Func _hostHasCollection($sCid)
	;if $sCid=="OHP01B61" Then _ArrayDisplay($aHostCollections,$sCid==)
	For $i=0 to UBound($aHostCollections,1)-1
        If $bSearchAbort Then Return SetError(0,1,False)
		If StringCompare($sCid,$aHostCollections[$i][9])==0 Then Return SetError(0,0,True)
	Next
	Return SetError(0,0,False)
EndFunc

Func _GuiAddRefresh()
	_Log("AddCollectionGui,Refresh")
	_GUICtrlStatusBar_SetText($idAddStatus,"Refreshing Apps...")
	$aAppsSMS=_SMSGetCollections()
	_ArraySort($aAppsSMS,0,0,0,0)
	_RefreshCollections()
    Sleep($iStatDelay)
	_GUICtrlStatusBar_SetText($idAddStatus,"Refreshing Apps...Done")
EndFunc

Func _GuiAddStates($bEnable)
    GUICtrlSetState($idAddSearch,$bEnable ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($idAddAppsLV,$bEnable ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($idAddSync,$bEnable ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($idAddWait,$bEnable ? $GUI_ENABLE : $GUI_DISABLE)
	GUICtrlSetState($idAddWait,$GUI_DISABLE)
    GUICtrlSetState($idAddRefresh,$bEnable ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($idAddDone,$bEnable ? $GUI_ENABLE : $GUI_DISABLE)
    GUICtrlSetState($idAddAdd,$bEnable ? $GUI_ENABLE : $GUI_DISABLE)

;~     If $bEnable Then
;~         GUICtrlSetState($idAddSync,$bGuiAddDefSync ? $GUI_ENABLE : $GUI_DISABLE)
;~     EndIf
    If $bWaitApp Then
        $bForceSyncLast=GuiCtrlRead($idAddSync)
        $bAddSync=True
        GUICtrlSetState($idAddSync, BitOR($GUI_CHECKED,$GUI_DISABLE))
    Else
        If $bEnable Then
            GUICtrlSetState($idAddSync, BitOR($bForceSyncLast,$GUI_ENABLE))
        Else
            GUICtrlSetState($idAddSync, $bForceSyncLast)
        EndIf
        $bAddSync=GuiCtrlRead($idAddSync)==$GUI_CHECKED
    EndIf

    If _GUICtrlListView_GetItemCount($hAddAppsLV)==0 Then
		_GUICtrlListView_BeginUpdate($hAddAppsLV)
		_GUICtrlListView_SetExtendedListViewStyle($hAddAppsLV,BitOr($LVS_EX_TWOCLICKACTIVATE,$WS_EX_CLIENTEDGE,$LVS_EX_DOUBLEBUFFER,$LVS_EX_FULLROWSELECT,$LVS_EX_GRIDLINES,$LVS_EX_GRIDLINES,$LVS_EX_CHECKBOXES))
		_GUICtrlListView_SetColumnWidth($hAddAppsLV,0,80)
		_GUICtrlListView_SetColumnWidth($hAddAppsLV,1,$LVSCW_AUTOSIZE_USEHEADER)
		_GUICtrlListView_JustifyColumn($hAddAppsLV,0,2)
		_GUICtrlListView_EndUpdate($hAddAppsLV)
    EndIf
EndFunc

Func _RefreshCollections()
    Local $bQueueSel
    _Log("RefreshHostCollections")
    ;_ArrayDisplay($aAppsSMS)
    _GUICtrlListView_BeginUpdate($hAddAppsLV)
    _GUICtrlListView_DeleteAllItems($hAddAppsLV)
    For $i=0 To UBound($aAppsSMS,1)-1
        $idItem=GUICtrlCreateListViewItem($aAppsSMS[$i][1]&'|'&$aAppsSMS[$i][0],$idAddAppsLV)
        If Mod($i, 2)==0 Then GUICtrlSetBkColor($idItem,0xEEEEEE)
        $bQueueSel=False
        For $j=0 To UBound($aAddQueue,1)-1
            If $aAppsSMS[$i][1]==$aAddQueue[$j][0] Then
                $bQueueSel=$aAddQueue[$j][2]
                _Log($aAddQueue[$j][0])
                ExitLoop
            EndIf
        Next
		if _hostHasCollection($aAppsSMS[$i][1]) Or $bQueueSel Then
			_GUICtrlListView_SetItemChecked($hAddAppsLV,$i,True)
            If $bQueueSel Then
                GUICtrlSetBkColor($idItem,0xAAAAEE)
            Else
                GUICtrlSetBkColor($idItem,0xEEEEAA)
            EndIf
		Else
			_GUICtrlListView_SetItemChecked($hAddAppsLV,$i,False)
		EndIf
    Next
    _GUICtrlListView_SetColumnWidth($hAddAppsLV,0,80)
    _GUICtrlListView_SetColumnWidth($hAddAppsLV,1,$LVSCW_AUTOSIZE_USEHEADER)
    _GUICtrlListView_JustifyColumn($hAddAppsLV,0,2)
    _GUICtrlListView_EndUpdate($hAddAppsLV)
EndFunc

Func _AddAppsSearch($sQuery)
    _Log("SearchApps,"&$sQuery)
    Local $iMax=0
    Local $aResults[0][3]
    For $i=0 To UBound($aAppsSMS,1)-1
        If $bSearchAbort Then Return SetError(0,1,False)
        $bMatch=False
        For $j=0 To 2
            If $bSearchAbort Then Return SetError(0,1,False)
            If StringInStr($aAppsSMS[$i][$j],$sQuery,0) Then $bMatch=True
        Next
        If $bMatch Then
            ReDim $aResults[$iMax+1][3]
            $aResults[$iMax][0]=$i
            $aResults[$iMax][1]=$aAppsSMS[$i][1]
            $aResults[$iMax][2]=$aAppsSMS[$i][0]
            $iMax+=1
        EndIf
    Next
    Return SetError(0,0,$aResults)
EndFunc

Func GuiAdd_WM_NOTIFY($hWnd, $iMsg, $wParam, $lParam)
    Local $hWndFrom, $iIDFrom, $iCode, $tNMHDR
    $tNMHDR = DllStructCreate($tagNMHDR, $lParam)
    $hWndFrom = HWnd(DllStructGetData($tNMHDR, "hWndFrom"))
    $iIDFrom = DllStructGetData($tNMHDR, "IDFrom")
    $iCode = DllStructGetData($tNMHDR, "Code")
    Local $sCid, $sDesc, $bCheck, $bAddHas, $iIdx, $iMax, $idItem
    Switch $hWndFrom
        Case $hAddAppsLV
            Switch $iCode
                Case $NM_CLICK, $NM_DBLCLK, $NM_RCLICK, $NM_RDBLCLK
                    Local $tInfo = DllStructCreate($tagNMITEMACTIVATE, $lParam)
                    Local $iIndex = DllStructGetData($tInfo, "Index")
                    If $iIndex <> -1 Then
                        Local $iX = DllStructGetData($tInfo, "X")
                        Local $iPart = 1
                        If _GUICtrlListView_GetView($hAddAppsLV) = 1 Then $iPart = 2 ;for large icons view
                        Local $aIconRect = _GUICtrlListView_GetItemRect($hAddAppsLV, $iIndex, $iPart)
                        If $iX < $aIconRect[0] And $iX >= 5 Then
                            Local $sCid=_GUICtrlListView_GetItemText($hAddAppsLV, $iIndex)
                            If _hostHasCollection($sCid) Then
                                _GUICtrlListView_SetItemChecked($hAddAppsLV, $iIndex,False)
                                Return 0
                            EndIf
                            $sDesc=_GUICtrlListView_GetItemText($hAddAppsLV, $iIndex,1)
                            $bCheck=_GUICtrlListView_GetItemChecked($hAddAppsLV, $iIndex)==False
                            $idItem=_GUICtrlListView_MapIndexToID($hAddAppsLV, $iIndex)
                            _Log($sDesc&' ('&$sCid&'): '&$bCheck,"GuiAdd_WM_NOTIFY")
                            $bAddHas=False
                            $iIdx=-1
                            $iMax=UBound($aAddQueue,1)
                            For $i=0 To $iMax-1
                                If $aAddQueue[$i][0]<>$sCid Then ContinueLoop
                                $bAddHas=True
                                $iIdx=$i
                                ExitLoop
                            Next
                            If $bAddHas Then
                                $aAddQueue[$iIdx][2]=$bCheck
                                If $bCheck Then
                                    GUICtrlSetBkColor($idItem,0xAAAAEE)
                                Else
                                    GUICtrlSetBkColor($idItem,0xEEEEAA)
                                EndIf
                            Else
                                ReDim $aAddQueue[$iMax+1][3]
                                $aAddQueue[$iMax][0]=$sCid
                                $aAddQueue[$iMax][1]=$sDesc
                                $aAddQueue[$iMax][2]=$bCheck
                            EndIf
                            Return 0
                        EndIf
                    EndIf
            EndSwitch
    EndSwitch
    Return $GUI_RUNDEFMSG
EndFunc

Func GuiAdd_WM_COMMAND($hWnd, $imsg, $iwParam, $ilParam)
    If BitShift($iwParam, 16) = $EN_CHANGE Then
        If $ilParam = $hAddSearch Then
            $sSearch=GUICtrlRead($idAddSearch)
            If $sSearch<>$sSearchLast Then
				If $bSearch Then $bSearchAbort=True
				_GUICtrlStatusBar_SetText($idAddStatus,"Searching...")
                While $bSearch
                    If $bSearchAbort And TimerDiff($tSearch)>=1000 Then
                        $bSearchAbort=False
                        $bSearch=False
                        ExitLoop
                    EndIf
                    Sleep(125)
                WEnd
                $tSearch=TimerInit()
				AdlibRegister("_GuiAddSearchProc")
            EndIf
        EndIf
    EndIf
    Return $GUI_RUNDEFMSG
EndFunc

Func _GuiAddSearchProc()
	AdlibUnRegister("_GuiAddSearchProc")
    Local $bQueueSel, $idItem
	$sSearchLast=$sSearch
	If $sSearch <> "" Then
		$aSearch=_AddAppsSearch($sSearch)
        If @extended==1 Then
            $bSearch=False
            $bSearchAbort=False
            Return
        EndIf
		_GUICtrlListView_BeginUpdate($hAddAppsLV)
		_GUICtrlListView_DeleteAllItems($hAddAppsLV)
        $bSearch=True
		For $i=0 To UBound($aSearch,1)-1
			;GUICtrlCreateListViewItem($aQuery[$i][0]&'|'&$aQuery[$i][1]&'|'&$aQuery[$i][2],$idAddAppsLV)
			$idItem=GUICtrlCreateListViewItem($aSearch[$i][1]&'|'&$aSearch[$i][2],$idAddAppsLV)
			If Mod($i, 2)==0 Then GUICtrlSetBkColor($idItem,0xEEEEEE)
            $bRet=_hostHasCollection($aSearch[$i][1])
            If @extended==1 Then
                $bSearch=False
                $bSearchAbort=False
                Return
            EndIf
            $bQueueSel=False
            For $j=0 To UBound($aAddQueue,1)-1
                If $aSearch[$i][1]==$aAddQueue[$j][0] Then
                    $bQueueSel=$aAddQueue[$j][2]
                    _Log($aAddQueue[$j][0])
                    ExitLoop
                EndIf
            Next
            if $bRet Or $bQueueSel Then
				_GUICtrlListView_SetItemChecked($hAddAppsLV,$i,True)
                If $bQueueSel Then
                    GUICtrlSetBkColor($idItem,0xAAAAEE)
                Else
                    GUICtrlSetBkColor($idItem,0xEEEEAA)
                EndIf
			Else
				_GUICtrlListView_SetItemChecked($hAddAppsLV,$i,False)
			EndIf
		Next
		;_GUICtrlListView_SetColumnWidth($hAddAppsLV,0,$LVSCW_AUTOSIZE)
		_GUICtrlListView_SetColumnWidth($hAddAppsLV,0,80)
		_GUICtrlListView_SetColumnWidth($hAddAppsLV,1,$LVSCW_AUTOSIZE_USEHEADER)
		_GUICtrlListView_JustifyColumn($hAddAppsLV,0,2)
		_GUICtrlListView_EndUpdate($hAddAppsLV)
	ElseIf $sSearch == "" Then
		_GUICtrlListView_BeginUpdate($hAddAppsLV)
		_GUICtrlListView_DeleteAllItems($hAddAppsLV)
		For $i=0 To UBound($aAppsSMS,1)-1
			$idItem=GUICtrlCreateListViewItem($aAppsSMS[$i][1]&'|'&$aAppsSMS[$i][0],$idAddAppsLV)
			;GUICtrlCreateListViewItem($i&'|'&$aAppsSMS[$i][1]&'|'&$aAppsSMS[$i][0],$idAddAppsLV)
			If Mod($i, 2)==0 Then GUICtrlSetBkColor($idItem,0xEEEEEE)
            $bRet=_hostHasCollection($aAppsSMS[$i][1])
            If @extended==1 Then
                ;_GUICtrlStatusBar_SetText($idAddStatus,"Ready")
                $bSearch=False
                $bSearchAbort=False
                Return
            EndIf
            $bQueueSel=False
            For $j=0 To UBound($aAddQueue,1)-1
                If $aAppsSMS[$i][1]==$aAddQueue[$j][0] Then
                    $bQueueSel=$aAddQueue[$j][2]
                    _Log($aAddQueue[$j][0])
                    ExitLoop
                EndIf
            Next
			if $bRet Or $bQueueSel Then
				_GUICtrlListView_SetItemChecked($hAddAppsLV,$i,True)
                If $bQueueSel Then
                    GUICtrlSetBkColor($idItem,0xAAAAEE)
                Else
                    GUICtrlSetBkColor($idItem,0xEEEEAA)
                EndIf
			Else
				_GUICtrlListView_SetItemChecked($hAddAppsLV,$i,False)
			EndIf
		Next
		;_GUICtrlListView_SetColumnWidth($hAddAppsLV,0,$LVSCW_AUTOSIZE)
		_GUICtrlListView_SetColumnWidth($hAddAppsLV,0,80)
		_GUICtrlListView_SetColumnWidth($hAddAppsLV,1,$LVSCW_AUTOSIZE_USEHEADER)
		_GUICtrlListView_JustifyColumn($hAddAppsLV,0,2)
		_GUICtrlListView_EndUpdate($hAddAppsLV)
		$bSearch=False
    EndIf
    $bSearch=False
	_GUICtrlStatusBar_SetText($idAddStatus,"Ready")
EndFunc


; Let's base64 CryptProtectData our creds 1st.
;ClipPut(_Base64Encode(_CryptProtectData("")))

;self.wmic=wmi.WMI("wapsccm01.ds.ohnet",namespace="root\sms\site_ohp",user=u,password=p)
;~ Global $aSccmSrv[]=[]
;~ Local $oResults
;~ $oLocator = ObjCreate("WbemScripting.SWbemLocator")
;_EnsureSMS()

;~ $aRet=_SMSGetCollections()
;~ _ArraySort($aRet)
;~ _ArrayDisplay($aRet)
;~ $ci='dt220833'
;~ $aRet=_DevGetCollections($ci,True)
;~ _ArrayDisplay($aRet)



;~ $colItems = $objtobrowse.buildinproperty
;~     For $objItem In $colItems
;~         _Log($objItem.<Name> & " - " & $objItem.<Value> & @CRLF)
;~     Next
;~ EndIf

; Cleanup
;For $i=0 To $aSccmSrv[0][0]-1
;~     If Not IsObj($aSccmSrv[$i][1]) Then ContinueLoop
;~     $aSccmSrv[$i][1].Close
;~ Next


Func _CCM_Event()
    If Not IsObj($ogCcmEvent) Then Return
EndFunc

Func ohAuth_loadToken()
    Dim $g_aAuth[1][2]
    $g_aAuth[0][0]=IniRead($sAuthIni,"ohAuth","opid",-1)
    $g_aAuth[0][1]=IniRead($sAuthIni,"ohAuth","token",-1)
    If $g_aAuth[0][0]==-1 Then Return SetError(1,0,0)
    If $g_aAuth[0][1]==-1 Then Return SetError(2,0,0)
    If Not _AuthValidateToken($g_aAuth[0][0],$g_aAuth[0][1]) Then Return SetError(3,0,0)
    Return SetError(0,0,1)
EndFunc

