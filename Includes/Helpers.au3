#include-once
#include <Array.au3>
#include "Common\CryptProtect.au3"
#include "Common\ArrayMultiColSort.au3"
#include "Common\InfinityCommon.au3"
#include "Common\CRC.au3"
# Helper Funcs
Global $goWMI[0][2], $oLocator, $g_oSMS, $gociLocator



Func _SMSGetHostColIds($sHost)
	_EnsureSMS()
    $iTimerRet=TimerInit()
    $iTimer=TimerInit()
    $oFCM=$g_oSMS.ExecQuery("Select SMS_Collection.CollectionId from SMS_FullCollectionMembership, SMS_Collection where name = '"&$sHost&"' and SMS_FullCollectionMembership.CollectionID=SMS_Collection.CollectionID", "WQL", $wbemFlagReturnImmediately)
    _Log("oFCM:"&TimerDiff($iTimer),'_SMSGetDevCollections')
    Local $iMax=0
    Local $aArray[0]
    For $objItem In $oFCM
        $iMax=UBound($aArray,1)
        ReDim $aArray[$iMax+1]
        $aArray[$iMax]=$objItem.CollectionId
    Next
    _Log("vRet:"&TimerDiff($iTimerRet),'_SMSGetHostColIds')
    Return $aArray
EndFunc

Func _SMSGetDevCollections($ci)
    $iTimerRet=TimerInit()
    $aCollIds=_SMSGetHostColIds($ci)
    $iTimer=TimerInit()
    Local $iMax=UBound($aCollIds,1)-1
    Local $sQuery="Select CollectionID,ModelName,CollectionName from SMS_DeploymentSummary Where"
    For $i=0 To $iMax
        $sQuery&=' CollectionId = "'&$aCollIds[$i]&'"'
        If $i<$iMax Then $sQuery&=' or '
    Next
    _Log("sQuery:"&TimerDiff($iTimer),'_SMSGetDevCollections')
    $iTimer=TimerInit()
    $oQuery=$g_oSMS.ExecQuery($sQuery,"WQL", $wbemFlagReturnImmediately)
    Local $aQuery[0][5]
    Local $iMax=0
    For $element In $oQuery
        $iMax=UBound($aQuery,1)
        ReDim $aQuery[$iMax+1][5]
        $aQuery[$iMax][0] = $element.ModelName
        $aQuery[$iMax][1] = $element.CollectionID
        $aQuery[$iMax][2] = $element.CollectionName
    Next
    _Log("oQuery:"&TimerDiff($iTimer),'_SMSGetDevCollections')
    _Log("vRet:"&TimerDiff($iTimerRet),'_SMSGetDevCollections')
    Return $aQuery
EndFunc

; Get Collections from SCCM
Func _SMSGetCollections($ci=Default,$cPath="/Production/APP_Applications")
    $iTimerRet=TimerInit()
	_EnsureSMS()
    $iTimer=TimerInit()
	if $ci<>Default Then
		$oResults=$g_oSMS.ExecQuery("Select Name,CollectionId,Comment from SMS_FullCollectionMembership, SMS_Collection where name = '"&$ci&"' and SMS_FullCollectionMembership.CollectionID=SMS_Collection.CollectionID", "WQL", $wbemFlagReturnImmediately)
    Else
        $oResults=$g_oSMS.ExecQuery("Select Name,CollectionId,Comment from SMS_Collection where ObjectPath = '"&$cPath&"'", "WQL", $wbemFlagReturnImmediately)
    EndIf
    _Log("oSMS: "&TimerDiff($iTimer),'_SMSGetCollections')
    Local $aArray[0][3]
    Local $iMax=0
    For $element In $oResults
        $iMax=UBound($aArray,1)
        ReDim $aArray[$iMax+1][3]
        $aArray[$iMax][0] = $element.Name
        $aArray[$iMax][1] = $element.CollectionID
        $aArray[$iMax][2] = $element.Comment
    Next
    _Log("vRet:"&TimerDiff($iTimerRet),'_SMSGetCollections')
    Return $aArray
EndFunc

; Get Collections From Host
Func _DevGetCollections($sHost,$bResolveCollectionIds=True)
    Local $colItems=""
    Local $sReturn=""
    Local $iExt=0
    Local $aReturn[0][16]
	Local $iMax=0
    $oWMI=_EnsureWMI($sHost,"/root/ccm/ClientSDK")
	; Get Applications from Host
    $iTimerRet=TimerInit()
    $iTimer=TimerInit()
    Local $colItems=$oWMI.ExecQuery("SELECT * FROM CCM_Application", "WQL", $wbemFlagReturnImmediately)
    If Not IsObj($colItems) Then
        _Log("oWmiError: "&$colItems,'_DevGetCollections')
        Return SetError(1,0,0)
    EndIf
    Local $aItems[0],$iMax
    For $objItem In $colItems
        $iMax=UBound($aItems,1)
        ReDim $aItems[$iMax+1]
        $aItems[$iMax]=$objItem
    Next
    _Log("HereX")
    _Log("oWmi: "&TimerDiff($iTimer),'_DevGetCollections')
    If $bResolveCollectionIds Then ; Get CollectionIDs from SMS
        _ArrayColInsert($aReturn,1)
        $aQuery=_SMSGetDevCollections($sHost)
        _Log("HereA")
    EndIf
    _Log("HereB")
    $iTimer=TimerInit()
    $ivDim=UBound($aReturn,2)
    _Log("HereC")
    Local $iT=0
    For $objItem In $colItems
        _Log("HereX:"&$iT)
        $iT+=1
        If Not IsObj($objItem) Then ContinueLoop
        $iMax=UBound($aReturn,1)
        ReDim $aReturn[$iMax+1][$ivDim]
        $aReturn[$iMax][0]=$objItem.Name
        $aReturn[$iMax][1]=$objItem.Description
        $aReturn[$iMax][2]=$objItem.Revision
        $aReturn[$iMax][3]=$objItem.InstallState
        $aReturn[$iMax][4]=WMIDateStringToDate($objItem.LastInstallTime)
        $aReturn[$iMax][5]=objToArray($objItem.AllowedActions,True)
        $aReturn[$iMax][6]=objToArray($objItem.InProgressActions,True)
        $aReturn[$iMax][7]=$objItem.IsMachineTarget
        $aReturn[$iMax][8]=$objItem.Id
;~ 		For $objClassProperty In $objItem.Properties_
;~ 			ConsoleWrite($objClassProperty.Name&@CRLF)
;~ 		Next
		$aReturn[$iMax][10]=$objItem.SoftwareVersion
		$aReturn[$iMax][11]=$objItem.EvaluationState
		$aReturn[$iMax][12]=$objItem.PercentComplete
        If $bResolveCollectionIds Then
            For $i=0 To UBound($aQuery,1)-1
                If $objItem.Id<>$aQuery[$i][0] Then ContinueLoop
                ;If $objItem.Name<>$aQuery[$i][2] Then
                ;    _Log($objItem.Id&': '&$objItem.Name&"<>"&$aQuery[$i][2],"_DevGetCollections")
                ;EndIf
                ;_Log($objItem.Id&': '&$aQuery[$i][1]&"<>"&$aQuery[$i][2],"_DevGetCollections")
                $aReturn[$iMax][9]=$aQuery[$i][1]
                $aReturn[$iMax][13]=$aQuery[$i][2]
            Next
        EndIf
        $aReturn[$iMax][15]=Hex(_CRC32($aReturn[$iMax][0]&$aReturn[$iMax][1]&$aReturn[$iMax][2]&$aReturn[$iMax][7]&$aReturn[$iMax][8]&$aReturn[$iMax][9]&$aReturn[$iMax][10]&$aReturn[$iMax][13]),8)
    Next
    _Log(UBound($aReturn,1)&@CRLF)
    If Not isArray($aReturn) Or UBound($aReturn,1)==0 Then
        _Log("No Results",'_DevGetCollections')
        Return SetError(1,0,0)
    EndIf
	Local $aSort[][]=[[6,1],[3,1],[0,0]]
	_ArrayMultiColSort($aReturn,$aSort)
    _Log(TimerDiff($iTimer),'_DevGetCollections')
    _Log("vRet:"&TimerDiff($iTimerRet),'_SMSGetDevCollections')
    Return SetError(0,$iExt,$aReturn)
EndFunc

;Func _CollectionRefresh

Func _GetCollection($sColl)
    _EnsureSMS()
    $oObj = $g_oSMS.Get('SMS_Collection.CollectionID="'&$sColl&'"')
    If Not IsObj($oObj) Then Return SetError(1,0,False)
    Return SetError(0,0,$oObj)
EndFunc

Func _CollectionAddResource($sHost,$sResourceId,$vColl=Default,$bNoWait=False,$bNoRefresh=False)
    If $vColl==Default Or Not IsObj($vColl) Then $vColl = _GetCollection($vColl)
    If Not IsObj($vColl) Then Return SetError(1,0,False)
    $oRule = $g_oSMS.Get("SMS_CollectionRuleDirect").SpawnInstance_()
    ;$oRule = $g_oSMS.Get("SMS_CollectionRuleQuery").SpawnInstance_()
    $oRule.RuleName=StringUpper(StringReplace($sHost,'.ds.ohnet',""))
    $oRule.ResourceClassName = "SMS_R_System"
    $oRule.ResourceID = $sResourceId
    $iRet=$vColl.AddMembershipRule($oRule)
    If $g_iWmiError Then
        SetError(1,$g_iWmiErrorExt)
        $g_iWmiErrorExt=0
        $g_sWmiError=""
        $g_iWmiError=0
        Return False
    EndIf
    _Log("AddMembershipRule,"&$vColl.CollectionId&","&$iRet&@CRLF,"_CollectionAddResource")
    If Not $bNoRefresh Then
        Sleep(5000)
        _CollectionRefresh($vColl)
    EndIf
    If Not $bNoWait Then Sleep(10000)
    Return SetError(0,0,True)
EndFunc

Func _CollectionRefresh($vColl)
    If $vColl==Default Or Not IsObj($vColl) Then $vColl = _GetCollection($vColl)
    If Not IsObj($vColl) Then Return SetError(1,0,False)
    $iRet=$vColl.RequestRefresh()
    If $g_iWmiError Then
        SetError(1,$g_iWmiErrorExt)
        $g_iWmiErrorExt=0
        $g_sWmiError=""
        $g_iWmiError=0
        Return False
    EndIf
    _Log("RequestRefresh,"&$vColl.CollectionId&","&$iRet&@CRLF,"_CollectionRefresh")
    Return SetError(0,0,$iRet)
EndFunc

Func _CollectionRemoveResource($sHost,$sResourceId,$vColl,$bNoWait=False,$bNoRefresh=False)
    If $vColl==Default Or Not IsObj($vColl) Then $vColl = _GetCollection($vColl)
    $oRule = $g_oSMS.Get("SMS_CollectionRuleDirect").SpawnInstance_()
    ;$oRule.RuleName=$sHost
    $oRule.ResourceClassName = "SMS_R_System"
    $oRule.ResourceID = $sResourceId
    $iRet=$vColl.DeleteMembershipRule($oRule)
    If $g_iWmiError Then
        SetError(1,$g_iWmiErrorExt)
        $g_iWmiErrorExt=0
        $g_sWmiError=""
        $g_iWmiError=0
        Return False
    EndIf
    _Log("DeleteMembershipRule,"&$vColl.CollectionId&","&$iRet&@CRLF,"_CollectionRemoveResource")
    If Not $bNoRefresh Then
        Sleep(5000)
        _CollectionRefresh($vColl)
    EndIf
    If Not $bNoWait Then Sleep(10000)
    Return SetError(0,0,True)
EndFunc

Func _HostSyncPolicies($sHost,$aSched=Default)
    If $aSched==Default Then $aSched=StringSplit("121|021|022|002",2)
    Local $oNS=_EnsureWMI($sHost,"/root/ccm")
    Local $oClass = $oNS.Get("SMS_Client")
    For $i=0 To UBound($aSched)-1
        _HostTriggerSchedule($sHost,$aSched[$i],$oClass,$oNS)
    Next
EndFunc

Func _HostTriggerSchedule($sHost,$sSched,$oClass=Default,$oNS=Default)
    If $oNS==Default Then $oNS=_EnsureWMI($sHost,"/root/ccm")
    If $oClass==Default Then $oClass = $oNS.Get("SMS_Client")
    $oParams = $oClass.Methods_("TriggerSchedule").inParameters.SpawnInstance_()
    $oParams.sScheduleID = "{00000000-0000-0000-0000-000000000"&$sSched&"}"
    $iRet=$oNS.ExecMethod("SMS_Client", "TriggerSchedule", $oParams)
    If $g_iWmiError Then
        SetError(1,$g_iWmiErrorExt)
        $g_iWmiErrorExt=0
        $g_sWmiError=""
        $g_iWmiError=0
        Return False
    EndIf
    ;$iRet=$oClass.TriggerSchedule("{00000000-0000-0000-0000-000000000"&$sSched&"}")
    _Log("TriggerSchedule,"&$sSched&","&$iRet&@CRLF,"_HostTriggerSchedule")
    Return SetError(0,0,True)
EndFunc

Func _HostResetPolicy($sHost,$bHard=False)
    If $bHard Then
        _Log('HostResetPol,Hard')
    Else
        _Log('HostResetPol')
    EndIf
    Local $oCCMNamespace=_EnsureWMI($sHost,"/root/ccm")
    $oInstance = $oCCMNamespace.Get("SMS_Client")
    $iRet=$oInstance.ResetPolicy($bHard ? 1 : 0)
    _Log("ResetPolicy,"&","&$iRet&@CRLF,"_HostResetPolicy")
    Return $iRet
EndFunc

Func _HostEvalPolicy($sHost)
    _Log('HostEvalPol')
    Local $oCCMNamespace=_EnsureWMI($sHost,"/root/ccm")
    $oInstance = $oCCMNamespace.Get("SMS_Client")
    $iRet=$oInstance.EvaluateMachinePolicy()
    _Log("EvalPolicy,"&","&$iRet&@CRLF,"_HostEvalPolicy")
EndFunc

Func _HostRetrievePolicy($sHost)
    _Log('HostEvalPol')
    Local $oCCMNamespace=_EnsureWMI($sHost,"/root/ccm")
    $oInstance = $oCCMNamespace.Get("SMS_Client")
    $iRet=$oInstance.RequestMachinePolicy(0)
    _Log("RetrievePolicy,"&","&$iRet&@CRLF,"_HostRetrievePolicy")
EndFunc


Func objToArray($oObj,$bStr=False)
    Local $aRet=[], $iMax=0
	If IsObj($oObj) or IsArray($oObj) Then
		For $oValue In $oObj
			ReDim $aRet[$iMax+1]
			$aRet[$iMax]=$oValue
			$iMax+=1
		Next
	Else
		Return VarGetType($oObj)
	EndIf
	if $bStr Then
		Return _ArrayToString($aRet)
	EndIf
    Return $aRet
EndFunc   ;==>objToArray

;-----------Community Funcs-----------

Func WMIDateStringToDate($dtmDate)
    Return (StringMid($dtmDate, 5, 2) & "/" & _
        StringMid($dtmDate, 7, 2) & "/" & StringLeft($dtmDate, 4) _
        & " " & StringMid($dtmDate, 9, 2) & ":" & StringMid($dtmDate, 11, 2) & ":" & StringMid($dtmDate,13, 2))
EndFunc

Func Array_Join($aArray, $sSeparator = " , ")
    Local $n, $sOut = ""
    If IsObj($aArray) Then
        For $value In $aArray
            $sOut &= $value & $sSeparator
        Next
        Return StringTrimRight($sOut, StringLen($sSeparator))
    Else
        For $n = 0 To UBound($aArray) - 1
            $sOut &= $aArray[$n] & $sSeparator
        Next
        Return StringTrimRight($sOut, StringLen($sSeparator))
    EndIf
EndFunc   ;==>Array_Join

; Convert the client (GUI) coordinates to screen (desktop) coordinates
Func ClientToScreen($hWnd, ByRef $x, ByRef $y)
    Local $stPoint = DllStructCreate("int;int")

    DllStructSetData($stPoint, 1, $x)
    DllStructSetData($stPoint, 2, $y)

    $iRet=DllCall("user32.dll", "int", "ClientToScreen", "hwnd", $hWnd, "ptr", DllStructGetPtr($stPoint))

    $x = DllStructGetData($stPoint, 1)
    $y = DllStructGetData($stPoint, 2)
    ; release Struct not really needed as it is a local
    $stPoint = 0
    Return $iRet
EndFunc   ;==>ClientToScreen

; Show at the given coordinates (x, y) the popup menu (hMenu) which belongs to a given GUI window (hWnd)
Func TrackPopupMenu($hWnd, $hMenu, $x, $y)
    Return DllCall("user32.dll", "int", "TrackPopupMenuEx", "hwnd", $hMenu, "int", 0, "int", $x, "int", $y, "hwnd", $hWnd, "ptr", 0)
EndFunc   ;==>TrackPopupMenu

Func _CheckHost($sHost,$iTimeout=10000)
    ; Validate Hostname/IP
    Local $aRet[3]
    $sHost=StringLower($sHost)
    if StringLeft($sHost,2)=="dt" or StringLeft($sHost,2)=="lt" or StringLeft($sHost,2)=="pr" and not StringInStr($sHost,".ds.ohnet") Then
        $sHost&=".ds.ohnet"
    EndIf
    $bIsIP=StringRegExp($sHost,"^(([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])\.){3}([0-9]|[1-9][0-9]|1[0-9]{2}|2[0-4][0-9]|25[0-5])$")
    $bIsHostname=StringRegExp($sHost,"^(([a-zA-Z0-9]|[a-zA-Z0-9][a-zA-Z0-9\-]*[a-zA-Z0-9])\.)*([A-Za-z0-9]|[A-Za-z0-9][A-Za-z0-9\-]*[A-Za-z0-9])$")
    If Not $bIsIP And Not $bIsHostname Then
        MsgBox(16,"Error", "Please enter a valid Hostname or IP.")
        Return False
    EndIf
    $sMsgPrefix="There was an error resolving the host."&@LF
    $sMsgSuffix=@LF&@LF&"Would you like to continue anyway? (This may make program unresponsive)"

    $aRes=_Resolve($sHost,$iTimeout)
    $iRet=Null
    If Not IsArray($aRes) Then
        $iRet=MsgBox(16,"Error", $sMsgPrefix&"Failed to get Hostname or IP from host specified."&$sMsgSuffix)
    Else
        $iErr=0
        If $aRes[0]==False Then $iErr+=10
        If $aRes[1]==False Then $iErr+=1
        If $iErr==1 Then
            $iRet=MsgBox(49,"Error",$sMsgPrefix&"Hostname lookup Failed."&$sMsgSuffix)
        ElseIf $iErr==10 Then
            $iRet=MsgBox(49,"Error", $sMsgPrefix&"IP lookup Failed."&$sMsgSuffix)
        ElseIf $iErr==11 Then
            $iRet=MsgBox(49,"Error", $sMsgPrefix&"Failed to get Hostname and IP."&$sMsgSuffix)
        EndIf
    EndIf
    If $iRet==2 Then Return False
    Return True
EndFunc

Global $ogCcmEvent
