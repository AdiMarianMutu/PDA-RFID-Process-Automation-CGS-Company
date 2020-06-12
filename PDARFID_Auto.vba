' Lowlevel Events
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdSHow As Long) As Long
Private Const SW_MAXIMIZE = 3
Private Const SW_MINIMIZE = 6
Private Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Private Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN = &H2
Private Const MOUSEEVENTF_LEFTUP = &H4
Private Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
' Create custom variable that holds two integers
Type POINTAPI
   x As Long
   y As Long
End Type



' Retrieves the IE tab handler using the Document.Title or URL property
Private Function GetIEHandler(Optional sTitle As String = "", Optional searchByURL As Boolean = False) As Object
On Error GoTo ErrorHandler:

Start:
    Dim Shell As Object
    Dim ie As Object
    Dim i As Variant
    
    Set Shell = CreateObject("Shell.Application")
    
    i = 0
    Set GetIEHandler = Nothing
    While i < Shell.Windows.Count And GetIEHandler Is Nothing
        Set ie = Shell.Windows.Item(i)
        If Not ie Is Nothing Then
            If TypeName(ie) = "IWebBrowser2" Then
                If ie.LocationURL <> "" And InStr(ie.LocationURL, "file://") <> 1 Then
                    If Len(Trim(sTitle)) = 0 Then
                        Set GetIEHandler = ie
                    Else
                        If searchByURL Then
                            If InStr(ie.LocationURL, sTitle) > 0 Then
                                Set GetIEHandler = ie
                            End If
                        Else
                            If InStr(ie.document.Title, sTitle) > 0 Then
                                Set GetIEHandler = ie
                            End If
                        End If
                    End If
                End If
            End If
        End If
        i = i + 1
    Wend
    Exit Function
    
ErrorHandler:
    ' Sometimes the "Permission Denied" exception can rise, if so we restart the function
    Application.Wait (Now + TimeValue("0:00:3"))
    GoTo Start
End Function

' Used to dispatch a keyboard event change
Private Function HTMLElementDispatchValue(ByRef objDocument, ByRef objElement As Object, ByVal strValue As String)
    ' Used to dispatch the "change" event to the item object
    Dim post As Object, evt As Object
    Set evt = objDocument.createEvent("keyboardevent")
    evt.initEvent "change", True, False
    
    Set post = objElement
        post.Value = strValue
        post.fireEvent "onchange"
    post.dispatchEvent evt
End Function

Private Function CheckRemedySynUpsStatus()
Start:
    ' If 'Remedy' tab is not open
    If GetIEHandler("Remedy") Is Nothing Then
        If MsgBox("[PDA & RFID PROCESS]" & vbNewLine & vbNewLine & " > 'BMC Remedy (Search)' Tab not open!" & vbNewLine & vbNewLine & "Press 'Cancel' to abort" & vbNewLine & vbNewLine & "N.B: If this message appeared after an internet connection drop or a similiar case:" & vbNewLine & " > Press cancel and start back from where you left", vbOKCancel) = vbOK Then
            GoTo Start
        End If
        
        ' Cancel action
        End
    End If
    ' Checks if the Syncreon session is still available
    If Not (GetIEHandler("syncreon Axional - Login") Is Nothing) Then
        If MsgBox("[PDA & RFID PROCESS]" & vbNewLine & vbNewLine & " > Your Syncreon session is not valid anymore, please log-in (on IE) and then press OK to continue or Cancel to abort", vbOKCancel) = vbCancel Then
            End
        End If
    End If
    ' If 'Syncreon' tab is not open
    If GetIEHandler("ECI Call In") Is Nothing Then
        If MsgBox("[PDA & RFID PROCESS]" & vbNewLine & vbNewLine & " > 'ECI Call In (comp_ww)' Tab not open!" & vbNewLine & vbNewLine & "Press 'Cancel' to abort" & vbNewLine & vbNewLine & "N.B: If this message appeared after an internet connection drop or a similiar case:" & vbNewLine & " > Press cancel and start back from where you left", vbOKCancel) = vbOK Then
            GoTo Start
        End If
        
        ' Cancel action
        End
    End If
    ' If 'UPS Tracking' link is not open we will open a new tab
    If GetIEHandler("UPS") Is Nothing Then
        With GetIEHandler()
            .Visible = True
            .Navigate "https://www.ups.com/track?loc=en_US&tracknum=<CGS>", CLng(2048)
        End With
        
        Application.Wait (Now + TimeValue("0:00:5"))
    End If
End Function

' Retrieves if the Syncreon tab is on the New Query tab
Private Function SyncreonSearchActive() As Boolean
    CheckRemedySynUpsStatus
    
    If GetIEHandler("https://axional1.syncreon.com/servlet/", True) Is Nothing Then
        SyncreonSearchActive = False
    Else
        SyncreonSearchActive = True
    End If
End Function

' Opens the Syncreon New Query tab
Private Function SyncreonNewSearch()
    CheckRemedySynUpsStatus

    If SyncreonSearchActive() = False Then
        GetIEHandler("ECI Call").document.getElementById("webapp_ui_button_button_3").Click
    End If
    
    CheckRemedySynUpsStatus
End Function

Private Function SyncreonSearchINC(inc As String)
    SyncreonNewSearch
    
    With GetIEHandler("ECI Call In")
        ' Waiting for the page to load
        While .ReadyState <> 4 Or .Busy: DoEvents: Wend
        Application.Wait (Now + TimeValue("0:00:2"))
        
        ' Used to check if Syncreon session is still valid
        CheckRemedySynUpsStatus
        
        HTMLElementDispatchValue .document, .document.getElementById("_QRY_eci_call_in.remedy"), inc
        .document.getElementById("webapp_ui_button_button_8").Click
    End With
End Function

Private Function UPSFormatDeliveryHour(ByVal hour) As String
    ' Converts to 24h
    hour = Replace(hour, ".", "")
    
    UPSFormatDeliveryHour = Format(hour, "HH:mm")
End Function
Private Function UPSUpdateStatusCheck() As String
On Error Resume Next
    
    ' Used only for the 'Update' status
    ' Sometimes UPS shows the 'Update' status and by clicking 'Detailed View' we can get more info
    ' sometimes successfully delivered status are placed in the 'Update' status, so we must check
    Dim u_status As String
    Dim cnt As Integer: cnt = 0
    Dim i As Integer
    Dim whileFlag As Boolean: whileFlag = True
    
    With GetIEHandler("UPS")
        While whileFlag
            u_status = .document.getElementById("stApp_ShpmtProg_LVP_milestone_name_" & cnt).innerText
            
            If InStr(u_status, "Delivered") Then
                delDate = .document.getElementById("stApp_ShpmtProg_LVP_milestone_" & cnt & "_date_1").innerText
                delHour = .document.getElementById("stApp_ShpmtProg_LVP_milestone_" & cnt & "_time_1").innerText
                    
                delDate = delDate & " " & UPSFormatDeliveryHour(delHour)
                
                GoTo finish
            End If
                
            ' If more than 15 rows, will raise the <manual_check_required> status
            If cnt > 15 Then
                whileFlag = False
            End If
                
            cnt = cnt + 1
        Wend
    End With
    
' No 'Delivered' note found
delDate = "Update"

finish:
    UPSUpdateStatusCheck = delDate
End Function
Private Function UPSGetDeliveryDate(ByVal trackingNumber) As String
    On Error GoTo ErrorHandler:
    
    delDate = ""
    delStatus = ""
    
    With GetIEHandler("UPS")
        ' Searches the Tracking Number
        .Navigate "https://www.ups.com/track?loc=en_US&tracknum=" & trackingNumber
                
        While .ReadyState <> 4 Or .Busy Or Not (TypeName(.document.getElementById("toTitle")) = "Null"): DoEvents: Wend
        CheckRemedySynUpsStatus
        
        delStatus = .document.getElementById("stApp_txtPackageStatus").innerText
        
        If InStr(LCase(delStatus), "delivered") Then
            ' Retrieves the date
            delDate = .document.getElementById("stApp_deliveredDate").innerText
            ' Retrieves the hour (12h) and converts to 24h
            delDate = delDate & " " & UPSFormatDeliveryHour(Mid(.document.getElementById("stApp_eodDate").innerText, InStr(.document.getElementById("stApp_eodDate").innerText, ":") - 2))

            UPSGetDeliveryDate = delDate
        ElseIf InStr(LCase(delStatus), "update") Then
            UPSGetDeliveryDate = UPSUpdateStatusCheck()
        Else
            UPSGetDeliveryDate = delStatus
        End If
    End With
    Exit Function
    
ErrorHandler:
    If Len(delDate) = 0 Then
        ' UPS could not locate the shipment details for this tracking number
        With GetIEHandler("UPS")
            If TypeName(.document.getElementById("stApp_error_alert_list0")) <> "Null" Then
                UPSGetDeliveryDate = .document.getElementById("stApp_error_alert_list0").innerText
                Exit Function
            End If
        End With
    End If
End Function

Private Function SyncreonCheckDeliveryStatus(ByRef statusVal As String) As String
    On Error GoTo ErrorHandler:
    
    CheckRemedySynUpsStatus
    
    ' Syncreon
    With GetIEHandler("ECI Call In")
        ' Waiting for the page to load
        While .ReadyState <> 4 Or .Busy: DoEvents: Wend
        CheckRemedySynUpsStatus
        
        trackingNumber = .document.getElementById("jrepapp_view_formauto_boxes_sqltable_sqltable_1_0_tracknbr_cell").innerText
        ' Extracts the tracking number from the bugged HTML href tag
        If InStr(trackingNumber, "<") > 0 Then
            trackingNumber = Mid(trackingNumber, InStr(trackingNumber, ">") + 1, 18)
        End If
        
        delStatus = .document.getElementById("jrepapp_view_formauto_boxes_sqltable_sqltable_1_0_carr_statdesc_cell").innerText
        
        If delStatus = "Entregado" Then
            CheckRemedySynUpsStatus
            
            delDate = .document.getElementById("jrepapp_view_formauto_boxes_sqltable_sqltable_1_0_dt_rec_cell").innerText
            
            statusVal = delStatus
            SyncreonCheckDeliveryStatus = Left(delDate, Len(delDate) - 3)
        ' If it was delivered through UPS but in Syncreon appears SKYNET as carrier
        ElseIf InStr(LCase(.document.getElementById("jrepapp_view_formauto_boxes_sqltable_sqltable_1_0_carrier_cell").innerText), "tipsa") = 0 Then
            CheckRemedySynUpsStatus
            
            ' UPS
            delDate = UPSGetDeliveryDate(trackingNumber)
            
            If InStr(LCase(delDate), "sender") = 0 And InStr(LCase(delDate), "returning") = 0 And InStr(LCase(delDate), "ups could not") = 0 And InStr(delDate, "/") = 0 And InStr(LCase(delDate), "update") = 0 Then
                ' Not yet delivered
                statusVal = "-1"
            ElseIf InStr(LCase(delDate), "sender") <> 0 Or InStr(LCase(delDate), "returning") <> 0 Then
                ' Returning to sender
                statusVal = delDate ' Returns the UPS comment status
            ElseIf InStr(LCase(delDate), "ups could not locate the shipment details for this tracking number") <> 0 Then
                ' Tracking code not active
                statusVal = "ups_trck_fail"
            ElseIf InStr(LCase(delDate), "update") <> 0 Then
                ' Manual check required
                statusVal = "manual_check"
            Else
                ' Delivered
                statusVal = "Delivered"
                SyncreonCheckDeliveryStatus = delDate
            End If
        Else
            ' TIPSA Courier but the package is still in transit
            statusVal = "-1"
        End If
        Exit Function
        
ErrorHandler:
        ' First we check for the "NO EXISTE" syncreon comment
        If Not (.document.getElementById("H__eci_call_in.comments_syncreon") Is Nothing) Then
            syncComment = .document.getElementById("H__eci_call_in.comments_syncreon").innerText
            
            ' If the "NO EXISTE" syncreon comment exists
            If Len(syncComment) <> 0 Then
                statusVal = syncComment
            ' Otherwhise we check the delivery date
            ElseIf Len(delDate) = 0 Then
                ' Sometimes the Syncreon page does not have the tracking code
                If Len(trackingNumber) = 0 Then
                    statusVal = "no_track"
                Else
                    ' Possible not yet delivered
                    statusVal = "-1"
                End If
            End If
            Exit Function
        ' Otherwhise we check the delivery date (used in case the "NO EXISTE" syncreon element does not exists)
        ElseIf Len(delDate) = 0 Then
            statusVal = "-1"
        End If
    End With
End Function

' Because Remedy does not allow "external" queries, we bypass this feature by using HTTP callbacks
Private Function RemedyOpenINC(ByVal inc As String, ByRef Remedy, Optional delStatus As String = "-1")
    With GetIEHandler("Remedy")
        ' Used to prevent the "Stay on this page/Leave this page" popup
        .document.parentWindow.execScript "window.onbeforeunload = null;", "Javascript"
        ' Open the new INC in a new tab
        .Navigate "https://extranet.inditex.com/arsys/servlet/ViewFormServlet?server=itxars&form=HPD:Help+Desk&qual=%271000000161%27%3D%22" & inc & "%22", CLng(2049)
        ' Closes the current open INC tab
        .Quit
    End With
    
    ' Wait for the new Remedy tab to load and assign the handler
    While Remedy Is Nothing: Set Remedy = GetIEHandler("Remedy"): DoEvents: Wend
    While Remedy.ReadyState <> 4 Or Remedy.Busy: DoEvents: Wend
    
    ' If the package was refused adds the notes to Remedy
'    If InStr(LCase(delStatus), "refused") > 0 Or InStr(LCase(delStatus), "sender") > 0 Then
'        With Remedy
'            Application.Wait (Now + TimeValue("0:00:5"))
'
'            ' Checks if the note ia already present
'            If InStr(.document.GetElementById("T301389614").InnerText, delStatus) = 0 Then
'                ' Adds the notes
'                HTMLElementDispatchValue .document, .document.GetElementById("arid_WIN_1_304247080"), "Courier Notes: " & delStatus
'                ' Checks the 'Public' checkbox
'                .document.GetElementById("WIN_1_rc1id1000000761").Click
'                ' Saves the changes
'                .document.GetElementById("WIN_1_301614800").Click
'            End If
'        End With
'    End If
End Function

' Function used to click on the Status Reason Item
Private Function RemedyProcessINC_statReasonNoFurthActReqClick()
    Dim curBack As POINTAPI
    Dim x As Long, y As Long
    
    With GetIEHandler("Remedy")
        ' Brings to top IE
        ShowWindow .hwnd, SW_MAXIMIZE
        
        ' Used to disable all the elements bar in order to correctly click on the Status Reason item
        If .MenuBar = True Then
            .MenuBar = False
        End If
        If .StatusBar = True Then
            .StatusBar = False
        End If
        If .Toolbar = True Then
            .Toolbar = False
        End If
    
        ' Creates an empty inputbox where to save the coordinates of the "No further action required" item (In Javascript)
        .document.parentWindow.execScript "var e = document.createElement('INPUT');e.setAttribute('id', 'adi_marian_mutu');e.setAttribute('type', 'hidden');document.body.appendChild(e);", "Javascript"
    
        ' Expands the Status Reason listbox
        .document.getElementById("arid_WIN_1_1000000881").Click
        
        ' Waits for the list to load
        While .ReadyState <> 4 Or .Busy: DoEvents: Wend
        Application.Wait (Now + TimeValue("0:00:5"))
        
        ' Retrieves and saves into the hidden input box the element coordinates (In Javascript)
        .document.parentWindow.execScript "try { var el = document.getElementsByClassName(""MenuTable"")(0).getElementsByClassName(""MenuTableBody"")(0).getElementsByClassName(""MenuEntryName"")(5);function getPosition(el){var xPos = 0;var yPos = 0;while (el) {if (el.tagName == ""BODY"") {var xScroll = el.scrollLeft || document.documentElement.scrollLeft;var yScroll = el.scrollTop || document.documentElement.scrollTop;xPos += (el.offsetLeft - xScroll + el.clientLeft);yPos += (el.offsetTop - yScroll + el.clientTop);} else {xPos += (el.offsetLeft - el.scrollLeft + el.clientLeft);yPos += (el.offsetTop - el.scrollTop + el.clientTop);}el = el.offsetParent;}return{x: xPos,y: yPos};} var _p = getPosition(el); document.getElementById(""adi_marian_mutu"").value = _p.x + "","" + _p.y; } catch(err) { window.alert(""<PDARFID_AUTO_JS_NoFurthActClick_ERROR>\n\n"" + err); }", "Javascript"
        
        ' Calculates the correct element coordinates
        coord_str = .document.getElementById("adi_marian_mutu").Value
        x = .Left + (Left(coord_str, InStr(coord_str, ",") - 1) + 50)
        y = .Top + (Right(coord_str, InStr(coord_str, ",") - 1) + 72)
        
        ' Saves the cursor coordinates
        GetCursorPos curBack
        ' Moves the cursor to the element coordinates
        SetCursorPos x, y
        ' Press and release the left mouse button
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        
        ' Take the cursor to the last known coordinates and minimize IE
        SetCursorPos curBack.x, curBack.y
        ShowWindow .hwnd, SW_MINIMIZE
    End With
End Function

' Return values:
' -> 0 = Tkt not closed
' -> 1 = Tkt closed
' -> 2 = Tkt already closed
Private Function RemedyProcessINC(ByVal inc As String, ByVal note As String) As Integer
    CheckRemedySynUpsStatus
    
    ' Used to refresh the Remedy tab handler
    Dim IERemedy As Object
    Set IERemedy = Nothing
    ' Searches the INC on Remedy
    RemedyOpenINC inc, IERemedy

    ' New Remedy page tab with the INC
    With IERemedy
        While .ReadyState <> 4 Or .Busy: DoEvents: Wend
        CheckRemedySynUpsStatus
        Application.Wait (Now + TimeValue("0:00:2"))
        
        remedyTktStatus = .document.getElementById("arid_WIN_1_7").Value
        If remedyTktStatus <> "Resolved" And remedyTktStatus <> "Closed" Then
            ' Changes the Status dropbox menu to Resolved
            HTMLElementDispatchValue .document, .document.getElementById("arid_WIN_1_7"), "Resolved"
            ' Changes the Status Reason dropbox menu to No Further Action Required
            RemedyProcessINC_statReasonNoFurthActReqClick
            
            ' Before saving the changes, checks if the Status Reason is "No Further Action Required"
            If .document.getElementById("arid_WIN_1_1000000881").Value <> "No Further Action Required" Then
                MsgBox ("[PDA & RFID PROCESS]" & vbNewLine & vbNewLine & "Unable to correctly click on ""No Further Action Required"" item!" & vbNewLine & vbNewLine & " > Click ""Ok"" to stop the script and try again" & vbNewLine & vbNewLine & "> If the problem persists, please call Adi")
                End
            End If
            
            ' Add the Delivery Date to the 'Notes' inputbox
            ' Before adding the Delivery Date note, checks if is already present
            delNote = "Delivered on: " & note
            dn = .document.getElementById("T301389614").innerText
            
            If InStr(dn, "Delivered on:") > 0 Then
                dn = Mid(dn, InStr(dn, "Delivered on:"))
                dn = Mid(dn, 1, InStr(dn, vbNewLine) - 1)
            End If
            
            If dn <> delNote Then
                HTMLElementDispatchValue .document, .document.getElementById("arid_WIN_1_304247080"), delNote
            End If
            
            ' Checks the 'Public' checkbox
            .document.getElementById("WIN_1_rc1id1000000761").Click
            
            ' Saves the changes
            .document.getElementById("WIN_1_301614800").Click
            
            ' Checks if the tkt was successfully saved
            Application.Wait (Now + TimeValue("0:00:5"))
            
            If TypeName(.document.getElementById("pbartable").getElementsByClassName("prompttext prompttexterr")(0)) <> "Nothing" Then
                RemedyProcessINC = 0
            Else
                RemedyProcessINC = 1
            End If
        Else
            RemedyProcessINC = 2
        End If
    End With
End Function

Private Function RemedyExtractSR(ByRef Remedy As Object) As String
    On Error GoTo ErrorHandler:

    With Remedy
        r = .document.getElementById("arid_WIN_1_1000000652").innerText
        
        Dim rgx As Object
        Set rgx = CreateObject("VBScript.RegExp"): rgx.Pattern = "[0-9]":
        If rgx.Test(r) Then
            ' Retrives the last SR from Remedy
            ' If the SR was logged multiple times, extracts only the last SR
            If Len(r) > 7 Then
                RemedyExtractSR = Right(r, 7)
            Else
                RemedyExtractSR = r
            End If
        Else
            RemedyExtractSR = "<no_SR>"
        End If
    End With
    Exit Function
    
ErrorHandler:
    ' If an unexpected error will rise
    RemedyExtractSR = "<unable_to_get_SR>"
End Function

Private Function GetColumnLetterFromCellValue(ByVal cellValue As String) As String
    On Error GoTo ErrorHandler:
    ret = Split(Cells(1, WorksheetFunction.Match(cellValue, ActiveWorkbook.Sheets(ActiveSheet.Name).Range("1:1"), 0)).Address(True, False), "$")
    
    GetColumnLetterFromCellValue = ret(0)
    Exit Function
    
ErrorHandler:
    ' If an unexpected error will rise
    MsgBox ("[PDA & RFID PROCESS]" & vbNewLine & vbNewLine & "Function: _GetColumnLetterFromCellValue_ *no value found" & vbNewLine & vbNewLine & " > Click ""Ok"" to stop the script and please call Adi to inform him about the problem")
    End
End Function


Public Sub PDARFIDAuto()
    ' Excel INC column header
    incColumn = GetColumnLetterFromCellValue("Incident ID*+")
    ' Excel Notes column header
    notesColumn = GetColumnLetterFromCellValue("Notes")
    ' Excel Service Request column header
    srColumn = Chr(Asc(notesColumn) + 2)
    ' Excel Delivery Date column header
    delDateColumn = Chr(Asc(notesColumn) + 3)

    ' Used to check if IE is running
    Dim ie As Object: Set ie = GetIEHandler
    If ie Is Nothing Then
        MsgBox ("[PDA & RFID PROCESS]" & vbNewLine & vbNewLine & "Please open Internet Explorer to continue")
        End
    Else
    ' Used to remove the Favorites Bar from IE because if is enabled the click option on the Status Reason item will not work correctly
        With ie
            If .FullScreen = False Then
                .FullScreen = True
                .FullScreen = False
            End If
        End With
    End If
    
    ' Checks if Remedy, Syncreon and the UPS Tracking tabs are open
    CheckRemedySynUpsStatus
    
    ' Changes the width of the "Notes" column
    If Columns(GetColumnLetterFromCellValue("Notes")).ColumnWidth < 29 Then
        Columns(GetColumnLetterFromCellValue("Notes")).ColumnWidth = 29
    End If
    ' Adds the column header if not present
    If Range(srColumn & 1) <> "Service Request" Then
        Columns(srColumn).HorizontalAlignment = xlCenter
        Range(srColumn & 1) = "Service Request"
        Columns(srColumn).ColumnWidth = 15
        Range(srColumn & 1).Interior.ColorIndex = 43
        Range(srColumn & 1).Borders.LineStyle = xlContinuous
    End If
    If Range(delDateColumn & 1) <> "Delivery Date" Then
        Columns(delDateColumn).HorizontalAlignment = xlCenter
        Range(delDateColumn & 1) = "Delivery Date"
        Columns(delDateColumn).ColumnWidth = 28
        Range(delDateColumn & 1).Interior.ColorIndex = 43
        Range(delDateColumn & 1).Borders.LineStyle = xlContinuous
    End If
    
    ' Contains the selected Excel rows by the user
    Dim exSelRows As Variant: Set exSelRows = Selection.Rows:
    
    ' Iterates all the selected rows
    For Each Rng In exSelRows
        ' Skip the hidden rows
        If Not Rng.EntireRow.Hidden Then
            inc = Range(incColumn & Rng.Row)
        
            ' Working feedback
            Range("A" & Rng.Row & ":Z" & Rng.Row).Interior.ColorIndex = 43
            Range(notesColumn & Rng.Row) = "<...working...>"
            Range(srColumn & Rng.Row) = ""
            Range(delDateColumn & Rng.Row) = ""
        
            ' Search the selected INC
            SyncreonSearchINC (inc)
            ' Will get the delivery status feedback
            Dim delStatus As String: delStatus = ""
            ' Retrieves the delivery date
            delDate = SyncreonCheckDeliveryStatus(delStatus)

            ' If the device was delivered, the INC will be processed
            If delStatus = "Delivered" Or delStatus = "Entregado" Then
                remResult = RemedyProcessINC(inc, delDate)
                
                ' If the remedy tkt is not successfully saved
                If remResult = 0 Then
                    Range(notesColumn & Rng.Row) = "<remedy_could_not_resolve_tkt>"
                ElseIf remResult = 1 Then
                    Range(notesColumn & Rng.Row) = "<remedy_tkt_resolved>"
                    
                    ' Changes the color of successfully received and saved tkts
                    Range(notesColumn & Rng.Row).Font.ColorIndex = 10
                    Range(srColumn & Rng.Row).Font.ColorIndex = 10
                    Range(delDateColumn & Rng.Row).Font.ColorIndex = 10
                Else
                    Range(notesColumn & Rng.Row) = "<remedy_tkt_already_resolved>"
                End If
                
                ' Extracts the SR
                Range(srColumn & Rng.Row) = RemedyExtractSR(GetIEHandler("Remedy"))
                ' Updates the Delivery Date column
                Range(delDateColumn & Rng.Row) = "Delivered on: " & delDate
            Else
                ' Updates the notes column
                ' If is not any of the possible cases below
                If InStr(delStatus, "NO EXISTE") = 0 And InStr(LCase(delStatus), "refused") = 0 And InStr(LCase(delStatus), "sender") = 0 And InStr(LCase(delStatus), "returning") = 0 Then
                    ' If the Syncreon page is empty
                    
                    If delStatus = "" Or delStatus = "null" Then
                        delStatus = "<sync_page_empty>" '"<syncreon_empty_or_manual_check_required>"
                        Range(notesColumn & Rng.Row).Font.ColorIndex = 3
                    ElseIf delStatus = "no_track" Then
                        delStatus = "<sync_trckcode_not_available>"
                        
                        ' Changes the color to highlight the tkt
                        Range(notesColumn & Rng.Row).Font.ColorIndex = 32
                    ElseIf delStatus = "ups_trck_fail" Then
                        delStatus = "<ups_trckcode_not_active>"
                        
                        ' Changes the color to highlight the tkt
                        Range(notesColumn & Rng.Row).Font.ColorIndex = 32
                    ElseIf delStatus = "manual_check" Then
                        delStatus = "<MANUAL_CHECK_REQUIRED>" & vbNewLine & "Please go to UPS > Detailed View  and check there for any useful information"
                        
                        ' Changes the color to highlight the tkt
                        Range(notesColumn & Rng.Row).Font.ColorIndex = 3
                    Else
                    ' If this point is reached, this means that the device is still not delivered
                        delStatus = "<not_yet_delivered>"
                    End If
                ' If the device was refused or returned to sender, we still must extract the Service Request number
                ElseIf InStr(LCase(delStatus), "refused") > 0 Or InStr(LCase(delStatus), "sender") > 0 Or InStr(LCase(delStatus), "returning") > 0 Or InStr(LCase(delStatus), "update") > 0 Then
                    ' Used to refresh the Remedy tab handler
                    Dim IERemedy As Object
                    Set IERemedy = Nothing
                    
                    ' Searches the INC on Remedy
                    RemedyOpenINC inc, IERemedy, delStatus
                    ' Delay to allow the page elements to load
                    Application.Wait (Now + TimeValue("0:00:3"))
                    ' Extracts the SR
                    Range(srColumn & Rng.Row) = RemedyExtractSR(IERemedy)
                End If
                
                ' Sometimes UPS Tracking leaves a non complete note "Returning" instead of "Returning to Sender"
                If LCase(delStatus) = "returning" Then
                    delStatus = "Returning to Sender"
                End If
                
                Range(notesColumn & Rng.Row) = delStatus
                
                If InStr(delStatus, "Sender") Then
                    ' Changes the color to highlight the tkt
                    Range(notesColumn & Rng.Row).Font.ColorIndex = 46
                    Range(srColumn & Rng.Row).Font.ColorIndex = 46
                End If
            End If
        
            ' If the status is "working" something went wrong...
            If Range(notesColumn & Rng.Row) = "<...working...>" Then
                Range(notesColumn & Rng.Row) = "<!something_went_wrong_MANUAL_CHECK_REQUIRED!>"
            End If
            
            ' Resets the working row color
            Range("A" & Rng.Row & ":Z" & Rng.Row).Interior.ColorIndex = xlNone
        End If
    Next Rng
End Sub
