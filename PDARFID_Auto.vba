' Lowlevel Events
Private Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdSHow As Long) As Long
Private Const SW_MAXIMIZE = 3
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
        GetIEHandler("ECI Call").document.getElementByID("webapp_ui_button_button_3").Click
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
        
        .document.getElementByID("_QRY_eci_call_in.remedy").innerText = inc
        .document.getElementByID("webapp_ui_button_button_8").Click
    End With
End Function

Private Function SyncreonCheckDeliveryStatus(ByRef statusVal As String) As String
    On Error GoTo ErrorHandler:
    
    CheckRemedySynUpsStatus
    
    ' Syncreon
    With GetIEHandler("ECI Call In")
        ' Waiting for the page to load
        While .ReadyState <> 4 Or .Busy: DoEvents: Wend
        CheckRemedySynUpsStatus
        
        delStatus = .document.getElementByID("jrepapp_view_formauto_boxes_sqltable_sqltable_1_0_carr_statdesc_cell").innerText
        
        If delStatus = "Delivered" Then
            ' Extracts the Tracking Number
            trackingNumber = .document.getElementByID("jrepapp_view_formauto_boxes_sqltable_sqltable_1_0_tracknbr_cell").innerText
            trackingNumber = Mid(trackingNumber, InStr(trackingNumber, ">") + 1, 18)
            
            ' UPS
            With GetIEHandler("UPS")
                ' Searches the Tracking Number
                .Navigate "https://www.ups.com/track?loc=en_US&tracknum=" & trackingNumber
                
                While .ReadyState <> 4 Or .Busy Or Not (TypeName(.document.getElementByID("toTitle")) = "Null"): DoEvents: Wend
                CheckRemedySynUpsStatus
                
                ' Retrieves the date
                delDate = .document.getElementByID("stApp_deliveredDate").innerText
                ' Retrieves the hour (12h) and converts to 24h
                delHour = Mid(.document.getElementByID("stApp_eodDate").innerText, InStr(.document.getElementByID("stApp_eodDate").innerText, ":") - 2)
                delHour = Replace(delHour, ".", "")
                delHour = Format(delHour, "HH:mm")
                
                delDate = delDate & " " & delHour
                
                statusVal = delStatus
                SyncreonCheckDeliveryStatus = delDate
            End With
        ElseIf delStatus = "Entregado" Then
            CheckRemedySynUpsStatus
            
            delDate = .document.getElementByID("jrepapp_view_formauto_boxes_sqltable_sqltable_1_0_dt_rec_cell").innerText
            
            statusVal = delStatus
            SyncreonCheckDeliveryStatus = Left(delDate, Len(delDate) - 3)
        Else
        ' If is not delivered
            CheckRemedySynUpsStatus
            
            statusVal = delStatus
        End If
        Exit Function
        
ErrorHandler:
        ' First we check for the "NO EXISTE" syncreon comment
        If Not (.document.getElementByID("H__eci_call_in.comments_syncreon") Is Nothing) Then
            syncComment = .document.getElementByID("H__eci_call_in.comments_syncreon").innerText
            
            ' If the "NO EXISTE" syncreon comment exists
            If Len(syncComment) <> 0 Then
                statusVal = syncComment
            ' Otherwhise we check the delivery date
            ElseIf Len(delDate) = 0 Then
                statusVal = "-1"
            End If
            
            Exit Function
        ' Otherwhise we check the delivery date (used in case the "NO EXISTE" syncreon element does not exists
        ElseIf Len(delDate) = 0 Then
            statusVal = "-1"
        End If
    End With
End Function

' Because Remedy does not allow "external" queries, we bypass this feature by using HTTP callbacks
Private Function RemedyOpenINC(ByVal inc As String, ByRef Remedy)
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
        .document.getElementByID("arid_WIN_1_1000000881").Click
        
        ' Waits for the list to load
        While .ReadyState <> 4 Or .Busy: DoEvents: Wend
        Application.Wait (Now + TimeValue("0:00:5"))
        
        ' Retrieves and saves into the hidden input box the element coordinates (In Javascript)
        '.Document.parentWindow.execScript "try { var el = document.getElementsByClassName(""MenuTable"")(0).getElementsByClassName(""MenuTableBody"")(0).getElementsByClassName(""MenuEntryName"")(5); var _x = 0; var _y = 0; while( el && !isNaN( el.offsetLeft ) && !isNaN( el.offsetTop ) ) { _x += el.offsetLeft - el.scrollLeft; _y += el.offsetTop - el.scrollTop; el = el.offsetParent; } document.getElementById(""adi_marian_mutu"").value = _x + "","" + _y; } catch(err) { window.alert(""<PDARFID_AUTO_JS_NoFurthActClick_ERROR>\n\n"" + err); }", "Javascript"
        .document.parentWindow.execScript "try { var el = document.getElementsByClassName(""MenuTable"")(0).getElementsByClassName(""MenuTableBody"")(0).getElementsByClassName(""MenuEntryName"")(5);function getPosition(el){var xPos = 0;var yPos = 0;while (el) {if (el.tagName == ""BODY"") {var xScroll = el.scrollLeft || document.documentElement.scrollLeft;var yScroll = el.scrollTop || document.documentElement.scrollTop;xPos += (el.offsetLeft - xScroll + el.clientLeft);yPos += (el.offsetTop - yScroll + el.clientTop);} else {xPos += (el.offsetLeft - el.scrollLeft + el.clientLeft);yPos += (el.offsetTop - el.scrollTop + el.clientTop);}el = el.offsetParent;}return{x: xPos,y: yPos};} var _p = getPosition(el); document.getElementById(""adi_marian_mutu"").value = _p.x + "","" + _p.y; } catch(err) { window.alert(""<PDARFID_AUTO_JS_NoFurthActClick_ERROR>\n\n"" + err); }", "Javascript"
        
        ' Calculates the correct element coordinates
        coord_str = .document.getElementByID("adi_marian_mutu").Value
        x = .Left + (Left(coord_str, InStr(coord_str, ",") - 1) + 50)
        y = .Top + (Right(coord_str, InStr(coord_str, ",") - 1) + 72)
        
        ' Saves the cursor coordinates
        GetCursorPos curBack
        ' Moves the cursor to the element coordinates
        SetCursorPos x, y
        ' Press and release the left mouse button
        mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
        mouse_event MOUSEEVENTF_LEFTUP, 0, 0, 0, 0
        
        ' Take the cursor to the last known coordinates
        SetCursorPos curBack.x, curBack.y
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
        
        remedyTktStatus = .document.getElementByID("arid_WIN_1_7").Value
        If remedyTktStatus <> "Resolved" And remedyTktStatus <> "Closed" Then
            ' Used to change and update the Status table
            Dim post As Object, evt As Object
            Set evt = .document.createEvent("keyboardevent")
            evt.initEvent "change", True, False

            ' Add the Delivery Date to the 'Notes' inputbox
            Set post = .document.getElementByID("arid_WIN_1_304247080")
                post.Value = "Delivered on: " & note
            post.dispatchEvent evt
            ' Checks the 'Public' checkbox
            .document.getElementByID("WIN_1_rc1id1000000761").Click
            ' Changes the Status dropbox menu to Resolved
            Set post = .document.getElementByID("arid_WIN_1_7")
                post.Value = "Resolved"
            post.dispatchEvent evt
            ' Changes the Status Reason dropbox menu to No Further Action Required
            RemedyProcessINC_statReasonNoFurthActReqClick
            
            ' Before saving the changes, checks if the Status Reason is "No Further Action Required"
            If .document.getElementByID("arid_WIN_1_1000000881").Value <> "No Further Action Required" Then
                MsgBox ("[PDA & RFID PROCESS]" & vbNewLine & vbNewLine & "Unable to correctly click on ""No Further Action Required"" item!" & vbNewLine & vbNewLine & " > Click ""Ok"" to stop the script and try again" & vbNewLine & vbNewLine & "> If the problem persists, please call Adi")
                'RemedyProcessINC = 0
                End
            End If
            
            ' Saves the changes
            .document.getElementByID("WIN_1_301614800").Click
            
            ' Checks if the tkt was successfully saved
            Application.Wait (Now + TimeValue("0:00:5"))
            
            If TypeName(.document.getElementByID("pbartable").getElementsByClassName("prompttext prompttexterr")(0)) <> "Nothing" Then
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
        r = .document.getElementByID("arid_WIN_1_1000000652").innerText
        
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


Public Sub PDARFIDAuto()
    ' Excel INC column header
    incColumn = "A"
    ' Excel Notes column header
    notesColumn = "F"
    ' Excel Service Request column header
    srColumn = "H"
    ' Excel Delivery Date column header
    delDateColumn = "I"
    
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
                ' If the remedy tkt is not successfully saved
                remResult = RemedyProcessINC(inc, delDate)
                If remResult = 0 Then
                    Range(notesColumn & Rng.Row) = "<remedy_could_not_resolve_tkt>"
                ElseIf remResult = 1 Then
                    Range(notesColumn & Rng.Row) = "<remedy_tkt_resolved>"
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
                If InStr(delStatus, "NO EXISTE") = 0 And InStr(delStatus, "refused") = 0 And InStr(delStatus, "sender") = 0 Then
                    ' If the Syncreon page is empty
                    If delStatus = "" Or delStatus = "null" Then
                        delStatus = "" '"<syncreon_empty_or_manual_check_required>"
                    Else
                    ' If this point is reached, this means that the device is still not delivered
                        delStatus = "<not_delivered>"
                    End If
                ' If the device was refused or returned to sender, we still must extract the Service Request number
                ElseIf InStr(delStatus, "refused") > 0 Or InStr(delStatus, "sender") > 0 Then
                    ' Used to refresh the Remedy tab handler
                    Dim IERemedy As Object
                    Set IERemedy = Nothing
                    ' Searches the INC on Remedy
                    RemedyOpenINC inc, IERemedy
                    ' Delay to allow the page elements to load
                    Application.Wait (Now + TimeValue("0:00:3"))
                    ' Extracts the SR
                    Range(srColumn & Rng.Row) = RemedyExtractSR(IERemedy)
                End If
                
                Range(notesColumn & Rng.Row) = delStatus
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
