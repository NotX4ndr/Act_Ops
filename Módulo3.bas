Attribute VB_Name = "Módulo3"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const BASE As String = "https://b2e-cloud.costa.it/B2EWeb/"
Private Const LOGIN_URL As String = BASE & "WebFrmLogin.aspx?ReturnUrl=%2fB2EWeb"
Private Const AGENCY_URL As String = BASE & "WebFrmAgency.aspx"
Private Const SEARCH_URL As String = BASE & "WebFrmBookingSearch.aspx?WithoutAgency=yes"
Private Const MAIN_URL As String = BASE & "WebFrmMain.aspx"
Private Const BKGINFO_URL As String = BASE & "WebFrmBkgInfo.aspx"

Private Const OP_COL As Long = 4
Private Const BKGDATE_COL As Long = 5
Private Const OPTDATE_COL As Long = 6
Private Const CRUISE_COL As Long = 7
Private Const STATUS_COL As Long = 8
Private Const PRICE_COL As Long = 9
Private Const HIST_USER_COL As Long = 2
Private Const HIST_USERDETAIL_COL As Long = 3
Private Const START_ROW As Long = 2

Private Const WAIT_BETWEEN_OPS As Long = 2

Private gSess As Object
Private gLastHTML As String
Private gProg As frmProgress

Public Sub ActualizarStatus_OPs()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.cells(ws.rows.Count, OP_COL).End(xlUp).Row
    If lastRow < START_ROW Then Exit Sub

    Dim startAt As Long
    startAt = AskStartRow(START_ROW, lastRow)
    If startAt = 0 Then Exit Sub

    Debug.Print Format$(Now, "dd/mm/yyyy hh:nn:ss") & " | ---- RUN START ----"
    Debug.Print Format$(Now, "dd/mm/yyyy hh:nn:ss") & " | Inicio desde fila " & startAt

    If gSess Is Nothing Then
        Set gSess = CreateObject("Scripting.Dictionary")
        gSess("cookie") = ""
        gSess("wsid") = ""
        gSess("logged") = False
    End If

    If Not EnsureLogin(gSess) Then
        MsgBox "No pude iniciar sesión en B2E. Revisa credenciales o acceso.", vbCritical
        Debug.Print Format$(Now, "dd/mm/yyyy hh:nn:ss") & " | Login fallido."
        Exit Sub
    End If

    Dim totalOps As Long, doneOps As Long
    totalOps = (lastRow - startAt + 1)
    doneOps = 0

    ProgressOpen "Actualizando OPs..."

    Application.ScreenUpdating = False
    On Error GoTo EH

    Dim r As Long, op As String, st As String, cruiseCode As String
    Dim optDate As String, bkgDate As String, priceVal As String
    Dim dOpt As Variant, dBkg As Variant, nPrice As Variant
    Dim ownerUserDetail As String
    Dim cancReason As String
    Dim histUser As String
    Dim histUserDetail As String
    Dim errNum As Long, errDesc As String

    For r = startAt To lastRow

        op = Trim$(CStr(ws.cells(r, OP_COL).Value))

        doneOps = doneOps + 1
        ProgressStep doneOps, totalOps, "Fila " & r & " | OP " & op

        If Len(op) > 0 Then

            LogLine "Fila " & r & " | OP=" & op & " | consultando..."

            gLastHTML = ""
            st = GetStatusByOP(op, gSess)
            cruiseCode = ParseCruiseCodeFromBooking(gLastHTML)
            bkgDate = ParseBkgDateFromBooking(gLastHTML)
            optDate = ParseOptDateFromBooking(gLastHTML)
            priceVal = ParseGrossValueFromBooking(gLastHTML)
            cancReason = ParseCancReasonFromBooking(gLastHTML)

            ownerUserDetail = GetOwnerUserDetailFromBooking(gLastHTML, gSess)
            ws.cells(r, 1).Value = ownerUserDetail
            
            DoEvents
            Sleep 700

            histUser = ""
            histUserDetail = ""

            If GetHistoryOptToBkdUsersFromBooking(op, gLastHTML, gSess, histUser, histUserDetail) Then
                ws.cells(r, HIST_USER_COL).Value = histUser
                ws.cells(r, HIST_USERDETAIL_COL).Value = histUserDetail
            Else
                ws.cells(r, HIST_USER_COL).ClearContents
                ws.cells(r, HIST_USERDETAIL_COL).ClearContents
            End If

            ClearCellComment ws.cells(r, STATUS_COL)

            If Len(st) = 0 Then
                LogLine "Fila " & r & " | OP=" & op & " | STATUS=VACIO"
                ws.cells(r, STATUS_COL).ClearContents
            Else
                ws.cells(r, STATUS_COL).Value = st
                ws.cells(r, CRUISE_COL).Value = cruiseCode

                If Len(bkgDate) > 0 Then
                    dBkg = DateFromDMY(bkgDate)
                    If IsDate(dBkg) Then
                        ws.cells(r, BKGDATE_COL).Value = dBkg
                    Else
                        ws.cells(r, BKGDATE_COL).Value = bkgDate
                    End If
                Else
                    ws.cells(r, BKGDATE_COL).ClearContents
                End If
                ws.cells(r, BKGDATE_COL).NumberFormat = "dd/mm/yyyy"

                If Len(optDate) > 0 Then
                    dOpt = DateFromDMY(optDate)
                    If IsDate(dOpt) Then
                        ws.cells(r, OPTDATE_COL).Value = dOpt
                    Else
                        ws.cells(r, OPTDATE_COL).Value = optDate
                    End If
                Else
                    ws.cells(r, OPTDATE_COL).ClearContents
                End If
                ws.cells(r, OPTDATE_COL).NumberFormat = "dd/mm/yyyy"

                If Len(priceVal) > 0 Then
                    nPrice = NormalizarNumero(priceVal)
                    If Len(CStr(nPrice)) > 0 Then
                        ws.cells(r, PRICE_COL).Value = CDbl(nPrice)
                    Else
                        ws.cells(r, PRICE_COL).ClearContents
                    End If
                Else
                    ws.cells(r, PRICE_COL).ClearContents
                End If
                ws.cells(r, PRICE_COL).NumberFormat = "#,##0.00 [$€-es-ES]"

                If (InStr(1, UCase$(st), "CXL", vbBinaryCompare) > 0 Or InStr(1, UCase$(st), "CX", vbBinaryCompare) > 0) Then
                    If Len(Trim$(cancReason)) > 0 Then
                        SetCellComment ws.cells(r, STATUS_COL), cancReason
                    End If
                End If

                LogLine "Fila " & r & " | OP=" & op & " | STATUS=" & st & " | CRUISE=" & cruiseCode & _
                        " | BKGDATE=" & bkgDate & " | OPTDATE=" & optDate & " | PRICE=" & priceVal & _
                        " | CANC_REASON=" & cancReason & " | OWNER=" & ownerUserDetail & _
                        " | HIST_USER=" & histUser & " | HIST_USERDETAIL=" & histUserDetail
            End If

            WaitSeconds WAIT_BETWEEN_OPS

        Else
            ClearCellComment ws.cells(r, STATUS_COL)
            ws.cells(r, HIST_USER_COL).ClearContents
            ws.cells(r, HIST_USERDETAIL_COL).ClearContents
        End If

    Next r

    Application.ScreenUpdating = True
    LogoutB2E gSess
    ProgressClose

    Debug.Print Format$(Now, "dd/mm/yyyy hh:nn:ss") & " | ---- RUN END ----"
    MsgBox "Listo. Status, Cruise, Fecha y Precio actualizados. Sesión cerrada.", vbInformation
    Exit Sub

EH:
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not gSess Is Nothing Then
        If gSess.Exists("logged") Then
            If gSess("logged") = True Then
                LogoutB2E gSess
            End If
        End If
    End If
    Application.ScreenUpdating = True
    ProgressClose
    On Error GoTo 0

    MsgBox "Error: " & errNum & " - " & errDesc, vbCritical
End Sub

Public Sub B2E_ClearSession()
    On Error Resume Next

    If Not gSess Is Nothing Then
        If gSess.Exists("logged") Then
            If gSess("logged") = True Then
                LogoutB2E gSess
            End If
        End If
    End If

    On Error GoTo 0
    Set gSess = Nothing
    MsgBox "Sesión limpiada y cerrada.", vbInformation
End Sub

Private Function EnsureLogin(ByVal sess As Object) As Boolean

    If sess.Exists("logged") Then
        If sess("logged") = True Then
            If Len(CStr(sess("cookie"))) > 0 And InStr(1, CStr(sess("cookie")), ".B2E-AUTH", vbTextCompare) > 0 Then
                EnsureLogin = True
                Exit Function
            End If
        End If
    End If

    Dim u As String, p As String

    Dim f As frmLogin
    Set f = New frmLogin

    f.txtUser.Value = ""
    f.txtPass.Value = ""
    f.Cancelled = True

    f.Show vbModal

    If f.Cancelled Then
        Unload f
        EnsureLogin = False
        Exit Function
    End If

    u = Trim$(CStr(f.txtUser.Value))
    p = CStr(f.txtPass.Value)

    Unload f

    If Len(u) = 0 Or Len(p) = 0 Then
        EnsureLogin = False
        Exit Function
    End If

    sess("cookie") = ""
    sess("wsid") = ""
    sess("logged") = False

    If Not LoginB2E(u, p, sess) Then
        On Error Resume Next
        If InStr(1, CStr(sess("cookie")), ".B2E-AUTH", vbTextCompare) > 0 Then
            sess("logged") = True
            LogoutB2E sess
        Else
            sess("cookie") = ""
            sess("wsid") = ""
            sess("logged") = False
        End If
        On Error GoTo 0
        Exit Function
    End If

    sess("logged") = True
    LogLine "Login OK. WSID=" & sess("wsid")
    EnsureLogin = True

End Function

Public Function LoginB2E(ByVal u As String, ByVal p As String, ByVal sess As Object) As Boolean

    Dim html As String
    html = HttpGet(LOGIN_URL, sess)

    Dim vs As String, vsg As String
    vs = HtmlHidden(html, "__VIEWSTATE")
    vsg = HtmlHidden(html, "__VIEWSTATEGENERATOR")

    If Len(vs) = 0 Or Len(vsg) = 0 Then Exit Function

    Dim body As String
    body = "winCurrentOpener=notvalid" _
        & "&defaultButton=sibLogin" _
        & "&__VIEWSTATE=" & UrlEnc(vs) _
        & "&__VIEWSTATEGENERATOR=" & UrlEnc(vsg) _
        & "&txtUsername=" & UrlEnc(u) _
        & "&txtPassword=" & UrlEnc(p) _
        & "&sibLogin.x=10&sibLogin.y=10"

    Dim resp As String
    resp = HttpPost(LOGIN_URL, body, sess)

    If InStr(1, CStr(sess("cookie")), ".B2E-AUTH", vbTextCompare) = 0 Then Exit Function

    Dim wsid As String
    wsid = ExtractWsid(resp)

    If Len(wsid) = 0 Then
        Dim agencyHtml As String
        agencyHtml = HttpGet(AGENCY_URL, sess)
        wsid = ExtractWsid(agencyHtml)
    End If

    sess("wsid") = wsid
    LoginB2E = True

End Function

Private Sub LogoutB2E(ByVal sess As Object)

    Dim html As String
    html = HttpGet(MAIN_URL, sess)
    If Len(html) = 0 Then GoTo CleanOnly

    Dim vs As String, vsg As String
    vs = HtmlHidden(html, "__VIEWSTATE")
    vsg = HtmlHidden(html, "__VIEWSTATEGENERATOR")

    If Len(vs) > 0 Then
        Dim body As String
        body = "__EVENTTARGET=" & UrlEnc("ddmMainMenu$ctl07") _
            & "&__EVENTARGUMENT=" _
            & "&__VIEWSTATE=" & UrlEnc(vs) _
            & "&__VIEWSTATEGENERATOR=" & UrlEnc(vsg)

        HttpPost MAIN_URL, body, sess
    End If

CleanOnly:
    sess("cookie") = ""
    sess("wsid") = ""
    sess("logged") = False

    LogLine "Sesión cerrada correctamente."

End Sub

Private Function GetStatusByOP(ByVal op As String, ByVal sess As Object) As String

    gLastHTML = ""
    
    Dim url As String
    url = SEARCH_URL

    If Len(CStr(sess("wsid"))) > 0 Then
        If InStr(1, url, "Wsid=", vbTextCompare) = 0 Then
            url = url & "&Wsid=" & UrlEnc(CStr(sess("wsid")))
        End If
    End If

    Dim html As String
    html = HttpGet(AddNoCacheToUrl(url), sess)
    If Len(html) = 0 Then Exit Function

    Dim vs As String, vsg As String
    vs = HtmlHidden(html, "__VIEWSTATE")
    vsg = HtmlHidden(html, "__VIEWSTATEGENERATOR")

    If Len(vs) = 0 Or Len(vsg) = 0 Then Exit Function

    Dim body As String
    body = "defaultButton=sibSearch" _
        & "&__EVENTTARGET=&__EVENTARGUMENT=" _
        & "&__VIEWSTATE=" & UrlEnc(vs) _
        & "&__VIEWSTATEGENERATOR=" & UrlEnc(vsg) _
        & "&txtBkgNum=" & UrlEnc(op) _
        & "&sibSearch.x=10&sibSearch.y=10"

    Dim resp As String
    resp = HttpPost(url, body, sess)
    If Len(resp) = 0 Then Exit Function

    ' --- NUEVO: asegurar que realmente estamos en el booking del OP solicitado ---
    resp = EnsureExactBookingPageFromSearch(op, resp, url, sess)

    If Len(resp) = 0 Then Exit Function

    ' Validación adicional (evita contexto reciclado / primer resultado)
    If Not BookingHtmlLooksLikeRequestedOP(resp, op) Then
        LogLine "WARN | OP=" & op & " | booking HTML no coincide con OP solicitado"
    End If

    gLastHTML = resp
    GetStatusByOP = ParseStatusFromBooking(resp)

End Function

Private Function ExtractExactBkgInfoUrlForOp(ByVal op As String, ByVal html As String) As String

    Dim reTr As Object, trs As Object, i As Long
    Set reTr = CreateObject("VBScript.RegExp")
    reTr.IgnoreCase = True
    reTr.Global = True
    reTr.Multiline = True
    reTr.Pattern = "<tr\b[^>]*>([\s\S]*?)</tr>"

    If Not reTr.Test(html) Then Exit Function
    Set trs = reTr.Execute(html)

    Dim opU As String
    opU = UCase$(Trim$(op))

    For i = 0 To trs.Count - 1
        Dim tr As String
        tr = CStr(trs(i).Value)

        If InStr(1, tr, "<td", vbTextCompare) = 0 Then GoTo NextRow
        If Not RowContainsExactOP(tr, opU) Then GoTo NextRow

        Dim u As String

        ' href directo con comillas dobles
        u = Re1(tr, "href=""([^""]*WebFrmBkgInfo\.aspx[^""]*)""")
        If Len(u) > 0 Then
            ExtractExactBkgInfoUrlForOp = HtmlDecode(u)
            Exit Function
        End If

        ' href directo con comillas simples
        u = Re1(tr, "href='([^']*WebFrmBkgInfo\.aspx[^']*)'")
        If Len(u) > 0 Then
            ExtractExactBkgInfoUrlForOp = HtmlDecode(u)
            Exit Function
        End If

        ' fallback: url metida en javascript dentro de la fila
        u = Re1(tr, "['""]([^'""]*WebFrmBkgInfo\.aspx[^'""]*)['""]")
        If Len(u) > 0 Then
            If InStr(1, u, "WebFrmBkgInfo.aspx", vbTextCompare) > 0 Then
                ExtractExactBkgInfoUrlForOp = HtmlDecode(u)
                Exit Function
            End If
        End If

NextRow:
    Next i

End Function

Private Function EnsureExactBookingPageFromSearch(ByVal op As String, ByVal searchRespHtml As String, ByVal searchUrl As String, ByVal sess As Object) As String

    Dim html As String
    html = searchRespHtml

    ' Si ya es booking page (tiene label de status / booking widget), devolver tal cual
    If IsBookingPageHtml(html) Then
        EnsureExactBookingPageFromSearch = html
        Exit Function
    End If

    ' Si hay grid/lista de resultados, intentar abrir fila exacta por OP
    Dim nextHtml As String
    nextHtml = OpenExactSearchResultRow(op, html, searchUrl, sess)

    If Len(nextHtml) > 0 Then
        EnsureExactBookingPageFromSearch = nextHtml
        Exit Function
    End If

    ' Fallback: intentar abrir por enlace directo al BkgInfo si aparece en el HTML de resultados
    Dim infoUrl As String
    infoUrl = ExtractExactBkgInfoUrlForOp(op, html)
    If Len(infoUrl) > 0 Then
        infoUrl = ResolveUrl(searchUrl, infoUrl)
        nextHtml = HttpGet(AddNoCacheToUrl(infoUrl), sess)
        If Len(nextHtml) > 0 Then
            EnsureExactBookingPageFromSearch = nextHtml
            Exit Function
        End If
    End If

    ' Último fallback: devolver lo recibido
    EnsureExactBookingPageFromSearch = html

End Function
Private Function OpenExactSearchResultRow(ByVal op As String, ByVal html As String, ByVal baseUrl As String, ByVal sess As Object) As String

    Dim rowHtml As String
    rowHtml = FindSearchResultRowByOP(html, op)
    If Len(rowHtml) = 0 Then
        LogLine "SEARCH | OP=" & op & " | no encontré fila exacta en grid"
        Exit Function
    End If

    LogLine "SEARCH | OP=" & op & " | fila exacta encontrada"

    ' 1) Intentar href directo (WebFrmBkgInfo.aspx?...)
    Dim href As String
    href = Re1(rowHtml, "href=""([^""]*WebFrmBkgInfo\.aspx[^""]*)""")
    If Len(href) = 0 Then
        href = Re1(rowHtml, "href='([^']*WebFrmBkgInfo\.aspx[^']*)'")
    End If

    If Len(href) > 0 Then
        href = HtmlDecode(href)
        OpenExactSearchResultRow = HttpGet(AddNoCacheToUrl(ResolveUrl(baseUrl, href)), sess)
        Exit Function
    End If

    ' 2) Intentar __doPostBack('target','arg')
    Dim evTarget As String, evArg As String
    If ExtractDoPostBackFromRow(rowHtml, evTarget, evArg) Then

        Dim postUrl As String
        postUrl = ExtractFormActionUrl(html)
        If Len(postUrl) = 0 Then
            postUrl = baseUrl
        Else
            postUrl = ResolveUrl(baseUrl, postUrl)
        End If

        Dim body As String
        body = BuildPostBackBodyFromHtml(html, "")
        body = UpsertPostField(body, "__EVENTTARGET", evTarget)
        body = UpsertPostField(body, "__EVENTARGUMENT", evArg)

        ' quitar clicks de imagen heredados si existen
        body = RemovePostField(body, "sibSearch.x")
        body = RemovePostField(body, "sibSearch.y")

        OpenExactSearchResultRow = HttpPost(postUrl, body, sess)
        Exit Function
    End If

    ' 3) Intentar javascript:Open... con URL dentro
    Dim jsUrl As String
    jsUrl = Re1(rowHtml, "['""]([^'""]*WebFrmBkgInfo\.aspx[^'""]*)['""]")
    If Len(jsUrl) > 0 Then
        jsUrl = HtmlDecode(jsUrl)
        OpenExactSearchResultRow = HttpGet(AddNoCacheToUrl(ResolveUrl(baseUrl, jsUrl)), sess)
        Exit Function
    End If

End Function
Private Function FindSearchResultRowByOP(ByVal html As String, ByVal op As String) As String

    Dim reTr As Object, ms As Object, i As Long
    Set reTr = CreateObject("VBScript.RegExp")
    reTr.IgnoreCase = True
    reTr.Global = True
    reTr.Multiline = True
    reTr.Pattern = "<tr\b[^>]*>([\s\S]*?)</tr>"

    If Not reTr.Test(html) Then Exit Function
    Set ms = reTr.Execute(html)

    Dim opNorm As String
    opNorm = UCase$(Trim$(op))

    For i = 0 To ms.Count - 1
        Dim tr As String
        tr = CStr(ms(i).Value)

        If InStr(1, tr, "<td", vbTextCompare) = 0 Then GoTo NextRow

        ' Debe contener el OP como texto exacto en una celda o link
        If RowContainsExactOP(tr, opNorm) Then
            FindSearchResultRowByOP = tr
            Exit Function
        End If

NextRow:
    Next i

End Function
Private Function RowContainsExactOP(ByVal rowHtml As String, ByVal opUpper As String) As Boolean

    Dim cells() As String
    cells = GetTdTextsFromRow(rowHtml)

    Dim i As Long
    For i = 0 To UBoundSafe(cells)
        Dim t As String
        t = UCase$(Trim$(cells(i)))
        If Len(t) = 0 Then GoTo NextCell

        ' exacto
        If StrComp(t, opUpper, vbBinaryCompare) = 0 Then
            RowContainsExactOP = True
            Exit Function
        End If

        ' token delimitado (evita falsos positivos tipo OP123 dentro de OP1234)
        If InStr(1, " " & Replace$(Replace$(t, "-", " "), "/", " ") & " ", " " & opUpper & " ", vbBinaryCompare) > 0 Then
            RowContainsExactOP = True
            Exit Function
        End If

NextCell:
    Next i

End Function
Private Function ExtractDoPostBackFromRow(ByVal rowHtml As String, ByRef evTarget As String, ByRef evArg As String) As Boolean

    evTarget = ""
    evArg = ""

    Dim r As Object, m As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.Multiline = True
    r.Pattern = "__doPostBack\(\s*'([^']*)'\s*,\s*'([^']*)'\s*\)"

    If r.Test(rowHtml) Then
        Set m = r.Execute(rowHtml)(0)
        evTarget = CStr(m.SubMatches(0))
        evArg = CStr(m.SubMatches(1))
        ExtractDoPostBackFromRow = (Len(evTarget) > 0)
        Exit Function
    End If

    r.Pattern = "__doPostBack\(\s*""([^""]*)""\s*,\s*""([^""]*)""\s*\)"
    If r.Test(rowHtml) Then
        Set m = r.Execute(rowHtml)(0)
        evTarget = CStr(m.SubMatches(0))
        evArg = CStr(m.SubMatches(1))
        ExtractDoPostBackFromRow = (Len(evTarget) > 0)
    End If

End Function
Private Function BookingHtmlLooksLikeRequestedOP(ByVal html As String, ByVal op As String) As Boolean

    Dim t As String
    t = UCase$(html)

    Dim opU As String
    opU = UCase$(Trim$(op))

    If Len(opU) = 0 Then Exit Function

    ' Búsqueda simple pero útil para evitar contexto del booking previo
    If InStr(1, t, UCase$(">" & op & "<"), vbBinaryCompare) > 0 Then
        BookingHtmlLooksLikeRequestedOP = True
        Exit Function
    End If

    If InStr(1, t, "VALUE=""" & opU & """", vbBinaryCompare) > 0 Then
        BookingHtmlLooksLikeRequestedOP = True
        Exit Function
    End If

    If InStr(1, t, opU, vbBinaryCompare) > 0 And IsBookingPageHtml(html) Then
        BookingHtmlLooksLikeRequestedOP = True
        Exit Function
    End If
End Function
Private Function UpsertPostField(ByVal body As String, ByVal fieldName As String, ByVal fieldValue As String) As String
    Dim cleaned As String
    cleaned = RemovePostField(body, fieldName)
    If Len(cleaned) > 0 Then cleaned = cleaned & "&"
    cleaned = cleaned & UrlEnc(fieldName) & "=" & UrlEnc(fieldValue)
    UpsertPostField = cleaned
End Function

Private Function RemovePostField(ByVal body As String, ByVal fieldName As String) As String
    If Len(body) = 0 Then
        RemovePostField = ""
        Exit Function
    End If

    Dim parts() As String
    parts = Split(body, "&")

    Dim i As Long
    Dim out As String
    Dim keyEnc As String
    keyEnc = LCase$(UrlEnc(fieldName))

    For i = LBound(parts) To UBound(parts)
        Dim p As String
        p = parts(i)

        Dim k As String
        If InStr(1, p, "=", vbBinaryCompare) > 0 Then
            k = LCase$(Left$(p, InStr(1, p, "=", vbBinaryCompare) - 1))
        Else
            k = LCase$(p)
        End If

        If k <> keyEnc Then
            If Len(out) > 0 Then out = out & "&"
            out = out & p
        End If
    Next i

    RemovePostField = out
End Function

Private Function IsBookingPageHtml(ByVal html As String) As Boolean
    If Len(html) = 0 Then Exit Function

    If InStr(1, html, "wucBooking_lblStatusValue", vbTextCompare) > 0 Then
        IsBookingPageHtml = True
        Exit Function
    End If

    If InStr(1, html, "WebFrmBkgInfo.aspx", vbTextCompare) > 0 And _
       InStr(1, html, "wucBooking_", vbTextCompare) > 0 Then
        IsBookingPageHtml = True
        Exit Function
    End If
End Function

Private Function GetOwnerUserDetailFromBooking(ByVal html As String, ByVal sess As Object) As String

    Dim infoUrl As String
    infoUrl = ExtractInfoUrl(html)

    If Len(infoUrl) = 0 Then
        Dim bookingWsid As String
        bookingWsid = ExtractBookingWsid(html)

        Dim contactName As String
        contactName = ParseContactNameFromBooking(html)

        If Len(bookingWsid) > 0 And Len(contactName) > 0 Then
            infoUrl = BKGINFO_URL & "?Wsid=" & UrlEnc(bookingWsid) & "&ContactName=" & UrlEnc(contactName)
        End If
    End If

    If Len(infoUrl) = 0 Then Exit Function

    infoUrl = ResolveUrl(SEARCH_URL, infoUrl)

    Dim infoHtml As String
    infoHtml = HttpGet(infoUrl, sess)
    If Len(infoHtml) = 0 Then Exit Function

    GetOwnerUserDetailFromBooking = ParseOwnerUserId(infoHtml)
    If Len(GetOwnerUserDetailFromBooking) = 0 Then GetOwnerUserDetailFromBooking = ParseOwnerUserDetail(infoHtml)

End Function

Private Function GetHistoryOptToBkdUsersFromBooking(ByVal op As String, ByVal html As String, ByVal sess As Object, ByRef histUser As String, ByRef histUserDetail As String) As Boolean
    Static sPrevCtxSig As String
    Static sPrevOp As String

    Const MAX_TRIES As Long = 3

    histUser = ""
    histUserDetail = ""

    Dim historyUrl As String
    Dim dialogHtml As String
    Dim bookingWsid As String

    Dim histHtml As String
    Dim histHtml2 As String
    Dim sigGet As String

    Dim postUrl As String
    Dim body As String
    Dim resp As String
    Dim sigPost As String

    Dim candidateUser As String
    Dim candidateReq As String

    Dim attempt As Long
    Dim suspiciousGet As Boolean
    Dim suspiciousPost As Boolean

    historyUrl = ""
    dialogHtml = ""
    bookingWsid = ""
    histHtml = ""
    histHtml2 = ""
    sigGet = ""
    postUrl = ""
    body = ""
    resp = ""
    sigPost = ""
    candidateUser = ""
    candidateReq = ""

    bookingWsid = ExtractBookingWsid(html)

    ' 1) Intentar abrir el diálogo de History desde el booking actual
    dialogHtml = OpenHistoryDialogFromBooking(html, sess)
    If Len(dialogHtml) > 0 Then
        historyUrl = ExtractHistoryUrl(dialogHtml)
    End If

    ' 2) Fallback: URL directa de HistorySelection por Wsid (contexto de sesión)
    If Len(historyUrl) = 0 Then
        If Len(bookingWsid) = 0 Then Exit Function
        historyUrl = "WebFrmHistorySelection.aspx?Wsid=" & UrlEnc(bookingWsid)
    End If

    ' 3) Último recurso: intentar extraerla del html original
    If Len(historyUrl) = 0 Then
        historyUrl = ExtractHistoryUrl(html)
    End If

    If Len(historyUrl) = 0 Then Exit Function

    historyUrl = ResolveUrl(SEARCH_URL, historyUrl)

    LogLine "HIST | OP=" & op & " | WSID_BOOKING=" & bookingWsid
    LogLine "HIST | OP=" & op & " | historyUrl=" & historyUrl

    For attempt = 1 To MAX_TRIES

        ' ---- GET #1 + GET #2 (F5 lógico sobre HistorySelection) ----
        histHtml = HttpGet(AddNoCacheToUrl(historyUrl), sess)
        If Len(histHtml) = 0 Then GoTo NextAttempt
        If Not HistoryHtmlMatchesCurrentBooking(histHtml, bookingWsid, op) Then
        LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | GET no coincide con booking actual (posible cache/contexto)"
        End If

        DoEvents
        Sleep 500

        histHtml2 = HttpGet(AddNoCacheToUrl(historyUrl), sess)
        If Len(histHtml2) > 0 Then histHtml = histHtml2

        sigGet = GetHistoryContextSignature(histHtml)

        LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | histHtml_len=" & Len(histHtml)
        LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | CTX_GET=" & sigGet

        suspiciousGet = False
        If Len(sigGet) > 0 And Len(sPrevCtxSig) > 0 Then
            If StrComp(sigGet, sPrevCtxSig, vbTextCompare) = 0 And StrComp(op, sPrevOp, vbTextCompare) <> 0 Then
                suspiciousGet = True
                LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | GET contexto reciclado (igual al OP previo " & sPrevOp & ")"
            End If
        End If

        ' Intentar parsear directamente el GET refrescado
        candidateUser = ""
        candidateReq = ""
        If ParseHistoryOptToBkdUsers_Strict(histHtml, candidateUser, candidateReq) Then
            If Not suspiciousGet Then
                histUser = candidateUser
                histUserDetail = candidateReq

                GetHistoryOptToBkdUsersFromBooking = True

                If Len(sigGet) > 0 Then sPrevCtxSig = sigGet
                sPrevOp = op

                LogLine "HIST | OP=" & op & " | FOUND(GET) | USER=" & histUser & " | REQ=" & histUserDetail
                Exit Function
            Else
                LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | GET parseado pero descartado por contexto reciclado"
            End If
        End If

        ' ---- Fallback POST (sibSearch) ----
        postUrl = ExtractFormActionUrl(histHtml)
        If Len(postUrl) = 0 Then
            postUrl = historyUrl
        Else
            postUrl = ResolveUrl(historyUrl, postUrl)
        End If

        body = BuildPostBackBodyFromHtml(histHtml, "sibSearch")
        If Len(body) = 0 Then
            body = BuildImageClickBodyByPostName(histHtml, "sibSearch")
        End If

        ' ddlBkgItems tiene disabled="disabled" y no se incluye en el POST automaticamente.
        ' Lo inyectamos manualmente con el valor selected del GET.
        Dim ddlVal As String
        ddlVal = Re1(histHtml, "<option\b[^>]*selected=""selected""[^>]*value=""(\d+)""")
        If Len(ddlVal) = 0 Then
            ddlVal = Re1(histHtml, "<option\b[^>]*value=""(\d+)""[^>]*selected=""selected""")
        End If
        If Len(ddlVal) > 0 Then
            body = UpsertPostField(body, "ddlBkgItems", ddlVal)
            LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | ddlBkgItems=" & ddlVal
        Else
            LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | ddlBkgItems no encontrado en histHtml"
        End If

        If Len(body) = 0 Then
            LogLine "HIST | OP=" & op & " | BODY_PRE_POST=" & Left$(body, 800)
            resp = HttpPost(postUrl, body, sess)
        End If

        If Len(resp) > 0 Then
            sigPost = GetHistoryContextSignature(resp)
            LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | resp_len=" & Len(resp)
            LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | CTX_POST=" & sigPost
        End If

        suspiciousPost = False
        If Len(sigPost) > 0 And Len(sPrevCtxSig) > 0 Then
            If StrComp(sigPost, sPrevCtxSig, vbTextCompare) = 0 And StrComp(op, sPrevOp, vbTextCompare) <> 0 Then
                suspiciousPost = True
                LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | POST contexto reciclado (igual al OP previo " & sPrevOp & ")"
            End If
        End If

        If Len(resp) > 0 Then
            candidateUser = ""
            candidateReq = ""
            If ParseHistoryOptToBkdUsers_Strict(resp, candidateUser, candidateReq) Then
                If Not suspiciousPost Then
                    histUser = candidateUser
                    histUserDetail = candidateReq

                    GetHistoryOptToBkdUsersFromBooking = True

                    If Len(sigPost) > 0 Then
                        sPrevCtxSig = sigPost
                    ElseIf Len(sigGet) > 0 Then
                        sPrevCtxSig = sigGet
                    End If
                    sPrevOp = op

                    LogLine "HIST | OP=" & op & " | FOUND(POST) | USER=" & histUser & " | REQ=" & histUserDetail
                    Exit Function
                Else
                    LogLine "HIST | OP=" & op & " | TRY=" & attempt & " | POST parseado pero descartado por contexto reciclado"
                End If
            End If
        End If

NextAttempt:
        If attempt < MAX_TRIES Then
            DoEvents
            Sleep 900
        End If
    Next attempt

    LogLine "HIST | OP=" & op & " | FOUND=False"
End Function

Private Function HistoryHtmlMatchesCurrentBooking(ByVal html As String, ByVal bookingWsid As String, ByVal op As String) As Boolean
    If Len(html) = 0 Then Exit Function

    If Len(bookingWsid) > 0 Then
        If InStr(1, html, bookingWsid, vbTextCompare) > 0 Then
            HistoryHtmlMatchesCurrentBooking = True
            Exit Function
        End If
    End If

    If Len(op) > 0 Then
        If InStr(1, UCase$(html), UCase$(op), vbBinaryCompare) > 0 Then
            HistoryHtmlMatchesCurrentBooking = True
            Exit Function
        End If
    End If

    ' Si no encontramos marcador, no invalidamos duro (algunas pantallas history no muestran OP)
    HistoryHtmlMatchesCurrentBooking = True
End Function

Private Function ParseHistorySelectedBkgItemValue(ByVal html As String) As String
    Dim blk As String, r As Object, m As Object
    blk = Re1(html, "<select\b[^>]*id=""ddlBkgItems""[^>]*>([\s\S]*?)</select>")
    If Len(blk) = 0 Then Exit Function

    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.Multiline = True
    r.Pattern = "<option\b[^>]*selected(?:\s*=\s*""selected"")?[^>]*value=""([^""]+)"""
    If r.Test(blk) Then
        Set m = r.Execute(blk)(0)
        ParseHistorySelectedBkgItemValue = Trim$(m.SubMatches(0))
    End If
End Function

Private Function ParseHistoryOptToBkdUsers_Strict(ByVal html As String, ByRef histUser As String, ByRef histUserDetail As String) As Boolean
    histUser = ""
    histUserDetail = ""

    If Len(html) = 0 Then Exit Function
    If InStr(1, html, "dgrHistoryList", vbTextCompare) = 0 Then Exit Function

    Dim tblHtml As String
    tblHtml = ExtractTableById(html, "dgrHistoryList")
    If Len(tblHtml) = 0 Then Exit Function

    Dim reTr As Object
    Set reTr = CreateObject("VBScript.RegExp")
    reTr.Global = True
    reTr.IgnoreCase = True
    reTr.Multiline = True
    reTr.Pattern = "<tr\b[^>]*>([\s\S]*?)</tr>"

    Dim trMatches As Object
    Set trMatches = reTr.Execute(tblHtml)

    Dim i As Long
    For i = 0 To trMatches.Count - 1
        Dim rowHtml As String
        rowHtml = trMatches(i).SubMatches(0)

        If InStr(1, rowHtml, "<td", vbTextCompare) = 0 Then GoTo NextRow

        Dim c() As String
        c = GetTdTextsFromRow(rowHtml)

        If UBoundSafe(c) < 6 Then GoTo NextRow

        Dim actionType As String
        Dim actionDesc As String
        Dim reqBy As String
        Dim actWho As String

        actionType = SafeArr(c, 2)
        actionDesc = SafeArr(c, 4)
        reqBy = SafeArr(c, 5)
        actWho = SafeArr(c, 6)

        If (InStr(1, actionType, "STS - CHG", vbTextCompare) > 0 Or InStr(1, actionType, "STATUS", vbTextCompare) > 0) _
           And IsHistoryOptToBkdEventText(actionDesc) Then

            histUserDetail = reqBy
            histUser = actWho

            If Len(histUser) = 0 Then histUser = reqBy
            If Len(histUserDetail) = 0 Then histUserDetail = reqBy

            ParseHistoryOptToBkdUsers_Strict = (Len(histUser) > 0 Or Len(histUserDetail) > 0)
            Exit Function
        End If

NextRow:
    Next i
End Function

Private Function ParseHistorySelectedBkgItemText(ByVal html As String) As String
    Dim blk As String, r As Object, m As Object
    blk = Re1(html, "<select\b[^>]*id=""ddlBkgItems""[^>]*>([\s\S]*?)</select>")
    If Len(blk) = 0 Then Exit Function

    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.Multiline = True
    r.Pattern = "<option\b[^>]*selected(?:\s*=\s*""selected"")?[^>]*>([\s\S]*?)</option>"
    If r.Test(blk) Then
        Set m = r.Execute(blk)(0)
        ParseHistorySelectedBkgItemText = CleanHtmlCellText(CStr(m.SubMatches(0)))
    End If
End Function

Private Function GetHistoryGridRowCount(ByVal html As String) As Long
    Dim tbl As String, r As Object, ms As Object, i As Long, n As Long
    tbl = ExtractTableById(html, "dgrHistoryList")
    If Len(tbl) = 0 Then Exit Function

    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = True
    r.Multiline = True
    r.Pattern = "<tr\b[^>]*>([\s\S]*?)</tr>"

    If r.Test(tbl) Then
        Set ms = r.Execute(tbl)
        For i = 0 To ms.Count - 1
            If InStr(1, ms(i).SubMatches(0), "<td", vbTextCompare) > 0 Then
                n = n + 1
            End If
        Next i
    End If

    GetHistoryGridRowCount = n
End Function

Private Function GetHistoryContextSignature(ByVal html As String) As String
    Dim v As String, t As String, n As Long
    v = ParseHistorySelectedBkgItemValue(html)
    t = ParseHistorySelectedBkgItemText(html)
    n = GetHistoryGridRowCount(html)

    GetHistoryContextSignature = Trim$(v) & " || " & Trim$(t) & " || rows=" & CStr(n)
End Function
Private Function OpenHistoryDialogFromBooking(ByVal bookingHtml As String, ByVal sess As Object) As String

    If Len(bookingHtml) = 0 Then Exit Function

    Dim postUrl As String
    postUrl = ExtractFormActionUrl(bookingHtml)
    If Len(postUrl) = 0 Then
        postUrl = SEARCH_URL
    Else
        postUrl = ResolveUrl(SEARCH_URL, postUrl)
    End If

    Dim btnName As String
    btnName = FindHistoryButtonPostName(bookingHtml)
    If Len(btnName) = 0 Then Exit Function

    Dim body As String
    body = BuildPostBackBodyFromHtml(bookingHtml, btnName)

    If Len(body) = 0 Then
        body = BuildImageClickBodyByPostName(bookingHtml, btnName)
    End If

    If Len(body) = 0 Then Exit Function

    OpenHistoryDialogFromBooking = HttpPost(postUrl, body, sess)

End Function

Private Function FindHistoryButtonPostName(ByVal html As String) As String

    Dim c As Variant
    For Each c In Array( _
        "sibHistory", _
        "ibHistory", _
        "btnHistory", _
        "cmdHistory", _
        "wucBooking$sibHistory", _
        "wucBooking$ibHistory", _
        "wucBooking$btnHistory", _
        "wucBooking$cmdHistory")
        
        If InStr(1, html, "name=""" & CStr(c) & """", vbTextCompare) > 0 _
           Or InStr(1, html, "id=""" & CStr(c) & """", vbTextCompare) > 0 Then
            FindHistoryButtonPostName = CStr(c)
            Exit Function
        End If
    Next c

    Dim r As Object, ms As Object, i As Long
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = True
    r.Multiline = True
    r.Pattern = "<input\b[^>]*>"

    If Not r.Test(html) Then Exit Function
    Set ms = r.Execute(html)

    For i = 0 To ms.Count - 1
        Dim tag As String
        Dim t As String
        Dim n As String
        Dim idv As String

        tag = CStr(ms(i).Value)
        t = LCase$(Trim$(HtmlAttr(tag, "type")))
        n = HtmlAttr(tag, "name")
        idv = HtmlAttr(tag, "id")

        If t = "image" Or t = "submit" Or t = "button" Then
            If InStr(1, tag, "hist", vbTextCompare) > 0 Or _
               InStr(1, tag, "History", vbTextCompare) > 0 Then

                If Len(n) > 0 Then
                    FindHistoryButtonPostName = n
                    Exit Function
                ElseIf Len(idv) > 0 Then
                    FindHistoryButtonPostName = idv
                    Exit Function
                End If
            End If
        End If
    Next i

End Function

Private Function BuildImageClickBodyByPostName(ByVal html As String, ByVal imgBtnName As String) As String

    Dim body As String
    body = BuildPostBackBodyFromHtml(html, imgBtnName)

    If Len(body) > 0 Then
        BuildImageClickBodyByPostName = body
        Exit Function
    End If

    body = ""

    Dim vs As String, vsg As String
    vs = HtmlHidden(html, "__VIEWSTATE")
    vsg = HtmlHidden(html, "__VIEWSTATEGENERATOR")

    If Len(vs) > 0 Then AppendPostField body, "__VIEWSTATE", vs
    If Len(vsg) > 0 Then AppendPostField body, "__VIEWSTATEGENERATOR", vsg

    If Not BodyHasField(body, "defaultButton") Then AppendPostField body, "defaultButton", imgBtnName
    If Not BodyHasField(body, "__EVENTTARGET") Then AppendPostField body, "__EVENTTARGET", ""
    If Not BodyHasField(body, "__EVENTARGUMENT") Then AppendPostField body, "__EVENTARGUMENT", ""

    If LCase$(imgBtnName) = "sibsearch" Then
        If Not BodyHasField(body, "gnWhen") Then AppendPostField body, "gnWhen", "rbAll"
    End If

    AppendPostField body, imgBtnName & ".x", "10"
    AppendPostField body, imgBtnName & ".y", "10"

    BuildImageClickBodyByPostName = body

End Function

Private Function ExtractTableById(ByVal html As String, ByVal tableId As String) As String
    Dim pId As Long
    pId = InStr(1, html, "id=""" & tableId & """", vbTextCompare)
    If pId = 0 Then Exit Function

    Dim pTableStart As Long
    pTableStart = InStrRev(LCase$(html), "<table", pId)
    If pTableStart = 0 Then Exit Function

    Dim pTableEnd As Long
    pTableEnd = InStr(pId, LCase$(html), "</table>", vbTextCompare)
    If pTableEnd = 0 Then Exit Function

    ExtractTableById = Mid$(html, pTableStart, (pTableEnd - pTableStart) + Len("</table>"))
End Function

Private Function GetTdTextsFromRow(ByVal rowHtml As String) As String()
    Dim out() As String
    ReDim out(-1 To -1)

    Dim reTd As Object
    Set reTd = CreateObject("VBScript.RegExp")
    reTd.Global = True
    reTd.IgnoreCase = True
    reTd.Multiline = True
    reTd.Pattern = "<td\b[^>]*>([\s\S]*?)</td>"

    Dim m As Object
    Dim n As Long
    n = -1

    For Each m In reTd.Execute(rowHtml)
        n = n + 1
        If n = 0 Then
            ReDim out(0 To 0)
        Else
            ReDim Preserve out(0 To n)
        End If
        out(n) = CleanHtmlCellText(CStr(m.SubMatches(0)))
    Next m

    GetTdTextsFromRow = out
End Function

Private Function CleanHtmlCellText(ByVal s As String) As String
    Dim reTag As Object
    Set reTag = CreateObject("VBScript.RegExp")
    reTag.Global = True
    reTag.IgnoreCase = True
    reTag.Multiline = True
    reTag.Pattern = "<[^>]+>"

    s = reTag.Replace(s, " ")

    s = Replace$(s, "&nbsp;", " ")
    s = Replace$(s, "&#160;", " ")
    s = Replace$(s, "&amp;", "&")
    s = Replace$(s, "&quot;", """")
    s = Replace$(s, "&apos;", "'")
    s = Replace$(s, "&#39;", "'")
    s = Replace$(s, "&lt;", "<")
    s = Replace$(s, "&gt;", ">")

    s = Replace$(s, vbCr, " ")
    s = Replace$(s, vbLf, " ")
    s = NormalizeSpaces(s)

    CleanHtmlCellText = s
End Function

Private Function NormalizeSpaces(ByVal s As String) As String
    Do While InStr(1, s, "  ", vbBinaryCompare) > 0
        s = Replace$(s, "  ", " ")
    Loop
    NormalizeSpaces = Trim$(s)
End Function

Private Function UBoundSafe(ByRef arr() As String) As Long
    On Error GoTo EH
    UBoundSafe = UBound(arr)
    Exit Function
EH:
    UBoundSafe = -1
End Function

Private Function SafeArr(ByRef arr() As String, ByVal idx As Long) As String
    If idx < LBound(arr) Or idx > UBound(arr) Then
        SafeArr = ""
    Else
        SafeArr = arr(idx)
    End If
End Function

Private Function AddNoCacheToUrl(ByVal u As String) As String
    Dim sep As String
    sep = IIf(InStr(1, u, "?", vbBinaryCompare) > 0, "&", "?")
    AddNoCacheToUrl = u & sep & "_ts=" & Replace$(Format$(Now, "yyyymmddhhnnss"), " ", "") & "_" & CLng(Timer * 1000)
End Function

Private Function ExtractInfoUrl(ByVal html As String) As String

    Dim u As String

    u = Re1(html, "(WebFrmBkgInfo\.aspx\?Wsid=B2E[0-9]+&ContactName=[^""'&\s<>]+)")
    If Len(u) > 0 Then
        ExtractInfoUrl = u
        Exit Function
    End If

    u = Re1(html, "(WebFrmBkgInfo\.aspx\?Wsid=B2E[0-9]+[^""'\s<>]*)")
    ExtractInfoUrl = u

End Function

Private Function ExtractHistoryUrl(ByVal html As String) As String

    Dim u As String

    ' Caso actual: iframe del diálogo ya renderizado
    u = Re1(html, "id=""dialog-body""[^>]*src=""([^""]*WebFrmHistorySelection\.aspx[^""]*)""")
    If Len(u) > 0 Then
        ExtractHistoryUrl = HtmlDecode(u)
        Exit Function
    End If

    ' src con comillas simples
    u = Re1(html, "id=""dialog-body""[^>]*src='([^']*WebFrmHistorySelection\.aspx[^']*)'")
    If Len(u) > 0 Then
        ExtractHistoryUrl = HtmlDecode(u)
        Exit Function
    End If

    ' URL directa en HTML
    u = Re1(html, "(WebFrmHistorySelection\.aspx\?Wsid=[^""'\s<>]+)")
    If Len(u) > 0 Then
        ExtractHistoryUrl = HtmlDecode(u)
        Exit Function
    End If

    ' URL en scripts JS (window.open / showModalDialog / openDialog / etc.)
    u = Re1(html, "window\.open\(\s*['""]([^'""]*WebFrmHistorySelection\.aspx[^'""]*)")
    If Len(u) > 0 Then
        ExtractHistoryUrl = HtmlDecode(u)
        Exit Function
    End If

    u = Re1(html, "showModalDialog\(\s*['""]([^'""]*WebFrmHistorySelection\.aspx[^'""]*)")
    If Len(u) > 0 Then
        ExtractHistoryUrl = HtmlDecode(u)
        Exit Function
    End If

    u = Re1(html, "open[a-zA-Z]*Dialog\(\s*['""]([^'""]*WebFrmHistorySelection\.aspx[^'""]*)")
    If Len(u) > 0 Then
        ExtractHistoryUrl = HtmlDecode(u)
        Exit Function
    End If

    ' URL escapada en JS/HTML (con &amp;)
    u = Re1(html, "['""]([^'""]*WebFrmHistorySelection\.aspx\?[^'""]+)['""]")
    If Len(u) > 0 Then
        If InStr(1, u, "WebFrmHistorySelection.aspx", vbTextCompare) > 0 Then
            ExtractHistoryUrl = HtmlDecode(u)
            Exit Function
        End If
    End If

    ExtractHistoryUrl = ""

End Function

Private Function ExtractBookingWsid(ByVal html As String) As String

    Dim v As String

    v = Re1(html, "WebFrmBkgInfo\.aspx\?Wsid=(B2E[0-9]+)")
    If Len(v) > 0 Then
        ExtractBookingWsid = v
        Exit Function
    End If

    v = Re1(html, "WebFrmPaxPayBkgNo\.aspx\?Wsid=(B2E[0-9]+)")
    If Len(v) > 0 Then
        ExtractBookingWsid = v
        Exit Function
    End If

    v = Re1(html, "\bWsid=(B2E[0-9]+)\b")
    ExtractBookingWsid = v

End Function

Private Function ParseContactNameFromBooking(ByVal html As String) As String

    Dim v As String

    v = Re1(html, "id=""wucBooking_lblContactNameValue""[^>]*>([^<]+)<")
    v = Trim$(HtmlDecode(v))
    If Len(v) > 0 Then
        ParseContactNameFromBooking = v
        Exit Function
    End If

    v = Re1(html, "value=""([^""]+)""[^>]*id=""wucBooking_txtContactName""")
    v = Trim$(HtmlDecode(v))
    If Len(v) > 0 Then
        ParseContactNameFromBooking = v
        Exit Function
    End If

    v = Re1(html, "id=""wucBooking_txtContactName""[^>]*value=""([^""]+)""")
    v = Trim$(HtmlDecode(v))
    If Len(v) > 0 Then
        ParseContactNameFromBooking = v
        Exit Function
    End If

    v = Re1(html, "ContactName=([^&""'<\s]+)")
    v = Trim$(HtmlDecode(v))
    If Len(v) > 0 Then
        ParseContactNameFromBooking = v
        Exit Function
    End If

    ParseContactNameFromBooking = ""

End Function

Private Function ParseOwnerUserId(ByVal html As String) As String
    Dim r As Object, m As Object, a As String, b As String
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.Multiline = True
    r.Pattern = "(?:id=""txtOwnerUserId""[^>]*value=""([^""]*)""|value=""([^""]*)""[^>]*id=""txtOwnerUserId"")"
    If r.Test(html) Then
        Set m = r.Execute(html)(0)
        a = ""
        b = ""
        On Error Resume Next
        a = m.SubMatches(0)
        b = m.SubMatches(1)
        On Error GoTo 0
        If Len(a) > 0 Then
            ParseOwnerUserId = a
        Else
            ParseOwnerUserId = b
        End If
    Else
        ParseOwnerUserId = ""
    End If
End Function

Private Function ParseOwnerUserDetail(ByVal html As String) As String
    Dim r As Object, m As Object, a As String, b As String
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.Multiline = True
    r.Pattern = "(?:id=""txtOwnerUserDetail""[^>]*value=""([^""]*)""|value=""([^""]*)""[^>]*id=""txtOwnerUserDetail"")"
    If r.Test(html) Then
        Set m = r.Execute(html)(0)
        a = ""
        b = ""
        On Error Resume Next
        a = m.SubMatches(0)
        b = m.SubMatches(1)
        On Error GoTo 0
        If Len(a) > 0 Then
            ParseOwnerUserDetail = a
        Else
            ParseOwnerUserDetail = b
        End If
    Else
        ParseOwnerUserDetail = ""
    End If
End Function

Private Function IsHistoryOptToBkdEventText(ByVal s As String) As Boolean

    Dim t As String
    t = UCase$(Trim$(s))

    If Len(t) = 0 Then Exit Function

    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace$(t, "  ", " ")
    Loop

    ' Casos exactos
    If InStr(1, t, "BKG STATUS CHANGED FROM OP TO BK", vbBinaryCompare) > 0 Then
        IsHistoryOptToBkdEventText = True
        Exit Function
    End If

    If InStr(1, t, "BKG STATUS CHANGED FROM OPT TO BKD", vbBinaryCompare) > 0 Then
        IsHistoryOptToBkdEventText = True
        Exit Function
    End If

    If InStr(1, t, "BKG STATUS CHANGED FROM OPT TO BK", vbBinaryCompare) > 0 Then
        IsHistoryOptToBkdEventText = True
        Exit Function
    End If

    ' Fallback flexible
    If InStr(1, t, "STATUS CHANGED", vbBinaryCompare) > 0 Then
        If ((InStr(1, t, "OPT", vbBinaryCompare) > 0 Or _
             InStr(1, t, " OP ", vbBinaryCompare) > 0 Or _
             Right$(t, 3) = " OP") And _
            (InStr(1, t, "BKD", vbBinaryCompare) > 0 Or _
             InStr(1, t, " BK ", vbBinaryCompare) > 0 Or _
             Right$(t, 3) = " BK")) Then

            IsHistoryOptToBkdEventText = True
            Exit Function
        End If
    End If

End Function

Private Function TryGetHistorySeqFromTds(ByVal tds As Collection, ByRef seqNum As Long) As Boolean

    Dim j As Long
    Dim maxJ As Long
    Dim s As String

    maxJ = 3
    If tds.Count < maxJ Then maxJ = tds.Count

    For j = 1 To maxJ
        s = Trim$(CStr(tds(j)))
        If Len(s) > 0 Then
            If IsNumeric(s) Then
                If InStr(1, s, ":", vbBinaryCompare) = 0 Then
                    seqNum = CLng(s)
                    TryGetHistorySeqFromTds = True
                    Exit Function
                End If
            End If
        End If
    Next j

    TryGetHistorySeqFromTds = False

End Function

Private Function LooksLikeHistoryUserDetail(ByVal s As String) As Boolean

    Dim t As String
    t = Trim$(s)

    If Len(t) = 0 Then Exit Function
    If InStr(1, t, "/", vbBinaryCompare) = 0 Then Exit Function
    If InStr(1, t, " ", vbBinaryCompare) > 0 Then Exit Function

    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.Multiline = False
    r.Pattern = "^[A-Za-z0-9]{1,8}/[A-Za-z0-9._-]+$"

    LooksLikeHistoryUserDetail = r.Test(t)

End Function

Private Function BuildPostBackBodyFromHtml(ByVal html As String, ByVal clickedButtonId As String) As String

    Dim body As String
    body = ""

    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = True
    r.Multiline = True

    Dim ms As Object
    Dim i As Long
    Dim tag As String
    Dim t As String
    Dim n As String
    Dim v As String
    Dim tagL As String

    r.Pattern = "<input\b[^>]*>"
    If r.Test(html) Then
        Set ms = r.Execute(html)

        For i = 0 To ms.Count - 1
            tag = CStr(ms(i).Value)
            tagL = LCase$(tag)

            t = LCase$(Trim$(HtmlAttr(tag, "type")))
            n = HtmlAttr(tag, "name")
            If Len(n) = 0 Then n = HtmlAttr(tag, "id")
            v = HtmlAttr(tag, "value")

            If Len(n) = 0 Then GoTo NextInput

            Select Case t
                Case "hidden"
                    AppendPostField body, n, v

                Case "text", "password", "search", "email", "tel", "number"
                    AppendPostField body, n, v

                Case "radio", "checkbox"
                    If InStr(1, tagL, " checked", vbTextCompare) > 0 Or InStr(1, tagL, " checked=", vbTextCompare) > 0 Then
                        AppendPostField body, n, v
                    End If

                Case "submit", "button", "image"

                Case Else
                    If Len(t) = 0 Then
                        AppendPostField body, n, v
                    End If
            End Select

NextInput:
        Next i
    End If

    r.Pattern = "<select\b[^>]*>[\s\S]*?</select>"
    If r.Test(html) Then
        Set ms = r.Execute(html)

        For i = 0 To ms.Count - 1
            tag = CStr(ms(i).Value)
            n = HtmlAttr(tag, "name")
            If Len(n) = 0 Then n = HtmlAttr(tag, "id")
            If Len(n) > 0 Then
                v = GetSelectedOptionValue(tag)
                If Not BodyHasField(body, n) Then
                    AppendPostField body, n, v
                End If
            End If
        Next i
    End If

    r.Pattern = "<textarea\b[^>]*>[\s\S]*?</textarea>"
    If r.Test(html) Then
        Set ms = r.Execute(html)

        For i = 0 To ms.Count - 1
            tag = CStr(ms(i).Value)
            n = HtmlAttr(tag, "name")
            If Len(n) = 0 Then n = HtmlAttr(tag, "id")
            If Len(n) > 0 Then
                v = Re1(tag, "<textarea\b[^>]*>([\s\S]*?)</textarea>")
                v = HtmlDecode(v)
                If Not BodyHasField(body, n) Then
                    AppendPostField body, n, v
                End If
            End If
        Next i
    End If

    If Not BodyHasField(body, "defaultButton") Then AppendPostField body, "defaultButton", clickedButtonId
    If Not BodyHasField(body, "__EVENTTARGET") Then AppendPostField body, "__EVENTTARGET", ""
    If Not BodyHasField(body, "__EVENTARGUMENT") Then AppendPostField body, "__EVENTARGUMENT", ""

    If LCase$(clickedButtonId) = "sibsearch" Then
        If Not BodyHasField(body, "gnWhen") Then AppendPostField body, "gnWhen", "rbAll"
    End If

    AppendPostField body, clickedButtonId & ".x", "10"
    AppendPostField body, clickedButtonId & ".y", "10"

    BuildPostBackBodyFromHtml = body

End Function
Private Function ExtractTdValuesFromRow(ByVal rowHtml As String) As Collection

    Dim c As New Collection

    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = True
    r.Multiline = True
    r.Pattern = "<td\b[^>]*>([\s\S]*?)</td>"

    If r.Test(rowHtml) Then
        Dim ms As Object
        Set ms = r.Execute(rowHtml)

        Dim i As Long
        For i = 0 To ms.Count - 1
            c.Add HtmlCellText(CStr(ms(i).SubMatches(0)))
        Next i
    End If

    Set ExtractTdValuesFromRow = c

End Function

Private Function HtmlCellText(ByVal s As String) As String

    Dim t As String
    t = s

    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = True
    r.Multiline = True
    r.Pattern = "<[^>]+>"

    t = r.Replace(t, "")
    t = Replace$(t, "&nbsp;", " ")
    t = HtmlDecode(t)
    t = Replace$(t, vbCr, " ")
    t = Replace$(t, vbLf, " ")

    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace$(t, "  ", " ")
    Loop

    HtmlCellText = Trim$(t)

End Function

Private Function GetSelectedOptionValue(ByVal selectBlockHtml As String) As String

    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = True
    r.Multiline = True
    r.Pattern = "<option\b[^>]*>"

    If Not r.Test(selectBlockHtml) Then
        GetSelectedOptionValue = ""
        Exit Function
    End If

    Dim ms As Object
    Set ms = r.Execute(selectBlockHtml)

    Dim i As Long
    Dim firstVal As String
    firstVal = ""

    For i = 0 To ms.Count - 1
        Dim tag As String
        tag = CStr(ms(i).Value)

        Dim v As String
        v = HtmlAttr(tag, "value")
        If Len(firstVal) = 0 Then firstVal = v

        If InStr(1, LCase$(tag), " selected", vbTextCompare) > 0 Or InStr(1, LCase$(tag), " selected=", vbTextCompare) > 0 Then
            GetSelectedOptionValue = v
            Exit Function
        End If
    Next i

    GetSelectedOptionValue = firstVal

End Function

Private Function BodyHasField(ByVal body As String, ByVal fieldName As String) As Boolean
    BodyHasField = (InStr(1, "&" & body & "&", "&" & UrlEnc(fieldName) & "=", vbTextCompare) > 0)
End Function

Private Sub AppendPostField(ByRef body As String, ByVal k As String, ByVal v As String)
    If Len(body) > 0 Then body = body & "&"
    body = body & UrlEnc(k) & "=" & UrlEnc(v)
End Sub

Private Function HtmlAttr(ByVal tagHtml As String, ByVal attrName As String) As String

    Dim r As Object, m As Object
    Set r = CreateObject("VBScript.RegExp")
    r.IgnoreCase = True
    r.Global = False
    r.Multiline = False
    r.Pattern = attrName & "\s*=\s*(?:""([^""]*)""|'([^']*)'|([^\s>]+))"

    If r.Test(tagHtml) Then
        Set m = r.Execute(tagHtml)(0)
        If Len(CStr(m.SubMatches(0))) > 0 Then
            HtmlAttr = CStr(m.SubMatches(0))
        ElseIf Len(CStr(m.SubMatches(1))) > 0 Then
            HtmlAttr = CStr(m.SubMatches(1))
        Else
            HtmlAttr = CStr(m.SubMatches(2))
        End If
    Else
        HtmlAttr = ""
    End If

End Function

Private Function ExtractFormActionUrl(ByVal html As String) As String

    Dim u As String

    u = Re1(html, "<form\b[^>]*action=""([^""]+)""")
    If Len(u) > 0 Then
        ExtractFormActionUrl = HtmlDecode(u)
        Exit Function
    End If

    u = Re1(html, "<form\b[^>]*action='([^']+)'")
    If Len(u) > 0 Then
        ExtractFormActionUrl = HtmlDecode(u)
        Exit Function
    End If

    ExtractFormActionUrl = ""

End Function

Private Function ParseStatusFromBooking(ByVal html As String) As String

    Dim v As String
    v = Re1(html, "id=""wucBooking_lblStatusValue""[^>]*>([^<]+)<")
    v = Trim$(HtmlDecode(v))

    If Len(v) > 0 Then
        ParseStatusFromBooking = UCase$(v)
        Exit Function
    End If

    v = Re1(html, "\b(BKD|OPT|CXL|WTL|BK|OP|CX)\b")
    v = Trim$(v)

    If Len(v) > 0 Then
        ParseStatusFromBooking = UCase$(v)
    End If

End Function

Private Function ParseCruiseCodeFromBooking(ByVal html As String) As String
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Pattern = "Cruise\s*-\s*([A-Z0-9]+)"
    r.IgnoreCase = True
    r.Global = False
    If r.Test(html) Then
        ParseCruiseCodeFromBooking = Trim$(r.Execute(html)(0).SubMatches(0))
    Else
        ParseCruiseCodeFromBooking = ""
    End If
End Function

Private Function ParseBkgDateFromBooking(ByVal html As String) As String
    ParseBkgDateFromBooking = Re1(html, "id=""wucBooking_lblBkgDateValue""[^>]*>([^<]+)<")
End Function

Private Function ParseOptDateFromBooking(ByVal html As String) As String
    ParseOptDateFromBooking = Re1(html, "id=""wucBooking_lblOptDateValue""[^>]*>([^<]+)<")
End Function

Private Function ParseCancReasonFromBooking(ByVal html As String) As String
    Dim v As String
    v = Re1(html, "id=""wucBooking_lblCancReasonValue""[^>]*>([^<]+)<")
    v = Trim$(HtmlDecode(v))
    ParseCancReasonFromBooking = v
End Function

Private Function ParseGrossValueFromBooking(ByVal html As String) As String

    Dim v As String
    v = Re1(html, "id=""wucPricing_lblGrossValue""[^>]*>([^<]+)<")

    v = Replace$(v, "EUR", "")
    v = Replace$(v, vbCr, "")
    v = Replace$(v, vbLf, "")
    v = Replace$(v, " ", "")

    ParseGrossValueFromBooking = Trim$(v)

End Function

Public Function ExtractWsid(ByVal html As String) As String
    ExtractWsid = Re1(html, "Wsid=(B2E[0-9]+)")
End Function

Private Function HttpGet(ByVal url As String, ByVal sess As Object) As String
    HttpGet = HttpSendFollow("GET", url, "", sess)
End Function

Private Function HttpPost(ByVal url As String, ByVal body As String, ByVal sess As Object) As String
    HttpPost = HttpSendFollow("POST", url, body, sess)
End Function

Private Function HttpSendFollow(ByVal method As String, ByVal url As String, ByVal body As String, ByVal sess As Object) As String

    Dim curUrl As String
    curUrl = url

    Dim curMethod As String
    curMethod = method

    Dim curBody As String
    curBody = body

    Dim hops As Long
    hops = 0

    Dim resp As String
    resp = ""

    Do

        Dim h As Object
        Set h = CreateObject("WinHTTP.WinHTTPRequest.5.1")

        On Error Resume Next
        h.Option(6) = False
        On Error GoTo 0

        h.Open curMethod, curUrl, False
        If Len(CStr(sess("cookie"))) > 0 Then h.SetRequestHeader "Cookie", CStr(sess("cookie"))
        h.SetRequestHeader "User-Agent", "Mozilla/5.0"
        h.SetRequestHeader "Cache-Control", "no-cache, no-store, must-revalidate"
        h.SetRequestHeader "Pragma", "no-cache"
        h.SetRequestHeader "Expires", "0"

        If UCase$(curMethod) = "POST" Then
            h.SetRequestHeader "Content-Type", "application/x-www-form-urlencoded"
            h.Send curBody
        Else
            h.Send
        End If

        UpdateCookies h, sess
        resp = h.ResponseText

        Dim st As Long
        st = 0
        On Error Resume Next
        st = CLng(h.status)
        On Error GoTo 0

        If st = 301 Or st = 302 Or st = 303 Or st = 307 Or st = 308 Then

            Dim loc As String
            loc = ""
            On Error Resume Next
            loc = CStr(h.GetResponseHeader("Location"))
            On Error GoTo 0

            If Len(loc) = 0 Then Exit Do

            curUrl = ResolveUrl(curUrl, loc)
            curMethod = "GET"
            curBody = ""

            hops = hops + 1
            If hops >= 8 Then Exit Do

        Else
            Exit Do
        End If

    Loop

    HttpSendFollow = resp

End Function

Private Function ResolveUrl(ByVal currentUrl As String, ByVal location As String) As String

    Dim loc As String
    loc = Trim$(location)

    If Len(loc) = 0 Then
        ResolveUrl = currentUrl
        Exit Function
    End If

    If LCase$(Left$(loc, 7)) = "http://" Or LCase$(Left$(loc, 8)) = "https://" Then
        ResolveUrl = loc
        Exit Function
    End If

    If Left$(loc, 1) = "/" Then
        ResolveUrl = GetUrlRoot(currentUrl) & loc
        Exit Function
    End If

    If InStr(1, loc, "B2EWeb/", vbTextCompare) > 0 Then
        If Left$(loc, 1) = "." Then
            ResolveUrl = GetUrlFolder(currentUrl) & Mid$(loc, 3)
        Else
            ResolveUrl = GetUrlRoot(currentUrl) & "/" & loc
        End If
        Exit Function
    End If

    ResolveUrl = GetUrlFolder(currentUrl) & loc

End Function

Private Function GetUrlRoot(ByVal u As String) As String
    Dim p As Long, p2 As Long
    p = InStr(1, u, "://", vbTextCompare)
    If p = 0 Then
        GetUrlRoot = ""
        Exit Function
    End If
    p2 = InStr(p + 3, u, "/")
    If p2 = 0 Then
        GetUrlRoot = u
    Else
        GetUrlRoot = Left$(u, p2 - 1)
    End If
End Function

Private Function GetUrlFolder(ByVal u As String) As String
    Dim p As Long
    p = InStrRev(u, "/")
    If p = 0 Then
        GetUrlFolder = u & "/"
    Else
        GetUrlFolder = Left$(u, p)
    End If
End Function

Private Sub UpdateCookies(ByVal h As Object, ByVal sess As Object)

    Dim hdr As String
    On Error Resume Next
    hdr = h.GetResponseHeader("Set-Cookie")
    On Error GoTo 0

    If Len(hdr) = 0 Then Exit Sub

    Dim jar As Object
    Set jar = CreateObject("Scripting.Dictionary")

    Dim existing As String
    existing = CStr(sess("cookie"))

    If Len(existing) > 0 Then
        Dim parts() As String, j As Long, kv As String, k As String, v As String
        parts = Split(existing, ";")
        For j = LBound(parts) To UBound(parts)
            kv = Trim$(parts(j))
            If InStr(1, kv, "=", vbBinaryCompare) > 0 Then
                k = Split(kv, "=")(0)
                v = Mid$(kv, Len(k) + 2)
                If Len(k) > 0 Then jar(k) = v
            End If
        Next j
    End If

    Dim lines() As String, i As Long
    lines = Split(hdr, vbLf)

    For i = LBound(lines) To UBound(lines)
        Dim ln As String
        ln = Trim$(lines(i))
        If InStr(1, ln, "=", vbBinaryCompare) > 0 Then
            Dim pair As String
            pair = Split(ln, ";")(0)

            Dim kk As String, vv As String
            kk = Split(pair, "=")(0)
            vv = Mid$(pair, Len(kk) + 2)

            jar(kk) = vv
        End If
    Next i

    Dim out As String, key As Variant
    For Each key In jar.keys
        If Len(out) > 0 Then out = out & "; "
        out = out & key & "=" & jar(key)
    Next key

    sess("cookie") = out

End Sub

Private Function HtmlHidden(ByVal html As String, ByVal nameOrId As String) As String
    HtmlHidden = Re1(html, "name=""" & nameOrId & """[^>]*value=""([^""]*)""")
    If Len(HtmlHidden) = 0 Then HtmlHidden = Re1(html, "id=""" & nameOrId & """[^>]*value=""([^""]*)""")
End Function

Private Function Re1(ByVal s As String, ByVal pat As String) As String
    Dim r As Object
    Set r = CreateObject("VBScript.RegExp")
    r.Pattern = pat
    r.IgnoreCase = True
    r.Global = False
    r.Multiline = True
    If r.Test(s) Then Re1 = r.Execute(s)(0).SubMatches(0) Else Re1 = ""
End Function

Private Function UrlEnc(ByVal s As String) As String
    Dim i As Long, ch As Integer, o As String
    For i = 1 To Len(s)
        ch = AscW(Mid$(s, i, 1))
        Select Case ch
            Case 48 To 57, 65 To 90, 97 To 122
                o = o & ChrW(ch)
            Case 32
                o = o & "+"
            Case Else
                o = o & "%" & Right$("0" & Hex$(ch And &HFF), 2)
        End Select
    Next i
    UrlEnc = o
End Function

Private Function HtmlDecode(ByVal s As String) As String
    s = Replace$(s, "&amp;", "&")
    s = Replace$(s, "&lt;", "<")
    s = Replace$(s, "&gt;", ">")
    s = Replace$(s, "&quot;", """")
    s = Replace$(s, "&#39;", "'")
    HtmlDecode = s
End Function

Private Sub LogLine(ByVal s As String)
    Debug.Print Format$(Now, "dd/mm/yyyy hh:nn:ss") & " | " & s
End Sub

Private Sub WaitSeconds(ByVal seconds As Long)
    Dim i As Long
    For i = 1 To seconds
        DoEvents
        Sleep 1000
    Next i
End Sub

Private Function AskStartRow(ByVal minRow As Long, ByVal maxRow As Long) As Long

    Dim s As String, v As Long

    s = InputBox("¿Desde qué fila quieres empezar?" & vbCrLf & _
                 "Mín: " & minRow & "  |  Máx: " & maxRow, _
                 "Inicio de evaluación", CStr(minRow))

    If Len(Trim$(s)) = 0 Then
        AskStartRow = 0
        Exit Function
    End If

    If Not IsNumeric(s) Then
        MsgBox "Debes ingresar un número de fila.", vbExclamation
        AskStartRow = 0
        Exit Function
    End If

    v = CLng(s)
    If v < minRow Or v > maxRow Then
        MsgBox "Fila fuera de rango. Debe estar entre " & minRow & " y " & maxRow & ".", vbExclamation
        AskStartRow = 0
        Exit Function
    End If

    AskStartRow = v

End Function

Private Function DateFromDMY(ByVal s As String) As Variant

    Dim t As String, p() As String
    t = Trim$(s)
    p = Split(t, "/")
    If UBound(p) <> 2 Then
        DateFromDMY = ""
        Exit Function
    End If

    Dim d As Integer, m As Integer, y As Integer
    d = CInt(p(0))
    m = CInt(p(1))
    y = CInt(p(2))

    DateFromDMY = DateSerial(y, m, d)

End Function

Private Function NormalizarNumero(ByVal s As String) As String

    Dim t As String
    t = Trim$(s)

    t = Replace$(t, "€", "")
    t = Replace$(t, "EUR", "")
    t = Replace$(t, vbCr, "")
    t = Replace$(t, vbLf, "")
    t = Replace$(t, " ", "")

    If InStr(1, t, ",", vbBinaryCompare) > 0 And InStr(1, t, ".", vbBinaryCompare) > 0 Then
        If InStrRev(t, ",") > InStrRev(t, ".") Then
            t = Replace$(t, ".", "")
            t = Replace$(t, ",", ".")
        Else
            t = Replace$(t, ",", "")
        End If
    Else
        If InStr(1, t, ",", vbBinaryCompare) > 0 And InStr(1, t, ".", vbBinaryCompare) = 0 Then
            t = Replace$(t, ",", ".")
        End If
    End If

    NormalizarNumero = t

End Function

Private Sub ClearCellComment(ByVal c As Range)
    On Error Resume Next
    If Not c.Comment Is Nothing Then c.Comment.Delete
    c.ClearComments
    On Error GoTo 0
End Sub

Private Sub SetCellComment(ByVal c As Range, ByVal txt As String)
    If Len(Trim$(txt)) = 0 Then Exit Sub
    On Error Resume Next
    c.AddComment Trim$(txt)
    If Err.Number = 0 Then
        c.Comment.Visible = False
    End If
    On Error GoTo 0
End Sub

Private Sub ProgressOpen(ByVal title As String)
    On Error Resume Next
    Set gProg = New frmProgress
    gProg.InitProgress title
    gProg.Show vbModeless
    DoEvents
    On Error GoTo 0
End Sub

Private Sub ProgressStep(ByVal current As Long, ByVal total As Long, Optional ByVal detail As String = "")
    On Error Resume Next
    If Not gProg Is Nothing Then
        gProg.SetProgress current, total, detail
        DoEvents
    End If
    On Error GoTo 0
End Sub

Private Sub ProgressClose()
    On Error Resume Next
    If Not gProg Is Nothing Then
        Unload gProg
        Set gProg = Nothing
    End If
    DoEvents
    On Error GoTo 0
End Sub

