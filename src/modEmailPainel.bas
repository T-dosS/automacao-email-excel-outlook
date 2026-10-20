Attribute VB_Name = "modEmailPainel"
'==========================================================================================
' Automação de e-mail do painel · Excel + Outlook (VBA)
' https://github.com/T-dosS/automacao-email-excel-outlook
'
' Um clique no botão do painel e o Outlook ABRE um e-mail pronto para você revisar:
'   - assunto com o título, o período e o status;
'   - resumo em texto com os indicadores escolhidos e os pontos de atenção;
'   - imagem do painel no corpo do e-mail (lê bem no celular);
'   - anexo .xlsx só com a aba do painel, com os valores congelados no momento do envio
'     (gráficos viram imagens; sem fórmulas, macros, nomes ou ligações com a planilha original).
' Nada é enviado sem você clicar em "Enviar" no Outlook.
'
' Requisitos: Excel para Windows (2010 ou mais novo) e Outlook clássico configurado.
'             O "novo Outlook" não aceita automação por VBA.
'==========================================================================================
Option Explicit

'============================== CONFIGURAÇÃO · ajuste para a sua planilha ==============================
Private Const TITULO As String = "Painel de Resultados"
Private Const ABA_PAINEL As String = "PAINEL"
Private Const AREA_PAINEL As String = "A1:O35"        ' área que vira imagem e anexo
Private Const CEL_PERIODO As String = "C5"            ' texto do período, ex.: "Outubro/2026"
Private Const CEL_STATUS As String = "L3"             ' texto do status, ex.: "ATENÇÃO" ("" = sem status)
' Indicadores do resumo no corpo do e-mail: "Rótulo|Célula;Rótulo|Célula;..." (o texto é lido como aparece na tela)
Private Const RESUMO As String = "Status|L3;Meta do mês|B10;Realizado|E10;% da meta|E11;Falta para a meta|H10;Projeção de fechamento|K10"
Private Const AREA_ALERTAS As String = "C14:C16"      ' pontos de atenção, um por célula ("" = sem)
Private Const LINHA_ACOES As Long = 7                  ' linha dos botões, oculta durante a captura (0 = nenhuma)
Private Const PREFIXO_BOTAO As String = "btn_"         ' formas com esse prefixo no nome somem do print
Private Const SENHA_PAINEL As String = ""              ' senha de proteção da aba do painel, se houver
Private Const ABA_DEST As String = "DESTINATARIOS"     ' colunas: B Nome · C E-mail · D Envio · E Ativo
Private Const DEST_PRIMEIRA As Long = 6
Private Const DEST_ULTIMA As Long = 55
Private Const CEL_ABERTURA As String = "C58"           ' primeira frase do e-mail
Private Const COR_TITULO As String = "#1F4E79"
Private Const LARGURA_IMG_EMAIL As Long = 900          ' largura da imagem no corpo do e-mail (px)
Private Const ESCALA_IMG As Double = 1.6               ' >1 = imagem maior e mais nítida
'=======================================================================================================

Private Const CID_IMG As String = "painel"
Private m_wbTemp As Workbook                           ' pasta temporária aberta (fechada também em caso de erro)

' Macro do botão do painel.
Public Sub EnviarPainelPorEmail()
    On Error GoTo Falha
    GerarEmailPainel False
    Exit Sub
Falha:
    Application.ScreenUpdating = True
    MsgBox "Não foi possível montar o e-mail do painel." & vbCrLf & vbCrLf & Err.Description & vbCrLf & vbCrLf & _
           "Verifique se o Outlook clássico está instalado e configurado com a sua conta.", vbExclamation, "Enviar painel por e-mail"
End Sub

' apenasTeste = True: monta tudo, grava o corpo em HTML ao lado dos arquivos, devolve um resumo
' e DESCARTA o e-mail (não envia e não salva rascunho). Útil para testar sem risco.
Public Function GerarEmailPainel(Optional ByVal apenasTeste As Boolean = False) As String
    Dim etapa As String
    On Error GoTo Erro
    etapa = "preparar"
    Dim pn As Worksheet: Set pn = ThisWorkbook.Worksheets(ABA_PAINEL)
    Dim periodo As String: periodo = Trim$(pn.Range(CEL_PERIODO).Text)
    Dim status As String: If CEL_STATUS <> "" Then status = Trim$(pn.Range(CEL_STATUS).Text)
    Dim base As String: base = NomeArquivo(TITULO & "_" & periodo)
    Dim pasta As String: pasta = Environ$("TEMP") & "\EmailPainel_" & Format$(Now, "yyyymmdd_hhnnss")
    Dim arqImg As String: arqImg = pasta & "\" & base & ".png"
    Dim arqXls As String: arqXls = pasta & "\" & base & ".xlsx"
    MkDir pasta

    Application.Calculate
    etapa = "gerar a imagem do painel": ExportarImagemPainel pn, arqImg
    etapa = "gerar o anexo com a aba do painel": ExportarAbaPainel pn, arqXls

    etapa = "abrir o Outlook"
    Dim ol As Object, m As Object, att As Object
    Set ol = CreateObject("Outlook.Application")
    Set m = ol.CreateItem(0)                                   ' olMailItem
    etapa = "montar o e-mail"
    PreencherDestinatarios m
    m.Subject = TITULO & " · " & periodo & IIf(status <> "", " · Status: " & status, "")
    Set att = m.Attachments.Add(arqImg, 1, 0)                   ' imagem usada no corpo (Content-ID)
    att.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x3712001F", CID_IMG
    att.PropertyAccessor.SetProperty "http://schemas.microsoft.com/mapi/proptag/0x7FFE000B", True
    m.Attachments.Add arqXls
    Dim corpo As String: corpo = MontarCorpo(pn, periodo)

    If apenasTeste Then
        m.HTMLBody = corpo
        GravarTexto pasta & "\corpo_email.html", corpo
        GerarEmailPainel = "assunto=" & m.Subject & "|para=" & m.To & "|cc=" & m.CC & "|cco=" & m.BCC & _
                           "|anexos=" & m.Attachments.Count & "|img=" & arqImg & "|xlsx=" & arqXls & "|html=" & pasta & "\corpo_email.html"
        m.Close 1                                               ' olDiscard
    Else
        m.Display                                               ' abre para revisão (carrega a sua assinatura)
        m.HTMLBody = corpo & m.HTMLBody
        On Error Resume Next                                    ' o Outlook já copiou os anexos
        Kill arqImg: Kill arqXls: RmDir pasta
    End If
    Exit Function
Erro:
    Dim nErr As Long, dErr As String: nErr = Err.Number: dErr = Err.Description
    On Error Resume Next
    Application.CutCopyMode = False
    OcultarAcoes ThisWorkbook.Worksheets(ABA_PAINEL), False
    If Not m_wbTemp Is Nothing Then m_wbTemp.Close False: Set m_wbTemp = Nothing
    ThisWorkbook.Activate
    Application.ScreenUpdating = True: Application.DisplayAlerts = True
    On Error GoTo 0
    If apenasTeste Then
        GerarEmailPainel = "ERRO|" & etapa & "|" & nErr & "|" & dErr
    Else
        Err.Raise nErr, "GerarEmailPainel", "Etapa: " & etapa & " · " & dErr
    End If
End Function

'------------------------------------------------------------------------------------------ imagem
Private Sub ExportarImagemPainel(ByVal pn As Worksheet, ByVal arq As String)
    Dim r As Range: Set r = pn.Range(AREA_PAINEL)
    Dim co As ChartObject
    OcultarAcoes pn, True                                      ' botões não vão para o print
    r.CopyPicture Appearance:=xlPrinter, Format:=xlPicture     ' como impresso: sem grade, independe da tela
    OcultarAcoes pn, False
    Set m_wbTemp = Workbooks.Add(xlWBATWorksheet)
    Set co = m_wbTemp.Worksheets(1).ChartObjects.Add(0, 0, r.Width * ESCALA_IMG, r.Height * ESCALA_IMG)
    co.Chart.ChartArea.Format.Line.Visible = msoFalse
    co.Activate
    co.Chart.Paste
    With co.Chart.Shapes(co.Chart.Shapes.Count)
        .LockAspectRatio = msoFalse
        .Left = 0: .Top = 0
        .Width = co.Chart.ChartArea.Width: .Height = co.Chart.ChartArea.Height
    End With
    co.Chart.Export arq, "PNG"
    m_wbTemp.Close False: Set m_wbTemp = Nothing
    ThisWorkbook.Activate
End Sub

' Oculta/reexibe a linha de botões durante a captura.
Private Sub OcultarAcoes(ByVal pn As Worksheet, ByVal ocultar As Boolean)
    Dim sh As Shape, protegida As Boolean
    protegida = pn.ProtectContents
    If protegida Then pn.Unprotect SENHA_PAINEL
    If LINHA_ACOES > 0 Then pn.Rows(LINHA_ACOES).Hidden = ocultar
    For Each sh In pn.Shapes
        If Left$(sh.Name, Len(PREFIXO_BOTAO)) = PREFIXO_BOTAO Then sh.Visible = IIf(ocultar, msoFalse, msoTrue)
    Next sh
    If protegida Then pn.Protect SENHA_PAINEL, True, True, True
End Sub

'------------------------------------------------------------------------------------------ anexo .xlsx (só a aba do painel)
' Monta um arquivo novo por colagem especial (larguras, valores e formatos) em vez de copiar a aba inteira:
' assim o anexo não leva fórmulas, nomes ou ligações, e não depende dos nomes ocultos da pasta original.
Private Sub ExportarAbaPainel(ByVal pn As Worksheet, ByVal arq As String)
    Dim alertas As Boolean: alertas = Application.DisplayAlerts
    Application.DisplayAlerts = False
    Dim src As Range: Set src = pn.Range(AREA_PAINEL)
    Set m_wbTemp = Workbooks.Add(xlWBATWorksheet)
    Dim ws As Worksheet: Set ws = m_wbTemp.Worksheets(1)
    ws.Name = Left$(ABA_PAINEL, 31)
    ws.Cells.Interior.Color = RGB(255, 255, 255)

    ' 1) larguras, valores e formatos
    src.Copy
    ws.Range("A1").PasteSpecial xlPasteColumnWidths
    ws.Range("A1").PasteSpecial xlPasteValues
    ws.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    Dim r As Long
    For r = 1 To src.Rows.Count: ws.Rows(r).RowHeight = pn.Rows(src.Row + r - 1).RowHeight: Next r
    If LINHA_ACOES > 0 Then ws.Rows(LINHA_ACOES - src.Row + 1).Hidden = True

    ' 2) cores dos formatos condicionais ficam fixas; só as barras de dados continuam condicionais
    Dim c As Range, s As Range, i As Long
    For Each s In src.Cells
        If s.FormatConditions.Count > 0 Then
            If s.Address = s.MergeArea.Cells(1, 1).Address Then
                Set c = ws.Cells(s.Row - src.Row + 1, s.Column - src.Column + 1).MergeArea
                c.Font.Color = s.DisplayFormat.Font.Color
                c.Font.Bold = s.DisplayFormat.Font.Bold
                If s.DisplayFormat.Interior.ColorIndex <> xlColorIndexNone Then c.Interior.Color = s.DisplayFormat.Interior.Color
            End If
        End If
    Next s
    For i = ws.Cells.FormatConditions.Count To 1 Step -1
        If ws.Cells.FormatConditions(i).Type <> xlDatabar Then ws.Cells.FormatConditions(i).Delete
    Next i
    ws.Cells.Validation.Delete

    ' 3) imagens e gráficos (gráficos viram imagens); botões ficam de fora
    Dim sh As Shape, p As Object
    ws.Activate: ws.Range("A1").Select
    For Each sh In pn.Shapes
        If (sh.Type = msoPicture Or sh.Type = msoChart) And Left$(sh.Name, Len(PREFIXO_BOTAO)) <> PREFIXO_BOTAO Then
            sh.CopyPicture Appearance:=xlPrinter, Format:=xlPicture
            Set p = ws.Pictures.Paste
            p.Left = sh.Left - src.Left: p.Top = sh.Top - src.Top: p.Width = sh.Width: p.Height = sh.Height
        End If
    Next sh

    ' 4) sem ligações com a planilha original
    Dim lk As Variant, k As Long
    lk = m_wbTemp.LinkSources(xlExcelLinks)
    If Not IsEmpty(lk) Then
        For k = LBound(lk) To UBound(lk): m_wbTemp.BreakLink lk(k), xlLinkTypeExcelLinks: Next k
    End If
    With ws.PageSetup
        .Orientation = xlLandscape: .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1: .CenterHorizontally = True
    End With
    ActiveWindow.DisplayGridlines = False
    ws.Range("A1").Select
    m_wbTemp.SaveAs Filename:=arq, FileFormat:=xlOpenXMLWorkbook
    m_wbTemp.Close False: Set m_wbTemp = Nothing
    Application.DisplayAlerts = alertas
    ThisWorkbook.Activate
End Sub

'------------------------------------------------------------------------------------------ e-mail
Private Sub PreencherDestinatarios(ByVal m As Object)
    Dim ws As Worksheet, r As Long, para As String, cc As String, cco As String, em As String, tipo As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ABA_DEST)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    For r = DEST_PRIMEIRA To DEST_ULTIMA
        em = Trim$(CStr(ws.Cells(r, 3).Value))
        If em <> "" And UCase$(Trim$(CStr(ws.Cells(r, 5).Value))) <> "NÃO" Then   ' Ativo vazio = Sim
            tipo = UCase$(Trim$(CStr(ws.Cells(r, 4).Value)))                     ' Envio vazio = Para
            If tipo = "CÓPIA OCULTA" Then
                cco = cco & em & ";"
            ElseIf tipo = "CÓPIA" Then
                cc = cc & em & ";"
            Else
                para = para & em & ";"
            End If
        End If
    Next r
    If para <> "" Then m.To = para
    If cc <> "" Then m.CC = cc
    If cco <> "" Then m.BCC = cco
End Sub

Private Function MontarCorpo(ByVal pn As Worksheet, ByVal periodo As String) As String
    Dim abertura As String, h As String, item As Variant, partes() As String, c As Range
    abertura = "Olá,"
    On Error Resume Next
    abertura = CStr(ThisWorkbook.Worksheets(ABA_DEST).Range(CEL_ABERTURA).Value)
    On Error GoTo 0
    If Trim$(abertura) = "" Then abertura = "Olá,"
    h = "<div style='font-family:Segoe UI,Arial,sans-serif;font-size:10.5pt;color:#1D1819'>"
    h = h & "<p>" & Esc(abertura) & "</p>"
    h = h & "<p style='margin:12px 0 8px 0;font-size:14pt;font-weight:bold;color:" & COR_TITULO & "'>" & Esc(TITULO) & " · " & Esc(periodo) & "</p>"
    h = h & "<table cellpadding='4' cellspacing='0' style='border-collapse:collapse;font-size:10pt'>"
    For Each item In Split(RESUMO, ";")
        partes = Split(item, "|")
        If UBound(partes) = 1 Then h = h & Linha(partes(0), "<b>" & Esc(pn.Range(partes(1)).Text) & "</b>")
    Next item
    h = h & "</table>"
    If AREA_ALERTAS <> "" Then
        Dim lista As String
        For Each c In pn.Range(AREA_ALERTAS).Cells
            If Trim$(c.Text) <> "" Then lista = lista & "<li>" & Esc(c.Text) & "</li>"
        Next c
        If lista <> "" Then h = h & "<p style='margin:12px 0 4px 0'><b>Pontos de atenção</b></p><ul style='margin-top:0'>" & lista & "</ul>"
    End If
    h = h & "<p><img src='cid:" & CID_IMG & "' width='" & LARGURA_IMG_EMAIL & "' alt='" & Esc(TITULO) & "'></p>"
    h = h & "<p style='color:#5F5759;font-size:9pt'>Em anexo: a aba do painel em Excel, com os valores do momento do envio.</p></div><br>"
    MontarCorpo = h
End Function

Private Function Linha(ByVal rot As String, ByVal valor As String) As String
    Linha = "<tr><td style='color:#5F5759;border-bottom:1px solid #E7E2E3;padding-right:16px'>" & Esc(rot) & "</td>" & _
            "<td style='border-bottom:1px solid #E7E2E3'>" & valor & "</td></tr>"
End Function

Private Function Esc(ByVal t As String) As String
    t = Replace(t, "&", "&amp;"): t = Replace(t, "<", "&lt;"): t = Replace(t, ">", "&gt;")
    Esc = t
End Function

Private Function NomeArquivo(ByVal t As String) As String
    Dim ch As Variant
    For Each ch In Array("\", "/", ":", "*", "?", """", "<", ">", "|", " · ", " ")
        t = Replace(t, ch, "_")
    Next ch
    NomeArquivo = t
End Function

Private Sub GravarTexto(ByVal arq As String, ByVal txt As String)
    Dim st As Object: Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open: st.WriteText txt: st.SaveToFile arq, 2: st.Close
End Sub
