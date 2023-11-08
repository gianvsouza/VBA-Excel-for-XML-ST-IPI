Attribute VB_Name = "ImportarXML"
Public Sub LerXML()
    ' Esta macro faz a leitura de um arquivo XML e registra os valores ST e IPI em uma planilha EXEL;
    ' Autor: Gian Vitor de Souza
    ' Empresa: OURIDIESEL
    ' Data: 30-10-2023
    ' Data de atualiza��o: 08-11-2023 (Incluimos uma condi��o If...End If selecionar o n� ICMS10 e IPITribut se existente)
    
    Dim XDoc As Object, root As Object, dest As Object, prod As Object
    Dim ws As Worksheet, i As Long, j As Long
    Dim filename As Variant
    
    Set XDoc = CreateObject("MSXML2.DOMDocument")
    XDoc.async = False: XDoc.validateOnParse = False
    
    filename = Application.GetOpenFilename(FileFilter:="Arquivos XML (*.xml), *.xml", Title:="Escolha um arquivo XML para abrir")
    
    If filename <> False Then
        XDoc.Load filename
    Else
        Exit Sub
    End If

    Set root = XDoc.DocumentElement
    Set ws = ThisWorkbook.Sheets("XML") 'altere aqui o nome da planilha onde voc� quer copiar os dados
            i = 2 'linha inicial da planilha
        For Each dest In root.SelectNodes("//nfeProc/NFe/infNFe/det") 'percorre todos os n�s det dentro de infNFe
            j = 1 'coluna inicial da planilha
                Set prod = dest.SelectSingleNode("prod") 'seleciona o n� prod dentro de det
                    ws.Cells(i, j) = prod.SelectSingleNode("cProd").Text 'copia o valor do n� cProd para a planilha
                        j = j + 1 'avan�a uma coluna na planilha
                    ws.Cells(i, j) = prod.SelectSingleNode("qCom").Text 'copia o valor do n� qCom para a planilha
                        j = j + 1 'avan�a uma coluna na planilha
                    ws.Cells(i, j) = prod.SelectSingleNode("vUnCom").Text 'copia o valor do n� vUnCom para a planilha
                        j = j + 1 'avan�a uma coluna na planilha
                    ws.Cells(i, j) = prod.SelectSingleNode("vProd").Text 'copia o valor do n� vProd para a planilha
                        j = j + 1 'avan�a uma coluna na planilha
                            'aqui come�a o bloco If...End If
                        If Not dest.SelectSingleNode("imposto/ICMS/ICMS10/CST") Is Nothing Then 'verifica se o n� ICMS10/CST existe
                            ws.Cells(i, j) = dest.SelectSingleNode("imposto/ICMS/ICMS10/CST").Text 'copia o valor do n� CST dentro de ICMS10 para a planilha
                                j = j + 1 'avan�a uma coluna na planilha
                            ws.Cells(i, j) = dest.SelectSingleNode("imposto/ICMS/ICMS10/vICMSST").Text 'copia o valor do n� vICMSST dentro de ICMS10 para a planilha
                                j = j + 1 'avan�a uma coluna na planilha
                        End If
                            'aqui termina o bloco If...ElseIf...End If
                            
                            'aqui come�a o bloco If...End If para o n� IPI
                        If Not dest.SelectSingleNode("imposto/IPI/IPITrib/vIPI") Is Nothing Then 'verifica se o n� IPITrib/vIPI existe
                            ws.Cells(i, j) = dest.SelectSingleNode("imposto/IPI/IPITrib/vIPI").Text 'copia o valor do n� vIPI dentro de IPITrib para a planilha
                        End If
            i = i + 1 'avan�a uma linha na planilha
        Next dest
End Sub
