Attribute VB_Name = "ImportarXML"
Public Sub LerXML()
    ' Esta macro faz a leitura de um arquivo XML e registra os valores ST e IPI em uma planilha EXEL;
    ' Autor: Gian Vitor de Souza
    ' Empresa: OURIDIESEL
    ' Data: 30-10-2023
    ' Data de atualização: 08-11-2023 (Incluimos uma condição If...End If selecionar o nó ICMS10 e IPITribut se existente)
    
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
    Set ws = ThisWorkbook.Sheets("XML") 'altere aqui o nome da planilha onde você quer copiar os dados
            i = 2 'linha inicial da planilha
        For Each dest In root.SelectNodes("//nfeProc/NFe/infNFe/det") 'percorre todos os nós det dentro de infNFe
            j = 1 'coluna inicial da planilha
                Set prod = dest.SelectSingleNode("prod") 'seleciona o nó prod dentro de det
                    ws.Cells(i, j) = prod.SelectSingleNode("cProd").Text 'copia o valor do nó cProd para a planilha
                        j = j + 1 'avança uma coluna na planilha
                    ws.Cells(i, j) = prod.SelectSingleNode("qCom").Text 'copia o valor do nó qCom para a planilha
                        j = j + 1 'avança uma coluna na planilha
                    ws.Cells(i, j) = prod.SelectSingleNode("vUnCom").Text 'copia o valor do nó vUnCom para a planilha
                        j = j + 1 'avança uma coluna na planilha
                    ws.Cells(i, j) = prod.SelectSingleNode("vProd").Text 'copia o valor do nó vProd para a planilha
                        j = j + 1 'avança uma coluna na planilha
                            'aqui começa o bloco If...End If
                        If Not dest.SelectSingleNode("imposto/ICMS/ICMS10/CST") Is Nothing Then 'verifica se o nó ICMS10/CST existe
                            ws.Cells(i, j) = dest.SelectSingleNode("imposto/ICMS/ICMS10/CST").Text 'copia o valor do nó CST dentro de ICMS10 para a planilha
                                j = j + 1 'avança uma coluna na planilha
                            ws.Cells(i, j) = dest.SelectSingleNode("imposto/ICMS/ICMS10/vICMSST").Text 'copia o valor do nó vICMSST dentro de ICMS10 para a planilha
                                j = j + 1 'avança uma coluna na planilha
                        End If
                            'aqui termina o bloco If...ElseIf...End If
                            
                            'aqui começa o bloco If...End If para o nó IPI
                        If Not dest.SelectSingleNode("imposto/IPI/IPITrib/vIPI") Is Nothing Then 'verifica se o nó IPITrib/vIPI existe
                            ws.Cells(i, j) = dest.SelectSingleNode("imposto/IPI/IPITrib/vIPI").Text 'copia o valor do nó vIPI dentro de IPITrib para a planilha
                        End If
            i = i + 1 'avança uma linha na planilha
        Next dest
End Sub
