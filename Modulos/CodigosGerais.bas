Attribute VB_Name = "CodigosGerais"
'PROCEDIMENTO PARA CONTAR OS ITENS DA LISTVIEW
Sub CONTARITENS()
usrPrincipal.CONTADOR_REGISTROS.Caption = usrPrincipal.lstTelefones.ListItems.Count
End Sub
'PROCEDIMENTO PARA CONTAR OS ITENS DA LISTVIEW NA STATUSBAR'
Sub CONTARITENSNALISTA()
usrPrincipal.StatusBar1.Panels(1).Text = "Registros: " & usrPrincipal.lstTelefones.ListItems.Count
End Sub

