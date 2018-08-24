Attribute VB_Name = "CopiaDeBackup"
'Public Function BackupCopy()
'Dim fso As FileSystemObject
'Dim sSourcePath As String
'' Caminho de rede onde o arquivo principal está salvo
'
'Dim sSourceFile As String
''Nome do arquivo original
'
'Dim sBackupPath As String
''Caminho onde será feita a cópia
'
'Dim sBackupFile As String
''Nome do novo arquivo
'
'sSourcePath = "C:\Controle de Senhas\"
'sSourceFile = "BASEDEDADOS.MDB"
'sBackupPath = "C:\Controle de Senhas\BACKUP\"
'
'
'
'sBackupFile = "BackupBaseDeDados " & Format(Date, "yyyy-mm-dd ") & ".mdb"
''No nome de arquivo acima, coloquei para salvar com a data e  diariamente salvará um arquivo novo, mas isso pode
''ser feito por mês ou um único arquivo que será sobreposto toda vez que a função for chamada
'
'Set fso = New FileSystemObject
'
'fso.CopyFile sSourcePath & sSourceFile, sBackupPath & sBackupFile, True
'Set fso = Nothing
'
'Beep
'MsgBox "Backup realizado com sucesso em: " & Chr(13) & Chr(13) & sBackupPath & Chr(13) & Chr(13) & "O Nome do Backup é: " & Chr(13) & Chr(13) & sBackupFile, vbInformation, "Backup Completo"
'
'End Function
