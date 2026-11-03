Attribute VB_Name = "mdlCriaBackup"
Option Explicit

Sub SaveWorkbookBackup()

Dim AWB As Workbook, BackupFileName As String, i As Long, ok As Boolean

On Error GoTo NotAbleToSave
    
Set AWB = ActiveWorkbook

'Atribuir caminho completo do arquivo ao longo do nome do arquivo para a variável BackupFileName
BackupFileName = AWB.FullName

'Verificando se o arquivo é salvo
'Se o arquivo não é salvo, em seguida, salvando o arquivo
If AWB.path = "" Then
    'Exibindo Salvar como caixa de diálogo para salvar arquivos
    Application.Dialogs(xlDialogSaveAs).Show
Else
    
    'Removendo a extensão de arquivo do nome do arquivo
    i = 0
    While InStr(i + 1, BackupFileName, ".") > 0
    'Encontre a extensão do arquivo
        i = InStr(i + 1, BackupFileName, ".")
    Wend
    
    If i > 0 Then BackupFileName = Left(BackupFileName, i - 1)
    
    'Adicionando back up de extensão ".bak" com nome de arquivo
    BackupFileName = BackupFileName & ".bak"
    ok = False
    
    With AWB
        .Save
        'Criando Backup de arquivo
        .SaveCopyAs BackupFileName
        ok = True
    End With
End If

NotAbleToSave:
'Código para tratamento de erros
    Set AWB = Nothing
    If Not ok Then
        MsgBox "Cópia de backup não salva!", vbExclamation, ThisWorkbook.Name
    End If
End Sub

Sub SaveWorkbookBackupToFloppy()

Dim AWB As Workbook, BackupFileName As String, i As Long, ok As Boolean
Dim DriveName As String

On Error GoTo NotAbleToSave
    
'Especifique o caminho para fazer backup em unidade
DriveName = "D:\"

'Inicializando as variáveis
Set AWB = ActiveWorkbook
BackupFileName = AWB.Name
ok = False

'Verificando se o arquivo é salvo
'Se o arquivo não é salvo, em seguida, salvando o arquivo
If AWB.path = "" Then
    'Exibindo Salvar como caixa de diálogo para salvar arquivos
    Application.Dialogs(xlDialogSaveAs).Show
Else
    ' Nunca apaga o backup anterior antes de tentar gravar a nova cópia.
    If StrComp(AWB.FullName, DriveName & BackupFileName, vbTextCompare) = 0 Then
        Err.Raise vbObjectError + 525, , "O destino do backup é o próprio arquivo aberto."
    End If
    With AWB
        .Save
        'Criando o arquivo de backup
        .SaveCopyAs DriveName & BackupFileName
        ok = True
    End With
End If

NotAbleToSave:
'Código para tratamento de erros
    Set AWB = Nothing
    
    If Not ok Then
        MsgBox "Cópia de backup não salva!", vbExclamation, ThisWorkbook.Name
    End If
End Sub

