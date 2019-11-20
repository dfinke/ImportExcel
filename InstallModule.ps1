$fullPath = 'C:\Program Files\WindowsPowerShell\Modules\ImportExcel'

Robocopy . $fullPath /mir /XD .vscode .git examples data /XF appveyor.yml azure-pipelines.yml .gitattributes .gitignore