$fullPath = 'C:\Program Files\WindowsPowerShell\Modules\ImportExcel'

Robocopy . $fullPath /mir /XD .vscode .git examples testimonials images spikes /XF appveyor.yml .gitattributes .gitignore