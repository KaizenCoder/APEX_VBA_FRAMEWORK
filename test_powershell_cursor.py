import subprocess

subprocess.run([
    'powershell', '-ExecutionPolicy', 'Bypass',
    '-Command', 'Set-Content -Path "C:\\Temp\\test_cursor_full.txt" -Value "Bonjour depuis Cursor PowerShell !"'
])
