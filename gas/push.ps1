# push Apps Script โดยอัตโนมัติ
Copy-Item "..\concrete-webap.js" ".\Code.js" -Force
clasp push --force
Write-Host "✅ Push สำเร็จ!" -ForegroundColor Green
