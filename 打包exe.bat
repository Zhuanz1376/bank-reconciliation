@echo off
chcp 65001 >nul
REM ============================================================
REM  打包exe.bat —— 一键把银行对账系统打包成 Windows 可执行文件
REM  用法：在 Windows 电脑上双击运行（需已安装 Python 3.9+）
REM  产物：dist\BankRecon.exe（可复制给客户，改名为“银行对账系统.exe”即可）
REM ============================================================

echo [1/3] 安装依赖（openpyxl + pyinstaller）...
python -m pip install -r requirements.txt pyinstaller || goto :err

echo [2/3] 开始打包（-F 单文件  -w 无黑窗口，示例数据一并打包）...
python -m PyInstaller -F -w -n BankRecon ^
  --add-data "示例数据_单位银行存款日记账.csv;." ^
  --add-data "示例数据_银行对账单.csv;." ^
  --add-data "示例数据_期初余额调节表.csv;." ^
  银行对账系统.py || goto :err

echo [3/3] 完成！
echo   产物位置: dist\BankRecon.exe
echo   发给客户时可改名为: 银行对账系统.exe
pause
exit /b 0

:err
echo 打包失败，请检查上方错误信息。
pause
exit /b 1
