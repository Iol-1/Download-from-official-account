@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在打包，请稍候...
pyinstaller --noconfirm --noconsole --onefile ^
    --icon=icon.png ^
    --add-data "icon.png;." ^
    --name "公众号文章保存工具" ^
    --hidden-import=lxml ^
    --hidden-import=lxml.etree ^
    --hidden-import=lxml.html ^
    --hidden-import=docx ^
    --hidden-import=docx.oxml ^
    --hidden-import=docx.oxml.ns ^
    --hidden-import=bs4 ^
    --hidden-import=PIL ^
    wechat_to_word.py
echo.
if exist "dist\公众号文章保存工具.exe" (
    echo 打包成功！
    echo 文件位置: dist\公众号文章保存工具.exe
    echo.
    echo 正在复制到当前目录...
    copy /Y "dist\公众号文章保存工具.exe" "." >nul
    echo 已复制到: %~dp0公众号文章保存工具.exe
) else (
    echo 打包失败，请检查错误信息。
)
pause
