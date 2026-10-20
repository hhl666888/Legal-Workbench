@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 正在启动合同审阅智能体...
if not exist .env if not exist .env.example (
  echo.
  echo 未找到 .env.example 配置文件。请先在 .env.example 中填写 DEEPSEEK_API_KEY。
  pause
  exit /b 1
)
node server.js
pause
