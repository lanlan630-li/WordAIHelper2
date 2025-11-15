    @echo off
    title WordAI Server Startup
    echo ========================================
    echo Starting WordAI Server...
    echo ========================================
    cd /d C:\Users\27213\WordAIHelper
    
    echo.
    echo Running pm2 resurrect...
    pm2 resurrect
    
    echo.
    echo ========================================
    echo Task finished. This window will close in 10 seconds.
    echo You can check the server status by running 'pm2 list' in CMD.
    echo ========================================
    timeout /t 10 /nobreak
    exit
    