@echo off
cd /d "%~dp0"
"C:\Users\cjbbe\anaconda3\python.exe" -c "import server; server.app.run(host='0.0.0.0', port=8899, debug=False)" > run_health_8899.out.log 2> run_health_8899.err.log
