@echo off
cd C:\Taalex\Dashboard
py gerar_dados.py
py gerar_dados_cambio.py
py gerar_dados_financeiro.py
git add dados_dashboard.json dados_cambio.json dados_financeiro.json dashboard_temu.html dashboard_cambio.html dashboard_financeiro.html
git commit -m "atualizar dados %date% %time%"
git push origin main
echo Pronto! Dashboards atualizados.
