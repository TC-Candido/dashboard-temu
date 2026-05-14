@echo off
cd C:\Taalex\Dashboard
py gerar_dados.py
py gerar_dados_cambio.py
git add dados_dashboard.json dados_cambio.json dashboard_temu.html dashboard_cambio.html
git commit -m "atualizar dados %date% %time%"
git push origin main
echo Pronto! Dashboards atualizados.
