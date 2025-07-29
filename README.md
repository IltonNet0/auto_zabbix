# 🖥️ Automação de Hosts no Zabbix via Selenium

Este script automatiza tarefas administrativas no Zabbix, como clonagem de hosts, adição de grupos e verificação de duplicidade, utilizando `undetected_chromedriver`, `selenium` e `pandas`. Ele é especialmente útil para gerenciar uma grande quantidade de dispositivos com base em planilhas do Excel.

## 📌 Funcionalidades

- Login automático na interface do Zabbix
- Clonagem de hosts existentes com novos nomes/IPs
- Adição de grupos/categorias a hosts
- Verificação de hosts duplicados por IP
- Atualização automática da planilha de origem com status

## 🧰 Tecnologias Utilizadas

- [Python 3.8+]
- [Selenium]
- [undetected-chromedriver]
- [pandas]
- [openpyxl]

## 📁 Requisitos

- Google Chrome compatível com a versão usada no `undetected_chromedriver` (por padrão: `version_main=138`)
- WebDriver adequado (instalado automaticamente com `undetected_chromedriver`)
- Arquivo Excel: `MW CORP e GOV.xlsx` com a aba `Inclusão 2025 (PRTG+Zabbix) - N`
