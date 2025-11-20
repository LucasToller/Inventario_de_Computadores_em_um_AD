"""
Inventario de Computadores do setor de Tecnologia da Informacao - Joyson Safety Systems (Balneario Picarras)

Objetivo:
    - Coletar informacoes de computadores via AD (unidade 251 - Balneario Picarras)
    - Salvar inventario em TXT detalhadamente por maquina + XLSX (planilha) com todas as informacoes coletadas 
    - Coleta informacoes detalhadas de hardware, sistema e rede
    - Facilitar no mapeamento das informacoes para gestao da T.I

Requisitos:
    - Python 3.x
    - Biblioteca openpyxl instalada > pip install openpyxl
    - WinRM (Windows Remote Management) habilitado + permissoes administrativas
    - RSAT (Ferramentas de Administracao de Servidor Remoto) instalado (para uso do Get-ADComputer e Get-ADUser no PowerShell)
    - Ajustar o caminho de saida da planilha/TXT conforme planta/unidade

Autor:
    Lucas Toller Gutmann - Estagiario TI Joyson Balneaio Picarras cursando Ciencia da Computacao na Univali Campus Itajai
"""

import os
import locale
from datetime import datetime
from functools import partial

import execucao_powershell
from menu import selecionar_computadores_por_menu
from configuracoes import carregar_configuracoes
from tratamento_dados import coletar_dados_computador
from saida_relatorios import gerar_txt_por_computador, gerar_planilha

try:
    locale.setlocale(locale.LC_TIME, "pt_BR.utf8")
except locale.Error:
    pass


# Caminhos do projeto - ("S:\\IT\\COMUM\\Inventario_TI\src\python\main.py")
BASE_DIRETORIO = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
CAMINHO_SCRIPT_POWERSHELL_COLETA = os.path.join(BASE_DIRETORIO, "src", "powershell", "coleta_inventario.ps1")
CAMINHO_SCRIPT_POWERSHELL_LISTAR_COMPUTADORES = os.path.join(BASE_DIRETORIO, "src", "powershell", "listar_computadores.ps1")


# Funcao principal
def main():
    
    # Carrega as configuracoes
    configuracoes = carregar_configuracoes(BASE_DIRETORIO)

    # Caminhos saidas
    pasta_saidas = configuracoes["caminhos"]["pasta_saida_principal"]
    pasta_planilhas = os.path.join(pasta_saidas, configuracoes["caminhos"]["subpasta_planilhas"])
    pasta_txts = os.path.join(pasta_saidas, configuracoes["caminhos"]["subpasta_txts"])
    pasta_falhas = os.path.join(pasta_saidas, configuracoes["caminhos"]["subpasta_falhas"])
    os.makedirs(pasta_planilhas, exist_ok=True)
    os.makedirs(pasta_txts, exist_ok=True)
    os.makedirs(pasta_falhas, exist_ok=True)

    # Unidades Organizacionais (OUs) definidas no arquivo JSON
    ous = dict(configuracoes["unidades_organizacionais"])

    # Lista os computadores da OU
    listar_computadores_da_ou = partial(
        execucao_powershell.listar_computadores, 
        caminho_script_listar_computadores=CAMINHO_SCRIPT_POWERSHELL_LISTAR_COMPUTADORES
    )

    # Selecao via menu 
    computadores = selecionar_computadores_por_menu(ous, listar_computadores_da_ou)
    if not computadores:
        print("[INFO] Nenhum computador selecionado. Encerrando.")
        return 
    
    total_computadores = len(computadores)
    print(f"\n[INFO] {total_computadores} computadores encontrados ou selecionados para inventariar.\n")

    data_execucao = datetime.now().strftime("%d-%m-%Y_%H-%M")
    caminho_xlsx = os.path.join(pasta_planilhas, f"Computadores_{data_execucao}.xlsx")

    todos_dados = []
    for idx, computador in enumerate(computadores, start=1):
        print(f"[INFO] ({idx}/{total_computadores}) Coletando dados de {computador} ...") #Imprime no terminal em qual computador esta do total. Exemplo: 20/130
        try:
            dados = coletar_dados_computador(computador, CAMINHO_SCRIPT_POWERSHELL_COLETA, BASE_DIRETORIO)
            
            # Filtro de dados vazios
            campos_chave = ["Fabricante", "Modelo", "SO", "Processador"]
            if all(dados.get(c, "") in ["", "N/D", None] for c in campos_chave):
                caminho_log = os.path.join(pasta_falhas, "falhas_captura_de_dados.txt")
                with open(caminho_log, "a", encoding="utf-8") as log:
                    log.write(f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Falha na captura de dados: {computador}\n")
                    log.write("Motivo: Dados vazios retornados pela sessao remota.\n" + "-"*60 + "\n")
                continue

            todos_dados.append(dados)
            
            
            # TXT por maquina
            gerar_txt_por_computador(dados, pasta_txts, computador, data_execucao)

        except Exception as e:
            caminho_log = os.path.join(pasta_falhas, "falhas_captura_de_dados.txt")
            with open(caminho_log, "a", encoding="utf-8") as log:
                log.write(f"[{datetime.now().strftime('%d/%m/%Y %H:%M:%S')}] Falha ao conectar com: {computador}\n")
                log.write(f"Erro: {str(e)}\n" + "-"*60 + "\n")
            continue
        
        
    # Planilha - Apos finalizar a coleta dos dados de todos os computadores
    gerar_planilha(todos_dados, caminho_xlsx)
    print(f"\n[OK] Planilha gerada salva em: {caminho_xlsx}\n")


# Execucao
if __name__ == "__main__":
    main() 