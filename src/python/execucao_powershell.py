import os
import subprocess
import json
import platform

# Funcao que remove dominio e normaliza para comparacao
def normalizar_nome_computador(nome: str) -> str:
    nome_normalizado = (nome or "").strip().lower()
    partes_nome = nome_normalizado.split(".")
    
    # Se parece IPv4 (4 partes so com digitos), mantem inteiro
    if len(partes_nome) == 4 and all(p.isdigit() for p in partes_nome):
        return nome_normalizado
    
    # Caso contrario, remove dominio
    return partes_nome[0]


# Funcao para saber se o alvo e a propria maquina (host que esta rodando o script)
def eh_maquina_local(nome_alvo: str) -> bool:
    nomes_locais = {
        normalizar_nome_computador(os.environ.get("COMPUTERNAME")),
        normalizar_nome_computador(platform.node()),
        normalizar_nome_computador("localhost"),
        normalizar_nome_computador("127.0.0.1"),
    }
    return normalizar_nome_computador(nome_alvo) in nomes_locais


# Executa o script PowerShell: local se for a propria maquina, WinRM caso contrario
def executar_script_remoto(computador: str, caminho_script_powershell: str) -> str:
    if not caminho_script_powershell:
        raise FileNotFoundError("Caminho do script PowerShell nao encontrado.")
    
    alvo_eh_local = eh_maquina_local(computador)
    
    caminho_powershell_escapado = caminho_script_powershell.replace("'", "''")
    
    if alvo_eh_local:
        
         # Execucao LOCAL (sem WinRM) para o proprio host
         print(f"[INFO] Executando script localmente em '{computador}' ...")
         comando_powershell = [
              "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
              "-File", caminho_script_powershell
         ]
    else:
        
        # Execucao REMOTA via WinRM
        print(f"[INFO] Executando script remotamente em '{computador}' ...")
        comando_powershell = [
        "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
        f"Invoke-Command -ComputerName '{computador}' -FilePath '{caminho_powershell_escapado}' -ErrorAction Stop"      
        ]
    
    try:
        resultado = subprocess.run(comando_powershell, capture_output=True, text=True, timeout=120)
        saida = (resultado.stdout or "").strip()
    
        if (resultado.returncode != 0 or not saida) and alvo_eh_local:
            comando_contingencia = [
                "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command",
                f"$ErrorActionPreference='Stop'; & '{caminho_powershell_escapado}' | Out-String"
            ]
         
            resultado = subprocess.run(comando_contingencia, capture_output=True, text=True, timeout=120)
            saida = (resultado.stdout or "").strip()
    
    except subprocess.TimeoutExpired:
        acao_execucao = "coletar localmente" if alvo_eh_local else "executar script remoto"
        raise RuntimeError(f"Tempo esgotado ao {acao_execucao} em {computador}.")

    if resultado.returncode != 0 or not saida:
        erro = (resultado.stderr or "").strip() or "sem detalhes"
        modo_execucao = "local" if alvo_eh_local else "remoto"
        raise RuntimeError(f"Falha ao executar script ({modo_execucao}) em {computador}: {erro}")
    
    return saida
    

# Lista computadores no AD dentro de uma Unidade Organizacional (OU)
def listar_computadores(dn_unidade_organizacional: str, caminho_script_listar_computadores: str) -> list[str]:
    """
    Retorna lista de computadores da OU especificada no AD.
    """
    comando_powershell = [
        "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", caminho_script_listar_computadores, "-SearchBase", dn_unidade_organizacional
    ]
    
    try:
        resultado = subprocess.run(comando_powershell, capture_output=True, text=True, timeout=60)
    except subprocess.TimeoutExpired:
        raise RuntimeError("Tempo esgotado ao listar computadores da OU.")
    
    if resultado.returncode != 0:
        raise RuntimeError(f"Falha ao listar computadores: {resultado.stderr.strip()}")
    return [linha.strip() for linha in resultado.stdout.splitlines() if linha.strip()]


# Nome completo do usuario no AD
def obter_nome_usuario(usuario_ad: str, caminho_script_obter_nome: str) -> str:
    if not usuario_ad or usuario_ad == "N/D":
        return "N/D"
    comando_powershell = [
        "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
        "-File", caminho_script_obter_nome, "-Sam", usuario_ad, "-ErrorAction", "SilentlyContinue"
    ]
    
    try:
        resultado = subprocess.run(comando_powershell, capture_output=True, text=True, timeout=30)
    except subprocess.TimeoutExpired:
        return "N/D"
    
    saida = (resultado.stdout or "").strip()
    return saida if resultado.returncode == 0 and saida else "N/D"

# Converte saida JSON > Python
def converter_json_para_python(texto_json: str):
    if not texto_json:
        return {}
    try:
        json_sanitizado = texto_json.lstrip("\ufeff").strip()
        return json.loads(json_sanitizado) if json_sanitizado else {}
    except json.JSONDecodeError:
        return {}