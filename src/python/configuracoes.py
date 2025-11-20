import os
import json
from typing import Dict

# Funcao para carregar e validar as configuracoes definidas em: config\configuracoes.json
def carregar_configuracoes(base_diretorio: str) -> Dict:
    caminho_configuracoes = os.path.join(base_diretorio, "config", "configuracoes.json")
    
    # Le o arquivo JSON
    try:
        with open(caminho_configuracoes, "r", encoding="utf-8") as f:
            configuracoes = json.load(f)
    except FileNotFoundError:
        raise FileNotFoundError("Arquivo 'config/configuracoes.json' nao encontrado.")
    except json.JSONDecodeError as erro:
        raise ValueError(f"JSON de configuracoes invalido: {erro}")
    
    # Valida a estrutura minima
    faltas = []
    if "caminhos" not in configuracoes:
        faltas.append("caminhos")
    else:
        for k in ("pasta_saida_principal", "subpasta_planilhas", "subpasta_txts", "subpasta_falhas"):
            if not configuracoes["caminhos"].get(k):
                faltas.append(f"caminhos.{k}")
                
    if "unidades_organizacionais" not in configuracoes or not isinstance(configuracoes["unidades_organizacionais"], dict) or not configuracoes["unidades_organizacionais"]:
        faltas.append("unidades_organizacionais (minimo 1 OU)")
    
    if faltas:
        raise ValueError("Configuracoes ausentes no JSON: " + ", ".join(faltas))
    
    # Execucao
    if "execucao" not in configuracoes or "max_workers" not in configuracoes["execucao"]:
        configuracoes.setdefault("execucao", {}).setdefault("max_workers", 1)
        
    return configuracoes 