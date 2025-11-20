from typing import Callable, Dict, List

# Trata e padroniza os dados informados pelo tecnico da T.I no menu de opcoes. Exclui os itens duplicados
def _normalizar_nomes_alvo(texto: str) -> list[str]:
    
    # Aceita virgula, ponto-e-virgula e quebra de linha, ignora vazios e repeticoes
    bruto = texto.replace(";", ",").replace("\n", ",").strip()
    nomes = [p.strip() for p in bruto.split(",") if p.strip()]
    
    # Remove duplicados preservando ordem
    vistos, saida = set(), []
    for n in nomes:
        if n not in vistos:
            vistos.add(n)
            saida.append(n)
    return saida


# Exibe menu para o tecnico da T.I escolher os computadores para inventariar
def selecionar_computadores_por_menu(ous: Dict[str, str],listar_computadores_por_ou: Callable[[str], List[str]],) -> List[str]:
    """
    A funcao permite tres opcoes:
    1) Inventariar todas as maquinas de todas as Unidades Organizacionais informadas (buscando via PowerShell).
    2) Informar manualmente os nomes das maquinas.
    3) Cancelar
    
    Dependendo da escolha:
    - Opcao 1: chama 'listar_computadores' para cada OU e retorna a lista consolidada (sem duplicados).
    - Opcao 2: le e normaliza os nomes digitados via '_normalizar_nomes_alvo'.
    - Opcao 3: Cancela a operacao de rodar o script.
    """
    print("========================================================================")
    print(" Inventario de Computadores T.I - Selecao de Opcoes ")
    print("========================================================================")
    print("1) Inventariar TODA a planta - Todas as Unidades Organizacionais(OUs) listadas")
    print("2) Inventariar APENAS maquina(s) especifica(s)")
    print("3) Cancelar")
    print("------------------------------------------")
    print("Unidades Organizacionais(OUs) cadastradas:", ", ".join(ous.keys()))
    opcao = input("Escolha uma das opcoes 1, 2 ou 3: ").strip()
  
    # Opcao 1: Todas as Unidades Organizacionais
    if opcao == "1":
        computadores: List[str] = []
        for nome_ou, dn_ou in ous.items():
            try:
                print(f"[INFO] Listando computadores da Unidade Organizacional: {nome_ou} ...")
                computadores += listar_computadores_por_ou(dn_ou)
            except Exception as e:
                print(f"[AVISO] Falha ao listar '{nome_ou}': {e}")
                
        # remove vazios e duplicados
        vistos, saida = set(), []
        for c in computadores:
            if c and c not in vistos:
                vistos.add(c)
                saida.append(c)
        return saida

    # Opcao 2: Computadores Especificos
    if opcao == "2":
        print("\nDigite o(s) nome(s) da(s) maquina(s) separados por virgula:")
        print("Exemplos: D251D00113, D251D00114")

        entrada = input("Maquina(s): ").strip()
        nomes = _normalizar_nomes_alvo(entrada)
        
        if not nomes:
            print("[AVISO] Nenhum nome informado. Encerrando.")
        else:
            print(f"[INFO] Maquina(s) alvo: {', '.join(nomes)}")
        return nomes

    print("[AVISO] Opcao invalida ou cancelada. Encerrando.")
    return []