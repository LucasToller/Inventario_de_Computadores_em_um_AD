# Inventário de Computadores 

Coleta automatizada de informações de hardware, sistema operacional e rede de estações Windows via **PowerShell Remoting (WinRM)**, consolidando os resultados em **planilha XLSX** e **TXTs individuais** por máquina.

> **Objetivo:** facilitar o mapeamento e a gestão de ativos de computadores para TI da unidade **251 – Balneário Piçarras**.
> **Observação (v1.0.0):** a coleta de **impressoras foi removida** nesta versão, pois não estavam retornando as impressoras mapeadas no servidor de impressão

> **Versão:** 1.0.0 

---

## Sumário

- [Arquitetura & Fluxo](#arquitetura--fluxo)
- [Estrutura do Repositório](#estrutura-do-repositório)
- [Requisitos](#requisitos)
- [Configuração (config/configuracoes.json)](#configuração-configconfiguracoesjson)
- [Como Executar](#como-executar)
- [Saídas Geradas](#saídas-geradas)
- [Campos Coletados](#campos-coletados)
- [Personalizações Comuns](#personalizações-comuns)
- [Solução de Problemas (FAQ)](#solução-de-problemas-faq)
- [Roadmap v2.0](#roadmap-v110)
- [Créditos](#créditos)
- [Licença](#licença)

---

## Arquitetura & Fluxo

1. **Python** (máquina de administração) lê as configurações do arquivo **`config/configuracoes.json`**.
2. A partir das OUs definidas no JSON, o Python lista os computadores do AD (via `Get-ADComputer`).
3. Para cada host:
    - Se o alvo **é a máquina local**, o script **executa localmente** o `coleta_inventario.ps1` com `-File`.
    - Caso contrário, executa **remotamente** via **WinRM** usando `Invoke-Command`.
    - **Contingência local:** se a execução local por `-File` não retornar saída, é feita **reexecução local alternativa**
    via `-Command` + `& 'script'` (fallback para cenários em que o `-File` não imprime stdout corretamente).
4. O PowerShell devolve um **JSON padronizado** (CS, OS, BIOS, etc.).
5. O Python trata/normaliza e gera:
   - **TXT** detalhado por máquina;
   - **XLSX** com visão consolidada (com cabeçalho congelado, AutoFilter e formatações numéricas).

---

## Estrutura do Repositório

```
InventarioTI/
├─ config/
│  └─ configuracoes.json          # único ponto de ajuste (pastas de saída, OUs, execucao)
├─ docs/
│  └─ imagens/                    # screenshots para documentação (opcional)
├─ saidas/                        # diretório pai definido no JSON
│  ├─ planilhas/                  # .xlsx gerados
│  ├─ txts/                       # relatórios .txt por máquina
│  └─ falhas_captura_de_dados/    # logs de falhas por host
├─ src/
│  ├─ powershell/
│  │  ├─ coleta_inventario.ps1       # coleta (retorna JSON padronizado)
│  │  ├─ listar_computadores.ps1     # lista computadores de uma OU
│  │  └─ obter_nome_usuario.ps1      # resolve DisplayName no AD
│  └─ python/
│     ├─ main.py                      # orquestração principal
│     ├─ execucao_powershell.py       # local vs remoto (WinRM) + utilidades
│     ├─ tratamento_dados.py          # normalizações/formatos dos campos
│     ├─ saida_relatorios.py          # geração de TXT e XLSX
│     ├─ menu.py                      # seleção por OUs ou nomes específicos
│     └─ configuracoes.py             # loader/validador do JSON (separado do main)
└─ README.md
```

> Os nomes dos arquivos podem variar conforme sua organização, mas a separação por **responsabilidade** é recomendada.

---

## Requisitos

**Máquina de administração (onde roda o Python):**
- Windows 10/11 ou Windows Server
- **Python 3.x** (recomendado 3.10+)
- Biblioteca Python: `openpyxl`
- **RSAT** instalado (módulo Active Directory) para `Get-ADComputer` e `Get-ADUser`
- PowerShell com permissão para remoting (cliente)

**Estações/servidores alvo:**
- **WinRM habilitado** e acessível (GPO recomendada)
- Conta de execução com permissões suficientes (WMI/CIM/WinRM)

---

## Configuração (`config/configuracoes.json`)

Todas as **configurações editáveis** ficam centralizadas neste arquivo:

```json
{
  "caminhos": {
    "pasta_saida_principal": "\\\\servidor\\compartilhado\\InventarioTI\\saidas",
    "subpasta_planilhas": "planilhas",
    "subpasta_txts": "txts",
    "subpasta_falhas": "falhas_captura_de_dados"
  },
  "unidades_organizacionais": {
    "Desktop": "OU=Desktop,OU=Computers,DC=empresa,DC=local",
    "Laptop": "OU=Laptop,OU=Computers,DC=empresa,DC=local",
    "ShopFloor": "OU=ShopFloor,OU=Computers,DC=empresa,DC=local"
  },
  "execucao": {
    "max_workers": 1
  }
}
```

> **Observações:**
> - **Pastas de saída**: a planilha e os TXTs são gravados dentro de `pasta_saida_principal` usando os nomes de subpastas.
> - **OUs**: ajuste os DNs conforme a planta/empresa-alvo.

---

## Como Executar

Rode a partir da pasta `src/python`:

```powershell
cd src\python
python .\main.py
```

- Escolha no **menu** entre: inventariar todas as OUs configuradas ou informar máquinas específicas.
- A coleta da **máquina local** (onde o script Python está rodando) é suportada automaticamente.

---

## Saídas Geradas

- **Planilha (.xlsx)**: `Computadores_DD-MM-YYYY_HH-MM.xlsx` (cabeçalho congelado, AutoFilter, formatação numérica/percentual)
- **Relatórios .txt por máquina**: `HOST_DD-MM-YYYY_HH-MM.txt`

Exemplo de trechos no TXT:

```
[IDENTIFICACAO]
Fabricante: Dell
Modelo: BP-061F
Numero de Serie: XXXXXXXX

[ARMAZENAMENTO]
Armazenamento Total do Disco C: 222.46 GB
Utilizado: 83.89 GB (37.71%)
Disponivel: 138.56 GB (62.29%)
Tipo de Armazenamento: SSD (223.57 GB), SSD (29.12 GB)
```

---

## Campos Coletados

- **Usuário:** `Usuario_Logado`, `Usuario_Nome_Completo`  
- **Identificação:** `Computador`, `Fabricante`, `Modelo`, `Numero_Serie`  
- **Sistema:** `SO`, `Versao_SO`, `Data_Instalacao_SO`  
- **Boot:** `Ultimo_Boot`, `Tempo_Inicializacao`, `Tempo_Total_Ligado`  
- **Processador/GPU:** `Processador`, `Placa_Video`  
- **Memória:** `Qtd_Modulos_RAM`, `Total_GB_RAM`, `Geracao_RAM`  
- **Armazenamento (C:)**: `DiscoC_Total_GB`, `DiscoC_Usado_GB`, `DiscoC_Uso_Pct`, `DiscoC_Livre_GB`, `DiscoC_Livre_Pct`  
- **Tipos de disco físico:** `Tipo_Armazenamento`  
- **Rede:** `IP_Principal`, `MAC`, `Velocidade_Rede`, `Placa_Rede`

> Em caso de valores inválidos de série (ex.: “default string”), o script normaliza e retorna **N/D**.

---

## Personalizações Comuns

- **Pastas de saída:** ajuste apenas o **JSON**; o `main.py` lê tudo de `config/configuracoes.json`.
- **Locale/PT-BR:** o script tenta `pt_BR.utf8`; se sua máquina não tiver, pode usar `Portuguese_Brazil.1252`.
- **Campos adicionais:** para coletar mais itens (ex.: BaseBoard/CSProduct), adicione no `coleta_inventario.ps1` e mapeie no `tratamento_dados.py`.

---

## Solução de Problemas (FAQ)

**1) `JSONDecodeError` ou dados vazios**  
• Geralmente é falha de WinRM/credencial. Valide conectividade: `Test-WsMan <hostname>`  
• Verifique permissões administrativas no alvo.  
• Confirme se `coleta_inventario.ps1` existe no caminho esperado.

**2) `Get-ADComputer` não encontrado**  
• Instale **RSAT** e habilite o módulo ActiveDirectory na máquina de administração.

**3) `Ultimo_Boot`/`Tempo_Total_Ligado` aparecem “N/D”**  
• Alguns hosts retornam `LastBootUpTime` em formato inesperado. O Python já tenta normalizar; se persistir, valide WMI/CIM do alvo.

**4) Número de série “default string”**  
• Placeholder de fabricante no SMBIOS. O script filtra automaticamente e retorna **N/D** para evitar falsos positivos.

**5) Execução bloqueada por política**  
• A chamada usa `-ExecutionPolicy Bypass` na sessão. Algumas GPOs podem bloquear: alinhe com a equipe de AD.

**6) Coleta na máquina local não retorna dados**
• O projeto possui **contingência local** (reexecução por `-Command`).
• Garanta que o caminho do `coleta_inventario.ps1` é válido e acessível pelo Python.
• Teste rápido no PowerShell: `& "<caminho>\coleta_inventario.ps1" | ConvertTo-Json -Depth 4`

---

## Roadmap v1.1.0 - Versão de atualização com novas funcionalidades do script

- Execução **paralela** (usar `execucao.max_workers` do JSON).
- Planilhas: exibir totais (sucesso/falha) ao final da coleta.
- Gerar **dashboard** automático (exemplo: Pandas + gráficos) a partir do XLSX.

---

## Créditos

- Autor: **Lucas Toller Gutmann** – Estagiário TI Joyson Balneário Piçarras/ Estudante de Ciência da Computação na Univali - Câmpus Itajaí 

---

## Licença

**Sem Licenças**.
