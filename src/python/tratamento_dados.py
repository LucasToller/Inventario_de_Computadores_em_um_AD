import os
from datetime import datetime
import execucao_powershell

# Funcao: Coleta e trata dados individualmente de cada computador coletado
def coletar_dados_computador(computador: str, CAMINHO_SCRIPT_POWERSHELL: str, BASE_DIRETORIO: str):
    
    dados = {"Computador": computador}
    try:
        
        # Retorno dos dados captados pelo script em PowerShell, de forma bruta (sem nenhum tratamento)
        retorno_dados_coletados_powershell = execucao_powershell.executar_script_remoto(computador, CAMINHO_SCRIPT_POWERSHELL)
        dados_brutos = execucao_powershell.converter_json_para_python(retorno_dados_coletados_powershell)
        if not dados_brutos:
            raise Exception("Sessao remota retornou nula (maquina offline ou sem permissao).")


        # Tratamento dos dados  
        dados_computador = dados_brutos.get("CS", {})
        dados_sistema_operacional = dados_brutos.get("OS", {})
        dados_bios = dados_brutos.get("BIOS", {})
        dados_inicializacao = dados_brutos.get("Boot", {})
        dados_processador = dados_brutos.get("CPU", {})
        dados_placa_video = dados_brutos.get("GPU", [])
        dados_memoria_ram = dados_brutos.get("RAM", [])
        dados_armazenamento = dados_brutos.get("Volumes", [])
        dados_discos = dados_brutos.get("Discs", [])
        dados_rede = dados_brutos.get("Net", {})
        dados_ip = dados_brutos.get("IP", {})
        dados_placa_rede = dados_brutos.get("Net", [])
        
        
        # Nome completo do usuario no AD
        CAMINHO_POWERSHELL_OBTER_NOME_USUARIO = os.path.join(BASE_DIRETORIO, "src", "powershell", "obter_nome_usuario.ps1")
        usuario_ad = str(dados_computador.get("UserName") or "").strip()
        dados ["Usuario_Logado"] = (usuario_ad.split("\\")[-1] if usuario_ad else "N/D")
        dados["Usuario_Nome_Completo"] = (execucao_powershell.obter_nome_usuario(dados.get("Usuario_Logado", "N/D"), CAMINHO_POWERSHELL_OBTER_NOME_USUARIO)
            if dados["Usuario_Logado"] not in ["", "N/D"] else "N/D")
        
            
        # Identificacao do computador - Fabricante, modelo e numero de serie
        fabricante = str(dados_computador.get("Manufacturer") or "").strip()
        modelo = str(dados_computador.get("Model") or "").strip()
        serie_bios = str(dados_bios.get("SerialNumber") or "").strip()
        
        # Filtra valores invalidos/placeholder que alguns fabricantes gravam no SMBIOS
        espacos_reservados = {
        "", "default string", "to be filled by o.e.m.", "to be filled by oem",
        "system serial number", "not available", "unavailable", "n/a", "none", "unknown"
        }
        
        serie_bios_normalizada = serie_bios.lower().replace(".", "").replace(" ", "")
        serie_invalida = (
            serie_bios_normalizada in {p.replace(" ", "").replace(".", "") for p in espacos_reservados} 
            or (serie_bios != "" and serie_bios.strip("0") == "") 
            or (len(serie_bios) >= 6 and len(set(serie_bios)) == 1))  
        
        numero_serie_backup = str(dados_brutos.get("CSProduct", {}).get("IdentifyingNumber") or "").strip()
        serie_final = serie_bios if not serie_invalida and serie_bios else (numero_serie_backup or "N/D")
        dados["Fabricante"] = fabricante or "N/D"
        dados["Modelo"] = modelo or "N/D"
        dados["Numero_Serie"] = serie_final
        
        
        # Sistema Operacional - SO, versao do SO e data de instalacao do sistema operacional
        dados["SO"] = dados_sistema_operacional.get("Caption", "N/D")
        dados["Versao_SO"] = dados_sistema_operacional.get("Version", "N/D")
        dados["Data_Instalacao_SO"] = dados_sistema_operacional.get("InstallDate", "N/D")
        
        data_instalacao = dados["Data_Instalacao_SO"]
        if isinstance(data_instalacao, str) and "Date(" in data_instalacao:
            try:
                # Extrai milissegundos e converte corretamente
                timestamp_ms = int(data_instalacao.split("(")[1].split(")")[0])
                data_instalacao_real = datetime.fromtimestamp(timestamp_ms / 1000)
                dados["Data_Instalacao_SO"] = data_instalacao_real.strftime("%A, %d de %B de %Y %H:%M:%S")
            except Exception:
                dados["Data_Instalacao_SO"] = "N/D"


        # Ultimo boot e tempo ligado - Data da ultima vez que foi inicializado e o tempo total ligado
        dados["Ultimo_Boot"] = "N/D"
        dados["Tempo_Total_Ligado"] = "N/D"

        # Fontes brutas: evento WMI
        carimbo_evento_boot = str(dados_inicializacao.get("TimeCreated", "") or "").strip()
        carimbo_wmi_boot = str(dados_sistema_operacional.get("LastBootUpTime", "") or "").strip()

        carimbo_boot_preferencial = carimbo_evento_boot or carimbo_wmi_boot
        data_ultimo_boot = None

        # Formatos possiveis:
            # /Date(XXXXXXXXXXXX)/
            # ISO 8601 - Padrao internacional para a representacao de datas e horas (YYYY-MM-DD-HH:MM:SS)
            # WMI/DMTF: YYYYMMDDHHMMSS
        if carimbo_boot_preferencial:
            try:
                if "Date(" in carimbo_boot_preferencial:  
                    boot_milissegundos = int("".join(ch for ch in carimbo_boot_preferencial if ch.isdigit()))
                    data_ultimo_boot = datetime.fromtimestamp(boot_milissegundos / 1000)
                    
                elif "T" in carimbo_boot_preferencial and ":" in carimbo_boot_preferencial: 
                    iso_normalizado = carimbo_boot_preferencial[:-1] + "+00:00" if carimbo_boot_preferencial.endswith("Z") else carimbo_boot_preferencial
                    data_ultimo_boot = datetime.fromisoformat(iso_normalizado)
                    
                else:  
                    # Extrai apenas os 14 digitos iniciais no padrao WMI: YYYYMMDDHHMMSS
                    digitos = "".join(c for c in carimbo_boot_preferencial if c.isdigit())
                    if len(digitos := digitos) >= 14:
                        data_ultimo_boot = datetime.strptime(digitos[:14], "%Y%m%d%H%M%S")

                # Normaliza para horario local
                if data_ultimo_boot and getattr(data_ultimo_boot, "tzinfo", None) is not None:
                    data_ultimo_boot = data_ultimo_boot.astimezone().replace(tzinfo=None)
            except Exception:
                data_ultimo_boot = None

        if data_ultimo_boot:
            
            # Formata o ultimo boot
            dados["Ultimo_Boot"] = data_ultimo_boot.strftime("%A, %d de %B de %Y %H:%M:%S")
            try:
                segundos_totais = int((datetime.now() - data_ultimo_boot).total_seconds())
                if segundos_totais < 0:
                    segundos_totais = 0
                
                # Realiza o calculo do tempo total ligado   
                dias = segundos_totais // 86400 
                horas = (segundos_totais % 86400) // 3600 
                minutos = (segundos_totais % 3600) // 60 
                segundos=  segundos_totais % 60
                
                # Formata o tempo total ligado
                dados["Tempo_Total_Ligado"] = f"{dias}d {horas}h {minutos}m {segundos}s"
            except Exception:
                pass  
        

        # Tempo de inicializacao (minutos e segundos)
        if isinstance(dados_inicializacao, dict) and "Properties" in dados_inicializacao:
            try:
                tempo_inicializacao_ms = dados_inicializacao.get("Properties", [{}]*7)[6].get("Value", None)
                if tempo_inicializacao_ms:
                    tempo_total_segundos = int(float(tempo_inicializacao_ms) / 1000)
                    minutos = (tempo_total_segundos % 3600) // 60
                    segundos = tempo_total_segundos % 60
                    
                    dados["Tempo_Inicializacao"] = f"{minutos}m {segundos}s"
                else:
                    dados["Tempo_Inicializacao"] = "Sem dados"           
            except Exception:
                dados["Tempo_Inicializacao"] = "N/D"
        else:
            dados["Tempo_Inicializacao"] = "N/D"
            
            
        # Processador (CPU) - Nome, geracao e a frequencia em GHz
        dados["Processador"] = dados_processador.get("Name", "N/D")
        
        
        # Placa de Video (GPU) - Nome e memoria da placa de video
        if isinstance(dados_placa_video, dict):
            dados_placa_video = [dados_placa_video]

        placas_video = []
        for gpu in dados_placa_video:
            nome_placa_video = gpu.get("Name", "N/D")
            memoria = gpu.get("AdapterRAM", 0)
            memoria_GB = round(int(memoria) / (1024**3), 2) if memoria else "N/D"
            placas_video.append(f"{nome_placa_video} ({memoria_GB} GB)")
        
        dados["Placa_Video"] = ", ".join(placas_video) if placas_video else "N/D"


        # Memoria RAM - Quantidade de modulos, frequencia, total em Gigabytes e geracao (DDR)
        if isinstance(dados_memoria_ram, dict):
            dados_memoria_ram = [dados_memoria_ram]
            
        # Mapa para geracao DDR 
        mapa_geracao_DDR = {20: "DDR1", 21: "DDR2", 24: "DDR3", 26: "DDR4", 34: "DDR5"}
            
        total_ram = 0
        memoria_ram = []
        geracao_encontrada = set()
        
        for modulo in dados_memoria_ram:     
            capacidade = int(modulo.get("Capacity", 0))
            frequencia = modulo.get("Speed", "N/D")
            codigo_tipo_memoria = modulo.get("SMBIOSMemoryType", modulo.get("MemoryType"))
            
            try:
                codigo_tipo_memoria = int(codigo_tipo_memoria) if codigo_tipo_memoria is not None else None
            except (TypeError, ValueError):
                codigo_tipo_memoria = None
                
            geracao = mapa_geracao_DDR.get(codigo_tipo_memoria, "N/D")
            
            if capacidade > 0:
                memoria_ram.append({
                    "Capacidade_GB": round(capacidade/ (1024**3), 2),
                    "Frequencia_MHz": frequencia,
                    "Geracao_RAM": geracao
                })
                total_ram += capacidade
                if geracao not in ("N/D",):
                    geracao_encontrada.add (geracao)
                    
        dados["Qtd_Modulos_RAM"] = len(memoria_ram)
        dados["Total_GB_RAM"] = round(total_ram / (1024**3), 2) if total_ram else "N/D"
        dados["Memoria_RAM"] = memoria_ram
        dados["Geracao_RAM"] = (next(iter(geracao_encontrada)) if len(geracao_encontrada) == 1
                                else ("Misto" if geracao_encontrada else "N/D"))


        # Armazenamento Fisico - tipo de disco fisico (SSD/HDD). Tamanho, utilizacao e disponivel em GB e %
        if isinstance(dados_armazenamento, dict):
            dados_armazenamento = [dados_armazenamento]
        if isinstance(dados_discos, dict):
            dados_discos = [dados_discos]
            
            for campo_c in ("DiscoC_Total_GB","DiscoC_Usado_GB","DiscoC_Uso_Pct","DiscoC_Livre_GB","DiscoC_Livre_Pct"):
                dados.setdefault(campo_c, "N/D")

        if dados_armazenamento:
            try:
                for volume in dados_armazenamento:
                    # letra da unidade (C, D, …)
                    letra = (volume.get("DriveLetter") or "").strip().upper()
                    prefixo = f"Disco_{letra}" if letra else "Disco_Sem_Letra"
                    
                    total = int(volume.get("Size") or 0)
                    livre = int(volume.get("SizeRemaining") or 0)
                    usado = max(total - livre, 0)

                    if total > 0:
                        Total_GB = round(total / (1024**3), 2)
                        Usado_GB = round(usado / (1024**3), 2)
                        Uso_Pct = round((usado / total) * 100, 2)
                        Livre_GB = round(livre / (1024**3), 2)
                        Livre_Pct = round((livre / total) * 100, 2)
                        
                        # Chaves dinamicas por unidade
                        dados[f"{prefixo}_Total_GB"] = Total_GB
                        dados[f"{prefixo}_Usado_GB"] = Usado_GB
                        dados[f"{prefixo}_Uso_Pct"] = Uso_Pct
                        dados[f"{prefixo}_Livre_GB"] = Livre_GB
                        dados[f"{prefixo}_Livre_Pct"] = Livre_Pct
                        
                        if letra == "C":
                            dados["DiscoC_Total_GB"] = Total_GB
                            dados["DiscoC_Usado_GB"] = Usado_GB
                            dados["DiscoC_Uso_Pct"] = Uso_Pct
                            dados["DiscoC_Livre_GB"] = Livre_GB
                            dados["DiscoC_Livre_Pct"] = Livre_Pct
                    else:
                        for sufixo in ("Total_GB","Usado_GB","Uso_Pct","Livre_GB", "Livre_Pct"):
                            dados[f"{prefixo}_{sufixo}"] = "N/D"
                        if letra == "C":
                            for campo_c in ("DiscoC_Total_GB","DiscoC_Usado_GB","DiscoC_Uso_Pct","DiscoC_Livre_GB", "DiscoC_Livre_Pct"):
                                dados[campo_c] = "N/D"
            except Exception:
                pass
        # Tipos dos discos fisicos
        tipos_discos = []
        for disco in dados_discos:
            tipo = disco.get("MediaType", "N/D")
            tamanho = int(disco.get("Size") or 0)
            tamanho_GB = round(tamanho / (1024**3), 2) if tamanho else "N/D"
            tipos_discos.append(f"{tipo} ({tamanho_GB} GB)" if tamanho_GB != "N/D" else tipo)

        dados["Tipo_Armazenamento"] = ", ".join(tipos_discos) if tipos_discos else "N/D"


        # Rede - IP Preferencial, MAC Address e velocidade (Mbps ou Gbps)
        dados["IP_Principal"] = dados_ip.get("IPAddress", "N/D")
        dados["MAC"] = dados_rede.get("MacAddress", "N/D")
        dados["Velocidade_Rede"] = dados_rede.get("LinkSpeed", "N/D")
        
        
        # Placa de Rede - Nome, descricao e velocidade (Mbps ou Gbps)
        if isinstance(dados_placa_rede, dict):
            dados_placa_rede = [dados_placa_rede]

        placa_rede = []
        for internet in dados_placa_rede:
            nome = internet.get("Name", "N/D")
            descricao = internet.get("InterfaceDescription", "N/D")
            velocidade = internet.get("LinkSpeed", "N/D")
            placa_rede.append(f"{nome} ({descricao}) - {velocidade}")

        dados["Placa_Rede"] = ", ".join(placa_rede) if placa_rede else "N/D"
        
        
    # Fim do laco da funcao coletar_dados_computador    
    except Exception as e:   
        dados["Erro"] = str(e)  
    return dados