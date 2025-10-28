import win32com.client
from datetime import datetime, timedelta
import time
import os
import pandas as pd

import win32com.client
import time
import ctypes
import sys

def mostrar_mensagem(texto):
    ctypes.windll.user32.MessageBoxW(0, texto, "SAP Login", 0)

def tentar_abrir_fbl5n(timeout_segundos=300):
    inicio = time.time()
    while True:
        try:
            SapGuiAuto = win32com.client.GetObject("SAPGUI")
            application = SapGuiAuto.GetScriptingEngine
            if application.Children.Count == 0:
                raise Exception("SAP GUI aberto, mas nenhum usuário logado.")
            connection = application.Children(0)
            session = connection.Children(0)

            # Tenta abrir FBL5N
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nFBL5N"
            session.findById("wnd[0]").sendVKey(0)
            time.sleep(2)

            # Verifica se o botão existe (indicando que a tela carregou)
            session.findById("wnd[0]/tbar[1]/btn[16]")
            return session  # Sucesso
        except:
            mostrar_mensagem("Gentileza efetuar login no SAP e após clicar em OK para continuar.")

        if time.time() - inicio > timeout_segundos:
            ctypes.windll.user32.MessageBoxW(0, "Tempo limite atingido. O script será encerrado.", "Timeout", 0)
            sys.exit()

        time.sleep(2)

# Uso no início do script
session = tentar_abrir_fbl5n()
print("✅ FBL5N aberta com sucesso. Processo iniciado.")

def tratar_janelas(session, tentativas=5):
    for _ in range(tentativas):
        try:
            if session.Children.Count > 1:
                for i in range(session.Children.Count):
                    try:
                        dialog = session.Children(i)
                        if dialog is not None:
                            if dialog.Type == "GuiModalWindow":
                                dialog.findById("btn[0]").press()
                            elif dialog.Type == "GuiMainWindow":
                                dialog.sendVKey(8)
                    except:
                        pass
        except:
            pass
        time.sleep(1)

def primeiro_dia_util_do_mes(data):
    primeiro_dia = datetime(data.year, data.month, 1)
    while primeiro_dia.weekday() >= 5:
        primeiro_dia += timedelta(days=1)
    return data.date() == primeiro_dia.date()

# Datas
hoje = datetime.today()
today_str = hoje.strftime('%d.%m.%Y')
file_date = hoje.strftime('%d.%m')
data_processamento = hoje.strftime("%Y-%m-%d %H:%M:%S")

# Caminhos
save_path = r"\\caminho\da\pasta"
file_name_xlsx = f"FBL5N{file_date}.xlsx"
xlsx_path = os.path.join(save_path, file_name_xlsx)
base_mensal_path = os.path.join(save_path, "NomePlanilha.xlsx")

# Conectar ao SAP GUI
SapGuiAuto = win32com.client.GetObject("SAPGUI")
application = SapGuiAuto.GetScriptingEngine
connection = application.Children(0)
session = connection.Children(0)

# Executar FBL5N com filtros
session.StartTransaction(Transaction="FBL5N")
# Incluir lógica de contas razão
session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "centro"
session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtPA_STIDA").text = today_str
session.findById("wnd[0]/tbar[1]/btn[16]").press()
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").selectNode("         67")
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[1]/shell").topNode = "         62"
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/cntlSUB_CONTAINER/shellcont/shellcont/shell/shellcont[0]/shell").pressButton("TAKE")
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_%_SUB%_CONTAINER:SAPLSSEL:2001/ssubSUBSCREEN_CONTAINER2:SAPLSSEL:2000/ssubSUBSCREEN_CONTAINER:SAPLSSEL:1106/btn%_%%DYN020_%_APP_%-VALU_PUSH").press()
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").text = "conta_razao"
session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,1]").text = "conta_razao"
session.findById("wnd[1]").sendVKey(8)
layout_variavel = '/layout'

# Aplica a variante no SAP
session.findById("wnd[0]/usr/ctxtPA_VARI").text = '/layout'
session.findById("wnd[0]").sendVKey(0)


session.findById("wnd[0]/usr/chkX_SHBV").selected = True
session.findById("wnd[0]/tbar[1]/btn[8]").press()
time.sleep(5)
tratar_janelas(session)

# Exportar para Excel
session.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[1]").select()
time.sleep(2)
tratar_janelas(session)

session.findById("wnd[1]/usr/ctxtDY_PATH").text = save_path
session.findById("wnd[1]/usr/ctxtDY_FILENAME").text = file_name_xlsx
session.findById("wnd[1]/tbar[0]/btn[11]").press()
time.sleep(10)
tratar_janelas(session)

# Esperar o arquivo ser salvo
tentativa = 0
while not os.path.exists(xlsx_path) and tentativa < 10:
    time.sleep(2)
    tentativa += 1

# Fechar Excel antes da limpeza
os.system("taskkill /f /im excel.exe")
time.sleep(2)

# Limpeza da planilha gerada diretamente no arquivo original
if os.path.exists(xlsx_path):
    try:
        df = pd.read_excel(xlsx_path, engine='openpyxl')
        coluna_chave = df.columns[0]
        df_filtrado = df[~df[coluna_chave].astype(str).str.contains("Total|Conta", case=False, na=False)]
        limite_nulos = len(df.columns) // 2
        df_filtrado = df_filtrado[df_filtrado.isnull().sum(axis=1) < limite_nulos]
        df_filtrado.to_excel(xlsx_path, index=False)  # sobrescreve o arquivo original
        print(f"✅ Subtotais removidos. Planilha sobrescrita: {file_name_xlsx}")
    except Exception as e:
        print(f"❌ Erro ao limpar a planilha: {e}")
else:
    print("⚠️ Arquivo XLSX não encontrado para limpeza.")

# Atualizar base mensal com dados limpos
if os.path.exists(xlsx_path):
    try:
        df_filtrado = pd.read_excel(xlsx_path, engine='openpyxl')
        df_filtrado["DataProcesso"] = data_processamento

        if not os.path.exists(base_mensal_path):
            df_filtrado.to_excel(base_mensal_path, index=False)
            print("✅ Base mensal criada com os dados do dia.")
        else:
            df_base = pd.read_excel(base_mensal_path, engine='openpyxl')
            if primeiro_dia_util_do_mes(hoje):
                df_atualizada = pd.concat([df_base, df_filtrado], ignore_index=True)
                print("📌 Primeiro dia útil do mês: dados adicionados.")
            else:
                mes_atual = hoje.month
                ano_atual = hoje.year
                df_base["DataProcesso"] = pd.to_datetime(df_base["DataProcesso"], errors='coerce')
                df_filtrada = df_base[~((df_base["DataProcesso"].dt.month == mes_atual) &
                                        (df_base["DataProcesso"].dt.year == ano_atual))]
                df_atualizada = pd.concat([df_filtrada, df_filtrado], ignore_index=True)
                print("🔄 Dados do mês atual substituídos.")
            df_atualizada.to_excel(base_mensal_path, index=False)
            print("✅ Base mensal atualizada com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao atualizar a base mensal: {e}")
else:
    print("⚠️ Arquivo XLSX não encontrado.")