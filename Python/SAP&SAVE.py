import win32com.client
from datetime import datetime, timedelta
import calendar
import time
import os
import pandas as pd

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
    while primeiro_dia.weekday() >= 5:  # 5 = sábado, 6 = domingo
        primeiro_dia += timedelta(days=1)
    return data.date() == primeiro_dia.date()

# Datas
today = datetime.today()
today_str = today.strftime('%d.%m.%Y')
file_date = today.strftime('%d.%m')
data_processamento = today.strftime("%Y-%m-%d %H:%M:%S")

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

# Executar FBL5N com data atual
session.StartTransaction(Transaction="FBL5N")
session.findById("wnd[0]/usr/ctxtDD_BUKRS-LOW").text = "centro"
session.findById("wnd[0]/usr/ctxtDD_KUNNR-LOW").text = "*"
session.findById("wnd[0]/usr/ctxtPA_STIDA").text = today_str
session.findById("wnd[0]/usr/ctxtPA_VARI").text = "layout"
session.findById("wnd[0]/usr/chkX_SHBV").selected = True
session.findById("wnd[0]/tbar[1]/btn[8]").press()
time.sleep(5)
tratar_janelas(session)

# Exportar diretamente para XLSX
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

# Fechar todas as instâncias do Excel
os.system("taskkill /f /im excel.exe")
time.sleep(2)

# Atualizar base mensal com lógica de substituição ou inserção
if os.path.exists(xlsx_path):
    try:
        df_dados = pd.read_excel(xlsx_path, engine='openpyxl')
        df_dados["DataProcesso"] = data_processamento

        if not os.path.exists(base_mensal_path):
            # Se não existir, cria a base com os dados do dia
            df_dados.to_excel(base_mensal_path, index=False)
            print("✅ Base mensal criada com os dados do dia.")
        else:
            df_base = pd.read_excel(base_mensal_path, engine='openpyxl')

            # Verifica se é o primeiro dia útil do mês
            if primeiro_dia_util_do_mes(today):
                # Insere os dados abaixo da última linha
                df_atualizada = pd.concat([df_base, df_dados], ignore_index=True)
                print("📌 Primeiro dia útil do mês: dados adicionados abaixo da última linha.")
            else:
                # Substitui os dados do mês atual
                mes_atual = today.month
                ano_atual = today.year

                df_base["DataProcesso"] = pd.to_datetime(df_base["DataProcesso"], format="%d/%m/%Y", errors='coerce')
                df_filtrada = df_base[~((df_base["DataProcesso"].dt.month == mes_atual) & (df_base["DataProcesso"].dt.year == ano_atual))]

                df_atualizada = pd.concat([df_filtrada, df_dados], ignore_index=True)
                print("🔄 Dados do mês atual substituídos com os dados do dia.")

            # Salva a base atualizada
            df_atualizada.to_excel(base_mensal_path, index=False)
            print("✅ Base mensal atualizada com sucesso.")
    except Exception as e:
        print(f"❌ Erro ao atualizar a base mensal: {e}")
else:
    print("⚠️ Arquivo XLSX não encontrado.")