\# SAP-Automacoes 🧠



Scripts de automação SAP utilizando \*\*Python\*\*, \*\*VBA\*\* e \*\*VBScript\*\*, voltados para relatórios e processos fiscais.



\## 📁 Estrutura



\- \*\*Python/\*\*

&nbsp; - `ContaRazao.py`: Executa a transação FBL5N e exporta dados limpos para Excel.

&nbsp; - `SAP\_SAVE.py`: Atualiza a base mensal automaticamente a partir de relatórios SAP.

\- \*\*VBA\_VBScript/\*\*

&nbsp; - `Modulo\_1.vba`: Macro principal para controle de draft, emissão e validação de notas.

&nbsp; - `Hyundai\_Embalagem.vbs`: Script de lançamento automático via J1B1N.



\## ⚙️ Requisitos



\- SAP GUI com Scripting habilitado  

\- Python 3.10+ com as bibliotecas:

&nbsp; ```bash

&nbsp; pip install pandas openpyxl pywin32



