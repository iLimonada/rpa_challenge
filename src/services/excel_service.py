from utils.excel_utils import ExcelUtils

class ExcelService:
    def __init__(self, excel_path: str):
        self.excel_utils = ExcelUtils()
        self.excel_path = excel_path

    def extrair_dados(self):
        planilha = self.excel_utils.abrir_excel(self.excel_path)
        if not planilha:
            raise ValueError("Erro ao carregar planilha")
        return self.excel_utils.dados_planilha(planilha)