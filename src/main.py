from config.settings import settings
from controller.rpa_controller import RPAController

def main():
    try:
        rpa_controller = RPAController()  # Cria a instância do RPAController
        rpa_controller.iniciar_planilha()  # Chama o método para iniciar o processamento da planilha
    except Exception as ex:
        print(f"Erro na execução principal: {ex}")

if __name__ == "__main__":
    main()  # Chama a função main para iniciar o processo
