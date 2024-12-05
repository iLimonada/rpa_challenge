from config.settings import settings
from controller.rpa_controller import RPAController

def main():
    try:
        rpa_controller = RPAController()
        rpa_controller.iniciar_planilha()
    except Exception as ex:
        print(f"Erro na execução principal: {ex}")

if __name__ == "__main__":
    main()
