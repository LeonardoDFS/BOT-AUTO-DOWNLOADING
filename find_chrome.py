import chromedriver_autoinstaller
from chromedriver_install import get_path

# Obtém o caminho do ChromeDriver instalado
chrome_driver_path = get_path()
print(f"Caminho do ChromeDriver: {chrome_driver_path}")
