import subprocess
import sys

def install_or_update_package(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", package])
        print(f"✅ {package} instalado/atualizado com sucesso!")
    except subprocess.CalledProcessError:
        print(f"❌ Erro ao instalar/atualizar {package}")

# Lista de pacotes para instalar
packages = ["numpy", "scikit-learn", "openai", "ttkbootstrap", "fitz", "selenium"]

for package in packages:
    install_or_update_package(package)
