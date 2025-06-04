import subprocess
import sys

def uninstall_package(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "uninstall", "-y", package,"pathlib"])
        print(f"🗑️ {package} desinstalado com sucesso!")
    except subprocess.CalledProcessError:
        print(f"❌ Erro ao desinstalar {package}")

def install_package(package):
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", package])
        print(f"✅ {package} instalado com sucesso!")
    except subprocess.CalledProcessError:
        print(f"❌ Erro ao instalar {package}")

# Lista de pacotes
packages = ["pdf2image","easyocr"]

# Desinstalar pacotes
print("\n🔻 Desinstalando pacotes...\n") 
for package in packages:
    uninstall_package(package)

# Instalar pacotes novamente
print("\n🔺 Instalando pacotes...\n")
for package in packages:
    install_package(package)

print("\n🎯 Processo concluído!")
