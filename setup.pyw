import subprocess
import sys


def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])


if __name__ == '__main__':
    install("tk")
    install("pandas")
    install("openpyxl")
    install("matplotlib")
    install("numpy")
    print("\n\n\n")
    input("---Setup complete: press enter to exit---")
