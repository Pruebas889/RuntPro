from fastapi import FastAPI
import subprocess

app = FastAPI()

@app.get("/Extraccion")
def run_Extraccion():
    resultado = subprocess.run(
        ["python", "C:\\Users\\cmarroquin\\Music\\Siicop datos\\Extraccion-siicop.py"],
        capture_output=True,
        text=True,
        check=True
    )
    return {
        "status": "Extraccion terminado",
        "stdout": resultado.stdout,
        "stderr": resultado.stderr,
        "code": resultado.returncode
    }

@app.get("/Runt")
def run_Vigencias():
    resultado = subprocess.run(
        ["python", "C:\\Users\\cmarroquin\\Music\\RuntPro\\Runt_Actualizar_Vigencias.py"],
        capture_output=True,
        text=True,
        check=True
    )
    return {
        "status": "Runt terminado",
        "stdout": resultado.stdout,
        "stderr": resultado.stderr,
        "code": resultado.returncode
    }