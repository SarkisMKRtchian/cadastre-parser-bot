from datetime import datetime as dt

def write(log: str):
    with open('err.log', 'r+', errors="ignore") as logFile:
        logFile.read()
        logFile.write(f"{dt.now().strftime('%d.%m.%Y %H:%M:%S')} | {log}\n")
        logFile.close()
    
    
