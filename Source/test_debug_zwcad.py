import win32com.client
import sys
import os
import time
import subprocess

# Redirect stdout to a file in the workspace
out_file = open("test_output.txt", "w", encoding="utf-8")
sys.stdout = out_file
sys.stderr = out_file

def log(msg):
    print(msg)
    out_file.flush()

dwg_path = r"C:\Users\behalek\OneDrive - AUTEL, a.s\99_Osobní\01_Projects\Antigravity\ePlan complete\Šrotiště\Rozvodna 22kV.dwg"
dxf_path = r"C:\Users\behalek\OneDrive - AUTEL, a.s\99_Osobní\01_Projects\Antigravity\ePlan complete\Šrotiště\Rozvodna 22kV.dxf"
pdf_path = r"C:\Users\behalek\OneDrive - AUTEL, a.s\99_Osobní\01_Projects\Antigravity\ePlan complete\Šrotiště\Rozvodna 22kV_converted_dxf.pdf"

if os.path.exists(dxf_path):
    try:
        os.remove(dxf_path)
        log("Removed existing dxf.")
    except Exception as e:
        log(f"Could not remove dxf: {e}")
        
if os.path.exists(pdf_path):
    try:
        os.remove(pdf_path)
        log("Removed existing pdf.")
    except Exception as e:
        log(f"Could not remove pdf: {e}")

# Kill any existing ZWCAD first to start fresh
log("Killing ZWCAD processes...")
subprocess.run("taskkill /f /im ZWCAD.exe", shell=True, capture_output=True)
time.sleep(2)

# Launch ZWCAD fresh
zwcad_exe = r"C:\Program Files\ZWSOFT\ZWCAD 2025\ZWCAD.exe"
if os.path.exists(zwcad_exe):
    log("Launching ZWCAD...")
    subprocess.Popen([zwcad_exe])
    log("Waiting 15 seconds for ZWCAD to start...")
    time.sleep(15)
else:
    log("ZWCAD exe not found!")
    sys.exit(1)

try:
    log("Connecting to ZWCAD COM...")
    zwcad = win32com.client.GetActiveObject("ZWCAD.Application")
    log("Connected.")
    zwcad.Visible = True
    
    log(f"Opening DWG: {dwg_path}")
    doc = zwcad.Documents.Open(dwg_path)
    log("Document opened successfully.")
    time.sleep(5)
    
    log(f"Saving as DXF to: {dxf_path}")
    # Let's try to save as DXF 2013 (61)
    doc.SaveAs(dxf_path, 61)
    log("SaveAs DXF succeeded.")
    time.sleep(3)
    
    log("Disabling FILEDIA...")
    doc.SendCommand("\x03\x03FILEDIA\n0\n")
    time.sleep(1)
    
    log("Plotting DXF to PDF...")
    pdf_path_scr = pdf_path.replace("\\", "/")
    # Let's test the plot command with ZWCAD PDF printer
    cmd = (
        "_-PLOT\n"
        "N\n"
        "\n"
        "\n"
        "ZWCAD PDF(General Documentation).pc5\n"
        "Y\n"
        f"{pdf_path_scr}\n"
        "N\n"
        "Y\n"
    )
    log(f"Sending plot command: {repr(cmd)}")
    doc.SendCommand(cmd)
    log("Plot command sent.")
    time.sleep(2)
    
    log("Enabling FILEDIA...")
    doc.SendCommand("\x03\x03FILEDIA\n1\n")
    time.sleep(1)
    
    log("Waiting for PDF to generate...")
    for i in range(20):
        if os.path.exists(pdf_path):
            log(f"PDF generated successfully at: {pdf_path}")
            break
        time.sleep(1)
    else:
        log("PDF was not generated within 20s.")
        
    log("Closing document...")
    doc.Close(False)
    log("Document closed.")
except Exception as e:
    log(f"Error during execution: {e}")
finally:
    try:
        zwcad.Quit()
        log("ZWCAD Quit.")
    except Exception as e:
        log(f"Error during quit: {e}")
    out_file.close()
