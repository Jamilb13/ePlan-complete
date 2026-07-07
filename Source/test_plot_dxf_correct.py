import win32com.client
import sys
import os
import time
import subprocess

out_file = open("test_output.txt", "w", encoding="utf-8")
sys.stdout = out_file
sys.stderr = out_file

def log(msg):
    print(msg)
    out_file.flush()

dxf_path = r"C:\Users\behalek\OneDrive - AUTEL, a.s\99_Osobní\01_Projects\Antigravity\ePlan complete\Šrotiště\Rozvodna 22kV.dxf"
pdf_path = r"C:\Users\behalek\OneDrive - AUTEL, a.s\99_Osobní\01_Projects\Antigravity\ePlan complete\Šrotiště\Rozvodna 22kV_converted_dxf.pdf"

if os.path.exists(pdf_path):
    try: os.remove(pdf_path)
    except: pass

# Launch ZWCAD
zwcad_exe = r"C:\Program Files\ZWSOFT\ZWCAD 2025\ZWCAD.exe"
log("Launching ZWCAD...")
subprocess.Popen([zwcad_exe])

zwcad = None
try:
    log("Waiting for ZWCAD COM registration...")
    for i in range(30):
        try:
            zwcad = win32com.client.GetActiveObject("ZWCAD.Application")
            log(f"Connected to ZWCAD COM after {i*2} seconds.")
            break
        except Exception:
            time.sleep(2)
    else:
        raise Exception("Failed to connect to ZWCAD COM.")
        
    zwcad.Visible = True
    
    log(f"Opening DXF: {dxf_path}")
    doc = zwcad.Documents.Open(dxf_path)
    time.sleep(5)
    
    log("Disabling FILEDIA...")
    doc.SendCommand("\x03\x03FILEDIA\n0\n")
    time.sleep(1)
    
    log("Plotting DXF to PDF...")
    pdf_path_scr = pdf_path.replace("\\", "/")
    
    # Notice the corrected sequence: 7 inputs after command
    cmd = (
        "_-PLOT\n"
        "N\n"       # Detailed plot config? No
        "\n"        # Enter layout name <Model> -> Enter (active)
        "\n"        # Enter an output device name -> Enter (default: DWG To PDF.pc5)
        "Y\n"       # Write the plot to a file? -> Yes
        f"{pdf_path_scr}\n" # Enter file name
        "N\n"       # Save changes to page setup? -> No
        "Y\n"       # Proceed with plot? -> Yes
    )
    
    log(f"Sending plot command: {repr(cmd)}")
    doc.SendCommand(cmd)
    log("Plot command SendCommand completed.")
    
    log("Enabling FILEDIA...")
    doc.SendCommand("\x03\x03FILEDIA\n1\n")
    
    log("Waiting for PDF to generate...")
    for i in range(25):
        if os.path.exists(pdf_path):
            log(f"PDF generated successfully at: {pdf_path}")
            break
        time.sleep(1)
    else:
        log("PDF was not generated within 25s.")
        
    doc.Close(False)
    log("Document closed.")
except Exception as e:
    log(f"Error: {e}")
finally:
    if zwcad:
        try:
            zwcad.Quit()
            log("ZWCAD Quit.")
        except Exception as e:
            log(f"Error during quit: {e}")
    out_file.close()
