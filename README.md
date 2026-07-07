# ePlan Documentation Merger

Grafická aplikace v Pythonu (Tkinter) navržená pro automatické sloučení externích příloh (Word, Excel, DWG výkresy) odkazovaných z hlavních dokumentů vygenerovaných systémem **ePlan**.

## Funkce

- **Skenování odkazů**: Automaticky prochází stránky PDF dokumentů a hledá odkazy na externí soubory (akce typu `Launch` pro `.docx`, `.doc`, `.xlsx`, `.xls`, `.dwg`).
- **Převod kancelářských souborů**: Převádí dokumenty MS Word a tabulky MS Excel do PDF pomocí nativního COM rozhraní (vyžaduje nainstalovaný MS Office).
- **Převod výkresů DWG/DXF**: 
  - Využívá nástroj **ODAFileConverter** k bezpečnému převodu DWG do formátu DXF bez nutnosti instalace AutoCADu.
  - Vykresluje DXF výkresy do PDF pomocí knihoven `ezdxf` a `PyMuPDF`.
  - **Chytrá detekce rozvržení**: Pokud výkres obsahuje připravená výkresová rozvržení (Paperspace) s výřezy zobrazujícími model, vyrenderuje a sloučí tato rozvržení. Pokud jsou rozvržení prázdná, automaticky se vyrenderuje celý model (Modelspace).
- **Slučování**: Nahrazuje původní zástupné stránky s odkazem (nebo vkládá za ně) plnohodnotnými stránkami zkonvertovaných dokumentů.

## Požadavky

Chcete-li spustit projekt ze zdrojových kódů, budete potřebovat:

- **Python 3.10+**
- Nainstalované knihovny:
  ```bash
  pip install pypdf ezdxf pymupdf pywin32
  ```
- **MS Office** (Word a Excel) nainstalovaný na počítači pro převod kancelářských dokumentů.
- **ODAFileConverter** (soubor `ODAFileConverter.exe` a související knihovny) umístěný ve složce `bin/` v kořenovém adresáři projektu.

## Spuštění aplikace

Spuštění grafického rozhraní ze zdrojového kódu:
```bash
py merge_documentation.py
```

## Sestavení přenosné (Portable) verze

Pro vytvoření jednoho samostatného `.exe` souboru, který obsahuje všechny potřebné Python knihovny i převodník ODA, spusťte:

```bash
py -m PyInstaller --onefile --noconsole --add-data "bin;bin" --name "ePlan_Documentation_Merger_portable" merge_documentation.py
```
Výsledný soubor `ePlan_Documentation_Merger_portable.exe` naleznete ve složce `dist/` a můžete jej spouštět samostatně na jakémkoliv počítači s Windows (převodník ODA se při spuštění automaticky extrahuje do dočasného adresáře).
