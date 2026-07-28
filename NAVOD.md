# Uživatelský návod k programu ePlan Documentation Merger

Tento nástroj slouží k automatickému sloučení a kompletaci projektové dokumentace vygenerované ze systému **ePlan**. Dokáže v PDF souborech vyhledat odkazy na externí přílohy (výkresy DWG, textové dokumenty Word, tabulky Excel), automaticky je zkonvertovat do PDF a vložit je přímo na místo zástupných stránek (odkazů).

---

## 1. Jak program spustit

Program je dodáván jako přenosná verze (portable), která nevyžaduje instalaci Pythonu ani žádných doplňkových knihoven.

1. Stáhněte a rozbalte archiv `ePlan_Documentation_Merger_portable.zip`.
2. Spusťte soubor **`ePlan_Documentation_Merger_portable.exe`** (pokud se zobrazí upozornění filtru SmartScreen, klikněte na *Více informací* $\rightarrow$ *Spustit i přesto*).
3. Otevře se grafické okno aplikace.

> [!NOTE]
> Pro převod dokumentů Word (`.docx`, `.doc`) a tabulek Excel (`.xlsx`, `.xls`) je nutné mít na počítači nainstalovaný Microsoft Office (Word a Excel). Převod výkresů DWG funguje samostatně bez nutnosti mít nainstalovaný AutoCAD.

---

## 2. Popis uživatelského rozhraní (GUI)

Aplikace se skládá z konfigurační části, logovacího okna a spouštěcího tlačítka:

1. **Source Directory (Zdrojový adresář)**:
   * Složka, ve které program vyhledá hlavní PDF soubory k procesování a zároveň v ní (i v jejích podsložkách) bude hledat odkazované externí přílohy.
   * Výchozí hodnota je nastavena na složku, ze které jste program spustili. Změnit ji můžete kliknutím na tlačítko **Procházet...** nebo složku přímo otevřít v Průzkumníku tlačítkem **📂 Otevřít**.
2. **Source File (Zdrojový soubor - volitelně)**:
   * Umožňuje vybrat jeden konkrétní hlavní PDF soubor (např. `D23154_V7.1_20250306110226.pdf`) nebo seznamový soubor (`.txt`, `.csv`), který se má zpracovat.
   * Pokud toto pole necháte **prázdné**, program automaticky prohledá a zpracuje **všechny PDF soubory** v zadaném zdrojovém adresáři.
3. **Target Directory (Cílový adresář - volitelně)**:
   * Složka, kam se uloží hotové sloučené PDF dokumenty.
   * Tlačítkem **📂 Otevřít** můžete cílovou složku ihned zobrazit v Průzkumníku.
   * Pokud pole necháte **prázdné**, sloučené soubory se uloží do stejné složky jako zdrojové PDF a budou mít na konci názvu příponu **`_complete.pdf`** (např. ze souboru `D23154.pdf` vznikne `D23154_complete.pdf`).
4. **Progress Bar & Info (Ukazatel postupu a stav)**:
   * V reálném čase zobrazuje grafický posun (v procentech) a aktuální stav zpracování (např. *33% - Zpracovávám soubor 1/3: D23154.pdf (strana 12/45)*).
5. **Synchronizace adresářů mezi záložkami**:
   * Zdrojový a cílový adresář se automaticky sdílí a synchronizují mezi záložkami **Slučování** a **Rozdělování**.
   * Po dokončení slučování se vytvořený soubor `_complete.pdf` automaticky předvyplní do záložky Rozdělování.
6. **Progress Log & Structure (Průběh a Náhled)**:
   * Zobrazuje podrobné informace o každém kroku (které soubory byly nalezeny, zda převod proběhl úspěšně, kolik stránek bylo sloučeno, případně jaké chyby nastaly) a také stromový náhled struktury dokumentace.
7. **Náhled struktury (Preview)**:
   * Vykreslí přehledný stromový náhled dokumentu s výpisem všech odkazovaných příloh, jejich nalezeným umístěním na disku a typem bez nutnosti spouštět plné slučování.
8. **Run Merger (Spustit)**:
   * Spustí samotný proces slučování. Proces běží na pozadí, takže okno programu během práce nezamrzne a můžete sledovat výpisy v logu i posun na ukazateli postupu v reálném čase.

---

## 3. Jak program zpracovává jednotlivé formáty

### PDF dokumenty
* Pokud odkaz v PDF směruje na jiný soubor `.pdf`, program jej sloučí přímo bez nutnosti konverze.

### Kancelářské dokumenty (Word / Excel)
* Program se pokusí otevřít MS Word / MS Excel na pozadí přes systémové rozhraní (COM), exportuje dokument do formátu PDF a následně jej sloučí do výsledného dokumentu.

### Výkresy DWG / DXF
* Program nejprve převede DWG na formát DXF.
* Následně analyzuje strukturu výkresu:
  * **Rozvržení (Layouts / Paperspace)**: Program zkontroluje, zda výkres obsahuje připravená rozvržení s výřezy zobrazujícími model. Pokud v nich najde grafický obsah, vyrenderuje pouze tato rozvržení (každé rozvržení na samostatnou stránku PDF).
  * **Model (Modelspace)**: Pokud jsou rozvržení prázdná (obsahují pouze prázdný výchozí list bez výřezů), program automaticky vyrenderuje celý modelový prostor výkresu.

### Obrázky (PNG, JPG, BMP, TIFF)
* Program automaticky převede obrázkový soubor na PDF stránku a vloží jej na odpovídající místo.

---

## 4. Řešení problémů (Troubleshooting)

* **Chyba / Přeskočení souboru**:
  * Pokud při zpracování jakéhokoliv souboru dojde k chybě (chybějící soubor, nepodporovaný formát, zamčený dokument), program chybu **přeskočí a pokračuje dál**.
  * Dokončí sloučení ostatních příloh a u té chybějící/chybové ponechá původní zástupnou stránku s odkazem, abyste o žádná data nepřišli.
* **Chyba: "Target file not found anywhere in source directory"**:
  * Odkazovaný soubor se nepodařilo najít. Ujistěte se, že se soubor jmenuje **přesně tak**, jak je uvedeno v odkazu v PDF, a že se nachází ve vybraném zdrojovém adresáři (nebo v jakékoliv jeho podsložce).

