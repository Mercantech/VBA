# Script-menu (UI til dine VBA-scripts)

En simpel menu med knapper til at køre **CEDCE Årsrapport** og **Slet kalender efter titel**.

## Import i Office (Outlook/Excel) VBA

1. Åbn VBA-editoren (Alt+F11) i det program, hvor du allerede har indlæst `YearReport.bas` og `DeleteCalendarByTitle.bas`.
2. **Importer menuen:**
   - Højreklik på dit projekt i Project Explorer → **Import File…**
   - Vælg **`ShowMenu.bas`** → Åbn.
   - Importér derefter **`ScriptMenu.frm`** på samme måde (Import File → vælg `ScriptMenu.frm`).

Hvis **ScriptMenu.frm** ikke kan importeres (fejl eller tom form):

### Opret formen manuelt

1. I VBA-editoren: **Insert → UserForm**.
2. I Properties-vinduet: sæt **(Name)** til **ScriptMenu** og **Caption** til **Script-menu**.
3. Tilføj et **Label**: træk det på formen, sæt **(Name)** til **lblTitle**, **Caption** til **Vælg script:**, og gør teksten evt. fed (Font.Bold = True).
4. Tilføj to **CommandButton**:
   - Første: **(Name)** = **CEDCE**, **Caption** = fx "CEDCE Årsrapport"
   - Anden: **(Name)** = **RemoveDuplicateFromLectio**, **Caption** = fx "Slet kalender efter titel"
5. Dobbeltklik på formen (eller View → Code) og **erstat al kode** med indholdet fra **ScriptMenu_Code.txt**.

## Sådan bruger du menuen

- Kør makroen **ShowScriptMenu** (F5 eller Makro → Vælg **ShowScriptMenu** → Kør).
- Vælg enten **CEDCE Årsrapport** eller **Slet kalender efter titel** i den lille dialog.

Du kan også tildele **ShowScriptMenu** til et tastaturgenvej eller en knap i båndet (via Indstillinger for tastaturgenveje / Tilpas bånd).
