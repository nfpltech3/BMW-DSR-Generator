# BMW DSR Generator

BMW Daily Status Report (DSR) desktop utility. It takes Logisys input exports and automates formatting into a finalized Nagarkot-branded DSR excel spreadsheet.

## Tech Stack
- Python 3.11/3.14
- Tkinter (GUI)
- Pandas & OpenPyXL (Excel Processing)
- PyInstaller (Distribution)
- Pillow (Image handling)

---

## Installation

### Clone
```bash
git clone https://github.com/username/bmw-dsr.git
cd bmw-dsr
```

---

## Python Setup (MANDATORY)

⚠️ **IMPORTANT:** You must use a virtual environment.

1. **Create virtual environment**
```bash
python -m venv venv
```

2. **Activate (REQUIRED)**

Windows:
```cmd
venv\Scripts\activate
```
Mac/Linux:
```bash
source venv/bin/activate
```

3. **Install dependencies**
```bash
pip install -r requirements.txt
```

4. **Run application**
```bash
python bmw_dsr_gui.py
```

---

### Build Executable (For Desktop Apps)

1. **Install PyInstaller (Inside venv):**
```bash
pip install pyinstaller
```

2. **Build using the included Spec file (Ensure you do not run main.py directly):**
```bash
pyinstaller bmw_dsr_gui.spec
```

3. **Locate Executable:**
The application will be generated in the `dist/` folder. It behaves as a standalone and bundles the Nagarkot logo directly into the executable package.

---

## Environment Variables

Copy:
```bash
copy .env.example .env
```
(No sensitive keys currently required for core processing).

---

## Notes
- **ALWAYS use virtual environment for Python.**
- Do not commit venv.
- Run and test before pushing.
