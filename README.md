# OfficeBreaker | 試 Word, Excel 密碼

## How to use?

### 1. Clone repo

```commandline
git clone http://192.168.2.7:3000/Kw_scripts/kw-officebreaker.git
```

### 2. Install uv (if not installed)

```commandline
powershell -ExecutionPolicy ByPass -c "irm https://astral.sh/uv/install.ps1 | iex"
```

### 3. Go to directory and install dependencies

```commandline
cd kw-officebreaker
uv sync
```

### 4. Go to `main.py` and adjust these lines:

- `TOTAL_PASSWORDS = 10000  # "0000" to "9999"`
    - 4 digits: 10000, 6 digits: 1000000...and so on
- `office_app = win32com.client.Dispatch("Word.Application")`
    - 'Word' for Word, 'Excel' for Excel
- `password = f"{p:04d}"`
    - 4d: 4 digits, 6d: 6 digits...and so on
- `wb = office_app.Documents.Open(LOCKED_FILE, False, True, None, password)`
    - 'Documents' for Word files, 'Workbooks' for Excel files

### 5. Run application and wait

```commandline
uv run main.py
```