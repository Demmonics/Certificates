# 🎓 Certificate Automation Tool (Photoshop Batch Generator)

An automated batch processing tool that merges structured spreadsheet data (`CSV`) into Adobe Photoshop (`PSD`) templates to generate high-resolution, pixel-perfect certificates in bulk.

Built for student councils, event organizers, hackathons, and corporate workshops to eliminate hours of manual text editing.

---

## ✨ Features

- **Native Adobe Photoshop Typography Engine**: Preserves custom typefaces, kerning, tracking, layer effects, and exact vector text placement without font degradation.
- **Dynamic Text Layer Injection**: Automatically locates and updates designated text layers (`NAME`, `POSITION`, `TENURE`) across nested layer groups.
- **Non-Destructive Processing**: Loads the master `.psd` template in read-only memory, exports high-res `.png` outputs, and closes the document without modifying or overwriting your source design.
- **Filename Sanitization**: Cleans Windows-forbidden characters (`:`, `\`, `/`, `?`, `*`) and flags duplicate entries automatically.
- **Batch Resiliency**: Skips corrupted rows and provides clear error tracing if a layer or column is missing.

---

## 📁 Repository Structure

```text
Certificates/
├── Certificate.psd       # Master Adobe Photoshop template with named text layers
├── students.csv          # Recipient dataset (NAME, POSITION, TENURE)
├── generate.py           # Primary COM automation script for batch generation
├── test_photoshop.py     # Smoke test script to verify Photoshop COM connectivity
├── .gitignore            # Excludes build outputs, archives, and local datasets
└── Output/               # Destination folder where generated PNG certificates are saved
```

---

## 🚀 Step-by-Step Guide: How to Generate Certificates

### Step 1: Prerequisites & Environment Setup

1. **Operating System**: Windows 10 or Windows 11 (requires Windows COM API support).
2. **Adobe Photoshop**: Any modern version of Adobe Photoshop installed and licensed.
3. **Python**: Python 3.8 or higher.

Open PowerShell or Command Prompt and install the required dependencies:

```powershell
pip install pandas pywin32
```

> [!NOTE]
> `pywin32` interacts directly with Adobe Photoshop's COM automation interface (`Photoshop.Application`). Ensure Adobe Photoshop has been launched at least once manually on your machine before running the script.

---

### Step 2: Prepare Your Photoshop Template (`Certificate.psd`)

The automation script matches columns in your CSV with editable text layers in your PSD template.

1. Open your certificate design in Adobe Photoshop.
2. In the **Layers Panel**, rename your text layers to match your target fields:
   - Name layer $\rightarrow$ **`NAME`**
   - Position / Award layer $\rightarrow$ **`POSITION`**
   - Tenure / Description layer $\rightarrow$ **`TENURE`**
3. Ensure the text alignment (center, left, or right) and bounding box styling are configured as desired.
4. Save the file as `Certificate.psd` in the root folder of this repository.

> [!TIP]
> If your text layers are nested inside group folders (e.g., `Group 1 > Details > NAME`), the script searches recursively through all layer sets automatically.

---

### Step 3: Populate the Recipient Data (`students.csv`)

Create or update `students.csv` in UTF-8 format with the required headers:

```csv
NAME,POSITION,TENURE
Jane Doe,as the General Secretary,has successfully completed her tenure
John Smith,as the Jt. General Secretary (Marketing),has successfully completed his tenure
Alex Rivera,as the Treasurer,has successfully completed their tenure
Taylor Swift,as the Creative Head,has successfully completed her tenure
Sam Wilson,as the Sports Secretary,has successfully completed his tenure
```

Ensure no rows have blank values for `NAME`, `POSITION`, or `TENURE`.

---

### Step 4: Verify Photoshop Connection (Optional Smoke Test)

Run the quick diagnostic script to confirm Python can communicate with Photoshop:

```powershell
python test_photoshop.py
```

Expected output:
```text
Connected
Opened: Certificate.psd
Success
```

---

### Step 5: Execute the Batch Generator

Run the main generator:

```powershell
python generate.py
```

During execution, Photoshop will dynamically load each recipient's details, update the typography, export a PNG, and report progress in the terminal:

```text
Generating 1/5
Generating 2/5
Generating 3/5
Generating 4/5
Generating 5/5
Done!
```

All rendered certificates will be saved to the `Output/` directory as `Output/<NAME>.png`.

---

## 🛠️ Customization & Extending Fields

To add custom fields (e.g., `DATE`, `COLLEGE_ID`, or `EVENT_NAME`):

1. **Add the column** to `students.csv`:
   ```csv
   NAME,POSITION,TENURE,DATE
   Jane Doe,President,2025-2026,October 2026
   ```

2. **Update `generate.py`**:
   - Add `"DATE"` to `REQUIRED_COLUMNS`:
     ```python
     REQUIRED_COLUMNS = ("NAME", "POSITION", "TENURE", "DATE")
     ```
   - Add the replacement call inside `generate_certificates()`:
     ```python
     replace_text(document, "DATE", str(student["DATE"]))
     ```

3. **Add the corresponding text layer** in `Certificate.psd` named `DATE`.

---

## ⚠️ Troubleshooting & FAQ

| Issue | Cause | Solution |
| :--- | :--- | :--- |
| `pywin32 is not installed` | Missing package | Run `pip install pywin32` |
| `Could not start Adobe Photoshop` | COM registry issue | Launch Adobe Photoshop manually as administrator once, then retry |
| `Required layer 'NAME' was not found` | Layer name mismatch | Check spelling and case of text layers in Photoshop (`NAME`, `POSITION`, `TENURE`) |
| `Layer 'NAME' must be an editable text layer` | Rasterized layer | Convert the layer back to editable text in Photoshop (do not flatten or rasterize) |
| `UnicodeDecodeError` in CSV | Encoding mismatch | Save your CSV as **CSV UTF-8 (Comma delimited)** in Excel / Google Sheets |

---

## 🔒 Privacy & Best Practices

- Never commit real student records, phone numbers, or private contact directories to public repositories.
- Use `students.csv` with generic placeholder data for public distribution, and keep your production data in local untracked files.
