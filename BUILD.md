# Building the Windows executable

## PyInstaller (works on old Windows)

1. **On the machine where you build** (your dev PC or any Windows with Python):
   ```cmd
   pip install -r requirements.txt -r requirements-build.txt
   pyinstaller servicio_tool.spec --noconfirm
   ```

2. **Output:** a **single** `dist\servicio_tool.exe`. No Python and no other libraries are needed on the target PC — only this exe.

3. **On the target PC:** Copy `servicio_tool.exe` to a folder, then put `config.json`, `calis.xlsx`, and the `groups\` folder (with your Excel files) in that **same folder**. Run the exe. It will read/write files there.

### Old Windows (e.g. Windows 7)

- Build with **Python 3.8** or **3.9** on a Windows 7–compatible machine if you need to support Windows 7 (Python 3.10+ dropped Win7).
- The single exe may take a few seconds to start on very old PCs (it unpacks to a temp folder). If antivirus flags it, add an exception or use a signed build.

### Antivirus

PyInstaller exes are sometimes flagged. If that happens: add an exception, or use the onedir build and sign the exe if you have a certificate.
