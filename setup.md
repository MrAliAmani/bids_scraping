# Shadow PC Setup Instructions

## 1. Install Miniconda
1. Connect to your Shadow PC
2. Download Miniconda for Windows: https://docs.conda.io/en/latest/miniconda.html
3. Run the Miniconda installer
   - Install for "Just Me (recommended)"
   - Use the default installation path (usually `C:\Users\[username]\Miniconda3`)
   - Check "Add Miniconda3 to my PATH environment variable"

## 2. Create Conda Environment
1. Open Command Prompt (cmd)
2. Create the bids environment:
   ```bash
   conda create -n bids python=3.12
   ```
3. Activate the environment:
   ```bash
   conda activate bids
   ```

## 3. Configure Paths
1. Find your Shadow PC username:
   - Open Command Prompt
   - Type `echo %USERNAME%` to see your username

2. Open `config.py` in a text editor
3. Update the PATHS dictionary with your username:
   ```python
   PATHS = {
       "PYTHON_EXECUTABLE": r"C:\Users\[your_username]\Miniconda3\python.exe",
       "CONDA_SCRIPTS": r"C:\Users\[your_username]\Miniconda3\Scripts",
       "CONDA_ENV_PATH": r"C:\Users\[your_username]\Miniconda3\envs\bids",
       "CONDA_ENV_NAME": "bids"
   }
   ```
   Replace `[your_username]` with your actual Shadow PC username.

## 4. Verify Setup
1. Open Command Prompt
2. Navigate to your project directory:
   ```bash
   cd path\to\your\project
   ```
3. Run the verification script:
   ```bash
   python verify_setup.py
   ```
4. If any issues are found, the script will provide specific instructions to fix them

## Common Issues and Solutions

### "Conda is not recognized"
- Close and reopen Command Prompt after installing Miniconda
- If still not working, manually add Miniconda to PATH:
  1. Search for "Environment Variables" in Windows
  2. Under "User variables", edit PATH
  3. Add these paths:
     ```
     C:\Users\[your_username]\Miniconda3
     C:\Users\[your_username]\Miniconda3\Scripts
     C:\Users\[your_username]\Miniconda3\Library\bin
     ```

### "Path not found" errors
- Double-check your username in the paths
- Verify Miniconda installed correctly
- Make sure you created the bids environment

### "Permission denied" errors
- Run Command Prompt as administrator
- Check file permissions in the Miniconda3 folder

## Verification Checklist
✅ Miniconda installed
✅ bids environment created
✅ Paths updated in config.py
✅ verify_setup.py runs successfully

Once all checks pass, you can run app.py to start the scraping dashboard.

## If you encounter any issues:
1. Run `verify_setup.py` to get detailed error messages
2. Check the exact paths on your system:
   ```bash
   conda info --base
   conda env list
   ```
