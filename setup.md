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
4. Install essential packages:
   ```bash
   # Web scraping essentials
   pip install selenium==4.25.0
   pip install selenium-stealth==1.0.6
   pip install undetected-chromedriver==3.5.5
   pip install webdriver-manager==4.0.2
   pip install beautifulsoup4==4.12.3
   pip install fake-useragent==1.5.1

   # Data processing
   pip install numpy==1.24.3
   pip install pandas==2.2.3
   pip install openpyxl==3.1.5
   pip install lxml==5.3.0

   # Web framework and sockets
   pip install flask==3.1.0
   pip install flask-socketio==5.4.1
   pip install python-socketio==5.11.4
   pip install python-engineio==4.10.1

   # Storage and file handling
   pip install minio==7.2.10
   pip install boto3==1.35.54
   pip install watchdog==6.0.0

   # UI automation and monitoring
   pip install pyautogui==0.9.54
   pip install pynput==1.7.7
   pip install keyboard==0.13.5
   pip install psutil==6.1.0

   # Utilities
   pip install python-dotenv==1.0.1
   pip install requests==2.32.3
   pip install fuzzywuzzy==0.18.0
   pip install python-Levenshtein==0.26.0
   pip install rich==13.9.2

   # Testing
   pip install pytest==8.3.3
   pip install pytest-mock==3.14.0
   ```

5. Install spaCy and its model:
   ```bash
   pip install spacy==3.8.2
   python -m spacy download en_core_web_sm
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

5. Run the command below to setupt the dependencies:
   ```bash
   pip install -r requirements.txt
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
