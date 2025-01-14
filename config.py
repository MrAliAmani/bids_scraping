# System Paths Configuration
PATHS = {
    "PYTHON_EXECUTABLE": r"C:\Users\AliAmani\Miniconda3\python.exe",
    "CONDA_SCRIPTS": r"C:\Users\AliAmani\Miniconda3\Scripts",
    "CONDA_ENV_PATH": r"C:\Users\AliAmani\Miniconda3\envs\bids",
    "CONDA_ENV_NAME": "bids"
}

# Verify these paths exist on startup
def verify_paths():
    import os
    missing_paths = []
    
    # Only verify paths, not the environment name
    paths_to_verify = {k: v for k, v in PATHS.items() if k != "CONDA_ENV_NAME"}
    
    for name, path in paths_to_verify.items():
        if not os.path.exists(path):
            missing_paths.append(f"{name}: {path}")
    
    if missing_paths:
        raise Exception(
            "Missing required paths:\n" + 
            "\n".join(missing_paths) + 
            "\n\nPlease update config.py with correct paths for your system."
        ) 