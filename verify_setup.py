import os
import subprocess
from config import PATHS, verify_paths

def check_conda_env():
    """Verify conda environment exists and has required packages"""
    try:
        # Check if conda is available
        subprocess.run(
            ["conda", "--version"], 
            check=True, 
            capture_output=True, 
            text=True
        )
        
        # Check if environment exists
        result = subprocess.run(
            ["conda", "env", "list"], 
            check=True, 
            capture_output=True, 
            text=True
        )
        
        env_name = PATHS["CONDA_ENV_NAME"]
        env_path = PATHS["CONDA_ENV_PATH"]
        
        # Check both the environment name and path
        if env_name not in result.stdout:
            print(f"❌ Conda environment '{env_name}' not found")
            print(f"📝 Create it using: conda create -n {env_name} python=3.8")
            return False
        
        if not os.path.exists(env_path):
            print(f"❌ Conda environment path not found: {env_path}")
            print(f"📝 Make sure the environment is created and the path is correct")
            return False
            
        print(f"✅ Conda environment '{env_name}' found and verified")
        return True
        
    except subprocess.CalledProcessError:
        print("❌ Conda is not installed or not in PATH")
        print("📝 Please install Miniconda from: https://docs.conda.io/en/latest/miniconda.html")
        return False

def main():
    print("🔍 Verifying system setup...")
    
    try:
        # First check conda environment
        conda_ok = check_conda_env()
        
        # Then check paths
        try:
            verify_paths()
            print("✅ All required paths exist")
            paths_ok = True
        except Exception as e:
            print(f"\n❌ Path verification failed: {str(e)}")
            paths_ok = False
        
        if conda_ok and paths_ok:
            print("\n✅ System setup looks good! You can now run app.py")
        else:
            print("\n❌ Please fix the above issues before running app.py")
            print("\n📝 Instructions for setting up paths:")
            print("1. Open config.py in a text editor")
            print("2. Update the PATHS dictionary with correct paths for your system")
            print("3. Common paths on Shadow PC:")
            print("   - PYTHON_EXECUTABLE: C:\\Users\\[username]\\Miniconda3\\python.exe")
            print("   - CONDA_SCRIPTS: C:\\Users\\[username]\\Miniconda3\\Scripts")
            print("   - CONDA_ENV_PATH: C:\\Users\\[username]\\Miniconda3\\envs\\bids")
            print("4. Run this script again to verify the updates")
            
    except Exception as e:
        print(f"\n❌ Setup verification failed: {str(e)}")

if __name__ == "__main__":
    main() 