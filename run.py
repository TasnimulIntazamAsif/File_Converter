import os
import sys
import subprocess
import venv
from pathlib import Path

def setup_environment():
    """Set up the virtual environment and install dependencies."""
    venv_path = Path(".venv")
    
    # Create virtual environment if it doesn't exist
    if not venv_path.exists():
        print("Creating virtual environment...")
        venv.create(venv_path, with_pip=True)
    
    # Get the path to the Python executable in the virtual environment
    if sys.platform == "win32":
        python_path = venv_path / "Scripts" / "python.exe"
        pip_path = venv_path / "Scripts" / "pip.exe"
    else:
        python_path = venv_path / "bin" / "python"
        pip_path = venv_path / "bin" / "pip"
    
    # Install requirements
    print("Installing dependencies...")
    subprocess.run([str(pip_path), "install", "-r", "requirements.txt"])
    
    return str(python_path)

def main():
    """Main function to run the application."""
    try:
        # Set up the environment
        python_path = setup_environment()
        
        # Run the application
        print("Starting the application...")
        subprocess.run([python_path, "src/main.py"])
        
    except Exception as e:
        print(f"An error occurred: {str(e)}")
        input("Press Enter to exit...")
        sys.exit(1)

if __name__ == "__main__":
    main() 