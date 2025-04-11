# Project Setup Instructions

This README provides detailed steps to set up the required environment for the project on both Mac and Windows systems, including installing dependencies and configuring the Chrome WebDriver.

## Prerequisites

Before proceeding, ensure the following are installed on your system:

- **Python 3.8+**
- **pip (Python package manager)**

## 1. Install Python and pip

### On Mac
1. Open the Terminal.
2. Install Homebrew (if not already installed):
   ```bash
   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
   ```
3. Install Python:
   ```bash
   brew install python3
   ```
4. Verify the installation:
   ```bash
   python3 --version
   pip3 --version
   ```

### On Windows
1. Download the Python installer from [Python.org](https://www.python.org/downloads/).
2. Run the installer and ensure "Add Python to PATH" is checked.
3. Verify the installation in Command Prompt:
   ```cmd
   python --version
   pip --version
   ```

## 2. Install Required Python Packages

Install the required dependencies by running the following command in your terminal (Mac) or Command Prompt (Windows):

```bash
pip install tkinter pandas selenium
```

## 3. Set Up ChromeDriver

### Determine Your Chrome Version

1. Open Chrome.
2. Navigate to `chrome://settings/help`.
3. Note the version of Chrome installed.

### Download ChromeDriver

1. Visit the [ChromeDriver downloads page](https://sites.google.com/chromium.org/driver/).
2. Download the version that matches your Chrome version.
3. Extract the downloaded file.

### Add ChromeDriver to PATH

#### On Mac
1. Move the `chromedriver` binary to `/usr/local/bin`:
   ```bash
   mv /path/to/chromedriver /usr/local/bin/
   ```
2. Verify the installation:
   ```bash
   chromedriver --version
   ```

#### On Windows
1. Move the `chromedriver.exe` file to a directory in your PATH, such as `C:\Windows\System32`.
2. Verify the installation:
   ```cmd
   chromedriver --version
   ```

## 4. Run the Project

1. Clone the project repository or download the project files.
2. Navigate to the project directory in your terminal or Command Prompt.
3. Run the main script:
   ```bash
   python main.py
   ```

## Troubleshooting

- If you encounter a "Driver not found" error, ensure `chromedriver` is correctly added to your PATH.
- For issues with dependencies, try reinstalling them:
  ```bash
  pip install --force-reinstall -r requirements.txt
  ```
- Ensure your Chrome browser is up-to-date.

## Additional Notes

- If the project involves GUI interactions with `tkinter`, ensure your display is configured properly.
- For automated browser tests, consider using a virtual environment to avoid conflicts:
  ```bash
  python -m venv env
  source env/bin/activate  # On Mac
  env\Scripts\activate    # On Windows
  
