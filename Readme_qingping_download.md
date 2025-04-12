# Qingping Project - Quick Start Guide

You are about to run a program called "qingping_download.py"
The program automates data download of multiple Qingping devices from qingpingiot.com (via Google Chrome)
This program will generate pop up which you should fill: email and password of QingpingIoT account, date range, and which device to download (all, online, offline) 
Make sure you have satisfied the requirement explained in "Readme" file
Make sure you have reliable internet connection

## Step-by-Step Instructions

1. **Open Terminal**
   ```
   Command + Space → type "Terminal" → Enter
   ```

2. **Navigate to the project directory**
   ```
   cd ~/Documents/GitHub/qingping_download
   ```

3. **Activate the virtual environment**
   ```
   source env/bin/activate
   ```
   You should see `(env)` appear at the beginning of your command prompt.

4. **Run the script**
   ```
   python qingping_download.py
   ```

5. **When finished, deactivate the environment**
   ```
   deactivate
   ```
