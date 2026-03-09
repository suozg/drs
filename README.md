# DRS

Программа створює зашифровану базу даних з окремих файлів msoffice і дозволяє здійснювати пошук по тексту.
Створена для заміни FileLocator для не-Windows систем.

![drs](screenshots/Screenshot%202025-06-21%2015.18.46.png)

## Збірка

1. Prerequisites
Python 3.11: Ensure Python 3.11 is installed.

2. Project Setup               
    2.1 Download project from Git or clone

    2.2 In PowerShell
   
        cd Project_DIR 

    2.3 Create a virtual environment
         
        python -m venv venv

    2.4 Activate the virtual environment
      
        source venv/bin/activate

4. Install Dependencies
Once the virtual environment is active (venv), run the following commands:

        python -m pip install --upgrade pip

Install required packages

        pip install wxPython pysqlcipher3-binary pyinstaller          

4. Verification and Building

    4.1 Test Run:
   
        python drs_wx.py

    4.2 If the application starts successfully, use PyInstaller to build the .exe file using your spec file:
   
        pyinstaller drs_wx.spec
