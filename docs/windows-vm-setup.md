# Install Windows VM + OpenDental on Mac (UTM)

## Step 1: Download Windows 11 ARM (Free)

1. Open Safari and go to: https://www.microsoft.com/en-us/software-download/windows11arm64
2. Click **"Download Now"** to get the VHDX file (~5GB)
3. Wait for download to finish

## Step 2: Create VM in UTM

1. Open **UTM** app on your Mac
2. Click **"Create a New Virtual Machine"**
3. Select **"Virtualize"**
4. Select **"Windows"**
5. Click **"Import VHDX Image"** and select the file you downloaded
6. Settings:
   - RAM: **4096 MB** (4GB minimum)
   - CPU Cores: **2**
   - Storage: **64 GB**
7. Click **"Save"** and then **"Play"** to start the VM
8. Follow Windows setup wizard (skip Microsoft account — click "I don't have internet")

## Step 3: Install OpenDental Inside the VM

Once Windows is running in UTM:

1. Open **Edge** browser inside the VM
2. Go to: **https://www.opendental.com/trial.html**
3. Download **TrialDownload-25-3-48.exe**
4. Right-click the downloaded file → **Run as Administrator**
5. Follow the installer:
   - Click Next through all steps
   - When asked for database: choose **"Demo database"**
   - This installs OpenDental + MySQL automatically
6. Launch OpenDental when installer finishes
7. At login screen, just press **Enter** (Admin with no password)

## Step 4: Install Python + Bot Inside the VM

Open **PowerShell** inside the VM and run:

```powershell
# Install Python
winget install Python.Python.3.11

# Restart PowerShell, then:
git clone https://github.com/svreddy-design/automation.git
cd automation
pip install -r requirements.txt

# Run the bot
python app.py
```

## Step 5: Test the Bot

1. Make sure OpenDental is running in the background
2. In the bot GUI, click **"Select App Executable"**
3. Browse to: `C:\Program Files\Open Dental\OpenDental.exe`
4. Fill in test patient data (or use defaults)
5. Click **"Run PyWinAuto"**
6. Watch the bot fill in all 12 fields automatically!

## Troubleshooting

- **VM is slow**: Increase RAM to 6GB or 8GB in UTM settings
- **OpenDental won't install**: Make sure you right-click → Run as Administrator
- **Bot can't find OpenDental**: Manually type the path in the app
- **Typing too fast**: Increase `typing_interval_ms` in config.json (try 100)
