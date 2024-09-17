# shotlist_creator2
This script is a powerful tool for DaVinci Resolve Studio users, specifically designed for those working on storyboards, VFX breakdowns, and notes. It allows you to export marker data and stills directly to an Excel file, helping streamline your workflow and improve communication with vendors, clients, and artists.

![ShotList_Creator2_DEMO_small](https://github.com/user-attachments/assets/90b118e0-e4ef-423e-b78c-cbef5ab188c4)

## Compatibility

macOS Ventura, Sonoma with DaVinci Resolve Studio 18, 19  
Windows 10, 11 Pro with DaVinci Resolve Studio 18, 19  
Python 3.10.9 (https://www.python.org/downloads/release/python-3109/)  


## Instruction Manual (for both systems):
There are two ways to use this script:

**1. From the box (pre-compiled executable):**
Unzip the downloaded folder to your preferred location. Inside the folder dist, you will find the executable file shotlist_creator2.  
  [download win pre-compiled executable](https://drive.google.com/file/d/1eepAVmB_wWZ88IEg-DoRQeF7_Bg0VyKg/view?usp=sharing)  
  [download mac pre-compiled executable](https://drive.google.com/file/d/1P7dFYC9Cu2k0ga3XAdWc2Rz75CFqR5xG/view?usp=drive_link)  

**2. Running directly from DaVinci Resolve Studio:**  
Copy the file shotlist_creator2.py to the DaVinci Resolve Utility scripts folder:  
*For macOS:* /Library/Application Support/Blackmagic Design/DaVinci Resolve/Fusion/Scripts/Utility/  
*For Windows:* C:\ProgramData\Blackmagic Design\DaVinci Resolve\Fusion\Scripts\Utility\  
Make sure the following Python modules are installed:  
*PySide6*  
*pynput*  
*Pillow*  
*xlsxwriter*  
*DaVinciResolveScript (comes with DaVinci Resolve Studio)*  

## How it works:
  1. Open DaVinci Resolve Studio and load your project.
  2. Go to keyboard customization and assign a key for "Next Marker" (Playback > Next Marker ("0")). This setup is required once. If you run shotlist_creator2.py directly from DaVinci Resolve Studio, you can modify the hotkey in the script and then assign it in the keyboard customization.
  3. Ensure that the album stills1 (in the color page) is empty. This is crucial for the script to function correctly.
  4. Run the script. A dialog box will prompt you to select options such as deleting stills from the album on the color page, setting the timeline timecode, choosing which metadata to extract, and defining the thumbnail size. The script will navigate through the timeline markers, capture thumbnails, and export the marker data and stills to an Excel file in your chosen folder.

[![Watch the video](https://img.youtube.com/vi/lGYmBYw0BuA/maxresdefault.jpg)](https://youtu.be/lGYmBYw0BuA)  

## Additional Tips:

1. For annotations, create a paint node in the Fusion page and add your notes there. Marker annotations and burn-in information will not be exported.
2. The exported file is optimized for size, making it easy to convert to PDF or upload to Google Sheets.
3. On macOS, ensure you grant Terminal accessibility access in Privacy settings.
4. On macOS, it’s recommended to launch DaVinci Resolve Studio from Contents-MacOS-Resolve for better performance.
5. This script works only with the Studio version of DaVinci Resolve.




## Support and Feedback

If this script saved you some time or you just love what it does, please feel free to share your thoughts and consider supporting my work as I continue my journey

### 💖 GitHub Sponsors
[Become a Sponsor](https://github.com/sponsors/natlrazfx)
### ☕ Buy Me a Coffee
[Buy Me a Coffee](https://www.buymeacoffee.com/natlrazfx)
### 💸 PayPal
[PayPal Me](https://paypal.me/natlrazfx)
### 👾 ByBit
119114169


## Cheers :) 

*Footage is provided by Freepik.*
