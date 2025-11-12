Dependency: ffmpeg.  
• Instructions: Place ffmpeg folder (e.g. "ffmpeg-8.0-full_build") (which must directly contain the "bin" folder) in the same directory as this PowerShell script.  

Dependency Download:  
FFMPEG Website: https://www.ffmpeg.org/download.html  
• FFMPEG Windows Builds from Gyan.dev: https://www.gyan.dev/ffmpeg/builds/  
• FFMPEG Windows Builds from BtbN: https://github.com/BtbN/FFmpeg-Builds/releases  

First-Time User of PowerShell Scripts?  
This code is NOT signed, mostly because it is unsupported. To enable PowerShell Scripts to be executed, just Google the error message you get, or follow the instructions here:  
• https://www.google.com/search?q=How+to+enable+PowerShell+Scripts+Windows+11  
• I'd always encourage looking over the code in Notepad, Notepad++, or your other favorite editor before running it. Functionality is outlined below.  
  
Functionality:  
1. Checks for dependency ffmpeg, erring with instructions if not found.  
2. Prompts user to "Copy Debug Info" from the video they wish to download (accessible by right-clicking the video).  
3. Prompts for the base filename. You may either accept the default, or provide a meaningful name. A files will use the chosen name, within a subfolder by the same name.  
4. Downloads the MP4 Video (in parts, then assembled). If an MP4 already exists at the same basename location, the user is prompted whether or not to re-download (to save time and bandwidth).  
5. Downloads the VTT Subtitles (in parts, then assembled).  
6. Converts the VTT Subtitles Format to SRT Subtitles Format (saving both formats).  
  
Disclaimer:  
This code was written by AI: It was rapidly developed using both ChatGPT and CoPilot.  
While I found it functional and gave it cursory reviews, it is unsupported. Feedback monitoring will be rare.   
Feel free to modify to improve as desired. No attribution is required.  
  
