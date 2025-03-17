# Music File Organizer

A PowerShell-based GUI application for organizing and managing your music collection.

## Features

The Music File Organizer provides several powerful features to help you organize your music files:

- **Organize files by artist**: Automatically sorts music files into folders based on the artist metadata
- **Find and delete empty folders**: Cleans up your music directory by removing empty folders
- **Find duplicates**: Identifies duplicate music files based on file checksums
- **Find and delete video files**: Identifies and offers to remove video files from your music collection
- **Find and delete image files**: Identifies and offers to remove image files from your music collection
- **Rename files by metadata**: Renames music files to the format "Artist - Song" using metadata
- **Generate playlist**: Creates an .m3u playlist with all music files from the selected folder
- **Test metadata**: Tests the extraction of metadata from a selected music file

## Requirements

- Windows operating system
- PowerShell 5.1 or later
- Administrative privileges (for some operations)

## Installation

1. Download the `Music_File_Organizer.ps1` script file
2. Right-click on the file and select "Run with PowerShell"
3. Or Download .exe and run from double click, installation not required.

If you encounter execution policy restrictions, you may need to run the following command in an administrative PowerShell window:

```powershell
Set-ExecutionPolicy -ExecutionPolicy Bypass -Scope Process
```

## Usage

1. Launch the application by running the script
2. Select your preferred language (Latvian, English, Lithuanian, or Estonian)
3. Click the "Select Folder" button to choose the music folder you want to organize
4. Select the action you want to perform from the options
5. Click the "START" button to begin the selected operation
6. View progress and logs in the application window

## How It Works

The application uses PowerShell and Windows Shell API to extract metadata from music files. It supports various audio formats including:

- MP3 (.mp3)
- FLAC (.flac)
- WAV (.wav)
- AAC (.aac)
- M4A (.m4a)
- OGG (.ogg)
- WMA (.wma)

## Supported Languages

The application interface is available in four languages:
- Latvian (lv)
- English (en)
- Lithuanian (lt)
- Estonian (ee)

## Technical Details

The application uses several Windows APIs to extract and process music file metadata:

- Shell.Application COM object for file property access
- Windows Media Player COM object as a fallback for metadata extraction
- MD5 hashing for duplicate file detection
- Standard PowerShell file management functions

## Support

If you find this tool useful, consider supporting the developer:

[![Buy Me A Coffee](https://cdn.buymeacoffee.com/buttons/v2/default-yellow.png)](https://buymeacoffee.com/latvietis)

## License

This project is available for free use and modification.

## Disclaimer

Always back up your music collection before performing bulk operations on your files. The developer is not responsible for any data loss that may occur from using this application.
