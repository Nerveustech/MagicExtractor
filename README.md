# MagicExtractor

**MagicExtractor** is a Windows service designed to automatically extract compressed files downloaded to the current user's Downloads folder, streamlining your file management process. After extraction, the original compressed files are automatically deleted to save space and keep your Downloads folder organized.

## Features

- **Automatic Extraction**: Monitors the Downloads folder and automatically extracts compressed files (ZIP, RAR, etc.) as soon as they are downloaded.
- **Automatic Deletion**: After successful extraction, the original compressed files are deleted automatically.
- **User-Friendly**: Operates seamlessly in the background with no manual intervention required.
- **Broad Compatibility**: Compatible with all recent versions of Windows.

## Installation

1. Download the installation package from [here]().
2. Run the installer and follow the on-screen instructions.
3. Once installed, the service will start working automatically.

### Manual Installation

If you prefer to install MagicExtractor manually, you can use the following commands in the terminal:

```powershell
MagicExtractor.exe --startup=auto install
MagicExtractor.exe start
```

### Removing the Service
To stop and remove the service, use the following commands:

```powershell
MagicExtractor.exe stop
MagicExtractor.exe remove
```

## Usage

After installation, no additional configuration is needed. MagicExtractor will begin monitoring the Downloads folder and automatically extract any compressed files.

## Contributing

If you would like to contribute to MagicExtractor, please follow these steps:

1. Fork the repository.
2. Create a new branch (`git checkout -b feature-name`).
3. Make your changes and add the files (`git add .`).
4. Commit your changes (`git commit -m 'Added new feature'`).
5. Push to the branch (`git push origin feature-name`).
6. Open a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for more details.