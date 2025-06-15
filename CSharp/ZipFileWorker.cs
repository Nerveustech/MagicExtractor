/*
 * MIT License
 *
 * Copyright (c) 2025 Andrea Michael M. Molino
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
*/

using System;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Threading.Tasks;


public class ZipFileWorker
{

    public async Task UnzipFile(string ZipFilePath)
    {
        MagicExtractor.WinAPI winAPI = new MagicExtractor.WinAPI();
        string Folder = Path.Combine(Path.GetDirectoryName(ZipFilePath), Path.GetFileNameWithoutExtension(ZipFilePath));

        DirectoryInfo DirectoryOfZip;

        if (Directory.Exists(Folder))
        {
            string IncrementedDirectoryName = GetIncrementedDirectoryName(Folder);
            DirectoryOfZip = Directory.CreateDirectory(Path.Combine(Path.GetDirectoryName(IncrementedDirectoryName), Path.GetFileNameWithoutExtension(IncrementedDirectoryName)));
        }
        else
        {
            DirectoryOfZip = Directory.CreateDirectory(Folder);
        }

        //FIXME: System.IO.IOException: 'The process cannot access the file 'C:\Users\Don\Downloads\Service.zip' because it is being used by another process.'
        using (FileStream zipFileStream = new FileStream(ZipFilePath, FileMode.Open, FileAccess.Read))
        using (ZipArchive zipArchive = new ZipArchive(zipFileStream, ZipArchiveMode.Read))
        {

            foreach (ZipArchiveEntry entry in zipArchive.Entries)
            {

                string destinationPath = Path.Combine(DirectoryOfZip.FullName, entry.FullName);


                if (entry.FullName.EndsWith("/"))
                {
                    Directory.CreateDirectory(destinationPath);
                    continue;
                }


                string directoryPath = Path.GetDirectoryName(destinationPath);
                if (!Directory.Exists(directoryPath))
                {
                    Directory.CreateDirectory(directoryPath);
                }


                await ExtractFile(entry, destinationPath);
                
            }
        }
        winAPI.SendFileToRecycleBin(ZipFilePath);
        //MagicExtractor.WinAPI.SendFileToRecycleBin(ZipFilePath);
    }

    // Asynchronous file extraction method
    private async Task ExtractFile(ZipArchiveEntry entry, string destinationPath)
    {
        // Open the entry stream and the output file stream
        using (Stream entryStream = entry.Open())
        using (FileStream fileStream = new FileStream(destinationPath, FileMode.Create, FileAccess.Write))
        {
            await entryStream.CopyToAsync(fileStream);
        }
        
    }

    private string GetIncrementedDirectoryName(string Folder)
    {
        string ParentDirectory = Path.GetDirectoryName(Folder);
        string DirectoryName = Path.GetFileName(Folder);

        int Index = 2;
        string DirectoryToIncrement = Folder;

        while (Directory.Exists(DirectoryToIncrement))
        {
            DirectoryToIncrement = Path.Combine(ParentDirectory, $"{DirectoryName} ({Index})");
            Index++;
        }
        return DirectoryToIncrement;
    }

}

/*
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MagicExtractor
{
    internal class ZipFileWorker
    {
    }
}
*/