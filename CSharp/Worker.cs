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
using System.Management;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;

namespace MagicExtractor
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly FileSystemWatcher _FileSystemWatcher;
        private string _FolderDownloads;
        private readonly ZipFileWorker _ZipFile;

        public Worker(ILogger<Worker> logger)
        {
            _logger = logger;
            _FolderDownloads = SetFolderDownloads();
            _FileSystemWatcher = new FileSystemWatcher
            {
                //Path = _FolderDownloads,
                Path = @"C:\Users\Don\Desktop\Codes\Tests",
                NotifyFilter = NotifyFilters.FileName | NotifyFilters.LastWrite,
                Filter = "*.zip",
                EnableRaisingEvents = true
            };

            _FileSystemWatcher.Created += OnFileCreated;
            //_fileSystemWatcher.Changed += OnFileChanged;
            //_fileSystemWatcher.Deleted += OnFileDeleted;
            //_fileSystemWatcher.Renamed += OnFileRenamed;
            _ZipFile = new ZipFileWorker();
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            _logger.LogInformation("Worker running at: {time}", DateTimeOffset.Now);
            //_logger.LogInformation($"User: {SetFolderDownloads()}");

            while (!stoppingToken.IsCancellationRequested) {
                await Task.Delay(1000, stoppingToken);
            }
        }

        private async void OnFileCreated(object sender, FileSystemEventArgs e)
        {
            //_logger.LogInformation($"User: {GetUsername()}");
            _logger.LogInformation($"File created: {e.FullPath}");
            await _ZipFile.UnzipFile(e.FullPath);
        }

        private void OnFileChanged(object sender, FileSystemEventArgs e)
        {
            _logger.LogInformation($"File changed: {e.FullPath}");
        }

        private void OnFileDeleted(object sender, FileSystemEventArgs e)
        {
            _logger.LogInformation($"File deleted: {e.FullPath}");
        }

        private void OnFileRenamed(object sender, RenamedEventArgs e)
        {
            _logger.LogInformation($"File renamed: {e.OldFullPath} to {e.FullPath}");
        }

        public override Task StopAsync(CancellationToken cancellationToken)
        {
            _logger.LogInformation("Stopping Worker...");
            _FileSystemWatcher.EnableRaisingEvents = false;
            _FileSystemWatcher.Created -= OnFileCreated;
            _FileSystemWatcher.Dispose();
            return base.StopAsync(cancellationToken);
        }

        private string SetFolderDownloads()
        {
#pragma warning disable CA1416 // Convalida compatibilità della piattaforma

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT UserName FROM Win32_ComputerSystem");
            ManagementObjectCollection collection = searcher.Get();
            string DomainUser = (string)collection.Cast<ManagementBaseObject>().First()["UserName"];

#pragma warning restore CA1416 // Convalida compatibilità della piattaforma

            string[] User = DomainUser.Split('\\');
            return Path.Combine(Path.GetPathRoot(Environment.SystemDirectory.Substring(0, 3)), $"Users\\{User[1]}\\Downloads");
        }
    }
}
