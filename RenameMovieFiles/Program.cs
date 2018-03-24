using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using Microsoft.WindowsAPICodePack.Shell;
using Microsoft.WindowsAPICodePack.Shell.PropertySystem;

namespace RenameMovieFiles
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var logDir = new DirectoryInfo(Directory.GetCurrentDirectory());
            string logFileName = Path.Combine(logDir.FullName, "log.txt");
            if (File.Exists(logFileName))
                File.Delete(logFileName);
            try
            {
                DirectoryInfo directoryInfo = Directory.GetParent(Directory.GetCurrentDirectory()).Parent;
                if (directoryInfo == null)
                    throw new DirectoryNotFoundException();

                string destDir = GetDirectory(args, directoryInfo.FullName);
                if (args.Length == 0)
                    destDir = Environment.GetFolderPath(Environment.SpecialFolder.MyVideos);
                CopyFilesFromDownload(destDir, destDir, "_partial");

                //newDir = createTestFiles(newDir);

                foreach (string path in Directory.GetFiles(destDir))
                {
                    string fileName = Path.GetFileNameWithoutExtension(path);
                    string fileExt = Path.GetExtension(path);
                    string newName = fileName;
                    newName = RemoveCrap(newName);
                    newName = ReplaceChars(newName);
                    newName = ReplaceStrings(newName);
                    newName = FormatDates(newName);
                    newName = CleanUp(newName);

                    Debug.WriteLine(newName);
                    string oldPath = Path.Combine(destDir, fileName + fileExt);
                    string newPath = Path.Combine(destDir, newName + fileExt);
                    File.Move(oldPath, GetUpdatedFileName(newPath));
                    File.SetAttributes(newPath, FileAttributes.Normal);
                    //updateMetadata(Path.Combine(newDir, newName), fileExt);
                }
            }
            catch (Exception e)
            {
                using (var writer = new StreamWriter(logFileName))
                {
                    Console.SetOut(writer);
                    Console.WriteLine(e.Message + Environment.NewLine + e.StackTrace);
                }
            }
        }

        private static string GetUpdatedFileName(string path)
        {
            var count = 1;
            string newSuffix = $"({count})";
            while (true)
            {
                if (!File.Exists(path))
                    return path;
                path = path.Replace(newSuffix, "");
                newSuffix = $"({count})";

                path = Path.Combine(Path.GetDirectoryName(path) ?? Environment.CurrentDirectory,
                           Path.GetFileNameWithoutExtension(path)) + newSuffix + Path.GetExtension(path);
                count++;
            }
        }

        private static void CopyFilesFromDownload(string fromFolder, string toFolder, string excludeFolder)
        {
            //string[] mediaExtensions =
            //{
            //    ".webm", ".mkv", ".flv", ".flv", ".vob", ".ogv", ".ogg", ".drc", ".gif", ".gifv", ".mng", ".avi",
            //    ".mov", ".qt", ".wmv", ".yuv", ".rm", ".rmvb", ".asf", ".amv", ".mp4", ".m4p", ".m4v", ".mpg", "mp2",
            //    ".mpeg", ".mpe", ".mpv", ".mpg", ".mpeg", ".m2v", ".m4v", ".svi", ".3gp", ".3g2", ".mxf", ".roq",
            //    ".nsv", ".flv", ".f4v", ".f4p", ".f4a", ".f4b"
            //};

            string[] mediaExtensions = {".mkv", ".avi", ".mp4", ".sub"};
            foreach (string fromPath in Directory.GetFiles(fromFolder, "*.*", SearchOption.AllDirectories))
            {
                string fromDirectory = Path.GetDirectoryName(fromPath);
                //fromDirectory.Equals(toFolder, StringComparison.OrdinalIgnoreCase)
                if (string.IsNullOrEmpty(fromDirectory))
                    continue;

                if (fromDirectory.Contains(excludeFolder) ||
                    //!fromPath.ToLowerInvariant().Contains(mediaExtensions)
                    !mediaExtensions.Any(s =>
                        fromPath.Contains(s) && !fromPath.EndsWith(".sub", StringComparison.OrdinalIgnoreCase)))
                    continue;

                if (fromPath == null) continue;
                string toPath = Path.Combine(toFolder, Path.GetFileName(fromPath));

                if (IsFileLocked(new FileInfo(fromPath)))
                    Debug.Print($"Skipping {fromPath} because file is open");
                else
                {
                    File.Move(fromPath, Path.Combine(fromPath, GetUpdatedFileName(toPath)));
                    Debug.Print(fromPath + " ----TO----" + toPath);
                }
            }
        }

        private static void updateMetadata(string fileNameWithoutExtension, string fileExt)
        {
            ShellFile file = ShellFile.FromFilePath(fileNameWithoutExtension + fileExt);
            ShellPropertyWriter writer = file.Properties.GetPropertyWriter();
            writer.WriteProperty(SystemProperties.System.Title, "test");
            writer.Close();
            //Debug.WriteLine(file.Properties.System.Title.Value);
            //file.Properties.System.Title.Value = "";// fileNameWithoutExtension;
            //FileInfo info = new FileInfo(fileName);
            //FileAttributes attributes = info.Attributes;
            //Debug.WriteLine(attributes.ToString());
        }

        private static bool IsFileLocked(FileInfo file)
        {
            FileStream stream = null;

            if (!file.Exists) return false;

            try
            {
                stream = file.Open(FileMode.Open, FileAccess.Read, FileShare.None);
            }
            catch (IOException)
            {
                //the file is unavailable because it is:
                //still being written to
                //or being processed by another thread
                //or does not exist (has already been processed)
                return true;
            }
            finally
            {
                stream?.Close();
            }

            //file is not locked
            return false;
        }

        private static string CreateTestFiles(string currDir)
        {
            string newDir = Path.Combine(currDir, "test");
            if (Directory.Exists(newDir))
            {
                Directory.Delete(newDir, true);
            }

            Directory.CreateDirectory(newDir);

            foreach (string path in Directory.GetFiles(currDir))
            {
                string fileName = Path.GetFileName(path);
                FileStream newFile = File.Create(Path.Combine(newDir, fileName));
                newFile.Close();
            }

            return newDir;
        }

        private static string CleanUp(string newName)
        {
            TextInfo textInfo = new CultureInfo("en-US", false).TextInfo;
            newName = textInfo.ToTitleCase(newName);
            newName = string.Join(" ", newName.Split(new[] {' '}, StringSplitOptions.RemoveEmptyEntries));
            return newName;
        }

        private static string GetDirectory(IReadOnlyList<string> args, string currDir)
        {
            if (args.Count > 0)
                currDir = args[0];
            if (!Directory.Exists(currDir))
                Console.WriteLine(currDir + " does not exist");

            return currDir;
        }

        private static string FormatDates(string newName)
        {
            string[] tmp = newName.Split(' ').ToArray().Select(x => x.ToString().Trim()).ToArray();
            if (tmp.Length > 0)
            {
                string lastVal = tmp[tmp.Length - 1];
                if (int.TryParse(lastVal, out int year) && Enumerable.Range(1900, 2100).Contains(year))
                {
                    lastVal = "(" + lastVal + ")";
                    tmp[tmp.Length - 1] = lastVal;
                }
            }

            newName = string.Join(" ", tmp);
            return newName;
        }

        private static string RemoveCrap(string fileName)
        {
            Match match = Regex.Match(fileName, @"\b\d{4}\b");
            string newName = fileName;
            if (match.Success)
            {
                newName = fileName.Substring(0, fileName.IndexOf(match.Value, StringComparison.Ordinal) + match.Length);
            }

            return newName;
        }

        private static string ReplaceStrings(string newName)
        {
            var replacementStrings = new List<string> {"1080p", "720p"};
            List<string> replacementStringsList = replacementStrings.ToList();

            return replacementStringsList.Aggregate(newName, (current, itm) => current.Replace(itm, ""));
        }

        private static string ReplaceChars(string newName)
        {
            const string replacementChars = "._()-[]";
            List<char> replacementCharsList = replacementChars.ToList();
            foreach (char chr in replacementCharsList)
            {
                newName = newName.Replace(chr, ' ');
            }

            return newName;
        }
    }
}