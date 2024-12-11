//=====================================================================================================================================//
//                              Below Lines Help to use Plugin with Measures                                                           //
//=====================================================================================================================================//

/*
[SwapShortcut]
Measure=Plugin
Plugin=YourStart
IniPath=#@#AppMeters.nek
[!CommandMeasure SwapShortcut "SwapShortcut 2 3"]

[CheckApps]
Measure=Plugin
Plugin=YourStart
TotalApps=84
Type=CheckApps
IniPath=#@#AllApps.nek
IconFolder=
NoMatchAction=

[ChangePath]
Measure=Plugin
Plugin=YourStart
IniPath=#@#AppMeters.nek
[!CommandMeasure ChangePath "ChangePath 2"]

[ChangeIcon]
Measure=Plugin
Plugin=YourStart
IniPath=#@#AppMeters.nek
[!CommandMeasure ChangeIcon "ChangeIcon 3"]

[ChangeName]
Measure=Plugin
Plugin=YourStart
IniPath=#@#AppMeters.nek
[!CommandMeasure ChangeName "ChangeName 2 NewName Here"]

[RemoveShortcut]
Measure=Plugin
Plugin=YourStart
[!CommandMeasure RemoveShortcut "RemoveShortcut 2"]

[AddShortcut]
Measure=Plugin
Plugin=YourStart
IconFolder="C:\\Users\\Nasir Shahbaz\\Desktop\\New folder"
DefaultIconPath="C:\\Users\\Nasir Shahbaz\\Desktop\\icon.ico"
TotalShortcuts=#TotalShortcuts#
WriteStringMeter=1
IniPath=#@#AppMeters.nek
OnCompleteAction=
[!CommandMeasure AddShortcut "addshortcut 1"]

[AllApps]
Measure=Plugin
Plugin=YourStart
TotalApps=#TotalApps#
IconFolder="C:\\Users\\Nasir Shahbaz\\Desktop\\New folder"
IniPath=#@#AllApps.nek
OnCompleteAction=
[!CommandMeasure AllApps "writeapps 1"]
*/


//=====================================================================================================================================//
//                                             Main  Code Start Here                                                                   //
//=====================================================================================================================================//
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Rainmeter;

namespace YourStart
{
    internal class Measure
    {
        private Rainmeter.API api;
        private string iconFolder;
        private string Type;
        private string iniPath;
        private string OnCompleteAction;
        private string NoMatchAction;
        private int totalApps;
        private List<AppInfo> appList;
        private int TotalShortcut;



        internal Measure()
        {
            appList = new List<AppInfo>();
        }

        internal void Reload(Rainmeter.API api, ref double maxValue)
        {
            this.api = api;
            iconFolder = api.ReadString("IconFolder", "").Trim();
            iniPath = api.ReadString("IniPath", "").Trim();
            OnCompleteAction = api.ReadString("OnCompleteAction", "").Trim();
            NoMatchAction = api.ReadString("NoMatchAction", "").Trim();
            totalApps = api.ReadInt("TotalApps", 0);
            TotalShortcut = api.ReadInt("TotalShortcut", 0);

            if (string.IsNullOrEmpty(iconFolder) || string.IsNullOrEmpty(iniPath))
            {
                api.Log(API.LogType.Error, "YourStart.dll: 'IconFolder', 'JsonPath', and 'IniPath' must be specified.");
                return;
            }


            if (!Directory.Exists(iconFolder))
            {
                try
                {
                    Directory.CreateDirectory(iconFolder);
                }
                catch (Exception ex)
                {
                    api.Log(API.LogType.Error, $"YourStart.dll: Failed to create directory '{iconFolder}': {ex.Message}");
                    return;
                }
            }

            // Read and validate the Type parameter
            Type = api.ReadString("Type", "").Trim();
            if (string.IsNullOrEmpty(Type))
            {
                api.Log(API.LogType.Warning, "YourStart.dll: 'Type' is not defined. Default behavior will apply.");
            }

            // Perform actions based on the Type parameter
            if (Type.Equals("CheckApps", StringComparison.OrdinalIgnoreCase))
            {
                ExecuteCheckApps();
            }
            else
            {
                api.Log(API.LogType.Debug, $"YourStart.dll: No specific actions defined for Type '{Type}'.");
            }
        }
        internal void Execute(string args)
        {
            if (string.IsNullOrEmpty(args))
            {
                LogError("No arguments provided.");
                return;
            }

            var arguments = args.Split(' ');
            if (arguments.Length < 2)
            {
                LogError("Invalid arguments. Supported commands are WriteApps, CheckApps, AddShortcut, SwapShortcut, RemoveShortcut, or ChangePath.");
                return;
            }

            string command = arguments[0].ToLower();
            switch (command)
            {
                case "writeapps":
                    ExecuteWriteApps();
                    break;

                case "addshortcut":
                    ExecuteAddShortcut();
                    break;

                case "swapshortcut":
                    ExecuteSwapShortcut(arguments);
                    break;

                case "removeshortcut":
                    ExecuteRemoveShortcut(arguments);
                    break;
                case "changeicon":
                    ChangeIconShortcut(arguments);
                    break;

                case "changepath":
                    ExecuteChangePath(arguments);
                    break;
                case "changename":
                    ExecuteChangeName(arguments);
                    break;

                default:
                    LogError($"Unknown command '{command}'.");
                    break;
            }
        }

        private void ExecuteChangeName(string[] arguments)
        {
            if (arguments.Length >= 3 && int.TryParse(arguments[1], out int index))
            {
                // Combine the rest of the arguments as the new name
                string newName = string.Join(" ", arguments.Skip(2));
                ChangeName(index, newName);
            }
            else
            {
                LogError("Invalid ChangeName arguments. Expected format: 'ChangeName index new name'.");
            }
        }
        private void ExecuteWriteApps()
        {
            if (string.IsNullOrEmpty(iniPath))
            {
                LogError("IniPath is not defined.");
                return;
            }

            UpdateAppList();
            SaveToIniFile(iniPath);
        }

        private void ExecuteCheckApps()
        {
            AppList();

            if (appList.Count != totalApps)
            {
                api.Execute(NoMatchAction);
            }
        }

        private void ExecuteAddShortcut()
        {
            AddShortcut();
        }

        private void ExecuteSwapShortcut(string[] arguments)
        {
            if (arguments.Length == 3 && int.TryParse(arguments[1], out int i1) && int.TryParse(arguments[2], out int i2))
            {
                SwapShortcut(i1, i2);
            }
            else
            {
                LogError("Invalid SwapShortcut arguments. Expected format: 'SwapShortcut i1 i2'.");
            }
        }

        private void ExecuteRemoveShortcut(string[] arguments)
        {
            if (arguments.Length == 2 && int.TryParse(arguments[1], out int index))
            {
                RemoveShortcut(index);
            }
            else
            {
                LogError("Invalid RemoveShortcut arguments. Expected format: 'RemoveShortcut index'.");
            }
        }

        private void ChangeIconShortcut(string[] arguments)
        {
            if (arguments.Length == 2 && int.TryParse(arguments[1], out int index))
            {
                ChangeIcon(index);
            }
            else
            {
                LogError("Invalid ChangeIcon arguments. Expected format: 'ChangeShortcut index'.");
            }
        }

        private void ExecuteChangePath(string[] arguments)
        {
            if (arguments.Length == 2 && int.TryParse(arguments[1], out int shortcutIndex))
            {
                ChangePath(shortcutIndex);
            }
            else
            {
                LogError("Invalid ChangePath arguments. Expected format: 'ChangePath index'.");
            }
        }

        private void LogError(string message)
        {
            api.Log(API.LogType.Error, $"YourStart.dll: {message}");
        }

        internal double Update()
        {
            return appList.Count;
        }
        //=====================================================================================================================================//
        //                                                Function of Check Apps List                                                          //
        //=====================================================================================================================================//
        private void AppList()
        {
            appList.Clear();
            string[] startMenuPaths = {
        Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu) + "\\Programs",
        Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) + "\\Programs"
    };

            var uniqueApps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var path in startMenuPaths)
            {
                if (Directory.Exists(path))
                {
                    foreach (var shortcut in Directory.GetFiles(path, "*.lnk", SearchOption.AllDirectories))
                    {
                        try
                        {
                            string exePath = GetShortcutTarget(shortcut);

                            // Validate the target
                            if (string.IsNullOrEmpty(exePath) || !File.Exists(exePath) || !exePath.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                            {
                                continue;
                            }

                            string appName = Path.GetFileNameWithoutExtension(exePath);
                            string uniqueKey = $"{appName.ToLowerInvariant()}_{exePath.ToLowerInvariant()}";

                            if (uniqueApps.Contains(uniqueKey)) continue;

                            // Add AppInfo without handling icons
                            appList.Add(new AppInfo
                            {
                                Name = appName,
                                Path = exePath,
                                IconPath = string.Empty // Leave IconPath empty since icons are skipped
                            });

                            uniqueApps.Add(uniqueKey);


                        }
                        catch (Exception ex)
                        {
                            api.Log(API.LogType.Error, $"YourStart.dll: Error processing shortcut '{shortcut}': {ex.Message}");
                        }
                    }
                }
            }
        }
        //=====================================================================================================================================//
        //                                                Function to Change Icon                                                              //
        //=====================================================================================================================================//
        private void ChangeIcon(int index)
        {
            api.Log(API.LogType.Debug, $"YourStart.dll: Attempting to change icon for shortcut {index}");

            // Open a file dialog to select an image
            string selectedFile = OpenFileDialog();
            if (string.IsNullOrEmpty(selectedFile))
            {
                api.Log(API.LogType.Warning, "YourStart.dll: No file selected. Operation aborted.");
                return;
            }

            api.Log(API.LogType.Debug, $"Selected file: {selectedFile}");

            // Read the INI file content
            var iniContent = ReadIniFile(iniPath);
            if (iniContent == null)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Failed to read INI file: {iniPath}");
                return;
            }

            // Update the ImageName for the specified shortcut
            string section = $"Shortcut_Icon_{index}";
            if (iniContent.ContainsKey(section))
            {
                iniContent[section]["ImageName"] = selectedFile;
                api.Log(API.LogType.Debug, $"YourStart.dll: Updated ImageName for {section} to {selectedFile}");

                // Write back the updated INI file
                WriteIniFile(iniPath, iniContent);

                // Refresh Rainmeter
                api.Execute("[!UpdateMeter *][!Redraw]");
            }
            else
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Section {section} not found in INI file.");
            }
        }

        private string OpenFileDialog()
        {
            string selectedFile = null;

            // Use a Windows Forms OpenFileDialog to select an image
            try
            {
                using (var dialog = new System.Windows.Forms.OpenFileDialog())
                {
                    dialog.Filter = "Image Files|*.bmp;*.jpg;*.jpeg;*.png;*.ico";
                    dialog.Title = "Select an Image File";

                    if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        selectedFile = dialog.FileName;
                    }
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Error while opening file dialog: {ex.Message}");
            }

            return selectedFile;
        }
        //=====================================================================================================================================//
        //                                                Function to Change ShortcutName                                                      //
        //=====================================================================================================================================//
        private void ChangeName(int index, string newName)
        {
            if (string.IsNullOrEmpty(iniPath))
            {
                LogError("IniPath is not defined.");
                return;
            }

            if (!File.Exists(iniPath))
            {
                LogError($"INI file '{iniPath}' not found.");
                return;
            }

            // Load the INI file
            var iniFile = ReadIniFile(iniPath);

            // Section for the Shortcut_String_X
            string section = $"Shortcut_String_{index}";

            if (!iniFile.ContainsKey(section))
            {
                LogError($"Section '{section}' does not exist in the INI file.");
                return;
            }

            // Update the "Text" key with the new name
            if (iniFile[section].ContainsKey("Text"))
            {
                iniFile[section]["Text"] = newName;

                // Corrected WriteIniFile call
                WriteIniFile(iniPath, iniFile);

                // Apply the changes dynamically
                api.Execute($"[!SetOption {section} Text \"{newName}\"][!UpdateMeter {section}][!Redraw]");
            }
            else
            {
                LogError($"Key 'Text' does not exist in section '{section}'.");
            }
        }
        //=====================================================================================================================================//
        //                                                Function to Swap Shortcut                                                            //
        //=====================================================================================================================================//

        private void SwapShortcut(int i1, int i2)
        {
            api.Log(API.LogType.Debug, $"YourStart.dll: Attempting to swap shortcuts {i1} and {i2}");

            // Load the INI content
            var iniContent = ReadIniFile(iniPath);
            if (iniContent == null)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Failed to read INI file: {iniPath}");
                return;
            }

            api.Log(API.LogType.Debug, "YourStart.dll: INI file loaded successfully.");

            // Validate required sections
            var requiredSections = new[] {
        $"Shortcut_BackGround_{i1}", $"Shortcut_Icon_{i1}", $"Shortcut_String_{i1}",
        $"Shortcut_BackGround_{i2}", $"Shortcut_Icon_{i2}", $"Shortcut_String_{i2}"
    };

            foreach (var section in requiredSections)
            {
                if (!iniContent.ContainsKey(section))
                {
                    api.Log(API.LogType.Error, $"YourStart.dll: Missing section [{section}] in {iniPath}");
                    return;
                }
            }

            api.Log(API.LogType.Debug, "YourStart.dll: All required sections found. Proceeding with swap.");

            // Swap background LeftMouseDownAction
            SwapValues(iniContent, $"Shortcut_BackGround_{i1}", $"Shortcut_BackGround_{i2}", "LeftMouseDownAction");
            api.Execute($"[!SetOption Shortcut_BackGround_{i1} LeftMouseDownAction \"{iniContent[$"Shortcut_BackGround_{i2}"]["LeftMouseDownAction"]}\"]");
            api.Execute($"[!SetOption Shortcut_BackGround_{i2} LeftMouseDownAction \"{iniContent[$"Shortcut_BackGround_{i1}"]["LeftMouseDownAction"]}\"]");



            // Swap icons
            SwapValues(iniContent, $"Shortcut_Icon_{i1}", $"Shortcut_Icon_{i2}", "ImageName");
            api.Execute($"[!SetOption Shortcut_Icon_{i1} ImageName \"{iniContent[$"Shortcut_Icon_{i2}"]["ImageName"]}\"]");
            api.Execute($"[!SetOption Shortcut_Icon_{i2} ImageName \"{iniContent[$"Shortcut_Icon_{i1}"]["ImageName"]}\"]");

            // Swap strings
            SwapValues(iniContent, $"Shortcut_String_{i1}", $"Shortcut_String_{i2}", "Text");
            api.Execute($"[!SetOption Shortcut_String_{i1} Text \"{iniContent[$"Shortcut_String_{i2}"]["Text"]}\"]");
            api.Execute($"[!SetOption Shortcut_String_{i2} Text \"{iniContent[$"Shortcut_String_{i1}"]["Text"]}\"]");

            // Write back to the INI file
            WriteIniFile(iniPath, iniContent);

            // Update Rainmeter
            api.Execute("[!UpdateMeter *][!Redraw]");
            api.Log(API.LogType.Debug, "YourStart.dll: Swap completed successfully.");
        }


        private void SwapValues(Dictionary<string, Dictionary<string, string>> iniContent, string section1, string section2, string key)
        {
            var temp = iniContent[section1][key];
            iniContent[section1][key] = iniContent[section2][key];
            iniContent[section2][key] = temp;
        }


        private Dictionary<string, Dictionary<string, string>> ParseIni(List<string> iniContent)
        {
            var result = new Dictionary<string, Dictionary<string, string>>();
            string currentSection = null;

            foreach (var line in iniContent)
            {
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith(";")) continue;

                var match = Regex.Match(line, @"^\[(.+)\]$");
                if (match.Success)
                {
                    currentSection = match.Groups[1].Value;
                    result[currentSection] = new Dictionary<string, string>();
                }
                else if (currentSection != null)
                {
                    var keyValue = line.Split(new[] { '=' }, 2);
                    if (keyValue.Length == 2)
                    {
                        result[currentSection][keyValue[0].Trim()] = keyValue[1].Trim();
                    }
                }
            }

            return result;
        }

        private void SwapSectionValues(Dictionary<string, Dictionary<string, string>> iniDict, string section1, string section2, string key)
        {
            if (iniDict[section1].ContainsKey(key) && iniDict[section2].ContainsKey(key))
            {
                var temp = iniDict[section1][key];
                iniDict[section1][key] = iniDict[section2][key];
                iniDict[section2][key] = temp;
            }
        }

        //=====================================================================================================================================//
        //                                                Function to Remove Shortcut                                                          //
        //=====================================================================================================================================//

        private void RemoveShortcut(int index)
        {
            api.Log(API.LogType.Debug, $"YourStart.dll: Attempting to remove shortcut {index}");

            // Load the INI content
            var iniContent = ReadIniFile(iniPath);
            if (iniContent == null)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Failed to read INI file: {iniPath}");
                return;
            }
            // Get and update the TotalShortcuts variable
            if (iniContent.ContainsKey("Variables") && iniContent["Variables"].ContainsKey("TotalShortcuts"))
            {
                if (int.TryParse(iniContent["Variables"]["TotalShortcuts"], out int totalShortcuts))
                {
                    totalShortcuts--; // Decrement the value
                    iniContent["Variables"]["TotalShortcuts"] = totalShortcuts.ToString();
                    api.Log(API.LogType.Debug, $"YourStart.dll: TotalShortcuts decremented to {totalShortcuts}.");
                }
                else
                {
                    api.Log(API.LogType.Error, "YourStart.dll: TotalShortcuts value is not a valid integer.");
                    return;
                }
            }
            else
            {
                api.Log(API.LogType.Error, "YourStart.dll: TotalShortcuts variable not found in INI file.");
                return;
            }

            // Validate required sections
            string shortcutPrefix = $"Shortcut_BackGround_{index}";
            if (!iniContent.ContainsKey(shortcutPrefix))
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Shortcut index {index} does not exist in the INI file.");
                return;
            }

            int lastIndex = iniContent.Keys
                .Where(k => k.StartsWith("Shortcut_BackGround_"))
                .Select(k => int.Parse(k.Split('_')[2]))
                .Max();

            // Collapse values
            for (int i = index; i < lastIndex; i++)
            {
                CopySectionValues(iniContent, i + 1, i);
            }

            // Remove last shortcut
            RemoveShortcutSections(iniContent, lastIndex);

            // Write back to the INI file
            WriteIniFile(iniPath, iniContent);

            // Update Rainmeter
            api.Execute("[!UpdateMeter *][!Redraw]");
            api.Log(API.LogType.Debug, "YourStart.dll: Shortcut removal completed successfully.");
        }
        private void RemoveShortcutSections(Dictionary<string, Dictionary<string, string>> iniContent, int index)
        {
            string[] sections = { "BackGround", "Icon", "String" };
            foreach (var section in sections)
            {
                string sectionKey = $"Shortcut_{section}_{index}";
                iniContent.Remove(sectionKey);
            }
        }
        private void CopySectionValues(Dictionary<string, Dictionary<string, string>> iniContent, int fromIndex, int toIndex)
        {
            string[] sections = { "BackGround", "Icon", "String" };
            foreach (var section in sections)
            {
                string fromSection = $"Shortcut_{section}_{fromIndex}";
                string toSection = $"Shortcut_{section}_{toIndex}";

                if (iniContent.ContainsKey(fromSection))
                {
                    iniContent[toSection] = new Dictionary<string, string>(iniContent[fromSection]);
                }
            }
        }
        //=====================================================================================================================================//
        //                                                Function to ChangePath                                                               //
        //=====================================================================================================================================//

        private void ChangePath(int index)
        {
            api.Log(API.LogType.Debug, $"YourStart.dll: Attempting to change path for Shortcut_BackGround_{index}");

            // Load the INI content
            var iniContent = ReadIniFile(iniPath);
            if (iniContent == null)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Failed to read INI file: {iniPath}");
                return;
            }

            // Validate the Shortcut_BackGround section exists
            string shortcutKey = $"Shortcut_BackGround_{index}";
            if (!iniContent.ContainsKey(shortcutKey))
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Shortcut_BackGround_{index} does not exist in INI file.");
                return;
            }

            // Open a file dialog for the user to select a file
            string selectedPath = ShowFileDialog();
            if (string.IsNullOrEmpty(selectedPath))
            {
                api.Log(API.LogType.Error, "YourStart.dll: No file was selected.");
                return;
            }

            api.Log(API.LogType.Debug, $"YourStart.dll: File selected: {selectedPath}");

            // Update the LeftMouseDownAction in the INI content
            iniContent[shortcutKey]["LeftMouseDownAction"] = $"\"{selectedPath}\"";

            // Write back to the INI file
            WriteIniFile(iniPath, iniContent);

            // Refresh Rainmeter
            api.Execute("[!Refresh]");
            api.Log(API.LogType.Debug, $"YourStart.dll: LeftMouseDownAction updated for {shortcutKey}.");
        }

        private string ShowFileDialog()
        {
            try
            {
                // Use OpenFileDialog to select a file
                using (var openFileDialog = new System.Windows.Forms.OpenFileDialog())
                {
                    openFileDialog.Filter = "All Files (*.*)|*.*";
                    openFileDialog.Title = "Select a File";

                    if (openFileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        return openFileDialog.FileName;
                    }
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Failed to open file dialog: {ex.Message}");
            }

            return null;
        }
        //=====================================================================================================================================//
        //                                               Read and Write Ini File                                                               //
        //=====================================================================================================================================//

        private Dictionary<string, Dictionary<string, string>> ReadIniFile(string filePath)
        {
            var iniContent = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);

            if (!File.Exists(filePath))
            {
                api.Log(API.LogType.Error, $"YourStart.dll: INI file not found: {filePath}");
                return null;
            }

            string currentSection = null;

            foreach (var line in File.ReadAllLines(filePath))
            {
                var trimmedLine = line.Trim();

                if (string.IsNullOrEmpty(trimmedLine) || trimmedLine.StartsWith(";"))
                {
                    // Skip empty lines or comments
                    continue;
                }

                if (trimmedLine.StartsWith("[") && trimmedLine.EndsWith("]"))
                {
                    // New section
                    currentSection = trimmedLine.Substring(1, trimmedLine.Length - 2);
                    if (!iniContent.ContainsKey(currentSection))
                    {
                        iniContent[currentSection] = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    }
                }
                else if (!string.IsNullOrEmpty(currentSection))
                {
                    // Key-value pair
                    var splitIndex = trimmedLine.IndexOf('=');
                    if (splitIndex > 0)
                    {
                        var key = trimmedLine.Substring(0, splitIndex).Trim();
                        var value = trimmedLine.Substring(splitIndex + 1).Trim();
                        iniContent[currentSection][key] = value;
                    }
                }
            }

            return iniContent;
        }

        private void WriteIniFile(string filePath, Dictionary<string, Dictionary<string, string>> iniContent)
        {
            try
            {
                using (var writer = new StreamWriter(filePath))
                {
                    foreach (var section in iniContent)
                    {
                        writer.WriteLine($"[{section.Key}]");
                        foreach (var kvp in section.Value)
                        {
                            writer.WriteLine($"{kvp.Key}={kvp.Value}");
                        }
                        writer.WriteLine(); // Add a blank line after each section
                    }
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Failed to write INI file: {ex.Message}");
            }
        }
        private List<string> GenerateIniContent(Dictionary<string, Dictionary<string, string>> iniDict)
        {
            var result = new List<string>();

            foreach (var section in iniDict)
            {
                result.Add($"[{section.Key}]");
                foreach (var kvp in section.Value)
                {
                    result.Add($"{kvp.Key}={kvp.Value}");
                }
                result.Add(""); // Add a blank line between sections
            }

            return result;
        }
        //=====================================================================================================================================//
        //                                                 Function Of AddShortcut                                                             //
        //=====================================================================================================================================//
        private void AddShortcut()
        {
            try
            {
                string defaultIconPath = api.ReadString("DefaultIconPath", "").Trim();
                if (string.IsNullOrEmpty(defaultIconPath) || !File.Exists(defaultIconPath))
                {
                    api.Log(API.LogType.Error, "YourStart.dll: 'DefaultIconPath' must be specified and exist for folder icons.");
                    return;
                }

                using (var fileDialog = new OpenFileDialog())
                {
                    fileDialog.Filter = "All Files|*.*";
                    fileDialog.CheckFileExists = false;
                    fileDialog.Title = "Select a File or Folder";

                    if (fileDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        string selectedPath = fileDialog.FileName;
                        bool isFolder = Directory.Exists(selectedPath);

                        if (!isFolder && !File.Exists(selectedPath))
                        {
                            api.Log(API.LogType.Error, $"YourStart.dll: Selected file or folder does not exist: {selectedPath}");
                            return;
                        }

                        string name = isFolder
                            ? new DirectoryInfo(selectedPath).Name
                            : Path.GetFileNameWithoutExtension(selectedPath);
                        string shortcutPath = selectedPath; // Use the path directly for files or folders
                        string iconPath = isFolder
                            ? defaultIconPath
                            : Path.Combine(iconFolder, $"{name}_{shortcutPath.GetHashCode()}.ico");

                        // Extract icon for files; use default icon for folders
                        if (!isFolder && !ExtractIconFromExe(shortcutPath, iconPath))
                        {
                            api.Log(API.LogType.Warning, $"YourStart.dll: Using default icon for '{shortcutPath}'.");
                            iconPath = defaultIconPath;
                        }

                        int currentTotalShortcuts = api.ReadInt("TotalShortcuts", 0);
                        int newShortcutIndex = currentTotalShortcuts + 1;
                        int writeStringMeter = api.ReadInt("WriteStringMeter", 0); // Read WriteStringMeter value

                        using (var iniWriter = new StreamWriter(iniPath, true))
                        {
                            iniWriter.WriteLine($"[Shortcut_BackGround_{newShortcutIndex}]");
                            iniWriter.WriteLine("Meter=Shape");
                            iniWriter.WriteLine("MeterStyle=All_Apps_BackGround");
                            iniWriter.WriteLine($"LeftMouseDownAction=[\"{shortcutPath}\"]");
                            iniWriter.WriteLine();

                            iniWriter.WriteLine($"[Shortcut_Icon_{newShortcutIndex}]");
                            iniWriter.WriteLine("Meter=Image");
                            iniWriter.WriteLine($"ImageName={iconPath}");
                            iniWriter.WriteLine("MeterStyle=All_Apps_Icons");
                            iniWriter.WriteLine();

                            // Write the string meter section only if WriteStringMeter is not 0
                            if (writeStringMeter != 0)
                            {
                                iniWriter.WriteLine($"[Shortcut_String_{newShortcutIndex}]");
                                iniWriter.WriteLine("Meter=String");
                                iniWriter.WriteLine($"Text={name}");
                                iniWriter.WriteLine("MeterStyle=All_Apps_Text");
                                iniWriter.WriteLine();
                            }
                        }

                        // Update the TotalShortcuts variable
                        UpdateVariableInIni(iniPath, "TotalShortcuts", newShortcutIndex.ToString());
                        api.Log(API.LogType.Notice, $"YourStart.dll: Shortcut '{name}' (folder: {isFolder}) added successfully.");
                        api.Execute(OnCompleteAction);
                    }
                    else
                    {
                        api.Log(API.LogType.Warning, "YourStart.dll: Shortcut addition was cancelled.");
                    }
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: An error occurred in AddShortcut - {ex.Message}");
            }
        }
        private void UpdateVariableInIni(string iniPath, string variableName, string variableValue)
        {
            var lines = File.ReadAllLines(iniPath);
            bool variableUpdated = false;

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].StartsWith($"{variableName}="))
                {
                    lines[i] = $"{variableName}={variableValue}";
                    variableUpdated = true;
                    break;
                }
            }

            if (!variableUpdated)
            {
                using (var iniWriter = new StreamWriter(iniPath, true))
                {
                    iniWriter.WriteLine($"{variableName}={variableValue}");
                }
            }
            else
            {
                File.WriteAllLines(iniPath, lines);
            }
        }
        //=====================================================================================================================================//
        //                                                 Function Of WriteAllApps                                                            //
        //=====================================================================================================================================//
        private void UpdateAppList()
        {
            appList.Clear();
            string[] startMenuPaths = {
                Environment.GetFolderPath(Environment.SpecialFolder.CommonStartMenu) + "\\Programs",
                Environment.GetFolderPath(Environment.SpecialFolder.StartMenu) + "\\Programs"
            };

            var uniqueApps = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var path in startMenuPaths)
            {
                if (Directory.Exists(path))
                {
                    foreach (var shortcut in Directory.GetFiles(path, "*.lnk", SearchOption.AllDirectories))
                    {
                        try
                        {
                            string exePath = GetShortcutTarget(shortcut);
                            if (!string.IsNullOrEmpty(exePath) && File.Exists(exePath) && exePath.EndsWith(".exe", StringComparison.OrdinalIgnoreCase))
                            {
                                string appName = Path.GetFileNameWithoutExtension(exePath);
                                string uniqueKey = $"{appName.ToLowerInvariant()}_{exePath.ToLowerInvariant()}";

                                if (uniqueApps.Contains(uniqueKey)) continue;

                                string appIconPath = Path.Combine(iconFolder, $"{appName}_{exePath.GetHashCode()}.ico");
                                bool iconExtracted = File.Exists(appIconPath) || ExtractIconFromExe(exePath, appIconPath);

                                if (iconExtracted)
                                {
                                    appList.Add(new AppInfo
                                    {
                                        Name = appName,
                                        Path = exePath,
                                        IconPath = appIconPath
                                    });
                                    uniqueApps.Add(uniqueKey);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            api.Log(API.LogType.Error, $"YourStart.dll: Error processing shortcut '{shortcut}': {ex.Message}");
                        }
                    }
                }
            }
        }
        private void SaveToIniFile(string iniPath)
        {
            try
            {
                var validAppList = appList
                    .Where(app => File.Exists(app.IconPath))
                    .OrderBy(app => app.Name)
                    .ToList();

                var iniEntries = new HashSet<string>(StringComparer.OrdinalIgnoreCase); // Track written entries
                int totalApps = validAppList.Count;

                using (var iniWriter = new StreamWriter(iniPath))
                {
                    // Write [Variables] section
                    iniWriter.WriteLine("[Variables]");
                    iniWriter.WriteLine($"AllApps={totalApps}");
                    iniWriter.WriteLine();

                    char currentGroup = '\0';
                    foreach (var app in validAppList)
                    {
                        char firstChar = char.ToUpper(app.Name[0]);
                        string groupName = char.IsLetter(firstChar) ? firstChar.ToString() : "#";

                        if (currentGroup != groupName[0])
                        {
                            currentGroup = groupName[0];

                            iniWriter.WriteLine($"[AppName_Divider__{groupName}]");
                            iniWriter.WriteLine("Meter=String");
                            iniWriter.WriteLine($"Text={groupName}");
                            iniWriter.WriteLine("MeterStyle=All_Apps_Text");
                            iniWriter.WriteLine();
                        }

                        string meterName = app.Name.Replace(" ", "_").Replace(".", "_");

                        if (!iniEntries.Contains(meterName))
                        {
                            iniWriter.WriteLine($"[AppName_BackGround_{meterName}]");
                            iniWriter.WriteLine("Meter=Shape");
                            iniWriter.WriteLine("MeterStyle=All_Apps_BackGround");
                            iniWriter.WriteLine($"LeftMouseDownAction=[\"{app.Path}\"]");

                            iniWriter.WriteLine($"[AppName_Icon_{meterName}]");
                            iniWriter.WriteLine("Meter=Image");
                            iniWriter.WriteLine($"ImageName={app.IconPath}");
                            iniWriter.WriteLine("MeterStyle=All_Apps_Icons");

                            iniWriter.WriteLine($"[AppName_String_{meterName}]");
                            iniWriter.WriteLine("Meter=String");
                            iniWriter.WriteLine($"Text={app.Name}");
                            iniWriter.WriteLine("MeterStyle=All_Apps_Text");
                            iniWriter.WriteLine();

                            iniEntries.Add(meterName); // Mark as written
                        }
                    }

                }
                api.Execute(OnCompleteAction);
            }

            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Error writing INI file: {ex.Message}");
            }
        }
        private string GetShortcutTarget(string shortcutPath)
        {
            IWshRuntimeLibrary.WshShell shell = new IWshRuntimeLibrary.WshShell();
            IWshRuntimeLibrary.IWshShortcut shortcut = (IWshRuntimeLibrary.IWshShortcut)shell.CreateShortcut(shortcutPath);
            return shortcut.TargetPath;
        }

        private bool ExtractIconFromExe(string exePath, string savePath)
        {
            try
            {
                // Extract the associated icon from the executable
                using (Icon icon = Icon.ExtractAssociatedIcon(exePath))
                {
                    if (icon != null)
                    {
                        // Convert the icon to a bitmap and save it
                        using (Bitmap bitmap = icon.ToBitmap())
                        {
                            bitmap.Save(savePath, System.Drawing.Imaging.ImageFormat.Png);
                        }
                        return true; // Successfully saved the icon
                    }
                    else
                    {
                        api.Log(API.LogType.Warning, $"YourStart.dll: No icon found for '{exePath}'.");
                        return false; // No icon found
                    }
                }
            }
            catch (Exception ex)
            {
                api.Log(API.LogType.Error, $"YourStart.dll: Failed to extract icon from '{exePath}': {ex.Message}");
                return false; // Error during extraction
            }
        }

        [DllImport("shell32.dll", EntryPoint = "ExtractIconEx", CharSet = CharSet.Auto)]
        private static extern int ExtractIconEx(string file, int index, out IntPtr largeIcon, out IntPtr smallIcon, int nIcons);

        [DllImport("user32.dll", EntryPoint = "DestroyIcon")]
        private static extern bool DestroyIcon(IntPtr hIcon);

        internal class AppInfo
        {
            public string Name { get; set; }
            public string Path { get; set; }
            public string IconPath { get; set; }
        }
        //=====================================================================================================================================//

    }



    public static class Plugin
    {
        static IntPtr StringBuffer = IntPtr.Zero;

        [DllExport]
        public static void Initialize(ref IntPtr data, IntPtr rm)
        {
            data = GCHandle.ToIntPtr(GCHandle.Alloc(new Measure()));
        }

        [DllExport]
        public static void Finalize(IntPtr data)
        {
            GCHandle.FromIntPtr(data).Free();

            if (StringBuffer != IntPtr.Zero)
            {
                Marshal.FreeHGlobal(StringBuffer);
                StringBuffer = IntPtr.Zero;
            }
        }

        [DllExport]
        public static void Reload(IntPtr data, IntPtr rm, ref double maxValue)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            measure.Reload(new API(rm), ref maxValue);
        }

        [DllExport]
        public static double Update(IntPtr data)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;
            return measure.Update();
        }



        [DllExport]
        public static void ExecuteBang(IntPtr data, [MarshalAs(UnmanagedType.LPWStr)] string args)
        {
            Measure measure = (Measure)GCHandle.FromIntPtr(data).Target;

            measure.Execute(args);

        }
    }
}
