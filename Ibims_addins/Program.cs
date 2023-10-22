using System;
using System.IO;

class Program
{
    static void Main(string[] args)
    {
        // Specify the source file path
        string current = Directory.GetCurrentDirectory();
        string sourceFilePath = Path.Combine(current, "IBIMSGen.dll");
        string addinFilePath = Path.Combine(current, "IBIMSGen.addin");
        string userProfileDirectory = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);


        // Set the current directory to the user's profile directory
        Directory.SetCurrentDirectory(userProfileDirectory);

        // Specify the destination file path
        string[] destinationFilePaths = Directory.GetDirectories(@"AppData\\Roaming\\Autodesk\\Revit\\Addins");

        try
        {
            for (int i = 0; i < destinationFilePaths.Length; i++)
            {
                string filePath = Path.Combine(destinationFilePaths[i], "IBIMSGen.dll");
                string addinPath = Path.Combine(destinationFilePaths[i], "IBIMSGen.addin");
                // Copy the file to the destination path
                File.Copy(sourceFilePath, filePath, true);
                File.Copy(addinFilePath, addinPath, true);
                Console.WriteLine("Files copied successfully!");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("Error occurred while copying the file: " + ex.Message);
            Console.WriteLine("Maybe a version of Revit is open. Close the revit before running the app.");
            Console.WriteLine("Or copy the files manually to \nC:\\users\\[your-user-name]\\AppData\\Autodesk\\Revit\\Addins\\[version]");
        }

        // Wait for user input before closing the console window (optional)
        Console.WriteLine("Press any key to exit.\n\n\n\nBy: Omar Elshaf3y | 2023 :)\n    m.me\\o.elshaf3y");
        Console.ReadKey();
    }
}