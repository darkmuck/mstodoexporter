using System.IO.Compression;
using System.Text.RegularExpressions;
using Microsoft.Data.Sqlite;
using Microsoft.Extensions.Configuration;

namespace mstodoexporter;

public class Program
{
    public static void Main(string[] args)
    {
        var configuration = new ConfigurationBuilder()
            .SetBasePath(AppContext.BaseDirectory)
            .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
            .AddCommandLine(args)
            .Build();

        string dbPath = configuration["dbPath"]!;
        string outputDir = configuration["outputDir"]!;
        bool clearOutputDirBeforeExport = configuration.GetValue<bool>("clearOutputDirBeforeExport");
        bool archiveOutput = configuration.GetValue<bool>("archiveOutput");
        bool removeOutputDirAfterArchive = configuration.GetValue<bool>("removeOutputDirAfterArchive");
        bool archiveOutputDirIfExistsBeforeExport = configuration.GetValue<bool>("archiveOutputDirIfExistsBeforeExport");
        bool nonInteractive = configuration.GetValue<bool>("nonInteractive");

        if (string.IsNullOrEmpty(dbPath) || !File.Exists(dbPath))
        {
            Console.WriteLine("Database file not found. Please check the 'dbPath' in appsettings.json or provide it as a command-line argument.");
            return;
        }

        if (string.IsNullOrEmpty(outputDir))
        {
            Console.WriteLine("Output directory not specified. Please check the 'outputDir' in appsettings.json or provide it as a command-line argument.");
            return;
        }

        if (archiveOutputDirIfExistsBeforeExport && Directory.Exists(outputDir))
        {
            Console.WriteLine("Output directory already exists.");
            ArchiveDirectory(outputDir, "exported_tasks_backup");
        }

        if (clearOutputDirBeforeExport && Directory.Exists(outputDir))
        {
            Console.WriteLine("Clearing existing output directory...");
            try
            {
                Directory.Delete(outputDir, true);
            }
            catch (IOException ex)
            {
                Console.WriteLine($"Error clearing output directory: {ex.Message}");
                if (!nonInteractive)
                {
                    Console.WriteLine("Do you want to continue? (y/n)");
                    var response = Console.ReadLine();
                    if (response == null || !response.Equals("y", StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine("Aborting export.");
                        return;
                    }
                }
                else
                {
                    Console.WriteLine("Aborting export due to error in non-interactive mode.");
                    return;
                }
            }
        }

        Directory.CreateDirectory(outputDir);

        var connectionString = new SqliteConnectionStringBuilder
        {
            DataSource = dbPath,
            Mode = SqliteOpenMode.ReadOnly
        }.ToString();

        using (var connection = new SqliteConnection(connectionString))
        {
            connection.Open();

            var taskFolders = GetTaskFolders(connection);
            var tasks = GetTasks(connection);
            var steps = GetSteps(connection);
            var attachments = GetAttachments(connection);

            foreach (var folder in taskFolders)
            {
                string folderPath = Path.Combine(outputDir, SanitizeFileName(folder.Name!));
                Directory.CreateDirectory(folderPath);

                foreach (var task in tasks)
                {
                    if (task.FolderId == folder.Id)
                    {
                        string taskFilePath = Path.Combine(folderPath, SanitizeFileName(task.Subject!) + ".md");
                        using (var writer = new StreamWriter(taskFilePath))
                        {
                            writer.WriteLine($"# {task.Subject}");
                            writer.WriteLine();

                            if (!string.IsNullOrEmpty(task.Body))
                            {
                                writer.WriteLine("## Notes");
                                writer.WriteLine(task.Body);
                                writer.WriteLine();
                            }

                            if (!string.IsNullOrEmpty(task.DueDate))
                            {
                                writer.WriteLine($"**Due:** {task.DueDate}");
                                writer.WriteLine();
                            }

                            if (!string.IsNullOrEmpty(task.ReminderDate))
                            {
                                writer.WriteLine($"**Reminder:** {task.ReminderDate}");
                                writer.WriteLine();
                            }

                            var taskSteps = steps.FindAll(s => s.TaskId == task.Id);
                            if (taskSteps.Count > 0)
                            {
                                writer.WriteLine("## Steps");
                                foreach (var step in taskSteps)
                                {
                                    writer.WriteLine($"- [{(step.Completed ? "x" : " ")}] {step.Subject}");
                                }
                                writer.WriteLine();
                            }

                            var taskAttachments = attachments.FindAll(a => a.TaskId == task.Id);
                            if (taskAttachments.Count > 0)
                            {
                                writer.WriteLine("## Attachments");
                                string assetsPath = Path.Combine(folderPath, "assets");
                                Directory.CreateDirectory(assetsPath);

                                foreach (var attachment in taskAttachments)
                                {
                                    if (string.IsNullOrEmpty(attachment.DisplayName))
                                    {
                                        continue;
                                    }

                                    string attachmentSubfolder = attachment.LocalId!;
                                    string attachmentSourcePath = Path.Combine(Path.GetDirectoryName(dbPath)!, "Attachments", attachmentSubfolder, attachment.DisplayName!);
                                    string attachmentDestPath = Path.Combine(assetsPath, attachment.DisplayName!);
                                    writer.WriteLine($"- [{attachment.DisplayName}](assets/{Uri.EscapeDataString(attachment.DisplayName!)})");

                                    try
                                    {
                                        if (File.Exists(attachmentSourcePath))
                                        {
                                            File.Copy(attachmentSourcePath, attachmentDestPath, true);
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Attachment not found: {attachmentSourcePath}");
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Error copying attachment: {ex.Message}");
                                    }
                                }
                                writer.WriteLine();
                            }
                        }
                    }
                }
            }
        }

        if (archiveOutput)
        {
            ArchiveDirectory(outputDir, "exported_tasks");

            if (removeOutputDirAfterArchive)
            {
                Directory.Delete(outputDir, true);
                Console.WriteLine("Removed output directory after archiving.");
            }
        }

        Console.WriteLine("Export complete!");
    }

    private static void ArchiveDirectory(string directoryPath, string archiveName)
    {
        string timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
        string zipPath = Path.Combine(Directory.GetParent(directoryPath)!.FullName, $"{archiveName}_{timestamp}.zip");
        if (File.Exists(zipPath))
        {
            File.Delete(zipPath);
        }
        ZipFile.CreateFromDirectory(directoryPath, zipPath);
        Console.WriteLine($"Directory '{directoryPath}' archived to '{zipPath}'");
    }

    private static List<TaskFolder> GetTaskFolders(SqliteConnection connection)
    {
        var folders = new List<TaskFolder>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT local_id, name FROM task_folders WHERE deleted = 0";
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                folders.Add(new TaskFolder
                {
                    Id = reader.GetString(0),
                    Name = reader.GetString(1)
                });

            }
        }
        return folders;
    }

    private static List<Task> GetTasks(SqliteConnection connection)
    {
        var tasks = new List<Task>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT local_id, task_folder_local_id, subject, body_content, due_date, reminder_datetime FROM tasks WHERE deleted = 0";
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                tasks.Add(new Task
                {
                    Id = reader.GetString(0),
                    FolderId = reader.GetString(1),
                    Subject = reader.GetString(2),
                    Body = reader.IsDBNull(3) ? null : reader.GetString(3),
                    DueDate = reader.IsDBNull(4) ? null : reader.GetString(4),
                    ReminderDate = reader.IsDBNull(5) ? null : reader.GetString(5)
                });

            }
        }
        return tasks;
    }

    private static List<Step> GetSteps(SqliteConnection connection)
    {
        var steps = new List<Step>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT task_local_id, subject, completed FROM steps WHERE deleted = 0";
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                steps.Add(new Step
                {
                    TaskId = reader.GetString(0),
                    Subject = reader.GetString(1),
                    Completed = reader.GetBoolean(2)
                });

            }
        }
        return steps;
    }

    private static List<Attachment> GetAttachments(SqliteConnection connection)
    {
        var attachments = new List<Attachment>();
        using (var command = connection.CreateCommand())
        {
            command.CommandText = "SELECT task_local_id, display_name, web_link, local_id FROM linked_entities WHERE deleted = 0";
            using var reader = command.ExecuteReader();
            while (reader.Read())
            {
                attachments.Add(new Attachment
                {
                    TaskId = reader.GetString(0),
                    DisplayName = reader.GetString(1),
                    WebLink = reader.IsDBNull(2) ? null : reader.GetString(2),
                    LocalId = reader.GetString(3)
                });

            }
        }
        return attachments;
    }

    private static string SanitizeFileName(string fileName)
    {
        string invalidChars = new string(Path.GetInvalidFileNameChars()) + new string(Path.GetInvalidPathChars());
        string regex = string.Format("[{0}]", Regex.Escape(invalidChars));
        return Regex.Replace(fileName, regex, "_");
    }
}

public class TaskFolder
{
    public string? Id { get; set; }
    public string? Name { get; set; }
}

public class Task
{
    public string? Id { get; set; }
    public string? FolderId { get; set; }
    public string? Subject { get; set; }
    public string? Body { get; set; }
    public string? DueDate { get; set; }
    public string? ReminderDate { get; set; }
}

public class Step
{
    public string? TaskId { get; set; }
    public string? Subject { get; set; }
    public bool Completed { get; set; }
}

public class Attachment
{
    public string? TaskId { get; set; }
    public string? DisplayName { get; set; }
    public string? WebLink { get; set; }
    public string? LocalId { get; set; }
}