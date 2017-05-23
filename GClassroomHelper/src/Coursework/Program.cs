using Google.Apis.Auth.OAuth2;
using Google.Apis.Classroom.v1;
using Google.Apis.Classroom.v1.Data;
using Google.Apis.Download;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.IO.Compression;


namespace Coursework
{
    class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/classroom.googleapis.com-dotnet-quickstart.json
        static string[] Scopes = {
            ClassroomService.Scope.ClassroomCoursesReadonly,
            ClassroomService.Scope.ClassroomCourseworkStudentsReadonly,
            ClassroomService.Scope.ClassroomCourseworkMeReadonly,
            ClassroomService.Scope.ClassroomProfileEmails,
            ClassroomService.Scope.ClassroomRostersReadonly,
            DriveService.Scope.DriveReadonly,
        };

        static string ApplicationName = "Classroom API .NET Quickstart";

        static ClassroomService classroomService;
        static DriveService driveService;

        static int lineLength = 50;

        static void Main(string[] args)
        {
            Console.OutputEncoding = Encoding.UTF8;

            try
            {
                var task = Task.Run(() => Run());

                task.Wait();
            }
            catch (Exception ex)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Exception: " + ex.Message);
            }

            Console.ResetColor();
            Console.WriteLine(new string('=', lineLength));

            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Tasks completed!");
           
            Console.ReadKey();
        }

        private static async Task Run()
        {
            UserCredential credential;

            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                string credPath = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
                credPath = Path.Combine(credPath, ".credentials/classroom.googleapis.com-dotnet-quickstart.json");

                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    "user",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;

                //Console.WriteLine($"Credential file saved to: {credPath}");
            }

            // Create Classroom API service.
            classroomService = new ClassroomService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            driveService = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });

            Console.ResetColor();
            Console.WriteLine("Requesting courses...");

            var coursesResult = classroomService.Courses.List().Execute();
            var courses = coursesResult.Courses;

            if (coursesResult != null)
            {
                Console.ResetColor();
                Console.WriteLine(new string('=', lineLength));

                for (int i = 0; i < courses.Count; i++)
                {
                    Console.Write($"{i + 1}. {courses[i].Name}");

                    if(!string.IsNullOrEmpty(courses[i].Section))
                    {
                        Console.WriteLine($" ({courses[i].Section})");
                    }
                    else
                    {
                        Console.WriteLine();
                    }
                }

                Console.WriteLine(new string('=', lineLength));
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Courses not found.");

                return;
            }

            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.Write("Course: ");

            var selectedCourseIndex = Select(1, courses.Count);
            var selectedCourse = courses[selectedCourseIndex - 1];

            Console.ResetColor();
            Console.WriteLine(new string('=', lineLength));
            Console.WriteLine("Requesting courseworks...");

            var courseWorksResult = classroomService.Courses.CourseWork.List(selectedCourse.Id).Execute();
            var courseWorks = courseWorksResult.CourseWork.OrderBy(p => p.CreationTime).ToList();

            if (courseWorks != null)
            {
                Console.ResetColor();
                Console.WriteLine(new string('=', lineLength));

                for (int i = 0; i < courseWorks.Count; i++)
                {
                    Console.WriteLine($"{i + 1}. {courseWorks[i].Title}");
                }

                Console.WriteLine(new string('=', lineLength));
            }
            else
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("CourseWorks not found.");

                return;
            }

            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.Write("Coursework: ");

            var selectedCourseWorkIndex = Select(1, courseWorks.Count);
            var selectedCourseWork = courseWorks[selectedCourseWorkIndex - 1];

            Console.ResetColor();
            Console.WriteLine(new string('=', lineLength));

            await DownloadSubmissionsAsync(selectedCourse, selectedCourseWork);
        }

        private static async Task DownloadSubmissionsAsync(Course course, CourseWork courseWork)
        {
            var courseWorkPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Personal),
                "Classroom",
                string.IsNullOrEmpty(course.Section) ? course.Name : $"{course.Name} ({course.Section})",
                courseWork.Title
                );

            RemoveDirectory(courseWorkPath);

            var submissionsResult = classroomService.Courses.CourseWork.StudentSubmissions.List(course.Id, courseWork.Id).Execute();
            var submissions = submissionsResult.StudentSubmissions;


            var tasks = new List<Task>();
            var archives = new List<string>();

            foreach (var submission in submissions)
            {
                if (submission.AssignmentSubmission?.Attachments != null && submission.AssignmentSubmission?.Attachments.Count != 0)
                {
                    var student = classroomService.Courses.Students.Get(course.Id, submission.UserId).Execute();

                    var studentAlias = $"{student.Profile.Name.FullName}__{student.Profile.EmailAddress.Split('@')[0]}";
                    var studentPath = Path.Combine(courseWorkPath, studentAlias);

                    Directory.CreateDirectory(studentPath);

                    

                    foreach (var attachment in submission.AssignmentSubmission.Attachments)
                    {
                        var request = driveService.Files.Get(attachment.DriveFile.Id);

                        var filePath = Path.Combine(studentPath, attachment.DriveFile.Title);

                        var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None);

                        
                        // Add a handler which will be notified on progress changes.
                        // It will notify on each chunk download and when the
                        // download is completed or failed.
                        request.MediaDownloader.ProgressChanged += (IDownloadProgress progress) =>
                        {
                            switch (progress.Status)
                            {
                                case DownloadStatus.Completed:
                                    {
                                        Console.ForegroundColor = ConsoleColor.Green;
                                        Console.WriteLine($"Done: {student.Profile.Name.FullName} => {attachment.DriveFile.Title}");
                                        fileStream.Close();
                                        break;
                                    }
                                case DownloadStatus.Failed:
                                    {
                                        Console.ForegroundColor = ConsoleColor.Red;
                                        Console.WriteLine($"Failed: {student.Profile.Name.FullName} => {attachment.DriveFile.Title}");
                                        fileStream.Close();
                                        break;
                                    }
                            }
                        };

                        tasks.Add(request.DownloadAsync(fileStream));
                        
                        if (attachment.DriveFile.Title.EndsWith(".zip", StringComparison.OrdinalIgnoreCase))
                        {
                            archives.Add(filePath);
                        }
                    }
                }
            }

            await Task.WhenAll(tasks);

            Console.ResetColor();
            Console.WriteLine(new string('=', lineLength));
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("Downloads completed!");

            await DecompressArchives(archives);
        }

        private static async Task DecompressArchives(IEnumerable<string> files)
        {
            Console.ResetColor();
            Console.WriteLine(new string('=', lineLength));

            Console.ForegroundColor = ConsoleColor.Magenta;
            Console.WriteLine("Do you want to decompress .zip archive files?");
            Console.WriteLine("Press any key to continue or close window to finish.");
            Console.ResetColor();
            Console.WriteLine(new string('=', lineLength));

            Console.ReadKey();

            foreach (var file in files)
            {
                var path = Path.Combine(Path.GetDirectoryName(file), Path.GetFileNameWithoutExtension(file));
                Directory.CreateDirectory(path);

                try
                {
                    ZipFile.ExtractToDirectory(file, path);

                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine($"Done: {file}");
                }
                catch (Exception ex)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine($"Failed: {file}");
                    Console.WriteLine($"Exception: {ex.Message}");
                }
            }
        }

        private static void RemoveDirectory(string courseWorkPath)
        {
            if (Directory.Exists(courseWorkPath))
            {
                DirectoryInfo di = new DirectoryInfo(courseWorkPath);

                foreach (FileInfo file in di.GetFiles())
                {
                    file.Delete();
                }

                foreach (DirectoryInfo dir in di.GetDirectories())
                {
                    dir.Delete(true);
                }

                Directory.Delete(courseWorkPath);
            }
        }

        private static int Select(int min, int max)
        {
            int i;
            var succeed = false;

            do
            {
                succeed = int.TryParse(Console.ReadLine(), out i);

                if (!succeed)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Please try again.");
                }

                if (i < min || i > max)
                {
                    succeed = false;

                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Please try again");
                }
            }
            while (!succeed);

            Console.ResetColor();

            return i;
        }
    }
}
