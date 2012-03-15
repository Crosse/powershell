using System;
using System.IO;
using System.IO.Packaging;

namespace Crosse.PowerShell.PackageManagement {
    public class PackageFile : IDisposable {

# region PackageProperties
        public string Category {
            get { return package.PackageProperties.Category; }
            set { package.PackageProperties.Category = value; }
        }


        public string ContentStatus {
            get { return package.PackageProperties.ContentStatus; }
            set { package.PackageProperties.ContentStatus = value; }
        }


        public string ContentType {
            get { return package.PackageProperties.ContentType; }
            set { package.PackageProperties.ContentType = value; }
        }


        public DateTime? Created {
            get { return package.PackageProperties.Created; }
            set { package.PackageProperties.Created = value; }
        }


        public string Creator {
            get { return package.PackageProperties.Creator; }
            set { package.PackageProperties.Creator = value; }
        }


        public string Description {
            get { return package.PackageProperties.Description; }
            set { package.PackageProperties.Description = value; }
        }


        public string Identifier {
            get { return package.PackageProperties.Identifier; }
            set { package.PackageProperties.Identifier = value; }
        }


        public string Keywords {
            get { return package.PackageProperties.Keywords; }
            set { package.PackageProperties.Keywords = value; }
        }


        public string Language {
            get { return package.PackageProperties.Language; }
            set { package.PackageProperties.Language = value; }
        }


        public string LastModifiedBy {
            get { return package.PackageProperties.LastModifiedBy; }
            set { package.PackageProperties.LastModifiedBy = value; }
        }


        public DateTime? LastPrinted {
            get { return package.PackageProperties.LastPrinted; }
            set { package.PackageProperties.LastPrinted = value; }
        }


        public DateTime? Modified {
            get { return package.PackageProperties.Modified; }
            set { package.PackageProperties.Modified = value; }
        }


        public string Revision {
            get { return package.PackageProperties.Revision; }
            set { package.PackageProperties.Revision = value; }
        }


        public string Subject {
            get { return package.PackageProperties.Subject; }
            set { package.PackageProperties.Subject = value; }
        }


        public string Title {
            get { return package.PackageProperties.Title; }
            set { package.PackageProperties.Title = value; }
        }


        public string Version {
            get { return package.PackageProperties.Version; }
            set { package.PackageProperties.Version = value; }
        }
#endregion

        public string FileName { get; internal set; }
        private Package package;

        public PackageFile(string fileName, FileMode mode) {
            package = Package.Open(fileName, mode);
            FileName = fileName;
        }

        public void Close() {
            if (package != null) {
                package.Close();
            }
            package = null;
            FileName = null;
        }

        public void Dispose() {
            Close();
        }
    }
}
