using System;
using System.IO;
using System.IO.Packaging;
using System.Linq;

namespace Crosse.PowerShell.PackageManagement {
    public class PackageFile {

# region Properties that mirror Package.PackageProperties
        public  string      Category        { get; set; }
        public  string      ContentStatus   { get; set; }
        public  string      ContentType     { get; set; }
        public  DateTime?   Created         { get; set; }
        public  string      Creator         { get; set; }
        public  string      Description     { get; set; }
        public  string      Identifier      { get; set; }
        public  string      Keywords        { get; set; }
        public  string      Language        { get; set; }
        public  string      LastModifiedBy  { get; set; }
        public  DateTime?   LastPrinted     { get; set; }
        public  DateTime?   Modified        { get; set; }
        public  string      Revision        { get; set; }
        public  string      Subject         { get; set; }
        public  string      Title           { get; set; }
        public  string      Version         { get; set; }
#endregion

        public  string      FileName        { get; internal set; }
        public  int         ItemCount       { get; internal set; }

        public PackageFile(string fileName, FileMode mode) {
            FileName = fileName;
            using (Package package = Package.Open(fileName, mode)) {
                Category        = package.PackageProperties.Category;
                ContentStatus   = package.PackageProperties.ContentStatus;
                ContentType     = package.PackageProperties.ContentType;
                Created         = package.PackageProperties.Created;
                Creator         = package.PackageProperties.Creator;
                Description     = package.PackageProperties.Description;
                Identifier      = package.PackageProperties.Identifier;
                Keywords        = package.PackageProperties.Keywords;
                Language        = package.PackageProperties.Language;
                LastModifiedBy  = package.PackageProperties.LastModifiedBy;
                LastPrinted     = package.PackageProperties.LastPrinted;
                Modified        = package.PackageProperties.Modified;
                Revision        = package.PackageProperties.Revision;
                Subject         = package.PackageProperties.Subject;
                Title           = package.PackageProperties.Title;
                Version         = package.PackageProperties.Version;

                ItemCount = package.GetParts().Count();
            }
        }

        public void Flush() {
            using (Package package = Package.Open(FileName, FileMode.Open)) {
                package.PackageProperties.Category          = Category;
                package.PackageProperties.ContentStatus     = ContentStatus;
                package.PackageProperties.ContentType       = ContentType;
                package.PackageProperties.Created           = Created;
                package.PackageProperties.Creator           = Creator;
                package.PackageProperties.Description       = Description;
                package.PackageProperties.Identifier        = Identifier;
                package.PackageProperties.Keywords          = Keywords;
                package.PackageProperties.Language          = Language;
                package.PackageProperties.LastModifiedBy    = LastModifiedBy;
                package.PackageProperties.LastPrinted       = LastPrinted;
                package.PackageProperties.Modified          = Modified;
                package.PackageProperties.Revision          = Revision;
                package.PackageProperties.Subject           = Subject;
                package.PackageProperties.Title             = Title;
                package.PackageProperties.Version           = Version;

                ItemCount = package.GetParts().Count();
            }
        }
    }
}
