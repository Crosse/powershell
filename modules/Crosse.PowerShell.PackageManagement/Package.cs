using System;
using System.Globalization;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Collections.Generic;

namespace Crosse.PowerShell.PackageManagement {
    public class PackageFile {

# region Properties that mirror Package.PackageProperties
        public  string      Category        { get; set; }
        public  string      ContentStatus   { get; set; }
        public  string      ContentType     { get; set; }
        public  DateTime?   Created         { get; set; }
        public  string      Creator         { get; set; }
        public  string      Description     { get; set; }
        public  Guid        Identifier      { get; set; }
        public  List<string>Keywords        { get; set; }
        public  CultureInfo Language        { get; set; }
        public  string      LastModifiedBy  { get; set; }
        public  DateTime?   LastPrinted     { get; set; }
        public  DateTime?   Modified        { get; set; }
        public  int         Revision        { get; set; }
        public  string      Subject         { get; set; }
        public  string      Title           { get; set; }
        public  Version     Version         { get; set; }
#endregion

        public  string      FileName        { get; internal set; }
        public  int         ItemCount       { get; internal set; }

        public PackageFile(string fileName, FileMode mode) {
            FileName = fileName;
            using (Package package = Package.Open(fileName, mode)) {
                ItemCount       = CountParts(package);
                Category        = package.PackageProperties.Category;
                ContentStatus   = package.PackageProperties.ContentStatus;
                ContentType     = package.PackageProperties.ContentType;
                Created         = package.PackageProperties.Created;
                Creator         = package.PackageProperties.Creator;
                Description     = package.PackageProperties.Description;
                LastModifiedBy  = package.PackageProperties.LastModifiedBy;
                LastPrinted     = package.PackageProperties.LastPrinted;
                Modified        = package.PackageProperties.Modified;
                Revision        = package.PackageProperties.Revision;
                Subject         = package.PackageProperties.Subject;
                Title           = package.PackageProperties.Title;

                if (String.IsNullOrEmpty(package.PackageProperties.Keywords)) {
                    Keywords = new List<string>();
                } else {
                    Keywords.AddRange(package.PackageProperties.Keywords.Split(','));
                }

                try {
                    Version  = new Version(package.PackageProperties.Version);
                } catch (Exception) {
                    this.Version = new Version();
                }

                try {
                    Language = new CultureInfo(package.PackageProperties.Language);
                } catch (ArgumentException) {
                    Language = null;
                }

                try {
                    Identifier = new Guid(package.PackageProperties.Identifier);
                } catch (Exception) {
                    Identifier = Guid.Empty;
                }
            }
        }

        public List<PackageItem> GetPackageItems() {
            List<PackageItem> parts = new List<PackageItem>();
            using (Package package = Package.Open(FileName, FileMode.Open)) {
                foreach (PackagePart part in package.GetParts()) {
                    if (part.Uri.OriginalString.StartsWith("/package/services/metadata/core-properties") ||
                            part.Uri.OriginalString.StartsWith("/_rels"))
                        continue;
                    else
                        parts.Add(new PackageItem(part));
                }
            }
            return parts;
        }

        public void Flush() {
            Keywords.Sort();
            using (Package package = Package.Open(FileName, FileMode.Open)) {
                ItemCount                                   = CountParts(package);
                package.PackageProperties.Category          = Category;
                package.PackageProperties.ContentStatus     = ContentStatus;
                package.PackageProperties.ContentType       = ContentType;
                package.PackageProperties.Created           = Created;
                package.PackageProperties.Creator           = Creator;
                package.PackageProperties.Description       = Description;
                package.PackageProperties.Identifier        = Identifier.ToString();
                package.PackageProperties.Keywords          = String.Join(",", Keywords.ToArray());
                package.PackageProperties.Language          = Language.ToString();
                package.PackageProperties.LastModifiedBy    = LastModifiedBy;
                package.PackageProperties.LastPrinted       = LastPrinted;
                package.PackageProperties.Modified          = Modified;
                package.PackageProperties.Revision          = Revision;
                package.PackageProperties.Subject           = Subject;
                package.PackageProperties.Title             = Title;
                package.PackageProperties.Version           = Version.ToString();
            }
        }

        private int CountParts(Package package) {
            int items = 0;
            foreach (ZipPackagePart part in package.GetParts()) {
                if (part.Uri.OriginalString.StartsWith("/package/services/metadata/core-properties") ||
                        part.Uri.OriginalString.StartsWith("/_rels"))
                    continue;
                else
                    items++;
            }
            return items;
        }
    }

    public class PackageItem {
        public Uri Uri { get; internal set; }
        public long UncompressedLength { get; internal set; }
        public CompressionOption CompressionOption { get; internal set; }

        internal PackageItem(PackagePart part) {
            Uri = part.Uri;
            CompressionOption = part.CompressionOption;
            using (Stream stream = part.GetStream(FileMode.Open, FileAccess.Read)) {
                UncompressedLength = stream.Length;
            }
        }
    }
}
