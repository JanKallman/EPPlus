using System;
using System.Drawing;
using System.IO;
using System.Reflection;

namespace EPPlusTest
{
    internal static class ResourceExtensions
    {
        /// <returns>The name of the resource.</returns>
        public static string GetResourceName(this EmbeddedResources resource)
        {
            var baseName = new Func<string>(() =>
            {
                switch (resource)
                {
                    case EmbeddedResources.BitmapImage:
                        return "BitmapImage.gif";
                    case EmbeddedResources.Test1:
                        return "Test1.jpg";
                    case EmbeddedResources.VectorDrawing:
                        return "Vector Drawing.wmf";
                    case EmbeddedResources.VectorDrawing2:
                        return "Vector Drawing2.wmf";
                    default:
                        throw new NotImplementedException($"Resource enum ${resource} not mapped to a file.");
                }
            });
            return $"EPPlusTest.Resources.{baseName()}";
        }

        /// <returns>The resource contents as a stream.</returns>
        public static Stream GetEmbeddedResource(this EmbeddedResources resource)
            => Assembly.GetExecutingAssembly().GetManifestResourceStream(GetResourceName(resource));

        /// <returns>The resource contents as an image.</returns>
        public static Image GetEmbeddedResourceAsImage(this EmbeddedResources resource)
            => new Bitmap(GetEmbeddedResource(resource));

        /// <summary>Copies the contents of the resource to a file and returns the path.</summary>
        /// <returns>The path the resource has been copied to.</returns>
        public static string GetEmbeddedResourceAsTempFile(this EmbeddedResources resource)
        {
            var fileName = Path.Combine(Scaffolding.WorksheetPath, $"{Guid.NewGuid()}-{resource.GetResourceName()}");
            using (var stream = File.OpenWrite(fileName))
            using (var resourceStream = GetEmbeddedResource(resource))
            {
                resourceStream.CopyTo(stream);
            }

            return fileName;
        }
    }
}