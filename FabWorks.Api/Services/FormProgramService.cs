using FabWorks.Core.Models;
using FabWorks.Core.PressBrake;
using System.Security.Cryptography;

namespace FabWorks.Api.Services
{
    public class FormProgramService
    {
        public FormProgram ParseFromFile(string filePath)
        {
            var pgm = FabWorks.Core.PressBrake.Program.Load(filePath);
            var hash = ComputeFileHash(filePath);

            return new FormProgram
            {
                ProgramFilePath = filePath,
                ContentHash = hash,
                ProgramName = pgm.ProgName ?? "",
                Thickness = pgm.MatThick > 0 ? pgm.MatThick : null,
                MaterialType = pgm.MatType.ToString(),
                KFactor = pgm.KFactor > 0 ? pgm.KFactor : null,
                BendCount = pgm.Steps.Count,
                UpperToolNames = string.Join(", ", pgm.UpperToolSets
                    .Select(t => t.Name).Where(n => !string.IsNullOrEmpty(n)).Distinct()),
                LowerToolNames = string.Join(", ", pgm.LowerToolSets
                    .Select(t => t.Name).Where(n => !string.IsNullOrEmpty(n)).Distinct()),
                SetupNotes = pgm.SetupNotes ?? ""
            };
        }

        private static string ComputeFileHash(string filePath)
        {
            using var sha = SHA256.Create();
            using var stream = File.OpenRead(filePath);
            var bytes = sha.ComputeHash(stream);
            return BitConverter.ToString(bytes).Replace("-", "").ToLowerInvariant();
        }
    }
}
