using FabWorks.Api.Services;
using Xunit;

namespace FabWorks.Tests
{
    public class FormProgramServiceTests
    {
        [Fact]
        public void ParseFromFile_SamplePgm_PopulatesMaterialType()
        {
            var service = new FormProgramService();
            var fp = service.ParseFromFile("TestData/sample.pgm");

            // ProgName is empty in the sample file, so verify MaterialType instead
            Assert.False(string.IsNullOrEmpty(fp.MaterialType));
        }

        [Fact]
        public void ParseFromFile_SamplePgm_PopulatesThickness()
        {
            var service = new FormProgramService();
            var fp = service.ParseFromFile("TestData/sample.pgm");
            Assert.NotNull(fp.Thickness);
            Assert.True(fp.Thickness > 0);
        }

        [Fact]
        public void ParseFromFile_SamplePgm_PopulatesBendCount()
        {
            var service = new FormProgramService();
            var fp = service.ParseFromFile("TestData/sample.pgm");
            Assert.True(fp.BendCount > 0);
        }

        [Fact]
        public void ParseFromFile_SamplePgm_PopulatesToolNames()
        {
            var service = new FormProgramService();
            var fp = service.ParseFromFile("TestData/sample.pgm");
            Assert.False(string.IsNullOrEmpty(fp.UpperToolNames));
            Assert.False(string.IsNullOrEmpty(fp.LowerToolNames));
        }

        [Fact]
        public void ParseFromFile_SamplePgm_ComputesContentHash()
        {
            var service = new FormProgramService();
            var fp = service.ParseFromFile("TestData/sample.pgm");
            Assert.NotNull(fp.ContentHash);
            Assert.Equal(64, fp.ContentHash.Length); // SHA256 hex = 64 chars
        }
    }
}
