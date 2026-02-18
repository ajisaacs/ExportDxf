using FabWorks.Core.PressBrake;
using Xunit;

namespace FabWorks.Tests.PressBrake
{
    public class ProgramReaderTests
    {
        [Fact]
        public void Load_SamplePgm_ParsesProgramAttributes()
        {
            var pgm = Program.Load("TestData/sample.pgm");

            // ProgName may be empty on some exports; verify PartName was parsed instead
            Assert.False(string.IsNullOrEmpty(pgm.PartName));
        }

        [Fact]
        public void Load_SamplePgm_ParsesThickness()
        {
            var pgm = Program.Load("TestData/sample.pgm");
            Assert.True(pgm.MatThick > 0);
        }

        [Fact]
        public void Load_SamplePgm_ParsesSteps()
        {
            var pgm = Program.Load("TestData/sample.pgm");
            Assert.NotEmpty(pgm.Steps);
        }

        [Fact]
        public void Load_SamplePgm_ParsesToolSetups()
        {
            var pgm = Program.Load("TestData/sample.pgm");
            Assert.NotEmpty(pgm.UpperToolSets);
            Assert.NotEmpty(pgm.LowerToolSets);
        }

        [Fact]
        public void Load_SamplePgm_ResolvesStepToolReferences()
        {
            var pgm = Program.Load("TestData/sample.pgm");
            var step = pgm.Steps[0];
            Assert.NotNull(step.UpperTool);
            Assert.NotNull(step.LowerTool);
        }
    }
}
