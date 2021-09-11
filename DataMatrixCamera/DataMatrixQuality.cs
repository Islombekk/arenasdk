using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace DataMatrix
{
    class DataMatrixQuality
    {
        private string fileName;
        private string decode;
        private string unusedErrorCorrection;
        private string symbolContrast;
        private string axialNonuniformity;
        private string gridNonuniformity;
        private string modulation;
        private string reflectanceMargin;
        private string fixedPatternDamage;
        private string additionalGradesKeys;
        private string quietZone;
        private string distortionAngle;
        private string scanGrade;

        public DataMatrixQuality()
        {
        }

        public DataMatrixQuality(string fileName, string decode, string unusedErrorCorrection, string symbolContrast, string axialNonuniformity, string gridNonuniformity, string modulation, string reflectanceMargin, string fixedPatternDamage, string additionalGradesKeys, string quietZone, string distortionAngle, string scanGrade)
        {
            this.fileName = fileName;
            this.decode = decode;
            this.unusedErrorCorrection = unusedErrorCorrection;
            this.symbolContrast = symbolContrast;
            this.axialNonuniformity = axialNonuniformity;
            this.gridNonuniformity = gridNonuniformity;
            this.modulation = modulation;
            this.reflectanceMargin = reflectanceMargin;
            this.FixedPatternDamage1 = fixedPatternDamage;
            this.additionalGradesKeys = additionalGradesKeys;
            this.quietZone = quietZone;
            this.distortionAngle = distortionAngle;
            this.scanGrade = scanGrade;
        }


        public string FileName { get => fileName; set => fileName = value; }
        public string Decode { get => decode; set => decode = value; }
        public string UnusedErrorCorrection { get => unusedErrorCorrection; set => unusedErrorCorrection = value; }
        public string SymbolContrast { get => symbolContrast; set => symbolContrast = value; }
        public string AxialNonuniformity { get => axialNonuniformity; set => axialNonuniformity = value; }
        public string GridNonuniformity { get => gridNonuniformity; set => gridNonuniformity = value; }
        public string Modulation { get => modulation; set => modulation = value; }
        public string ReflectanceMargin { get => reflectanceMargin; set => reflectanceMargin = value; }
        public string FixedPatternDamage { get => FixedPatternDamage1; set => FixedPatternDamage1 = value; }
        public string AdditionalGradesKeys { get => additionalGradesKeys; set => additionalGradesKeys = value; }
        public string QuietZone { get => quietZone; set => quietZone = value; }
        public string DistortionAngle { get => distortionAngle; set => distortionAngle = value; }
        public string ScanGrade { get => scanGrade; set => scanGrade = value; }
        public string FixedPatternDamage1 { get => fixedPatternDamage; set => fixedPatternDamage = value; }
    }
}
