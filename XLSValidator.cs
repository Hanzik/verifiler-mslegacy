using System;
using System.IO;
using NPOI.HSSF.UserModel;
using VerifilerCore;

namespace VerifilerMSLegacy {

	/// <summary>
	/// This validation step is using the NPOI library which can work
	/// both with .xls and .xlsx formats.
	/// 
	/// The error code produced by this validation is Error.Corrupted.
	/// </summary>
	public class XLSValidator : FormatSpecificValidator {

		public override int ErrorCode { get; set; } = Error.Corrupted;

		public override void Setup() {
			Name = "Microsoft Excel .xls files Verification";
			RelevantExtensions.Add(".xls");
			Enable();
		}

		public override void ValidateFile(string file) {

			FileStream stream = null;
			try {
				stream = File.Open(file, FileMode.Open, FileAccess.Read);
				HSSFWorkbook workbook = new HSSFWorkbook(stream);
			} catch (Exception e) {
				ReportAsError("File is corrupted: " + file + "; Message: " + e.Message);
				GC.Collect();
			} finally {
				stream?.Close();
			}
		}
	}
}