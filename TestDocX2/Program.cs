using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xceed.Words.NET;

namespace TestDocX2
{
    class Program
    {
        static void Main(string[] args)
        {
			ProcessDocuments(@"C:\GS01");
		}

		static public ConversionResult ConvertWorkInstructionsFromWordDoc(Attachment file)
		{
			//Work instructions are a documents that contain a set of questions/instructions.
			//Instructions are usually grouped under one or more sections / headings
			//It also contains other tables that can be ignored for this excercise (Safety guidelines, Tools, Procedures, Feedback sections, ect)
			//Write an engine that will extract work instructions from attached file examples
			//The documents vary and your engine must be flexible enough to handle all examples
			//Note, Work Instructions may span multiple tables
			//Validate the document and assign a score to rate the conversion success/confidence. Log issues found with the document (RuleViolations)
			//These documents may contain images related to an instruction. Extra points for extracting them successfully 
			//Do not use Microsoft.Office.Interop.Word: https://support.microsoft.com/en-za/help/257757/considerations-for-server-side-automation-of-office
			//There are many 3rd party libraries available, you are free to use any of them.  Allow for older formats such as 2007 .DOC 
			//One library example: http://cathalscorner.blogspot.com/2010/06/cathal-why-did-you-create-docx.html

			var conversionResult = new ConversionResult();
			conversionResult.Filename = file.FileName;

			try
			{ 
				using (var document = DocX.Load(new MemoryStream(file.FileBytes)))
				{

					// Query to determine all tables that is work instructions
					IEnumerable < Xceed.Document.NET.Table > tableQuery =
						from table in document.Tables
						from row in table.Rows
						from cell in row.Cells
						from Paragraphs in cell.Paragraphs
						where Paragraphs.Text.ToUpper().Contains("WORK INSTRUCTION")
						select table;

					// Execute the query.
					if (tableQuery.Count() == 0)
					{
						conversionResult.AddRuleViolation("NotFound", false, $"No working instructions found in a table in the file");
					}
					else
					{
						var workinstruction = new WorkInstruction();
						workinstruction.SourceFilename = file.FileName;
						workinstruction.InstructionsAsText = new List<WorkInstructionTextItem>();

						foreach (var t in tableQuery)
						{

							//get all valid rows in table 
							var	rowQuery = (from row in t.Rows
										where (row.Cells[0].Paragraphs[0].Text.Length > 0) && !(row.Cells[0].Paragraphs[0].Text.ToUpper().Contains("WORK INSTRUCTION")) 
											select  new {row, Text = row.Cells[0].Paragraphs[0].Text, isListItem = row.Cells[0].Paragraphs[0].IsListItem }
										).Distinct();

							string groupName = "";

							foreach (var r in rowQuery)

							{
							
								if (!r.isListItem) 
								{ groupName = r.Text; }
								else
								{
									workinstruction.InstructionsAsText.Add(new WorkInstructionTextItem() { WorkInstructionItemId = Guid.NewGuid(), Text = r.Text, GroupName = groupName });
									//Console.WriteLine($"Source:{file.FileName} - Text {r.Text}");

								}
							}

							conversionResult.WorkInstructions = workinstruction;

							//Conversion Score 
							var countRows = (from table in document.Tables
											 from rows in table.Rows
											 select rows).Count();
							conversionResult.ConversionScore = (int) (Convert.ToDouble(workinstruction.InstructionsAsText.Count) / Convert.ToDouble(countRows) * 100.00);

						}
					}
				}
			}
			catch (Exception ex)
			{ conversionResult.AddRuleViolation("RunTimeError", true, ex.Message); }



			return conversionResult;

			
		}

		static void ProcessDocuments(string path)
		{
			var docs = new List<ConversionResult>();
			var exceptions = new List<string>();

			var fileEntries = Directory.EnumerateFiles(path, "*.*", SearchOption.TopDirectoryOnly).Where(s => s.EndsWith(".doc") || s.EndsWith(".docx"));

			foreach (string filePath in fileEntries)
			{
				try
				{
					var attachment = new Attachment { FileBytes = File.ReadAllBytes(filePath), FileName = Path.GetFileName(filePath) };

					var docResult = ConvertWorkInstructionsFromWordDoc(attachment);
					docs.Add(docResult);

					var resultTypeName = "Success";

					if (docResult.Aborted)
					{
						resultTypeName = "Aborted";
					}
					else if (docResult.HasRuleViolations)
					{
						resultTypeName = "SuccessWithWarnings";
					}

					var outputFilePath = Path.Combine(path, resultTypeName, attachment.FileName);

					System.IO.Directory.CreateDirectory(Path.GetDirectoryName(outputFilePath));

					File.WriteAllBytes(outputFilePath, attachment.FileBytes);

					File.WriteAllText(outputFilePath + ".result.json", JsonConvert.SerializeObject(docResult, Newtonsoft.Json.Formatting.Indented));
				}

				catch (Exception ex)
				{
					var msg = $"Error processing filename '{filePath}': {ex.Message}";
					exceptions.Add(msg);
				}
			}

			Console.WriteLine( $"Processed {docs.Count} docs. {docs.Count(d => !d.Aborted)} successful, {docs.Count(d => d.Aborted)} failed. {docs.Count(d => d.HasRuleViolations)} had warnings");
			Console.ReadLine();
		}

		#region Models

		public class ConversionResult
		{
			public string Filename;

			public int ConversionScore;

			private List<RuleViolation> _ruleViolations = new List<RuleViolation>();

			public bool Aborted => _ruleViolations.Any((RuleViolation t) => t.IsCritical);

			public bool HasRuleViolations => _ruleViolations.Count > 0;

			public ReadOnlyCollection<RuleViolation> RuleViolations => new ReadOnlyCollection<RuleViolation>(_ruleViolations);

			public WorkInstruction WorkInstructions { get; set; } = new WorkInstruction();

			public void AddRuleViolation(string ruleName, bool abort, string warningText)
			{
				_ruleViolations.Add(new RuleViolation
				{
					Message = warningText,
					IsCritical = abort,
					Rule = ruleName
				});
			}
		}

		public class WorkInstruction
		{
			public List<WorkInstructionTextItem> InstructionsAsText { get; set; }
			public string SourceFilename { get; set; }
		}

		public class RuleViolation
		{
			public string Rule { get; set; }
			public string Message { get; set; }
			public bool IsCritical { get; set; }
		}

		public class WorkInstructionTextItem
		{
			public Guid WorkInstructionItemId { get; set; }

			/// <summary>
			/// Items are grouped by <see cref="P:dCode.Models.PlantMaintenance.WorkInstructionTextItem.GroupName" /> when displayed as a Survey
			/// </summary>
			public string GroupName { get; set; }

			public string Text { get; set; }

			public string SubText { get; set; }
		}

		public class Attachment
		{
			public string Title { get; set; }

			public byte[] FileBytes { get; set; }

			public string FileName { get; set; }

			public string MimeType { get; set; }

			public string Path { get; set; }

			public string ExternalId { get; set; }

			public long Size { get; set; }

			public DateTime CreatedDate { get; set; }
		}

		#endregion Models
	}
}
