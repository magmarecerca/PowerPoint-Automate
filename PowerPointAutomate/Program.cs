using PowerAutomation;

namespace PowerPointAutomate;

internal abstract class Program {
	private static void Main(string[] args) {
		if (args.Length == 0) {
			Console.WriteLine("Drag and drop the prizes CSV file to run the app.");
			goto program_end;
		}

		try {
			string csvPath = args[0];
			string rootPath = Path.GetDirectoryName(csvPath) ?? throw new InvalidOperationException();
			string templateFilePath = Path.Combine(rootPath, "TemplatePresentation.pptx");
			string saveFilePath = Path.Combine(rootPath, "ModifiedPresentation.pptx");
			string participantImagesFolderPath = Path.Combine(rootPath, "Participants");
			string logosFolderPath = Path.Combine(rootPath, "Logos");

			Markers markers = new();
			PowerPoint powerPoint = new(templateFilePath, markers);

			Console.WriteLine("Write the prizes template index:");
			string? readTemplateIndex = Console.ReadLine();
			if (readTemplateIndex == null)
				goto program_end;
			int templateIndex = int.Parse(readTemplateIndex);

			Generator generator = new(powerPoint, participantImagesFolderPath, logosFolderPath, templateIndex);
			generator.Generate(csvPath);
			generator.Export(saveFilePath);

			powerPoint.Dispose();
		}
		catch (Exception e) {
			Console.WriteLine(e);
		}

		program_end:
		Console.WriteLine("Program execution completed, press any key to exit...");
		Console.ReadKey();
	}
}