using PowerAutomation;

namespace PowerPointAutomate;

internal abstract class Program {
	private static void Main(string[] args) {
		if (args.Length == 0) {
			Console.WriteLine("Drag and drop the prizes CSV file to run the app.");
			return;
		}

		string csvPath = args[0];
		string rootPath = Path.GetDirectoryName(csvPath) ?? throw new InvalidOperationException();
		string templateFilePath = Path.Combine(rootPath, "TemplatePresentation.pptx");
		string saveFilePath = Path.Combine(rootPath, "ModifiedPresentation.pptx");
		string participantImagesFolderPath = Path.Combine(rootPath, "Participants");
		string logosFolderPath = Path.Combine(rootPath, "Logos");

		Markers markers = new();
		PowerPoint powerPoint = new(templateFilePath, markers);

		Generator generator = new(powerPoint, participantImagesFolderPath, logosFolderPath);
		generator.Generate(csvPath);
		generator.Export(saveFilePath);

		powerPoint.Dispose();
	}
}