using PowerAutomation;

namespace PowerPointAutomate;

internal abstract class Program {
	private static void Main() {
		const string templateFilePath = @"E:\TemplatePresentation.pptx";
		const string saveFilePath = @"E:\ModifiedPresentation.pptx";
		const string imagePath = @"E:\TemplateImage.jpg";
		const string csvPath = @"E:\Prizes.csv";

		Markers markers = new();
		PowerPoint powerPoint = new(templateFilePath, markers);

		Generator generator = new(powerPoint);
		generator.Generate(csvPath);
		generator.Export(saveFilePath);

		powerPoint.Dispose();
	}
}