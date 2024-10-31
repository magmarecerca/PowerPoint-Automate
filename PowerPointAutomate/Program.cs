using PowerAutomation;

namespace PowerPointAutomate;

internal abstract class Program {
	private static void Main() {
		const string templateFilePath = @"E:\TemplatePresentation.pptx";
		const string saveFilePath = @"E:\ModifiedPresentation.pptx";

		Markers markers = new();
		PowerPoint powerPoint = new(templateFilePath, markers);

		Slide template = powerPoint.GetSlide(1);
		Slide slide = powerPoint.DuplicateSlide(template);
		slide.SetTitle("Updated Title from Template!");
		slide.SetText("This slide was edited using C# from a template.");

		slide = powerPoint.DuplicateSlideAt(powerPoint.LastSlideNumber + 1, template);
		slide.SetTitle("Copy of first slide with different title.");
		slide.SetText("This slide was edited using C# from a template.");

		powerPoint.SaveAs(saveFilePath);

		powerPoint.Dispose();

		Console.WriteLine("Presentation modified and saved successfully at " + saveFilePath);
	}
}