using PowerAutomation;

namespace PowerPointAutomate;

internal abstract class Program {
	private static void Main() {
		const string templateFilePath = @"E:\TemplatePresentation.pptx";
		const string saveFilePath = @"E:\ModifiedPresentation.pptx";

		Markers markers = new();
		PowerPoint powerPoint = new(templateFilePath, markers);

		Slide slide = powerPoint.GetSlide(1);
		powerPoint.DuplicateSlideAt(3, slide);
		slide.SetTitle("Updated Title from Template!");
		slide.SetText("This slide was edited using C# from a template.");

		slide = powerPoint.GetSlide(3);
		slide.SetTitle("Copy of first slide with different title.");
		slide.SetText("This slide was edited using C# from a template.");

		powerPoint.SaveAs(saveFilePath);

		powerPoint.Dispose();

		Console.WriteLine("Presentation modified and saved successfully at " + saveFilePath);
	}
}