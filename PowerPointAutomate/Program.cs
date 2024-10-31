using PowerAutomation;

namespace PowerPointAutomate;

internal abstract class Program {
	private static void Main() {
		const string templateFilePath = @"E:\TemplatePresentation.pptx";
		const string saveFilePath = @"E:\ModifiedPresentation.pptx";

		PowerPoint powerPoint = new(templateFilePath);
		Markers markers = new();
		powerPoint.SetMarkers(markers);

		Slide slide = powerPoint.GetSlide(1);
		slide.SetTitle("Updated Title from Template!");
		slide.SetText("This slide was edited using C# from a template.");

		// Slide newSlide = powerPoint.CreateSlide(2, PpSlideLayout.ppLayoutText);
		// newSlide.SetTitle("New Slide Title");
		// newSlide.SetText("This is content added to a new slide.");

		powerPoint.SaveAs(saveFilePath);

		powerPoint.Dispose();

		Console.WriteLine("Presentation modified and saved successfully at " + saveFilePath);
	}
}