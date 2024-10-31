using Microsoft.Office.Interop.PowerPoint;
using PowerAutomation;

namespace PowerPointAutomate;

internal abstract class Program {
	private static void Main() {
		const string templateFilePath = @"E:\TemplatePresentation.pptx";
		const string saveFilePath = @"E:\ModifiedPresentation.pptx";

		PowerPoint powerPoint = new(templateFilePath);

		Slide slide = powerPoint.Slides[1];

		if (slide.Shapes.Count >= 2) {
			slide.Shapes[1].TextFrame.TextRange.Text = "Updated Title from Template!";
			slide.Shapes[2].TextFrame.TextRange.Text = "This slide was edited using C# from a template.";
		}

		Slide newSlide = powerPoint.Slides.Add(2, PpSlideLayout.ppLayoutText);
		newSlide.Shapes[1].TextFrame.TextRange.Text = "New Slide Title";
		newSlide.Shapes[2].TextFrame.TextRange.Text = "This is content added to a new slide.";

		powerPoint.SaveAs(saveFilePath);

		powerPoint.Dispose();

		Console.WriteLine("Presentation modified and saved successfully at " + saveFilePath);
	}
}