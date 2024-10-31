﻿using PowerAutomation;

namespace PowerPointAutomate;

internal abstract class Program {
	private static void Main() {
		const string templateFilePath = @"E:\TemplatePresentation.pptx";
		const string saveFilePath = @"E:\ModifiedPresentation.pptx";
		const string imagePath = @"E:\TemplateImage.jpg";

		Markers markers = new();
		PowerPoint powerPoint = new(templateFilePath, markers);

		Slide titleTemplate = powerPoint.GetSlide(1);
		Slide imageTemplate = powerPoint.GetSlide(3);

		Slide slide = powerPoint.DuplicateSlide(titleTemplate);
		slide.SetTitle("Updated Title from Template!");
		slide.SetText("This slide was edited using C# from a template.");

		slide = powerPoint.DuplicateSlideAt(powerPoint.LastSlideNumber + 1, imageTemplate);
		slide.SetTitle("Copy of first slide with different title.");
		slide.SetImage(imagePath);

		powerPoint.RemoveTemplateSlides();
		powerPoint.SaveAs(saveFilePath);

		powerPoint.Dispose();

		Console.WriteLine("Presentation modified and saved successfully at " + saveFilePath);
	}
}