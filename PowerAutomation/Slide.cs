using System.Drawing;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerAutomation;

public partial class Slide(Microsoft.Office.Interop.PowerPoint.Slide slide, Markers markers) : IComparable<Slide> {
	public void SetProjectTitle(string text) {
		GetTextRangeByMarker(markers.Title).Text = text;
	}

	public void SetAuthors(string text) {
		GetTextRangeByMarker(markers.Author).Text = text;
	}

	public void SetParticipantImage(string imagePath) {
		SetImageByMarker(markers.ParticipantImage, imagePath);
	}

	public void SetPrizeName(string text) {
		GetTextRangeByMarker(markers.PrizeName).Text = text;
	}

	public void SetPrizeLogo(string imagePath) {
		SetImageByMarker(markers.PrizeLogo, imagePath);
	}

	private TextRange GetTextRangeByMarker(string marker) {
		foreach (Shape shape in slide.Shapes) {
			if (!shape.HasTextFrame.AsBool() || !shape.TextFrame.HasText.AsBool())
				continue;

			if (!shape.TextFrame.TextRange.Text.Contains(marker))
				continue;

			return shape.TextFrame.TextRange;
		}

		throw new KeyNotFoundException($"No marker found for {marker}.");
	}

	private void SetImageByMarker(string marker, string imagePath) {
		foreach (Shape shape in slide.Shapes) {
			if (!shape.HasTextFrame.AsBool() || !shape.TextFrame.HasText.AsBool())
				continue;

			if (!shape.TextFrame.TextRange.Text.Contains(marker))
				continue;

			float left = shape.Left;
			float top = shape.Top;
			float width = shape.Width;
			float height = shape.Height;

			shape.Delete();

			RectangleF rectangle = CalculateImageSize(imagePath, left, top, width, height);
			slide.Shapes.AddPicture(imagePath, MsoTriState.msoFalse, MsoTriState.msoCTrue,
				rectangle.X, rectangle.Y, rectangle.Width, rectangle.Height);
			return;
		}

		throw new KeyNotFoundException($"No marker found for {marker}.");
	}

	private RectangleF CalculateImageSize(string imagePath, float left, float top, float width, float height) {
		SizeF imageSize = Helpers.GetImageSize(imagePath);

		float aspectRatio = imageSize.Width / imageSize.Height;
		float newWidth, newHeight;

		if (width / height > aspectRatio) {
			newHeight = height;
			newWidth = height * aspectRatio;
		} else {
			newWidth = width;
			newHeight = width / aspectRatio;
		}

		float newLeft = left + (width - newWidth) / 2;
		float newTop = top + (height - newHeight) / 2;

		return new RectangleF(newLeft, newTop, newWidth, newHeight);
	}

	public int SlideNumber => slide.SlideNumber;

	public bool IsTemplate {
		get {
			foreach (Shape shape in slide.Shapes) {
				if (!shape.TextFrame.HasText.AsBool())
					continue;

				if (TemplateRegex().IsMatch(shape.TextFrame.TextRange.Text))
					return true;
			}

			return false;
		}
	}

	public void MoveTo(int number) {
		slide.MoveTo(number);
	}

	internal void Delete() => slide.Delete();

	internal Slide Duplicate() {
		Microsoft.Office.Interop.PowerPoint.Slide newSlide = slide.Duplicate()[1];
		return new Slide(newSlide, markers);
	}

	public int CompareTo(Slide? other) {
		if (other == null)
			return -1;

		return SlideNumber.CompareTo(other.SlideNumber);
	}

	[GeneratedRegex(@"\{\{[^{}]+\}\}")]
	private static partial Regex TemplateRegex();
}