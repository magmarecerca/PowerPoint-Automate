using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerAutomation;

public partial class Slide(Microsoft.Office.Interop.PowerPoint.Slide slide, Markers markers) : IComparable<Slide> {
	public void SetTitle(string text) {
		GetTextRangeByMarker(markers.Title).Text = text;
	}

	public void SetText(string text) {
		GetTextRangeByMarker(markers.Author).Text = text;
	}

	private TextRange GetTextRangeByMarker(string marker) {
		foreach (Shape shape in slide.Shapes) {
			if (!shape.TextFrame.HasText.AsBool())
				continue;

			if (!shape.TextFrame.TextRange.Text.Contains(marker))
				continue;

			return shape.TextFrame.TextRange;
		}

		throw new KeyNotFoundException($"No marker found for {marker}.");
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