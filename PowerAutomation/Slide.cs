using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerAutomation;

public class Slide(Microsoft.Office.Interop.PowerPoint.Slide slide, Markers markers) {
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

	internal void Delete() {
		slide.Delete();
	}
}