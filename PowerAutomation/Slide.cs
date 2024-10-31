namespace PowerAutomation;

public class Slide(Microsoft.Office.Interop.PowerPoint.Slide slide) {
	public void SetTitle(string text) {
		slide.Shapes[1].TextFrame.TextRange.Text = text;
	}

	public void SetText(string text) {
		slide.Shapes[2].TextFrame.TextRange.Text = text;
	}
}