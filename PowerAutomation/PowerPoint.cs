using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace PowerAutomation;

public class PowerPoint : IDisposable {
	private readonly Application _pptApplication = new() {
		Visible = MsoTriState.msoTrue
	};

	private readonly Presentation _presentation;

	public Slides Slides => _presentation.Slides;

	public PowerPoint(string filePath) {
		_presentation = _pptApplication.Presentations.Open(filePath, WithWindow: MsoTriState.msoFalse);
	}

	public void SaveAs(string filePath) {
		_presentation.SaveAs(filePath);
	}

	public void Dispose() {
		_presentation.Close();
		_pptApplication.Quit();
	}
}