using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace PowerAutomation;

public class PowerPoint : IDisposable {
	private readonly Application _pptApplication = new() {
		Visible = MsoTriState.msoTrue
	};

	private readonly Presentation _presentation;

	private readonly List<Slide> _slides = [];

	public PowerPoint(string filePath) {
		_presentation = _pptApplication.Presentations.Open(filePath, WithWindow: MsoTriState.msoFalse);

		foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in _presentation.Slides) {
			_slides.Add(new Slide(slide));
		}
	}

	public Slide CreateSlide(int index, PpSlideLayout layout) {
		Microsoft.Office.Interop.PowerPoint.Slide slide = _presentation.Slides.Add(index, layout);
		_slides.Insert(index - 1, new Slide(slide));

		return _slides[index - 1];
	}
	public Slide GetSlide(int index) => _slides[index - 1];

	public void SaveAs(string filePath) {
		_presentation.SaveAs(filePath);
	}

	public void Dispose() {
		Console.WriteLine(_slides);
		_presentation.Close();
		_pptApplication.Quit();
	}
}