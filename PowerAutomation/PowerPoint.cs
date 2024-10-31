using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace PowerAutomation;

public class PowerPoint : IDisposable {
	private readonly Application _pptApplication = new() {
		Visible = MsoTriState.msoTrue
	};

	private readonly Presentation _presentation;
	private readonly SortedSet<Slide> _slides = [];
	private readonly Markers _markers = new();

	public PowerPoint(string filePath) {
		_presentation = _pptApplication.Presentations.Open(filePath, WithWindow: MsoTriState.msoFalse);

		foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in _presentation.Slides) {
			_slides.Add(new Slide(slide, _markers));
		}
	}

	public PowerPoint(string filePath, Markers markers) : this(filePath) => _markers = markers;

	public Slide CreateSlide(int index, PpSlideLayout layout) {
		Microsoft.Office.Interop.PowerPoint.Slide slide = _presentation.Slides.Add(index, layout);
		Slide newSlide = new(slide, _markers);
		_slides.Add(newSlide);

		return newSlide;
	}

	public Slide GetSlide(int number) {
		foreach (Slide slide in _slides.Where(slide => slide.GetSlideNumber() == number)) {
			return slide;
		}
		throw new ArgumentOutOfRangeException($"The slide with number {number} does not exist.");
	}

	public void RemoveSlide(Slide slide) {
		_slides.Remove(slide);
		slide.Delete();
	}

	public Slide DuplicateSlideAt(int number, Slide original) {
		Slide slide = original.Duplicate();
		slide.MoveTo(number);
		_slides.Add(slide);
		return slide;
	}

	public void SaveAs(string filePath) {
		_presentation.SaveAs(filePath);
	}

	public void Dispose() {
		_presentation.Close();
		_pptApplication.Quit();
	}
}