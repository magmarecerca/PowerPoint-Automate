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

	public int LastSlideNumber => _slides.Count;

	public PowerPoint(string filePath) {
		_presentation = _pptApplication.Presentations.Open(filePath, WithWindow: MsoTriState.msoFalse);

		foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in _presentation.Slides) {
			_slides.Add(new Slide(slide, _markers));
		}
	}

	public PowerPoint(string filePath, Markers markers) : this(filePath) => _markers = markers;

	public Slide GetSlide(int number) {
		foreach (Slide slide in _slides.Where(slide => slide.SlideNumber == number)) {
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

	public Slide DuplicateSlide(Slide original) {
		return DuplicateSlideAt(original.SlideNumber + 1, original);
	}

	public void SaveAs(string filePath) {
		_presentation.SaveAs(filePath);
	}

	public void Dispose() {
		_presentation.Close();
		_pptApplication.Quit();
	}

	/// <summary>
	/// <b>WARNING:</b> this method is destructive, you should call it right before exporting the project.
	/// </summary>
	public void RemoveTemplateSlides() {
		List<Slide> templateSlides = _slides.Where(slide => slide.IsTemplate).ToList();
		foreach (Slide slide in templateSlides)
			RemoveSlide(slide);
	}
}