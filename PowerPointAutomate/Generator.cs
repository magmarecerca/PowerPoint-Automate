using PowerAutomation;

namespace PowerPointAutomate;

public class Generator {
	private readonly PowerPoint _powerPoint;
	private readonly Slide _prizeTitleTemplate;
	private readonly Slide _prizeResultTemplate;

	public Generator(PowerPoint powerPoint) {
		_powerPoint = powerPoint;

		_prizeTitleTemplate = powerPoint.GetSlide(2);
		_prizeResultTemplate = powerPoint.GetSlide(3);
	}

	public void Generate(string csvPath) {

	}
}