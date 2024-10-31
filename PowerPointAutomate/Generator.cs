using System.Globalization;
using CsvHelper;
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
		List<Prize> prizes = GetPrizes(csvPath);

		for (int i = 0; i < prizes.Count; i++) {
			Prize prize = prizes[i];

			int newSlideNumber = _prizeTitleTemplate.SlideNumber + 2 + 2 * i;
			Slide titleSlide = _powerPoint.DuplicateSlideAt(newSlideNumber, _prizeTitleTemplate);
			titleSlide.SetPrizeName(prize.PrizeName);
			titleSlide.SetPrizeLogo(prize.PrizeLogo);

			Slide resultSlide = _powerPoint.DuplicateSlideAt(newSlideNumber + 1, _prizeResultTemplate);
			resultSlide.SetProjectTitle(prize.Title);
			resultSlide.SetAuthors(prize.Author);
			Console.WriteLine(prize.Author);
			resultSlide.SetParticipantImage(prize.ParticipantImage);
		}

		Console.WriteLine("Presentation modified successfully");
	}

	private List<Prize> GetPrizes(string csvPath) {
		List<Prize> prizes = [];

		using StreamReader reader = new(csvPath);
		using CsvReader csv = new(reader, CultureInfo.InvariantCulture);
		IEnumerable<Prize> records = csv.GetRecords<Prize>();
		prizes.AddRange(records);
		return prizes;
	}

	public void Export(string exportPath) {
		_powerPoint.RemoveTemplateSlides();
		_powerPoint.SaveAs(exportPath);

		Console.WriteLine($"Presentation saved successfully at {exportPath}");
	}
}