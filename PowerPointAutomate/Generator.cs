using System.Globalization;
using CsvHelper;
using PowerAutomation;

namespace PowerPointAutomate;

public class Generator(PowerPoint powerPoint, string participantImagesFolderPath, string logosFolderPath) {
	private readonly Slide _prizeTitleTemplate = powerPoint.GetSlide(2);
	private readonly Slide _prizeResultTemplate = powerPoint.GetSlide(3);

	public void Generate(string csvPath) {
		List<Prize> prizes = GetPrizes(csvPath);

		for (int i = 0; i < prizes.Count; i++) {
			Prize prize = prizes[i];

			int newSlideNumber = _prizeTitleTemplate.SlideNumber + 2 + 2 * i;
			Slide titleSlide = powerPoint.DuplicateSlideAt(newSlideNumber, _prizeTitleTemplate);
			titleSlide.SetPrizeName(prize.PrizeName);
			string prizeLogo = Path.Combine(logosFolderPath, prize.PrizeLogo);
			titleSlide.SetPrizeLogo(prizeLogo);

			Slide resultSlide = powerPoint.DuplicateSlideAt(newSlideNumber + 1, _prizeResultTemplate);
			resultSlide.SetProjectTitle(prize.Title);
			resultSlide.SetAuthors(prize.Author);
			string participantImage = Path.Combine(participantImagesFolderPath, prize.ParticipantImage);
			resultSlide.SetParticipantImage(participantImage);
			Console.WriteLine($"Generated slides for prize: {prize.PrizeName}");
		}

		Console.WriteLine("Presentation modified successfully");
	}

	private static List<Prize> GetPrizes(string csvPath) {
		List<Prize> prizes = [];

		using StreamReader reader = new(csvPath);
		using CsvReader csv = new(reader, CultureInfo.InvariantCulture);
		IEnumerable<Prize> records = csv.GetRecords<Prize>();
		prizes.AddRange(records);
		return prizes;
	}

	public void Export(string exportPath) {
		powerPoint.RemoveTemplateSlides();
		powerPoint.SaveAs(exportPath);

		Console.WriteLine($"Presentation saved successfully at {exportPath}");
	}
}