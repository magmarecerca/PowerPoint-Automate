using System.Globalization;
using CsvHelper;
using PowerAutomation;

namespace PowerPointAutomate;

public class Generator(PowerPoint powerPoint, string participantImagesFolderPath, string logosFolderPath, int templateIndex) {
	private readonly Slide _prizeTitleTemplate = powerPoint.GetSlide(templateIndex);
	private readonly Slide _prizeResultTemplate = powerPoint.GetSlide(templateIndex + 1);

	public void Generate(string csvPath) {
		List<Prize> prizes = GetPrizes(csvPath);
		CreatePrizeSlides(prizes);

		Console.WriteLine("Presentation modified successfully");
	}

	private int CalculatePrizeSlideIndex(int prizeIndex) =>
		_prizeTitleTemplate.SlideNumber + 2 + 2 * prizeIndex;

	private void CreatePrizeSlides(List<Prize> prizes) {
		for (int i = 0; i < prizes.Count; i++) {
			Prize prize = prizes[i];

			int newSlideNumber = CalculatePrizeSlideIndex(i);
			CreatePrizeTitleSlide(newSlideNumber, prize);
			CreatePrizeWinnerSlide(newSlideNumber, prize);

			Console.WriteLine($"({i + 1}/{prizes.Count}) Generated slides for prize: {prize.PrizeName}");
		}
	}

	private void CreatePrizeTitleSlide(int slideNumber, Prize prize) {
		Slide titleSlide = powerPoint.DuplicateSlideAt(slideNumber, _prizeTitleTemplate);
		titleSlide.SetPrizeName(prize.PrizeName);
		string prizeLogo = Path.Combine(logosFolderPath, prize.PrizeLogo);
		titleSlide.SetPrizeLogo(prizeLogo);
	}

	private void CreatePrizeWinnerSlide(int slideNumber, Prize prize) {
		Slide resultSlide = powerPoint.DuplicateSlideAt(slideNumber + 1, _prizeResultTemplate);

		if (string.IsNullOrEmpty(prize.Title)) {
			resultSlide.SetProjectTitle("Desert");
			resultSlide.RemoveTemplateElements();
			return;
		}
		resultSlide.SetProjectTitle(prize.Title);
		resultSlide.SetAuthors($"{prize.Author1}\n{prize.Author2}\n{prize.Author3}\n{prize.Author4}");
		string participantImage = Path.Combine(participantImagesFolderPath, prize.ParticipantImage);
		resultSlide.SetParticipantImage(participantImage);
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