using System.Drawing;
using Microsoft.Office.Core;
using System.Drawing;

namespace PowerAutomation;

internal static class Helpers {
	public static bool AsBool(this MsoTriState state) =>
		state switch {
			MsoTriState.msoTrue => true,
			MsoTriState.msoCTrue => true,
			MsoTriState.msoFalse => false,
			MsoTriState.msoTriStateToggle => throw new InvalidOperationException(
				"Cannot cast 'msoTriStateToggle' to bool."),
			MsoTriState.msoTriStateMixed => throw new InvalidOperationException(
				"Cannot cast 'msoTriStateMixed' to bool."),
			_ => throw new ArgumentOutOfRangeException(nameof(state), "Unknown MsoTriState value.")
		};

	public static SizeF GetImageSize(string imagePath) {
#pragma warning disable CA1416
		using Image img = Image.FromFile(imagePath);
		return new SizeF(img.Width, img.Height);
#pragma warning restore CA1416
	}
}