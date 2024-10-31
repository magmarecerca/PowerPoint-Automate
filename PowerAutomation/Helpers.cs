using Microsoft.Office.Core;

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
}