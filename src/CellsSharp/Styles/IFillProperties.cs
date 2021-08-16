using System.Collections.Generic;
using System.Drawing;
using JetBrains.Annotations;

namespace CellsSharp.Styles
{
	[PublicAPI]
	public interface IFillProperties
	{
		/// <summary>
		/// Gets an <see cref="IPatternFillProperties"/> object
		/// representing the properties of a pattern fill style.
		/// </summary>
		IPatternFillProperties PatternFill { get; }

		/// <summary>
		/// Gets an <see cref="IGradientFillProperties"/> object
		/// representing the properties of a gradient fill style.
		/// </summary>
		IGradientFillProperties GradientFill { get; }
	}

	[PublicAPI]
	public interface IPatternFillProperties
	{
		/// <summary>
		/// Gets or sets the pattern style of this pattern fill.
		/// </summary>
		PatternStyle PatternType { get; set; }

		/// <summary>
		/// Gets or sets the foreground color of this pattern fill.
		/// </summary>
		Color ForegroundColor { get; set; }

		/// <summary>
		/// Gets or sets the background color of this pattern fill.
		/// </summary>
		Color BackgroundColor { get; set; }
	}

	[PublicAPI]
	public interface IGradientFillProperties
	{
		/// <summary>
		/// Gets or sets the gradient style of this gradient fill.
		/// </summary>
		GradientStyle Type { get; set; }

		/// <summary>
		/// Gets or sets the angle of a linear gradient, in degrees,
		/// with 0 being straight to the right, and the angle increasing
		/// in the clockwise direction.
		/// </summary>
		double Degree { get; set; }

		/// <summary>
		/// Gets or sets the left convergence of a path gradient,
		/// as a ratio of the cell's width.
		/// </summary>
		/// <remarks>
		/// The convergence of a particular side of a path gradient
		/// is the distance inwards from that edge of the cell at
		/// which point the gradient will reach its final color. A
		/// value of 0 means the gradient immediately reaches its
		/// final color at that edge, and a value of 1 means the
		/// gradient will reach its final color at the opposite edge.
		/// </remarks>
		double Left { get; set; }

		/// <summary>
		/// Gets or sets the right convergence of a path gradient,
		/// as a ratio of the cell's width.
		/// </summary>
		/// <inheritdoc cref="Left" path="/remarks" />
		double Right { get; set; }

		/// <summary>
		/// Gets or sets the top convergence of a path gradient,
		/// as a ratio of the cell's width.
		/// </summary>
		/// <inheritdoc cref="Left" path="/remarks" />
		double Top { get; set; }

		/// <summary>
		/// Gets or sets the bottom convergence of a path gradient,
		/// as a ratio of the cell's width.
		/// </summary>
		/// <inheritdoc cref="Left" path="/remarks" />
		double Bottom { get; set; }

		/// <summary>
		/// Gets a collection of all the <see cref="IGradientStop"/>s
		/// in this gradient fill.
		/// </summary>
		IEnumerable<IGradientStop> Stops { get; }

		/// <summary>
		/// Removes all the stops from this gradient fill's stop
		/// collection.
		/// </summary>
		void RemoveAllStops();

		/// <summary>
		/// Adds a stop to this gradient fill at the provided
		/// <paramref name="distance"/> and <paramref name="color"/>.
		/// </summary>
		void AddStop(double distance, Color color);

		/// <summary>
		/// Sets stops for this gradient fill such that the provided
		/// <paramref name="colors"/> are spaced evenly between the
		/// start and end of the gradient.
		/// </summary>
		/// <remarks>
		/// If there are any stops already in this gradient fill when
		/// this method is called, they will be removed.
		/// </remarks>
		void SetStops(params Color[] colors);
	}

	[PublicAPI]
	public interface IGradientStop
	{
		/// <summary>
		/// Gets or sets the position of this gradient stop, with
		/// 0 being the start of the gradient and 1 being the end.
		/// </summary>
		double Position { get; set; }

		/// <summary>
		/// Gets or sets the color of the gradient at this stop.
		/// </summary>
		Color Color { get; set; }
	}
}