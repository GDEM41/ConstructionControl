using System;

namespace ConstructionControl
{
    internal static class CalculationCore
    {
        public const double QuantityEpsilon = 0.0001;

        public static bool IsOverage(double arrived, double planned)
            => arrived - planned > QuantityEpsilon;

        public static bool IsDeficit(double arrived, double planned)
            => planned - arrived > QuantityEpsilon;

        public static bool ShouldNotifySummaryDelta(double arrived, double planned, bool includeOverage, bool includeDeficit)
            => (includeOverage && IsOverage(arrived, planned))
               || (includeDeficit && IsDeficit(arrived, planned));

        public static double ClampToAvailable(double requested, double arrived, double mounted)
        {
            var allowed = Math.Max(0, arrived - mounted);
            var normalizedRequested = Math.Max(0, requested);
            return Math.Min(normalizedRequested, allowed);
        }

        public static bool HasDifference(double left, double right)
            => Math.Abs(left - right) > QuantityEpsilon;
    }
}
