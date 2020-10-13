using System;
using System.Collections.Generic;
using System.Linq;

namespace BingoCard
{
    public record CardData
    {
        private readonly List<string> _taken;

        public CardData(IEnumerable<string> data)
        {
            const int requiredDataSize = 25;
            var taken = data.Take(requiredDataSize).ToList();
            if (taken.Count != requiredDataSize)
                throw new ArgumentException($"Data only has {taken.Count} elements. {requiredDataSize} elements are required.");

            _taken = taken;
        }

        private CardRow _0 => new CardRow(_taken[0], _taken[1], _taken[2], _taken[3], _taken[4]);
        private CardRow _1 => new CardRow(_taken[5], _taken[6], _taken[7], _taken[8], _taken[9]);
        private CardRow _2 => new CardRow(_taken[10], _taken[11], _taken[12], _taken[13], _taken[14]);
        private CardRow _3 => new CardRow(_taken[15], _taken[16], _taken[17], _taken[18], _taken[19]);
        private CardRow _4 => new CardRow(_taken[20], _taken[21], _taken[22], _taken[23], _taken[24]);

        public IEnumerable<CardRow> Rows => new[] {_0, _1, _2, _3, _4};
    }

    public record CardRow(string B, string I, string N, string G, string O) { }
}