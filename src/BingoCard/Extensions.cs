using System;
using System.Collections.Generic;
using System.Linq;
using CliFx.Exceptions;

namespace BingoCard
{
    public static class Extensions
    {
        public static CommandException ToCommandException(this ExitCode target)
        {
            return new CommandException(target.Message, target.Value);
        }


        public static IEnumerable<T> TakeRandom<T>(this IEnumerable<T> target, int count)
        {
            var rnd = new Random(count.ToString().GetHashCode());

            return target
               .OrderBy(_ => rnd.Next())
               .Take(count);
        }
    }
}