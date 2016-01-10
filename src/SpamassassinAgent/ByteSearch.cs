using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SpamassassinAgent
{
    static class ByteSearch
    {

        public static int Locate(this byte[] self, byte[] candidate, int offset)
        {
            if (IsEmptyLocate(self, candidate))
            {
                return -1;
            }

            for (int i = offset; i < self.Length; i++)
            {
                if (!IsMatch(self, i, candidate))
                    continue;
                return i;
            }

            return -1;
        }

        static bool IsMatch(byte[] array, int position, byte[] candidate)
        {
            if (candidate.Length > (array.Length - position))
            {
                return false;
            }

            for (int i = 0; i < candidate.Length; i++)
            {
                if (array[position + i] != candidate[i])
                {
                    return false;
                }
            }

            return true;
        }

        static bool IsEmptyLocate(byte[] array, byte[] candidate)
        {
            return array == null
                || candidate == null
                || array.Length == 0
                || candidate.Length == 0
                || candidate.Length > array.Length;
        }
    }
}
