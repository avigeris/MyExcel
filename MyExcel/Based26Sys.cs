using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcel
{
    public class Based26Sys
    {
        public string To26Sys(int i)
        {
            string res = "";
            if (i < 0)
            {
             //   res += "-";
                i = -i;
            }

            int k = 0;
            int[] arr = new int[200];
            while (i > 25)
            {
                arr[k] = i / 26 - 1;
                k++;
                i = i % 26;
            }
            arr[k] = i;

 
            for (int j = 0; j <= k; j++)
            {
                res = res + ((char)('A' + arr[j])).ToString();
            }
            return res;
        }

        public int From26Sys(string ColumnHeader)
        {
            char[] arr = ColumnHeader.ToCharArray();
            int lenghtArr = arr.Length;
            int res = 0;
            for (int i = lenghtArr - 2; i >= 0; i--)
            {
                res = res + ((int)arr[i] - (int)'A' + 1) * Convert.ToInt32(Math.Pow(26, lenghtArr - i - 1));
            }
            res = res + ((int)arr[lenghtArr - 1] - (int)'A');
            return res;
        }
    }
}
