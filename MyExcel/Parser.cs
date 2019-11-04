using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MyExcel
{
    public class Parser
    {
        //типи лексем
        enum Types { NONE, DELIMITER, VARIABLE, NUMBER };
      
        private Dictionary<string, double> variable = new Dictionary<string, double>();
        private Dictionary<string, List<string>> dependent = new Dictionary<string, List<string>>();
        private string exp; //посилання на рядок виразу
        private int expIndx; //поточний індекс
        private string token; //поточна лексема
        private Types tokType; //тип лексеми
        public string str_error = null;
        private string currentName = null;

        public void SetCurrentName(string name)
        {
                currentName = name;
           // foreach (var key in dependent.Keys)
            //{
             //  if (dependent[key].Contains(currentName))
              // {
                 //   dependent[key].Remove(currentName);
              // }
            // }
        }

        public void removeVariable(string name)
        {
            if (variable.ContainsKey(name))
            {
                variable.Remove(name);
            }

        }

        public List<string> GetDependent(string name)
        {
            if (dependent.ContainsKey(name))
                return dependent[name];
            else
                return null;
        }

        public double Evaluate(string expstr)
        {
     
            if (!variable.ContainsKey(currentName))
            {
                variable.Add(currentName, 0);
            }

            double result;
            exp = expstr;
            expIndx = 0;
            GetToken();

            if (token == "")
            {
               // MessageBox.Show("Noexp error", "MiniExcel");
                variable[currentName] = 0.0;
                return 0.0;

            }

            EvalExp2(out result);

            if (token != "")
            {
                MessageBox.Show("Syntax error", "MiniExcel");
            }

            variable[currentName] = result;
            return result;
        }

        private void EvalExp2(out double result)
        {
            string op;
            double partialResult;
            EvalExp3(out result);
            while ((op = token) == "+" || op == "-")
            {
                GetToken();
                EvalExp3(out partialResult);
                switch (op)
                {
                    case "-":
                        result = result - partialResult;
                        break;
                    case "+":
                        result = result + partialResult;
                        break;
                }
            }
        }

        private void EvalExp3(out double result)
        {
            string op;
            double partialResult;
            partialResult = 0.0;
            EvalExp4(out result);
            while ((op = token) == "*" || op == "/" || op == "%" || op == "|")
            {
                GetToken();
                EvalExp4(out partialResult);
                switch (op)
                {
                    case "*":
                        result = result * partialResult;
                        break;
                    case "/":
                        if (partialResult == 0.0)
                        {
                            MessageBox.Show("Divbyzero error", "MiniExcel");
                            result = result / partialResult;
                            return;
                        }
                        result = result / partialResult;
                        break;
                    case "%":
                        if (partialResult == 0.0)
                        {
                            MessageBox.Show("Divbyzero error", "MiniExcel");
                            return;
                        }
                        result = (int)result % (int)partialResult;
                        break;
                    case "|":
                        if (partialResult == 0.0)
                        {
                            MessageBox.Show("Divbyzero error", "MiniExcel");
                            return;
                        }
                        result = (int)result / (int)partialResult;
                        break;
                }
            }
        }

        private void EvalExp4(out double result)
        {
            double partialResult;
            double ex;
            int t;
            EvalExp5(out result);
            if (token == "^")
            {
                GetToken();
                EvalExp4(out partialResult);
                ex = result;
                if (partialResult == 0.0)
                {
                    result = 1.0;
                    return;
                }
                for (t = (int)partialResult - 1; t > 0; t--)
                {
                    result = result * (double)ex;
                }
            }
        }

        private void EvalExp5(out double result)
        {
            string op;
            double partialResult;
            EvalExp6(out result);
            while ((op = token) == ">" || op == "<")
            {
                GetToken();
                EvalExp6(out partialResult);
                switch (op)
                {
                    case ">":
                        if (result < partialResult)
                        {
                            result = partialResult;
                        }
                        break;
                    case "<":
                        if (result > partialResult)
                        {
                            result = partialResult;
                        }
                        break;
                }
            }
        }

        private void EvalExp6(out double result)
        {
            string op;
            op = "";
            if ((tokType == Types.DELIMITER) && token == "+" || token == "-")
            {
                op = token;
                GetToken();
            }
            EvalExp7(out result);
            if (op == "-")
            {
                result = -result;
            }
        }

        private void EvalExp7(out double result)
        {
            if (token == "(")
            {
                GetToken();
                EvalExp2(out result);
                if (token != ")")
                {
                    MessageBox.Show("Unbalparens error", "MiniExcel");
                    return;
                }
                GetToken();
            }
            else
            {
                Atom(out result);
            }
        }

        private void Atom(out double result)
        {
            switch (tokType)
            {
                case Types.NUMBER:
                    try
                    {
                        result = Double.Parse(token);
                    }
                    catch (FormatException)
                    {
                        result = 0.0;
                        MessageBox.Show("Syntax error", "MiniExcel");
                    }
                    GetToken();
                    return;
                case Types.VARIABLE:
                    if (token == currentName)
                    {
                        MessageBox.Show("Recurr", "MiniExcel");
                        result = 0;
                        return;
                    }
                    if (variable.ContainsKey(token))
                    {
                        result = variable[token];
                    }
                    else
                    {
                        variable.Add(token, 0);
                        result = variable[token];
                    }
                    if (dependent.ContainsKey(token))
                    {
                        if (!dependent[token].Contains(currentName))
                            dependent[token].Add(currentName);
                    }
                    else
                    {
                        dependent.Add(token, new List<string>());
                        dependent[token].Add(currentName);
                    }

                    GetToken();
                    return;
                default:
                    result = 0.0;
                    MessageBox.Show("Syntax error", "MiniExcel");
                    break;
            }
        }

        //отримуємо наступну лексему
        private void GetToken()
        {
            tokType = Types.NONE;
            token = "";
            if (expIndx == exp.Length)
                return;
            while (expIndx < exp.Length && char.IsWhiteSpace(exp[expIndx]))
                ++expIndx;
            if (expIndx == exp.Length)
                return;
            if (IsDelim(exp[expIndx]))
            {
                token += exp[expIndx];
                expIndx++;
                tokType = Types.DELIMITER;
            }
            else if (Char.IsLetter(exp[expIndx]))
            {
                while (!IsDelim(exp[expIndx]))
                {
                    token += exp[expIndx];
                    expIndx++;
                    if (expIndx >= exp.Length)
                        break;
                }
                tokType = Types.VARIABLE;
            }
            else if (Char.IsDigit(exp[expIndx]))
            {
                while (!IsDelim(exp[expIndx]))
                {
                    token += exp[expIndx];
                    expIndx++;
                    if (expIndx >= exp.Length)
                        break;
                }
                tokType = Types.NUMBER;
            }
        }

        private bool IsDelim(char c)
        {
            if (("+-/*%=()^|><".IndexOf(c) != -1))
                return true;
            return false;
        }

    }
}
