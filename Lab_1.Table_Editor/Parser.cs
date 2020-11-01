using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Lab_1.Table_Editor
{
    class ParserException : ApplicationException
    {
        public ParserException(string str) : base(str) { }
        public override string ToString()
        {
            return Message;
        }
    }
    public class Parser
    {
        public Cell[,] Var;
        public int rows;
        public int columns;
        public static List<string> varsInFormula = new List<string>();
        public bool Error;
        public string message;
        enum Errors { SYNTAX, UNBALPARENS, NOEXP, RECURSIVE, OUTOFRANGE, MAXMINOUT};
        public Parser(Cell[,] vars, int rows, int columns)
        {
            Var = vars;
            this.rows = rows;
            this.columns = columns;
        }
        public double Eval(string exp)
        {
            Error = false;
            try
            {
                exp = RemoveWhite(exp);
                return Result(exp);
            }
            catch (ParserException ex)
            {
                Console.WriteLine(ex);
                return 0.0;
            }
        }
        void SyntaxErr(Errors error)
        {
            string[] err ={
                         "Синтаксична помилка",
                         "Дизбаланс дужок",
                         "Вираз відсутній",
                         "Рекурсивне обчислення",
                         "Не існує клітинки з заданим індексом",
                         "Неправильне оформлення max/min" };
            Error = true;
            message = err[(int)error];
            varsInFormula.Clear();
            throw new ParserException(err[(int)error]);
        }
        public double Result(string exp)
        {
            if (InvBrackets(exp) == false)
                SyntaxErr(Errors.UNBALPARENS);
            varsInFormula = FillingVarsInF(exp);
            varsInFormula = Parsing(varsInFormula);
            return CalculateR(ref varsInFormula);
        }
        public double CalculateR(ref List<string> varsInFormula)
        {
            List<string> Formula = new List<string>();
            List<string> Brackets = new List<string>();
            bool IsM = true;
            string MaxMin = "";
            bool IsB = true;
            foreach (var f in varsInFormula)
            {
                if ((f == "min") || (f == "mmin") ||
                    (f == "max") || (f == "mmax"))
                {
                    IsM = false;
                    MaxMin = f;
                }
            }
            if (IsM == false)
            {
                Formula.Clear();
                int i = 0;
                while ((varsInFormula[i] != MaxMin))
                {
                    Formula.Add(varsInFormula[i]);
                    i++;
                }
                i += 2;
                int j = 1;
                while (j != 0)
                {
                    Brackets.Add(varsInFormula[i]);
                    if (varsInFormula[i] == "(")
                        j++;
                    if (varsInFormula[i] == ")")
                        j--;
                    i++;
                }
                Brackets.RemoveAt(Brackets.Count() - 1);
                Formula.Add(MAXMIN(Brackets, MaxMin).ToString());
                while (i < varsInFormula.Count())
                {
                    Formula.Add(varsInFormula[i]);
                    i++;
                }
                varsInFormula = new List<string>(Formula);
                while (ExistsMetaSymbol(varsInFormula) && HasBrackets(varsInFormula))
                {
                    ConnectMinus(ref varsInFormula);
                    CalculateR(ref varsInFormula);
                }
                IsM = true;
            }
            if (IsM == true)
            {
                Brackets.Clear();
                Formula.Clear();
                foreach (var f in varsInFormula)
                {
                    if ((f == "(") || (f == ")"))
                        IsB = false;
                }
                if (IsB == false)
                {
                    int i = 0;
                    while (varsInFormula[i] != "(")
                    {
                        Formula.Add(varsInFormula[i]);
                        i++;
                    }
                    i++;
                    int j = 1;
                    while (j != 0)
                    {
                        Brackets.Add(varsInFormula[i]);
                        if (varsInFormula[i] == "(")
                            j++;
                        if (varsInFormula[i] == ")")
                            j--;
                        i++;
                    }
                    Brackets.RemoveAt(Brackets.Count() - 1);
                    Formula.Add(CalculateR(ref Brackets).ToString());
                    while (i < varsInFormula.Count())
                    {
                        Formula.Add(varsInFormula[i]);
                        i++;
                    }
                    varsInFormula = Formula;
                    while (ExistsMetaSymbol(varsInFormula) && HasBrackets(varsInFormula))
                    {
                        ConnectMinus(ref varsInFormula);
                        CalculateR(ref varsInFormula);
                    }
                    IsB = true;
                }
                if (IsB == true)
                {
                    for (int f = 0; f < varsInFormula.Count; f++)
                    {
                        if (varsInFormula[f] == "^")
                        {
                            varsInFormula[f - 1] = Math.Pow(Convert.ToDouble(varsInFormula[f - 1]), 
                                Convert.ToDouble(varsInFormula[f + 1])).ToString();
                            varsInFormula.RemoveAt(f + 1);
                            varsInFormula.RemoveAt(f);
                            f -= 2;
                        }
                    }
                }
            }
            ConnectMinus(ref varsInFormula);
            return Convert.ToDouble(varsInFormula[0]);
        }
        public double MAXMIN(List<string> varsInFormula, string MaxMin)
        {
            List<double> Formula = new List<double>();
            List<string> Brackets = new List<string>();
            int Koma = 0;
            for (int i = 0; i < varsInFormula.Count; i++)
            {
                if (varsInFormula[i] == ",")
                {
                    Koma++;
                }
            }
            if ((Koma != 1) && 
                ((MaxMin == "min") || 
                (MaxMin == "max")))
            {
                SyntaxErr(Errors.MAXMINOUT);
            }
            for (int i = 0; i < varsInFormula.Count; i++)
            {
                if (varsInFormula[i] == ",")
                {
                    Formula.Add(CalculateR(ref Brackets));
                    Brackets.Clear();
                    continue;
                }
                Brackets.Add(varsInFormula[i]);
            }
            Formula.Add(CalculateR(ref Brackets));
            if ((MaxMin == "max") || (MaxMin == "mmax"))
                return Enumerable.Max(Formula);
            else
                return Enumerable.Min(Formula);
        }
        private static bool IsPlus(String s)
        {
            return (s == "+");
        }
        private static bool IsNumder(String s)
        {
            if (s.Length >= 1)
                if ((s[0] > 47) && (s[0] < 58))
                {
                    return true;
                }
            if (s.Length >= 2)
                if ((s[0] == '-') && (s[1] > 47) && (s[1] < 58))
                {
                    return true;
                }
            return false;
        }
        public void ConnectMinus(ref List<string> formula)
        {
            for (int i = 0; i < formula.Count; i++)
            {
                if ((formula[i] == "-") && (i < (formula.Count - 1)))
                {
                    if (IsNumder(formula[i + 1]) == true)
                    {
                        formula[i + 1] = (Convert.ToDouble(formula[i + 1]) * (-1)).ToString();
                        formula.RemoveAt(i);
                        i--;
                    }
                }
            }
        }
        public List<string> Parsing(List<string> varsInFormula)
        {
            List<string> Formula = new List<string>();
            int curn = 0;
            int row = 0;
            int column = 0;
            bool ISZnak = true;
            foreach (var f in varsInFormula)
            {
                curn = 0;
                row = 0;
                column = 0;
                ISZnak = true;
                if ((f[0] > 64) && (f[0] < 91))
                {
                    for (int i = 0; i < f.Length; i++)
                    {
                        if ((f[i] > 64) && (f[i] < 91) && (ISZnak == true))
                        {
                            column = (column * 26) + f[i] - 64;
                        }
                        else if ((f[i] > 47) && (f[i] < 58))
                        {
                            row = (row * 10) + f[i] - 48;
                            ISZnak = false;
                        }
                        else
                        {
                            SyntaxErr(Errors.SYNTAX);
                        }
                    }
                    if ((rows >= row) && (columns >= column))
                        Formula.Add(Var[row - 1, column - 1].value.ToString());
                    else
                        SyntaxErr(Errors.OUTOFRANGE);
                }
                else if ((f[0] > 47) && (f[0] < 58))
                {
                    for (int i = 0; i < f.Length; i++)
                    {
                        if ((f[i] > 47) && (f[i] < 58))
                        {
                            curn = (curn * 10) + f[i] - 48;
                        }
                        else
                        {
                            SyntaxErr(Errors.SYNTAX);
                        }
                    }
                    Formula.Add(curn.ToString());
                }
                else if ((f == "(") || (f == ")") ||
                    (f == "+") || (f == "-") ||
                    (f == ",") || (f == "min") ||
                    (f == "mmin") || (f == "max") ||
                    (f == "mmax") || (f == "^"))
                {
                    Formula.Add(f);
                }
                else
                    SyntaxErr(Errors.SYNTAX);
            }
            Formula.RemoveAll(IsPlus);
            ConnectMinus(ref Formula);
            return Formula;
        }
        public List<string> FillingVarsInF(string exp)
        {
            List<string> Var = new List<string>();
            string curw = "";
            for (int i = 0; i < exp.Length; i++)
            {
                if ((exp[i].ToString() == "(") ||
                    (exp[i].ToString() == ")") ||
                    (exp[i].ToString() == "+") ||
                    (exp[i].ToString() == "-") ||
                    (exp[i].ToString() == ",") ||
                    (exp[i].ToString() == "^"))
                {
                    if (curw != "")
                    {
                        Var.Add(curw);
                        curw = "";
                    }
                    Var.Add(exp[i].ToString());
                }
                else
                {
                    curw += exp[i].ToString();
                }
            }
            if (curw != "")
            {
                Var.Add(curw);
            }
            return Var;
        }
        public bool HasBrackets(string exp)
        {
            for (int i = 0; i < exp.Length; i++)
            {
                if (exp[i].ToString() == ")")
                    return true;
            }
            return false;
        }
        public bool HasBrackets(List<string> exp)
        {
            for (int i = 0; i < exp.Count(); i++)
            {
                if (exp[i] == ")")
                    return true;
            }
            return false;
        }
        public bool InvBrackets(string exp)
        {
            int l = 0, r = 0;
            for (int i = 0; i < exp.Length; i++)
            {
                if (exp[i].ToString() == "(")
                    l++;
                if (exp[i].ToString() == ")")
                {
                    r++;
                }
            }
            if (l != r)
                return false;
            return true;
        }
        private string RemoveBrackets(string exp)
        {
            string ret = exp;
            if (exp != "")
            {
                string left = exp[0].ToString();
                string right = exp[exp.Length - 1].ToString();
                while (left == "(" && right == ")" && !ExistsMetaSymbol(ret) && (ret != ""))
                {
                    ret = ret.Remove(0, 1);
                    ret = ret.Remove(ret.Length - 1);
                    if (ret != "")
                    {
                        left = ret[0].ToString();
                        right = ret[ret.Length - 1].ToString();
                    }
                }
            }
            return ret;
        }
        private bool ExistsMetaSymbol(string exp)
        {
            int count = -1;
            int leftb = 0;
            int rightb = 0;
            for (int i = 0; i < exp.Length; i++)
            {
                if (exp[i] == '(')
                    leftb++;
                if (exp[i] == ')')
                    rightb++;
                if (rightb == leftb)
                    count++;
            }
            return count > 0;
        }
        private bool ExistsMetaSymbol(List<string> exp)
        {
            int count = -1;
            int leftb = 0;
            int rightb = 0;
            for (int i = 0; i < exp.Count(); i++)
            {
                if (exp[i] == "(")
                    leftb++;
                if (exp[i] == ")")
                    rightb++;
                if (rightb == leftb)
                    count++;
            }
            return count > 0;
        }
        private string RemoveWhite(string exp)
        {
            return exp.Replace(" ", "");
        }
    }
}
