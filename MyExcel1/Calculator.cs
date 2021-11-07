using Antlr4.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcel1
{
    public class Calculator
    {
        public static Dictionary<string, Cell> dictionary;
        public Calculator(Dictionary<string, Cell> D)
        {
            dictionary = D;
        }
        public double Evaluate(string expression, string CurrentCell, ref string ErrorCell)
        {
            try
            {
                var lexer = new LabCalcLexer(new AntlrInputStream(expression));
                lexer.RemoveErrorListeners();
                lexer.AddErrorListener(new ThrowExceptionErrorListener());

                var tokens = new CommonTokenStream(lexer);
                var parser = new LabCalcParser(tokens);

                parser.AddErrorListener(new ThrowExceptionErrorListener());

                var tree = parser.compileUnit();

                var visitor = new LabCalcVisitor(dictionary, CurrentCell);

                double res = visitor.Visit(tree);
                dictionary[CurrentCell].Val = res.ToString();
                return res;
            }
            catch (RecursiveException)
            {
                if (dictionary[CurrentCell].Val != "ErrorCycle")
                    ErrorCell += "[" + CurrentCell + "] рекурсія";
                dictionary[CurrentCell].Val = "ErrorCycle1";
                return 0;
            }
            catch (DivideZeroException)
            {
                if (dictionary[CurrentCell].Val != "ErrorDivZero")
                    ErrorCell += "[" + CurrentCell + "] ділення на 0";
                dictionary[CurrentCell].Val = "ErrorDivZero";
                return 0;
            }
            catch (Exception)
            {
                if (dictionary[CurrentCell].Val != "Error")
                    ErrorCell += "[" + CurrentCell + "] помилка";
                dictionary[CurrentCell].Val = "Error";
                return 0;
            }
        }
    }
}
