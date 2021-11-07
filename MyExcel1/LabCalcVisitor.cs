using Antlr4.Runtime.Misc;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcel1
{
    class RecursiveException : Exception
    {
    }
    class DivideZeroException : Exception
    {
    }
    class LabCalcVisitor : LabCalcBaseVisitor<double>
    {
        public Dictionary<string, Cell> tableIdentifier;
        public string CurrentCell;

        public LabCalcVisitor(Dictionary<string, Cell> dictionary, string cell)
        {
            tableIdentifier = dictionary;
            CurrentCell = cell;
        }

        public override double VisitCompileUnit(LabCalcParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }

        public override double VisitNumberExpr(LabCalcParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            Debug.WriteLine(result);
            return result;
        }

        public override double VisitIdentifierExpr(LabCalcParser.IdentifierExprContext context)
        {
            var result = context.GetText();
            if (!Recursive(CurrentCell, result))
            {
                if (!tableIdentifier[CurrentCell].CellReference.Contains(result))
                    tableIdentifier[CurrentCell].CellReference.Add(result);
                string StrValue = tableIdentifier[result].Val;
                if (StrValue == "Error")
                    throw new Exception();
                if (StrValue == "ErrorCycle")
                    throw new RecursiveException();
                if (StrValue == "ErrorDivZero")
                    throw new DivideByZeroException();
                double value = Convert.ToDouble(StrValue);
                Debug.WriteLine(value);
                return value;
            }
            else
            {
                throw new RecursiveException();
            }
        }

        public bool Recursive(string CurrentCell, string ReferedCell)
        {
            if (ReferedCell == CurrentCell)
                return true;
            if (tableIdentifier[ReferedCell].CellReference.Count != 0)
            {
                foreach (string i in tableIdentifier[ReferedCell].CellReference)
                {
                    if (Recursive(CurrentCell, i)) return true;
                }
                return false;
            }
            return false;
        }

        // our "( )"
        public override double VisitParenthesizedExpr(LabCalcParser.ParenthesizedExprContext context)//+
        {
            return Visit(context.expression());
        }

        // our "^" operator
        public override double VisitExponentialExpr(LabCalcParser.ExponentialExprContext context)//+
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            Debug.WriteLine("{0} ^ {1}", left, right);
            return System.Math.Pow(left, right);
        }

        private double WalkLeft(LabCalcParser.ExpressionContext context) //+
        {
            return Visit(context.GetRuleContext<LabCalcParser.ExpressionContext>(0));
        }

        private double WalkRight(LabCalcParser.ExpressionContext context)//+
        {
            return Visit(context.GetRuleContext<LabCalcParser.ExpressionContext>(1));
        }

        // our binary "+ -" operators
        public override double VisitAdditiveExpr(LabCalcParser.AdditiveExprContext context) //
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabCalcLexer.ADD)
            {
                Debug.WriteLine("{0} + {1}", left, right);
                return left + right;
            }
            else //LabCalculatorLexer.SUBTRACT
            {
                Debug.WriteLine("{0} - {1}", left, right);
                return left - right;
            }
        }

        // our unary operators "+ -"
        public override double VisitUnaryAdditiveExpr(LabCalcParser.UnaryAdditiveExprContext context)
        {
            var expression = Visit(context.expression());
            if (context.operatorToken.Type == LabCalcLexer.ADD)
            {
                Debug.WriteLine("+{0}", expression);
                return expression;
            }
            else
            {
                Debug.WriteLine("-{0}", expression);
                return -expression;
            }
        }

        // our binary "* /" operators
        public override double VisitMultiplicativeExpr(LabCalcParser.MultiplicativeExprContext context) // +
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == LabCalcLexer.MULTIPLY)
            {
                Debug.WriteLine("{0} * {1}", left, right);
                return (int)(left * right);
            }
            else //LabCalculatorLexer.DIVIDE
            {
                Debug.WriteLine("{0} / {1}", left, right);
                return (int)(left / right);
            }
        }

        // our "% /" operators
        public override double VisitModDivExpr(LabCalcParser.ModDivExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);

            if (context.operatorToken.Type == LabCalcLexer.MOD)
            {
                Debug.WriteLine("{0} % {1}", left, right);
                return left % right;
            }
            else //LabCalculatorLexer.DIV Integer division
            {
                Debug.WriteLine("{0} / {1}", left, right);
                return (int)(left / right);
            }
        }
    }
}
