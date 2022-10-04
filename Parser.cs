using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;

namespace MicroExcel
{

    class ThrowExceptionErrorListener : BaseErrorListener, IAntlrErrorListener<int>
    {
        //BaseErrorListener implementation
        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string
        msg, RecognitionException e)
        {
            throw new ArgumentException("Invalid Expression: {0}", msg, e);
        }
        //IAntlrErrorListener&lt;int&gt; implementation
        public void SyntaxError(IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg,
        RecognitionException e)
        {
            throw new ArgumentException("Invalid Expression: {0}", msg, e);
        }
    }

    public class MEParser : LabCalculatorBaseVisitor<double>
    {

        public enum Errors { NOERR, SYNTAX, UNBALPARENS, NOEXP, DIVBYZERO };
        public Errors tokErrors;
        //таблиця ідентифікаторів
        Dictionary<string, double> tableIdentifier = new Dictionary<string, double>();
        public override double VisitCompileUnit(LabCalculatorParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }
        public override double VisitNumberExpr(LabCalculatorParser.NumberExprContext context)
        {
            var result = double.Parse(context.GetText());
            //Debug.WriteLine(result);
            return result;
        }
        //IdentifierExpr
        public override double VisitIdentifierExpr(LabCalculatorParser.IdentifierExprContext context)
        {

            var result = context.GetText();
            double value;
            //видобути значення змінної з таблиці
            if (tableIdentifier.TryGetValue(result.ToString(), out value))
            {
                return value;
            }
            else
            {
                return 0.0;
            }
        }
        public override double VisitParenthesizedExpr(LabCalculatorParser.ParenthesizedExprContext context)
        {
            return Visit(context.expression());
        }
        public override double VisitExponentialExpr(LabCalculatorParser.ExponentialExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            //Debug.WriteLine("{0} ^ {1}", left, right);
            return System.Math.Pow(left, right);
        }

        public override double VisitAdditiveExpr([NotNull] LabCalculatorParser.AdditiveExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabCalculatorLexer.ADD)
            {
                //Debug.WriteLine("{0} + {1}", left, right);
                return left + right;
            }
            else
            {
                //Debug.WriteLine("{0} - {1}", left, right);
                return left - right;
            }
        }

        public override double VisitMaxMinExpr([NotNull] LabCalculatorParser.MaxMinExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabCalculatorLexer.MAX)
            {
                //Debug.WriteLine("MAX({0} , {1})", left, right);
                if (left > right) return left;
                return right;
            }
            else
            {
                //Debug.WriteLine("MIN({0} , {1})", left, right);
                if (left < right) return left;
                return right;
            }
        }

        public override double VisitDivModExpr([NotNull] LabCalculatorParser.DivModExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabCalculatorLexer.DIV)
            {
                //Debug.WriteLine("{0} div {1}", left, right);
                return Math.Truncate(left/right);
            }
            else //LabCalculatorLexer.SUBTRACT
            {
                //Debug.WriteLine("{0} mod {1}", left, right);
                return left % right;
            }
        }

        public override double VisitMultiplicativeExpr(LabCalculatorParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type == LabCalculatorLexer.MULTIPLY)
            {
                //Debug.WriteLine("{0} * {1}", left, right);
                return left * right;
            }
            else //LabCalculatorLexer.DIVIDE
            {
                //Debug.WriteLine("{0} / {1}", left, right);
                return left / right;
            }
        }
        private double WalkLeft(LabCalculatorParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LabCalculatorParser.ExpressionContext>(0));
        }
        private double WalkRight(LabCalculatorParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LabCalculatorParser.ExpressionContext>(1));
        }
        public double Evaluate(string expression)
        {
            var lexer = new LabCalculatorLexer(new AntlrInputStream(expression));
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(new ThrowExceptionErrorListener());
            var tokens = new CommonTokenStream(lexer);
            var parser = new LabCalculatorParser(tokens);
            var tree = parser.compileUnit();
            var visitor = new MEParser();
            return visitor.Visit(tree);

        }
    }
}
