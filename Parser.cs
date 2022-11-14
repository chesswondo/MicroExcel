using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Text;
using System.Threading.Tasks;
using Antlr4.Runtime;
using Antlr4.Runtime.Misc;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Globalization;

namespace MicroExcel
{

    class ThrowExceptionErrorListener : BaseErrorListener, IAntlrErrorListener<int>
    {
        //BaseErrorListener implementation
        public override void SyntaxError(IRecognizer recognizer, IToken offendingSymbol, int line, int charPositionInLine, string
        msg, RecognitionException e)
        {
            throw new ArgumentException("Invalid Expression: " + msg, e);
        }
        //IAntlrErrorListener&lt;int&gt; implementation
        public void SyntaxError(IRecognizer recognizer, int offendingSymbol, int line, int charPositionInLine, string msg,
        RecognitionException e)
        {
            throw new ArgumentException("Invalid Expression: " + msg, e);
        }
    }

    public class MEParser : LabCalculatorBaseVisitor<double>
    {

        private MicroExcel me;
        public enum Errors { NOERR, SYNTAX, UNBALPARENS, NOEXP, DIVBYZERO };
        public Errors tokErrors;
        public override double VisitCompileUnit(LabCalculatorParser.CompileUnitContext context)
        {
            return Visit(context.expression());
        }
        public override double VisitNumberExpr(LabCalculatorParser.NumberExprContext context)
        {
            //MessageBox.Show(context.GetText());
            var result = double.Parse(context.GetText(), CultureInfo.InvariantCulture);
            //Debug.WriteLine(result);
            return result;
        }
        //IdentifierExpr
        public override double VisitIdentifierExpr(LabCalculatorParser.IdentifierExprContext context)
        {
            string result = context.GetText();
            var rc = Regex.Split(result, @"\D+").Where(s => s != "").ToArray();
            int row = Convert.ToInt32(rc[0])-1;
            int col = Convert.ToInt32(rc[1])-1;
            if (row < 0 || row >= me.RowCount || col < 0 || col >= me.ColCount)
                throw new ArgumentException("Посилання на неіснуючу комірку");
            Cell cell = me.getCell(row, col);
            if (cell.Error)
                throw new ArgumentException("Посилання на помилкову комірку");

            if (cell.Calculed == true) 
                throw new ArgumentException("Знайдена циклічна залежність");
            return Evaluate(cell.Formula, me);
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
            if (left == 0 && right == 0)
                throw new ArgumentException("0^0");

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

        public override double VisitMultiplicativeExpr(LabCalculatorParser.MultiplicativeExprContext context)
        {
            var left = WalkLeft(context);
            var right = WalkRight(context);
            if (context.operatorToken.Type != LabCalculatorLexer.MULTIPLY && right == 0)
                throw new ArgumentException("Дiлення на 0");
            switch (context.operatorToken.Type)
            {
                case LabCalculatorLexer.MULTIPLY: return left * right;
                case LabCalculatorLexer.DIVIDE:   return left / right;
                case LabCalculatorLexer.DIV:      return Math.Truncate(left / right);
                case LabCalculatorLexer.MOD:      return left % right;
            }
            throw new ArgumentException("Помилковий вираз");
        }
        private double WalkLeft(LabCalculatorParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LabCalculatorParser.ExpressionContext>(0));
        }
        private double WalkRight(LabCalculatorParser.ExpressionContext context)
        {
            return Visit(context.GetRuleContext<LabCalculatorParser.ExpressionContext>(1));
        }
        public double Evaluate(string expression, MicroExcel me)
        {
            this.me = me;
            var lexer = new LabCalculatorLexer(new AntlrInputStream(expression));
            lexer.RemoveErrorListeners();
            lexer.AddErrorListener(new ThrowExceptionErrorListener());
            var tokens = new CommonTokenStream(lexer);
            var parser = new LabCalculatorParser(tokens);
            var tree = parser.compileUnit();
            //   var visitor = new MEParser();
            //   return visitor.Visit(tree);
            return MicroExcel.parser.Visit(tree);

        }
    }
}
