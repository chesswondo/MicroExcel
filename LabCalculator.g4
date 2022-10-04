grammar LabCalculator;

/*
* Parser Rules
*/
compileUnit : expression EOF;
expression :
LPAREN expression RPAREN #ParenthesizedExpr
//expression COMA expression #
|expression EXPONENT expression #ExponentialExpr
| expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
| expression operatorToken=(DIV | MOD) expression #DivModExpr
| expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr
| operatorToken=(MAX | MIN) LPAREN expression COMA expression RPAREN #MaxMinExpr
| NUMBER #NumberExpr
| IDENTIFIER #IdentifierExpr
;

/*
* Lexer Rules
*/
NUMBER : INT ('.' INT)?;
IDENTIFIER : [a-zA-Z]+[1-9][0-9]+;
INT : ('0'..'9')+;
//EXPONENT : '^';
EXPONENT : '$';
COMA : ',';
MULTIPLY : '*';
DIVIDE : '/';
SUBTRACT : '-';
ADD : '+';
DIV : 'div';
MOD : 'mod';
MAX : 'max';
MIN : 'min';
LPAREN : '(';
RPAREN : ')';
WS : [ \t\r\n] -> channel(HIDDEN);