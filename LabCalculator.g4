grammar LabCalculator;

/*
* Parser Rules
*/
//compileUnit : expression EOF;
compileUnit : expression;
expression :
LPAREN expression RPAREN                                              #ParenthesizedExpr
| <assoc=right> expression EXPONENT expression                        #ExponentialExpr
| expression operatorToken=(MULTIPLY | DIVIDE | DIV | MOD) expression #MultiplicativeExpr
| expression operatorToken=(ADD | SUBTRACT) expression                #AdditiveExpr
| operatorToken=(MAX | MIN) LPAREN expression COMA expression RPAREN  #MaxMinExpr
| NUMBER                                                              #NumberExpr
| IDENTIFIER                                                          #IdentifierExpr
;

/*
* Lexer Rules
*/
NUMBER : INT ('.' INT)?;
INT : ('0'..'9')+;
IDENTIFIER : [rR]('0'..'9')+[cC]('0'..'9')+;
EXPONENT : '^';
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