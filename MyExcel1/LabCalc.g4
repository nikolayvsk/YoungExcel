grammar LabCalc;

/*
* Parser Rules
*/

compileUnit : expression EOF;
expression:
  LPAREN expression RPAREN #ParenthesizedExpr
  |expression EXPONENT expression #ExponentialExpr
  |expression operatorToken=(MULTIPLY | DIVIDE) expression #MultiplicativeExpr
  | operatorToken=(ADD|SUBTRACT) expression #UnaryAdditiveExpr
  |expression operatorToken=(ADD | SUBTRACT) expression #AdditiveExpr
  | operatorToken=(MOD | DIV) LPAREN expression COMA expression RPAREN #ModDivExpr
  | NUMBER #NumberExpr
  | IDENTIFIER #IdentifierExpr
  ; 


/*
* Lexer Rules
*/

NUMBER : INT ('.' INT)?;
IDENTIFIER : [A-Z]+[1-9][0-9]*;

INT:('0'..'9')+;

MOD: 'mod';
DIV: 'div';
COMA: ',';
EXPONENT : '^';
MULTIPLY : '*';
DIVIDE : '/';
SUBTRACT : '-';
ADD : '+';
LPAREN : '(';
RPAREN : ')';

WS : [ \t\r\n] -> channel(HIDDEN);