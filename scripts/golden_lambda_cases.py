"""LAMBDA and its helper functions for the Excel golden fixture; see golden_cases.CASES."""

INPUTS: dict[str, list[list[object]]] = {
    "Lam": [[1, 2, 3], [4, 5, 6], [7, None, "x"], [10, 20, 30]],
}
RANGE = "Lam!A1:C2"
CASES: list[str | tuple[str, float]] = []


def cells(name: str, formula: str, rows: int, columns: int) -> list[str]:
    return [f"=INDEX({formula},{r},{c})" for r in range(1, rows + 1) for c in range(1, columns + 1)]


CASES += cells("BYROW", f"BYROW({RANGE},LAMBDA(r,SUM(r)))", 2, 1)
CASES += cells("BYROW", "BYROW(Lam!A1:C4,LAMBDA(r,MAX(r)-MIN(r)))", 4, 1)
CASES += cells("BYROW", "BYROW(Lam!A3:C3,LAMBDA(r,COUNTA(r)))", 1, 1)
CASES += cells("BYROW", 'BYROW(Lam!A1:C2,LAMBDA(row,TEXTJOIN("-",TRUE,row)))', 2, 1)
CASES += cells("BYCOL", f"BYCOL({RANGE},LAMBDA(c,SUM(c)))", 1, 3)
CASES += cells("BYCOL", "BYCOL(Lam!A1:C4,LAMBDA(c,AVERAGE(c)))", 1, 3)
CASES += cells("BYCOL", "BYCOL(Lam!A1:C3,LAMBDA(c,COUNT(c)))", 1, 3)
CASES += cells("MAP", f"MAP({RANGE},LAMBDA(x,x*2))", 2, 3)
CASES += cells("MAP", f"MAP({RANGE},Lam!A1:C2,LAMBDA(a,b,a+b))", 2, 3)
CASES += cells("MAP", "MAP({1,2,3},LAMBDA(x,x^2))", 1, 3)
CASES += cells("MAP", 'MAP({1;2;3},{10;20;30},LAMBDA(x,y,x&"-"&y))', 3, 1)
CASES += cells("MAP", "MAP(Lam!A3:C3,LAMBDA(x,ISNUMBER(x)))", 1, 3)
CASES += cells("MAP", 'MAP(Lam!A3:C3,LAMBDA(x,IF(ISBLANK(x),"blank",x)))', 1, 3)
CASES += cells("MAP", "MAP(Lam!A1:C3,Lam!A1:C2,LAMBDA(a,b,a+b))", 3, 3)
CASES += cells("MAP", "MAP({1,2},{1,2,3},LAMBDA(a,b,a+b))", 1, 3)
CASES += cells("MAP", "MAP({1,2,3},LAMBDA(x,{1,2}))", 1, 3)
CASES += cells("MAP", "MAP({1,2,3},LAMBDA(x,1/(x-2)))", 1, 3)
CASES += cells("MAP", 'MAP({1,2,3},LAMBDA(x,IFERROR(1/(x-2),"none")))', 1, 3)
CASES += cells("MAP", "MAP(Lam!A1:A4,LAMBDA(x,x))", 4, 1)
CASES += [
    "=REDUCE(0,Lam!A1:C2,LAMBDA(a,b,a+b))",
    "=REDUCE(1,Lam!A1:C2,LAMBDA(a,b,a*b))",
    '=REDUCE("",{"a","b","c"},LAMBDA(a,b,a&b))',
    "=REDUCE({1,2,3},{1,2},LAMBDA(a,b,a*b))",
    "=INDEX(REDUCE({1,2,3},{1,2},LAMBDA(a,b,a*b)),1,3)",
    "=REDUCE(Lam!A1:C2,LAMBDA(a,b,a+b))" if False else "=REDUCE(0,{5},LAMBDA(a,b,a-b))",
    "=REDUCE(100,{1,2,3},LAMBDA(a,b,a-b))",
    "=REDUCE(0,Lam!A3:C3,LAMBDA(a,b,a+N(b)))",
    "=REDUCE(0,{1,2,3},LAMBDA(a,b,IF(b=2,NA(),a+b)))",
    "=REDUCE(0,{1,2,3},LAMBDA(a,b,a+b))+1",
    "=REDUCE(MAX(Lam!A1:C2),Lam!A1:C2,LAMBDA(a,b,MIN(a,b)))",
]
CASES += cells("SCAN", "SCAN(0,{1,2,3,4},LAMBDA(a,b,a+b))", 1, 4)
CASES += cells("SCAN", "SCAN(1,Lam!A1:C2,LAMBDA(a,b,a*b))", 2, 3)
CASES += cells("SCAN", 'SCAN("",{"a","b","c"},LAMBDA(a,b,a&b))', 1, 3)
CASES += cells("SCAN", "SCAN(0,{1;2;3},LAMBDA(a,b,a+b))", 3, 1)
CASES += cells("SCAN", "SCAN(0,Lam!A3:C3,LAMBDA(a,b,a+N(b)))", 1, 3)
CASES += cells("MAKEARRAY", "MAKEARRAY(3,4,LAMBDA(r,c,r*c))", 3, 4)
CASES += cells("MAKEARRAY", "MAKEARRAY(2,2,LAMBDA(r,c,IF(r=c,1,0)))", 2, 2)
CASES += cells("MAKEARRAY", "MAKEARRAY(1,3,LAMBDA(r,c,c^2))", 1, 3)
CASES += [
    "=ROWS(MAKEARRAY(5,2,LAMBDA(r,c,r)))",
    "=COLUMNS(MAKEARRAY(5,2,LAMBDA(r,c,r)))",
    "=INDEX(MAKEARRAY(0,2,LAMBDA(r,c,r)),1,1)",
    "=INDEX(MAKEARRAY(2,0,LAMBDA(r,c,r)),1,1)",
    "=INDEX(MAKEARRAY(-1,2,LAMBDA(r,c,r)),1,1)",
    "=INDEX(MAKEARRAY(2,2,LAMBDA(r,c,{1,2})),1,1)",
    "=INDEX(MAKEARRAY(2.9,2,LAMBDA(r,c,r)),2,1)",
    "=ROWS(MAKEARRAY(2.9,2,LAMBDA(r,c,r)))",
    "=INDEX(MAKEARRAY(2,2,LAMBDA(r,c,1/0)),1,1)",
]
CASES += [
    "=LET(f,LAMBDA(x,x*2),f(5))",
    "=LET(f,LAMBDA(x,y,x+y),f(2,3))",
    "=LET(f,LAMBDA(x,x+1),f(f(f(1))))",
    "=LET(f,LAMBDA(x,x*2),SUM(MAP({1,2,3},f)))",
    "=LET(g,LAMBDA(x,x^2),INDEX(MAP({1,2,3},g),1,3))",
    "=LET(n,10,f,LAMBDA(x,x+n),f(1))",
    "=LET(f,LAMBDA(x,x*2),f(1,2))",
    "=LET(f,LAMBDA(x,y,x+y),f(1))",
    "=LET(f,LAMBDA(x,x),f({1,2,3}))",
    "=LET(f,LAMBDA(x,SUM(x)),f(Lam!A1:C2))",
    "=LET(f,LAMBDA(x,x*2),g,LAMBDA(y,f(y)+1),g(3))",
    "=LET(a,LAMBDA(x,x),a)",
    "=LET(a,LAMBDA(x,x),a(1)+a(2))",
    "=BYROW({1,2;3,4},LAMBDA(r,SUM(r)))",
    "=BYCOL({1,2;3,4},LAMBDA(c,SUM(c)))",
    "=SUM(BYROW({1,2;3,4},LAMBDA(r,PRODUCT(r))))",
    "=INDEX(BYROW({1,2;3,4},LAMBDA(r,r)),1,1)",
    "=INDEX(BYCOL({1,2;3,4},LAMBDA(c,c)),1,1)",
    "=INDEX(BYROW(5,LAMBDA(r,r*2)),1,1)",
    "=INDEX(BYROW(Lam!A1:C2,SUM),1,1)",
    "=INDEX(MAP({1,2},ABS),1,1)",
    "=ISOMITTED(1)",
    "=LAMBDA(x,x)",
]
