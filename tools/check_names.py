"""빠진 import 같은 '정의되지 않은 이름'을 빌드 전에 찾는다 (예: datetime 을 쓰면서 import 안 함)."""
import ast
import builtins
import sys

bad = 0
for path in sys.argv[1:]:
    tree = ast.parse(open(path, encoding="utf-8").read())
    defined = set(dir(builtins)) | {"__file__", "__name__"}
    for n in ast.walk(tree):
        if isinstance(n, (ast.Import, ast.ImportFrom)):
            for a in n.names:
                defined.add((a.asname or a.name).split(".")[0])
        elif isinstance(n, (ast.FunctionDef, ast.ClassDef, ast.AsyncFunctionDef)):
            defined.add(n.name)
        elif isinstance(n, ast.Name) and isinstance(n.ctx, (ast.Store, ast.Del)):
            defined.add(n.id)
        elif isinstance(n, ast.arg):
            defined.add(n.arg)
        elif isinstance(n, ast.ExceptHandler) and n.name:
            defined.add(n.name)
    used = {n.id for n in ast.walk(tree) if isinstance(n, ast.Name) and isinstance(n.ctx, ast.Load)}
    missing = sorted(used - defined)
    print(f"{path}: {'정의 안 된 이름 ' + str(missing) if missing else 'OK'}")
    bad += bool(missing)
sys.exit(1 if bad else 0)
