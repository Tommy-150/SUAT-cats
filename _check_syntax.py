import ast, sys
src = open('D:/Documents/OpenHanako/SUAT-cats/cats_manager.py', encoding='utf-8').read()
try:
    ast.parse(src)
    print('SYNTAX OK, lines:', src.count('\n'))
except SyntaxError as e:
    print('SYNTAX ERR:', e)
    sys.exit(1)
