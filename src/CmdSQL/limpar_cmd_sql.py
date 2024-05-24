import pandas as pd
cmd_sql = []
for i in open('LogComandos.sql','r'):
    cmd_sql.append(i.split(' '))

with open('ComandosLog.txt','a') as f:
    for lin in cmd_sql:
        for i in range(2, len(lin)):
            f.write(f' {lin[i]}')

