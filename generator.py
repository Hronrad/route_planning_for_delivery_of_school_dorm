for i in range(2, 7):
    n = int(input(f"{i}floor: "))
    if(i % 2 == 0):
        for j in range(1, n+1):
            if(j//10 == 0):
                st = f'0{j}'
            else:
                st = str(j)
            print(f'A{i}{st}')
        for k in range(1, n+1):
            j = n+1-k
            if(j//10 == 0):
                st = f'0{j}'
            else:
                st = str(j)
            print(f'B{i}{st}')
    else:
        for j in range(1, n+1):
            if(j//10 == 0):
                st = f'0{j}'
            else:
                st = str(j)
            print(f'B{i}{st}')
        for k in range(1, n+1):
            j = n+1-k
            if(j//10 == 0):
                st = f'0{j}'
            else:
                st = str(j)
            print(f'A{i}{st}')
