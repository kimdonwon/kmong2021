import os



fl=os.listdir('./image')
for i in fl:
    if 'A' in i:
        print(i)
