b=[4,5,6]
c=[[1,2],3,[4,5,6]]
d=[]

for n, i in enumerate(c):
    if len(str(i))>1:
        for j in i:
            d.append({b[n] :j})
            print((b[n],j))
    else:
        d.append({b[n] :i})
        print((b[n], i))

print(d)

a=[10,20,30,40,50]
for i in a:
    print(i)

print(list(enumerate(a)))