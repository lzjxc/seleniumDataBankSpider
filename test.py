dict1 = {'a':'aa','b':'bb'}
list2 = []
for k,v in dict1.items():
    print(k,v)
    list1=[k,v]
    list2.append(list1)
print(list2)
print(tuple(list2))