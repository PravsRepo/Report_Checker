# my_dict = {1:"string"}
# first_key = next(iter(my_dict))
# print(first_key)
# name = "NAGHUL PRANAV BAABURAM KAVITHA"
# word_count = len(name.split())
# print(f"Word count: {word_count}")


import pandas as pd
df = pd.DataFrame({"Name": ["ABIRAMI G", "ANISHA L", "SARGURU P", "ANSLEY SINGH A", "DHARSHAN B", "SANTHOSH P" ]})
result = ['Anisha', 'Ansley', 'Santhosh', 'Dharshan', 'Sarguru', 'CEILING', 'MOUNTED']
pattern = "|". join(result).lower()
matches = []
for name in df["Name"]:
    iter_name = name.lower().split()
    found = False
    for i in iter_name:
        if len(i)>1:
            if i in pattern:
                found = True
                matches.append(True)
                break
            else:
                matches.append(False)
        else:
            continue
print(matches)
        


# sub_string = "BHOOMIKHA P"
# pattern = "TABLE PRINCE BHOOMIKHA NAGHUL PRANAV NIRANJAN KUMAR DILLI BABU"
# a = sub_string.lower().split()
# b = pattern.lower().split()
# print(a)
# print(b)
# for i in a:
#     for j in b:
#         if i.__contains__(j):
#             print(f"Yes! {i} is containing.")
#         else:
#             print("No")







