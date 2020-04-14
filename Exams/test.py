# -*- coding: utf-8 -*-

# create dictionary
student_ids = { "홍길동": "13083301",
               "김미자": "11030104",
               "박은정": "11121994" }

# iterate over hashmap
for k, v in student_ids.items():
    print("Key: " + k + ", Value: " + v)

# get keys
print(student_ids.keys())

# get value for key
print(student_ids["박은정"])